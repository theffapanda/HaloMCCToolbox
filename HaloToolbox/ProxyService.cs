using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using Microsoft.Win32;
using Titanium.Web.Proxy;
using Titanium.Web.Proxy.EventArguments;
using Titanium.Web.Proxy.Models;

namespace HaloToolbox;

public class ProxyService : IDisposable
{
    private const string BanProcessorHost = "banprocessor.svc.halowaypoint.com";
    private const string BanSummaryPath = "/hmcc/bansummary";

    // ── Configuration ─────────────────────────────────────────────────────────
    public int Port { get; set; } = 8888;
    private const int CertificateSetupTimeoutMs = 30000;
    private static readonly SemaphoreSlim CertificateSetupLock = new(1, 1);

    // ── State ─────────────────────────────────────────────────────────────────
    public bool IsRunning { get; private set; }

    // Raised on the thread-pool; callers must marshal to the UI thread
    public event EventHandler<ProxyCaptureEntry>? OnRequestCaptured;

    // Raised when the WinHTTP elevation (UAC) is declined — provides the manual command
    public event EventHandler<string>? WinHttpManualSetRequired;

    // Raised when a CascadeMatchmaking session is persisted to disk
    public event EventHandler? OnMatchSessionSaved;

    // Raised when the proxy learns or changes the current player identity.
    public event EventHandler? OnPlayerIdentityChanged;

    // Raised when observed squad context changes between solo / party / unknown.
    public event EventHandler? OnRejoinContextChanged;

    // Raised when crash-restore mode arms or clears so UI-owned helpers can stand down.
    public event EventHandler<bool>? OnCrashRestorePendingChanged;

    // Raised when the proxy learns the active PlayFab dedicated server.
    public event EventHandler<GameServerInfo?>? OnGameServerChanged;

    // Raised when MCC publishes its regional network tab measurements.
    public event EventHandler<NetworkRegionLatency?>? OnNetworkRegionLatencyChanged;

    // Raised when MPSD exposes per-member matchmaking region measurements.
    public event EventHandler<IReadOnlyList<MatchmakingPlayerPing>>? OnMatchmakingPlayerPingsObserved;

    public event EventHandler<SmartMatchWaitEstimate>? OnSmartMatchWaitEstimateChanged;
    public event EventHandler? OnSmartMatchWaitCancelled;

    private readonly object _banSpartanTokenLock = new();
    private readonly object _smartMatchAuthLock = new();
    private Dictionary<string, string> _smartMatchRequestHeaders = new(StringComparer.OrdinalIgnoreCase);
    private string _smartMatchServiceConfigId = "";
    private string _latestBanSpartanToken = "";
    private DateTimeOffset _latestBanSpartanTokenCapturedAtUtc;
    private string _latestBanSpartanTokenSourceHost = "";

    // In-memory copy of the saved matchmaking session — used for session discovery injection
    private SavedHandleInfo? _lastMatchSession;

    // Last observed squad session state — drives solo vs party handling and UI labeling.
    private RejoinSquadState? _lastSquadState;
    private SavedHandleInfo? _lastSquadHandle;

    // Set true when MCC exits and we have a saved match session.  Tells the proxy to
    // (a) do a JIT PUT+handle on the first sessiondirectory request, and
    // (b) force-replace session discovery results (even non-empty).
    // Cleared by ClearSavedMatchSession() or after timeout (5 minutes).
    private bool _pendingCrashRestore;
    private DateTime _pendingCrashRestoreStartedAt = DateTime.MinValue;
    private bool _jitHandleDone;    // prevents repeating activity handle POST
    private bool _jitPutDone;       // prevents repeating PUT /members/me
    private const int CRASH_RESTORE_TIMEOUT_MINUTES = 5;

    // Player's XUID — captured from session discovery URL or handles response
    private string _playerXuid = "";

    // Player's gamertag — captured from X-Xbl-Debug response headers
    private string _playerGamertag = "";

    public string CurrentPlayerGamertag => _playerGamertag;

    // Cached session body from INJECT[Member] — used to fake PUT responses
    private string? _cachedInjectedMatchBody;
    private string  _cachedInjectedMatchEtag = "";

    // When true, the next PUT {"members":{"me":null}} to CascadeMatchmaking is
    // rewritten to a harmless touch {"members":{"me":{}}} so the player stays
    // in the session on MPSD.  MCC's rejoin prompt (Dec 2022 update) checks
    // session membership when the user queues for matchmaking — blocking the
    // leave keeps that check valid.  One-shot: cleared after the first block.
    private bool _blockMatchLeave;

    // Ghost session mode: when enabled, fake MPSD responses so MCC thinks it's in
    // the session, while we simultaneously sync with real MPSD in the background.
    // Allows rejoin prompt to appear and function even if MCC temporarily lost session
    // membership. Disabled when background sync completes successfully.
    private bool _ghostSessionMode = false;
    private SavedHandleInfo? _ghostSession = null;
    private Task? _ghostSessionSyncTask;  // background sync task
    private bool _ghostSessionSyncSuccess = false;
    private string _ghostSessionOriginalConnectionGuid = "";
    private bool _ghostSessionGuidUpgraded = false;

    private const string PlaceholderConnectionGuid = "12345678-1234-1234-1234-123456789abc";

    // Game server redirection: cache the original game server info from RequestParty,
    // and redirect subsequent requests to use the same server (prevents PlayFab from
    // assigning a different server after restart, which breaks rejoin).
    private GameServerInfo? _cachedGameServerInfo = null;
    private GameServerInfo? _currentObservedGameServerInfo = null;
    private NetworkRegionLatency? _bestNetworkRegionLatency = null;
    private readonly Dictionary<string, MatchmakingPlayerPing> _matchmakingPlayerPings =
        new(StringComparer.OrdinalIgnoreCase);
    private bool _gameServerRedirectionActive = false;


    public bool IsGameServerRedirectionActive => _gameServerRedirectionActive;
    public string CurrentGameServerIp => _currentObservedGameServerInfo?.IPv4Address ?? "";
    public GameServerInfo? CurrentGameServerInfo => _currentObservedGameServerInfo;
    public NetworkRegionLatency? BestNetworkRegionLatency => _bestNetworkRegionLatency;
    public RejoinSessionMode CurrentRejoinMode => _lastSquadState?.Mode ?? RejoinSessionMode.Unknown;
    public string CurrentRejoinModeLabel => CurrentRejoinMode.ToDisplayLabel();
    public int CurrentSquadMemberCount => _lastSquadState?.MemberCount ?? 0;
    public ProxyService()
    {
        _cachedGameServerInfo = LoadPersistedGameServer();
        _lastSquadState = LoadPersistedSquadState();
        _lastSquadHandle = LoadPersistedHandle();
        LoadPersistedBanSpartanToken();
    }

    public void CaptureBanSpartanToken(string token, DateTimeOffset capturedAtUtc, string sourceHost)
    {
        if (string.IsNullOrWhiteSpace(token))
            return;

        lock (_banSpartanTokenLock)
        {
            _latestBanSpartanToken = token;
            _latestBanSpartanTokenCapturedAtUtc = capturedAtUtc;
            _latestBanSpartanTokenSourceHost = sourceHost;
        }

        PersistBanSpartanToken(token, capturedAtUtc, sourceHost);
    }

    public bool TryGetLatestBanSpartanToken(
        out string token,
        out DateTimeOffset capturedAtUtc,
        out string sourceHost)
    {
        lock (_banSpartanTokenLock)
        {
            token = _latestBanSpartanToken;
            capturedAtUtc = _latestBanSpartanTokenCapturedAtUtc;
            sourceHost = _latestBanSpartanTokenSourceHost;

            if (string.IsNullOrWhiteSpace(token))
                return false;

            return true;
        }
    }

    public void ClearBanSpartanToken(string token)
    {
        lock (_banSpartanTokenLock)
        {
            if (!string.Equals(_latestBanSpartanToken, token, StringComparison.Ordinal))
                return;

            _latestBanSpartanToken = "";
            _latestBanSpartanTokenCapturedAtUtc = default;
            _latestBanSpartanTokenSourceHost = "";
        }

        try { File.Delete(RejoinFixPaths.BanSpartanTokenFile); } catch { }
    }

    private void PersistBanSpartanToken(string token, DateTimeOffset capturedAtUtc, string sourceHost)
    {
        try
        {
            RejoinFixPaths.EnsureRootDirectory();
            var payload = new PersistedBanSpartanToken
            {
                ProtectedToken = Convert.ToBase64String(ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(token),
                    optionalEntropy: null,
                    DataProtectionScope.CurrentUser)),
                CapturedAtUtc = capturedAtUtc,
                SourceHost = sourceHost,
            };

            File.WriteAllText(RejoinFixPaths.BanSpartanTokenFile, JsonSerializer.Serialize(payload));
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("ban-checker", $"Failed to persist encrypted banprocessor token: {ex.Message}");
        }
    }

    private void LoadPersistedBanSpartanToken()
    {
        try
        {
            string path = RejoinFixPaths.BanSpartanTokenFile;
            if (!File.Exists(path))
                return;

            var payload = JsonSerializer.Deserialize<PersistedBanSpartanToken>(File.ReadAllText(path));
            if (payload is null || string.IsNullOrWhiteSpace(payload.ProtectedToken))
                return;

            byte[] encrypted = Convert.FromBase64String(payload.ProtectedToken);
            string token = Encoding.UTF8.GetString(ProtectedData.Unprotect(
                encrypted,
                optionalEntropy: null,
                DataProtectionScope.CurrentUser));

            if (string.IsNullOrWhiteSpace(token))
                return;

            lock (_banSpartanTokenLock)
            {
                _latestBanSpartanToken = token;
                _latestBanSpartanTokenCapturedAtUtc = payload.CapturedAtUtc;
                _latestBanSpartanTokenSourceHost = payload.SourceHost;
            }
        }
        catch (Exception ex)
        {
            try { File.Delete(RejoinFixPaths.BanSpartanTokenFile); } catch { }
            RejoinFixDiagnostics.Warn("ban-checker", $"Failed to load encrypted banprocessor token: {ex.Message}");
        }
    }

    private sealed class PersistedBanSpartanToken
    {
        public string ProtectedToken { get; set; } = "";
        public DateTimeOffset CapturedAtUtc { get; set; }
        public string SourceHost { get; set; } = "";
    }

    /// <summary>Clears cached game server info and disables redirection.</summary>
    public void ClearGameServerRedirection()
    {
        _cachedGameServerInfo = null;
        _currentObservedGameServerInfo = null;
        _gameServerRedirectionActive = false;
        // Also clear the persisted file
        try { File.Delete(GetGameServerCacheFile()); } catch { /* ignore */ }
        OnGameServerChanged?.Invoke(this, null);
    }

    private void SetCurrentGameServer(GameServerInfo serverInfo)
    {
        if (string.IsNullOrWhiteSpace(serverInfo.IPv4Address))
            return;

        _cachedGameServerInfo = serverInfo;
        _currentObservedGameServerInfo = serverInfo;
        PersistGameServerToDisk(serverInfo);
        OnGameServerChanged?.Invoke(this, serverInfo);
    }

    /// <summary>Get the path to the cached game server file.</summary>
    private static string GetGameServerCacheFile() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                     "HaloMCCToolbox", "RejoinFix", "last-game-server.json");

    /// <summary>Save the game server info to disk so it survives proxy restart.</summary>
    public void PersistGameServerToDisk(GameServerInfo serverInfo)
    {
        try
        {
            var dir = Path.GetDirectoryName(GetGameServerCacheFile());
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir!);

            var json = JsonSerializer.Serialize(serverInfo, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(GetGameServerCacheFile(), json);
        }
        catch (Exception ex)
        {
            // Best-effort; don't break proxy if file write fails
            RejoinFixDiagnostics.Warn("game-server", $"Failed to persist cached game server: {ex.Message}");
            Debug.WriteLine($"Failed to persist game server: {ex.Message}");
        }
    }

    /// <summary>Load the persisted game server info from disk.</summary>
    private GameServerInfo? LoadPersistedGameServer()
    {
        try
        {
            var filePath = GetGameServerCacheFile();
            if (!File.Exists(filePath))
                return null;

            var json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<GameServerInfo>(json);
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("game-server", $"Failed to load cached game server: {ex.Message}");
            Debug.WriteLine($"Failed to load persisted game server: {ex.Message}");
            return null;
        }
    }

    private RejoinSquadState? LoadPersistedSquadState()
    {
        try
        {
            if (!File.Exists(RejoinFixPaths.LastSquadStateFile))
                return null;

            var json = File.ReadAllText(RejoinFixPaths.LastSquadStateFile);
            return JsonSerializer.Deserialize<RejoinSquadState>(json);
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("squad", $"Failed to load cached squad state: {ex.Message}");
            return null;
        }
    }

    private SavedHandleInfo? LoadPersistedHandle()
    {
        try
        {
            if (!File.Exists(RejoinFixPaths.LastHandleFile))
                return null;

            var json = File.ReadAllText(RejoinFixPaths.LastHandleFile);
            return JsonSerializer.Deserialize<SavedHandleInfo>(json);
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("capture", $"Failed to load cached activity handle: {ex.Message}");
            return null;
        }
    }

    /// <summary>Signals that MCC just crashed and we should force-inject on next discovery.</summary>
    public void SetPendingCrashRestore(SavedHandleInfo? matchSession = null)
    {
        _pendingCrashRestore = true;
        _pendingCrashRestoreStartedAt = DateTime.UtcNow;  // Start timeout clock
        _jitHandleDone = false;
        _jitPutDone = false;
        _blockMatchLeave = true;
        _cachedInjectedMatchBody = null;
        _cachedInjectedMatchEtag = "";

        // RACE CONDITION FIX: Accept matchSession parameter from caller (MainWindow)
        // to ensure we don't rely on _lastMatchSession timing. Prefer the stronger
        // restore source so a post-crash solo-looking capture cannot replace a
        // pre-crash squad capture.
        if (matchSession is not null &&
            (_lastMatchSession is null || IsBetterRestoreSource(matchSession, _lastMatchSession)))
        {
            _lastMatchSession = matchSession;
            RejoinFixDiagnostics.Info("restore", $"Recovered saved match session from UI state: {matchSession.TemplateName}/{matchSession.SessionShort}");
            Debug.WriteLine($"[RESTORE] Match session restored from parameter: {matchSession.TemplateName}/{matchSession.SessionShort}");
        }

        // Enable aggressive ghost session mode
        if (_lastMatchSession is not null)
        {
            _ghostSessionMode = true;
            _ghostSession = _lastMatchSession;
            _ghostSessionSyncSuccess = false;
            _ghostSessionOriginalConnectionGuid = _ghostSession.ConnectionGuid ?? "";
            _ghostSessionGuidUpgraded = false;

            // Start background sync immediately
            _ghostSessionSyncTask = AutoSyncGhostSessionAsync();
        }

        OnCrashRestorePendingChanged?.Invoke(this, true);
    }

    public void ClearGhostSessionMode()
    {
        _ghostSessionMode = false;
        _ghostSession = null;
        _ghostSessionSyncSuccess = false;
        _ghostSessionOriginalConnectionGuid = "";
        _ghostSessionGuidUpgraded = false;
    }

    public bool IsGhostSessionActive() => _ghostSessionMode;

    /// <summary>Check if crash restore timeout has expired and clear if needed.</summary>
    private void CheckAndClearPendingCrashRestoreTimeout()
    {
        if (_pendingCrashRestore && _pendingCrashRestoreStartedAt != DateTime.MinValue)
        {
            if ((DateTime.UtcNow - _pendingCrashRestoreStartedAt).TotalMinutes > CRASH_RESTORE_TIMEOUT_MINUTES)
            {
                RejoinFixDiagnostics.Warn("restore", $"Crash restore timed out after {CRASH_RESTORE_TIMEOUT_MINUTES} minutes; clearing saved session.");
                Debug.WriteLine($"[TIMEOUT] Crash restore exceeded {CRASH_RESTORE_TIMEOUT_MINUTES} minutes - clearing pending state");
                ClearSavedMatchSession();  // This will clear the flag and related state
            }
        }
    }

    // ── Internals ─────────────────────────────────────────────────────────────
    private ProxyServer?           _server;
    private ExplicitProxyEndPoint? _endpoint;

    // WinINet originals — restored on Stop()
    private int    _savedProxyEnable;
    private string _savedProxyServer   = "";
    private string _savedProxyOverride = "";

    // ── Domain filter ─────────────────────────────────────────────────────────
    //
    // IMPORTANT: use the most specific suffix possible.
    //
    // BeforeTunnelConnectRequest below sets DecryptSsl=false for any host NOT in this
    // list, so only these hosts are SSL-intercepted.  Everything else
    // (presence-heartbeat, userpresence, auth, …) becomes a plain TCP
    // tunnel with zero TLS overhead and zero interference with Xbox Live session state.
    private static readonly string[] _watchedDomains =
    [
        "halowaypoint.com",              // Halo Waypoint API + Spartan token extraction
        "sessiondirectory.xboxlive.com", // Xbox Live MPSD — session documents with skill data
        "smartmatch.xboxlive.com",       // Xbox Live SmartMatch queue ticket estimates
        "playfabapi.com",                // PlayFab matchmaking / telemetry
    ];

    private static readonly string[] _xboxShellProcessNames =
    [
        "GameBar",
        "GameBarFTServer",
        "GameBarPresenceWriter",
        "XboxApp",
        "XboxPcApp",
        "XboxPcAppFT",
        "XboxGameBarWidgets",
        "XboxGamingOverlay",
        "GamingServices",
        "GamingServicesNet",
    ];

    private static readonly string[] _systemProxyBypassHosts =
    [
        "localhost",
        "127.0.0.1",
        "<local>",
        "*.auth.xboxlive.com",
        "accounts.xboxlive.com",
        "achievements.xboxlive.com",
        "activityhub.xboxlive.com",
        "avty.xboxlive.com",
        "clubhub.xboxlive.com",
        "conversationhub.xboxlive.com",
        "gameclipsmetadata.xboxlive.com",
        "peoplehub.xboxlive.com",
        "privacy.xboxlive.com",
        "profile.xboxlive.com",
        "reputation.xboxlive.com",
        "social.xboxlive.com",
        "titlehub.xboxlive.com",
        "user.auth.xboxlive.com",
        "userpresence.xboxlive.com",
        "userstats.xboxlive.com",
        "xblmessaging.xboxlive.com",
        "presence-heartbeat.xboxlive.com",
        "xnotify.xboxlive.com",
    ];

    // Bodies to skip even within watched domains.
    // banprocessor.svc.halowaypoint.com → under halowaypoint.com → still watched;
    // skip its body to avoid any stall on the ban-check response.
    private static readonly string[] _bypassHosts =
    [
        "banprocessor",  // MCC ban-check — may use unusual response format; skip body read
    ];

    // HttpClient that bypasses the proxy for our own out-of-band MPSD reads.
    // UseProxy=false prevents our ETag-refresh GETs from being re-intercepted by ourselves.
    // Short timeout: this sits on the hot path of MCC's session PUT; 3 s is generous.
    private static readonly HttpClient _refreshClient = new(
        new HttpClientHandler { UseProxy = false })
    {
        Timeout = TimeSpan.FromSeconds(3),
    };

    // ── Certificate storage ───────────────────────────────────────────────────
    private static string CertStorePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox", "RejoinFix", "proxy-root.pfx");

    // ── Start / Stop ──────────────────────────────────────────────────────────
    public async Task StartAsync()
    {
        if (IsRunning) return;

        Directory.CreateDirectory(Path.GetDirectoryName(CertStorePath)!);

        ProxyServer? server = null;
        ExplicitProxyEndPoint? endpoint = null;
        bool winInetProxySet = false;

        try
        {
            server = await CreateProxyServerWithCertificateAsync();
            server.BeforeRequest  += OnBeforeRequestAsync;
            server.BeforeResponse += OnBeforeResponseAsync;

            endpoint = new ExplicitProxyEndPoint(IPAddress.Loopback, Port, decryptSsl: true);
            // BeforeTunnelConnectRequest lives on the endpoint, not the server.
            // It must be registered after the endpoint is created.
            endpoint.BeforeTunnelConnectRequest += OnBeforeTunnelConnectRequest;
            server.AddEndPoint(endpoint);
            await server.StartAsync();

            _server = server;
            _endpoint = endpoint;

            // WinINet proxy (no admin required)
            SetWinINetProxy($"127.0.0.1:{Port}");
            winInetProxySet = true;

            // WinHTTP proxy: Halo MCC uses WinHTTP, not WinINet.
            await TrySetWinHttpProxyAsync();

            IsRunning = true;
            RejoinFixDiagnostics.Info("proxy", $"Proxy started on 127.0.0.1:{Port}.");
        }
        catch
        {
            CleanupFailedStart(server, endpoint, winInetProxySet);
            throw;
        }
    }

    private static async Task<ProxyServer> CreateProxyServerWithCertificateAsync()
    {
        bool certificateWasMissing = !File.Exists(CertStorePath);
        if (certificateWasMissing)
            RejoinFixDiagnostics.Info("proxy", "Proxy root certificate is missing; creating and trusting a new certificate.");

        if (!await CertificateSetupLock.WaitAsync(TimeSpan.FromSeconds(5)))
            throw new TimeoutException("Titanium proxy certificate setup is already running. Wait a moment, then relaunch Rejoin Fix.");

        bool releaseLockInFinally = true;

        try
        {
            var setupTask = Task.Run(() =>
            {
                var server = new ProxyServer();
                server.CertificateManager.PfxFilePath = CertStorePath;
                server.CertificateManager.PfxPassword = "halointel-proxy";

                // Install CA into CurrentUser\Root; this can show a one-time Windows trust dialog.
                // Keep it off the WPF thread so a first-run certificate prompt cannot freeze Toolbox.
#pragma warning disable CS0618
                server.CertificateManager.EnsureRootCertificate(
                    userTrustRootCertificate:    true,
                    machineTrustRootCertificate: false,
                    trustRootCertificateAsAdmin: false);
#pragma warning restore CS0618

                return server;
            });

            if (await Task.WhenAny(setupTask, Task.Delay(CertificateSetupTimeoutMs)) != setupTask)
            {
                RejoinFixDiagnostics.Error("proxy", "Timed out while creating or trusting the Titanium proxy certificate.");
                releaseLockInFinally = false;
                _ = setupTask.ContinueWith(t =>
                {
                    _ = t.Exception;
                    CertificateSetupLock.Release();
                }, TaskScheduler.Default);
                throw new TimeoutException("Titanium proxy certificate setup timed out. Close any certificate/trust prompts, then relaunch Rejoin Fix.");
            }

            var proxyServer = await setupTask;
            if (certificateWasMissing)
                RejoinFixDiagnostics.Info("proxy", "Proxy root certificate was created and trusted.");

            return proxyServer;
        }
        finally
        {
            if (releaseLockInFinally)
                CertificateSetupLock.Release();
        }
    }

    private void CleanupFailedStart(ProxyServer? server, ExplicitProxyEndPoint? endpoint, bool restoreWinInetProxy)
    {
        try
        {
            if (restoreWinInetProxy)
                RestoreWinINetProxy();

            if (server is not null)
            {
                server.BeforeRequest  -= OnBeforeRequestAsync;
                server.BeforeResponse -= OnBeforeResponseAsync;
            }

            if (endpoint is not null)
                endpoint.BeforeTunnelConnectRequest -= OnBeforeTunnelConnectRequest;

            server?.Stop();
            server?.Dispose();
        }
        catch
        {
            // Startup failed already; cleanup is best effort so the original error stays visible.
        }

        _server = null;
        _endpoint = null;
        IsRunning = false;
    }

    public void Stop()
    {
        if (!IsRunning) return;

        RestoreWinINetProxy();
        TryResetWinHttpProxy(); // best-effort elevated netsh

        if (_server is not null)
        {
            _server.BeforeRequest  -= OnBeforeRequestAsync;
            _server.BeforeResponse -= OnBeforeResponseAsync;
            if (_endpoint is not null)
            {
                _endpoint.BeforeTunnelConnectRequest -= OnBeforeTunnelConnectRequest;
                _endpoint = null;
            }
            _server.Stop();
            _server.Dispose();
            _server = null;
        }

        IsRunning = false;
        RejoinFixDiagnostics.Info("proxy", "Proxy stopped and system proxy settings were restored.");
    }

    // ── Tunnel-connect filter ─────────────────────────────────────────────────
    //
    // With decryptSsl:true on the endpoint, Titanium would MITM every HTTPS
    // connection by default — including presence heartbeats and auth —
    // adding TLS handshake overhead that disrupts Xbox Live session state.
    //
    // Here we opt non-watched domains OUT of SSL decryption: they become plain
    // TCP tunnels (zero overhead).  Known Xbox shell clients also stay tunneled
    // on shared Xbox hosts so Game Bar can coexist with MCC rejoin capture.
    private Task OnBeforeTunnelConnectRequest(object sender, TunnelConnectSessionEventArgs e)
    {
        e.DecryptSsl = ShouldDecryptTunnel(
            e.HttpClient.Request.RequestUri.Host,
            e.HttpClient.ProcessId.Value);
        return Task.CompletedTask;
    }

    // ── Request intercept ─────────────────────────────────────────────────────
    //
    // For most MPSD writes (PUT to /sessions/) we do NOT read the request body:
    // MCC sends gzip-compressed PUT bodies, and if Titanium decompresses via
    // GetRequestBodyAsString() then SetRequestBodyString() re-sends plain text
    // with Content-Encoding: gzip still set — the MPSD server rejects it.
    //
    // For /handles POST/PUT we safely peek at the body using raw bytes:
    //   GetRequestBody()  → raw bytes (no decompression by Titanium)
    //   SetRequestBody()  → restores the exact same bytes unchanged
    //   TryDecompressGzip → decompress locally for display only
    // This gives us visibility into rejoin handle creation without touching the wire.
    private async Task OnBeforeRequestAsync(object sender, SessionEventArgs e)
    {
        var req = e.HttpClient.Request;
        if (!IsDomainWatched(req.RequestUri.Host)) return;

        if (req.Method == "PUT" &&
            req.RequestUri.AbsolutePath.Contains("/CascadeSquadSession/sessions/", StringComparison.OrdinalIgnoreCase))
        {
            var rawBytes = await e.GetRequestBody();
            e.SetRequestBody(rawBytes);
            string? contentEncoding = null;
            foreach (var header in req.Headers)
            {
                if (header.Name.Equals("Content-Encoding", StringComparison.OrdinalIgnoreCase))
                {
                    contentEncoding = header.Value;
                    break;
                }
            }
            string body = TryDecompressGzip(rawBytes, contentEncoding);
            ObserveSquadSessionDocument(req.Url, body, "request-put");
            TryUpgradeGhostSessionConnectionGuid(body);
        }

        // Ghost session interception: intercept MPSD requests during crash recovery
        if (_ghostSessionMode && _ghostSession is not null && IsRequestForGhostSession(req))
        {
            if (HandleGhostSessionRequest(req, e))
            {
                return;  // Request intercepted and handled
            }
        }

        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var h in req.Headers)
            headers[h.Name] = h.Value;

        var requestUri = req.RequestUri;
        if (string.Equals(req.Method, "GET", StringComparison.OrdinalIgnoreCase) &&
            requestUri is not null &&
            requestUri.Host.Equals(BanProcessorHost, StringComparison.OrdinalIgnoreCase) &&
            requestUri.AbsolutePath.StartsWith(BanSummaryPath, StringComparison.OrdinalIgnoreCase) &&
            headers.TryGetValue("X-343-Authorization-Spartan", out var banToken))
        {
            CaptureBanSpartanToken(banToken, DateTimeOffset.UtcNow, BanProcessorHost);
        }

        // ── Capture auth headers for background polling ────────────────────
        var entry = new ProxyCaptureEntry
        {
            Method         = req.Method,
            Url            = req.Url,
            Host           = req.RequestUri.Host,
            Path           = req.RequestUri.PathAndQuery,
            RequestHeaders = headers,
            RequestBody    = "",
        };

        if (req.RequestUri.Host.EndsWith("smartmatch.xboxlive.com", StringComparison.OrdinalIgnoreCase))
        {
            var segments = req.RequestUri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            int serviceConfigsIndex = Array.FindIndex(segments, x => x.Equals("serviceconfigs", StringComparison.OrdinalIgnoreCase));
            lock (_smartMatchAuthLock)
            {
                _smartMatchRequestHeaders = new Dictionary<string, string>(headers, StringComparer.OrdinalIgnoreCase);
                if (serviceConfigsIndex >= 0 && serviceConfigsIndex + 1 < segments.Length)
                    _smartMatchServiceConfigId = segments[serviceConfigsIndex + 1];
            }
        }

        // Safely capture request body for /handles endpoints (rejoin handle observation)
        if (req.HasBody && ShouldCaptureRequestBody(req))
        {
            var rawBytes = await e.GetRequestBody();
            e.SetRequestBody(rawBytes);  // restore exact bytes — no Content-Encoding change
            headers.TryGetValue("content-encoding", out var ce);
            entry.RequestBody = TryDecompressGzip(rawBytes, ce);

            // Persist to disk IMMEDIATELY — before awaiting the response — so a game
            // crash between the request and response doesn't lose the game session ref.
            if (req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
                req.RequestUri.AbsolutePath.Contains("/handles", StringComparison.OrdinalIgnoreCase))
            {
                PersistHandleToDisk(entry.RequestBody, headers);
            }

        }

        // ── Block match leave during crash restore ────────────────────────────
        // MCC cancels a queue either by deleting its SmartMatch ticket or by
        // leaving the associated CascadeMatchTicketSession.
        bool isSmartMatchDelete = req.Method == "DELETE" &&
            req.RequestUri.Host.EndsWith("smartmatch.xboxlive.com", StringComparison.OrdinalIgnoreCase);
        bool isTicketSessionLeave = req.Method == "PUT" &&
            req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.AbsolutePath.Contains("/CascadeMatchTicketSession/sessions/", StringComparison.OrdinalIgnoreCase) &&
            entry.RequestBody.Contains("\"me\":null", StringComparison.Ordinal);
        if (isSmartMatchDelete || isTicketSessionLeave)
            OnSmartMatchWaitCancelled?.Invoke(this, EventArgs.Empty);

        // Cap 16 proved: on startup MCC sends {"members":{"me":null}} to
        // CascadeMatchmaking to leave any leftover match.  This removes the
        // player from MPSD BEFORE the rejoin check (which fires when the user
        // enters matchmaking — Dec 2022 update).
        //
        // Cap 17 proved: SetRequestBodyString() silently fails when the
        // original request is gzip-compressed — the proxy sends plaintext
        // with Content-Encoding: gzip still set, so MPSD ignores/rejects it.
        //
        // Fix: short-circuit with e.Ok() — return a fake 200 to MCC so it
        // thinks the leave succeeded, but NEVER forward the leave to MPSD.
        // Player stays in the match session → rejoin prompt appears when
        // the user queues for matchmaking.
        // BLOCK MATCH LEAVE: Intercept any session leave ("me":null) while blocking is enabled
        // This works for both match sessions (CascadeMatchmaking) and queue sessions (CascadeMatchTicketSession)
        bool isLeaveRequest = req.Method == "PUT" &&
                              req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
                              req.RequestUri.AbsolutePath.Contains("/sessions/", StringComparison.OrdinalIgnoreCase) &&
                              !string.IsNullOrEmpty(entry.RequestBody) &&
                              entry.RequestBody.Contains("\"me\":null");

        if (isLeaveRequest)
        {
            Debug.WriteLine($"[LEAVE] Detected leave request. _blockMatchLeave={_blockMatchLeave}");
            if (_blockMatchLeave)
            {
                // Short-circuit: fake 200 to MCC, leave never reaches MPSD
                Debug.WriteLine("[LEAVE] BLOCKING — returning fake 200");
                e.Ok("{}");
                _blockMatchLeave = false; // one-shot: only block the first leave after crash

                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "BLOCK[MatchLeave]",
                    Url          = entry.Url,
                    Host         = entry.Host,
                    Path         = entry.Path,
                    RequestBody  = "Short-circuited leave PUT → fake 200 to MCC (never reached MPSD)",
                    StatusCode   = 200,
                    ResponseBody = $"Original body: {entry.RequestBody}",
                });

                return; // Skip ETag refresh, discovery, etc. — request is already handled
            }
        }

        // ETag refresh: if MCC is PUTting a session document with a stale If-Match header,
        // silently do a fresh GET of the same URL, extract the current ETag, and swap it
        // in before the PUT reaches MPSD.  This converts a 412 Precondition Failed into a
        // 200, bypassing MCC's missing retry-on-412 logic at crash-rejoin time.
        if (IsSessionPutWithIfMatch(req))
            await RefreshETagAsync(e, req);

        // Persist matchmaking session reference for crash-rejoin restoration.
        // When the player PUTs to a CascadeMatchmaking session, that means they're
        // joining a match — save the session ref + auth headers + connection GUID so we can POST a
        // handle for it after a crash.
        if (req.Method == "PUT" &&
            req.RequestUri.AbsolutePath.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase))
        {
            // Extract request body to get connection GUID
            string matchSessionBody = "";
            if (req.HasBody)
            {
                var rawBytes = await e.GetRequestBody();
                e.SetRequestBody(rawBytes);  // restore exact bytes
                headers.TryGetValue("content-encoding", out var ce);
                matchSessionBody = TryDecompressGzip(rawBytes, ce);
            }
            PersistMatchSessionToDisk(req.Url, headers, matchSessionBody);
        }

        // ── JIT crash restore — PASSIVE MODE ─────────────────────────────────
        // Phase A (JIT-Handle POST) and Phase B (JIT-PUT /members/me) are DISABLED.
        //
        // Capture 15 analysis proved:
        //   1. MCC never checks activity handles (zero GET /handles requests)
        //      → JIT-Handle POST was useless
        //   2. JIT-PUT overwrites the player's pre-crash connection GUID with a
        //      proxy-generated fake 20ms before MCC reads the session, potentially
        //      confusing MCC's rejoin state machine
        //   3. MPSD already returned the match session in discovery naturally
        //      (player was still an active member ~71s after crash)
        //
        // Passive mode: let MCC discover and read the match session with its
        // original pre-crash member state.  INJECT[Member] + FAKE[MatchPut]
        // still fire as fallbacks if the player was removed from the session.

        // Stash so OnBeforeResponseAsync can complete it
        e.UserData = entry;
    }

    /// <summary>Returns true for POST/PUT to known finite matchmaking endpoints.
    /// Captures PUT request bodies to session URLs so we can see what MCC writes
    /// (e.g., the 23-byte match session touch, squad session properties, etc.).</summary>
    private static bool ShouldCaptureRequestBody(Titanium.Web.Proxy.Http.Request req) =>
        (req.Method == "POST" || req.Method == "PUT") &&
        ((req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
          (req.RequestUri.AbsolutePath.Contains("/handles", StringComparison.OrdinalIgnoreCase) ||
           req.RequestUri.AbsolutePath.Contains("/sessions/", StringComparison.OrdinalIgnoreCase))) ||
         req.RequestUri.Host.EndsWith("smartmatch.xboxlive.com", StringComparison.OrdinalIgnoreCase));

    private static string ShortServerId(string serverId) =>
        string.IsNullOrWhiteSpace(serverId)
            ? "unknown"
            : serverId.Length <= 13
                ? serverId
                : $"{serverId[..13]}...";

    private static bool TryParsePlayFabGameServer(string body, out GameServerInfo serverInfo)
    {
        serverInfo = new GameServerInfo();

        if (!body.Contains("IPV4Address", StringComparison.OrdinalIgnoreCase) &&
            !body.Contains("ipv4Address", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        using var doc = JsonDocument.Parse(body);
        if (!TryFindGameServerElement(doc.RootElement, out var data))
            return false;

        serverInfo = new GameServerInfo
        {
            PartyId = GetStringProperty(data, "PartyId", "partyId"),
            ServerId = GetStringProperty(data, "ServerId", "serverId"),
            VmId = GetStringProperty(data, "VmId", "vmId"),
            IPv4Address = GetStringProperty(data, "IPV4Address", "ipv4Address"),
            FQDN = GetStringProperty(data, "FQDN", "fqdn"),
            Region = GetStringProperty(data, "Region", "region"),
            State = GetStringProperty(data, "State", "state"),
            BuildId = GetStringProperty(data, "BuildId", "buildId"),
            DTLSCertificateSHA2Thumbprint = GetStringProperty(data, "DTLSCertificateSHA2Thumbprint", "dtlsCertificateSHA2Thumbprint"),
            CachedAt = DateTime.UtcNow
        };

        if (TryGetPropertyAnyCase(data, out var ports, "Ports", "ports") &&
            ports.ValueKind == JsonValueKind.Array)
        {
            foreach (var port in ports.EnumerateArray())
            {
                serverInfo.Ports.Add(new GameServerPort
                {
                    Name = GetStringProperty(port, "Name", "name"),
                    Num = GetIntProperty(port, "Num", "num"),
                    Protocol = GetStringProperty(port, "Protocol", "protocol")
                });
            }
        }

        return !string.IsNullOrWhiteSpace(serverInfo.IPv4Address);
    }

    private static bool TryFindGameServerElement(JsonElement element, out JsonElement serverElement)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            if (TryGetPropertyAnyCase(element, out _, "IPV4Address", "ipv4Address"))
            {
                serverElement = element;
                return true;
            }

            if (TryGetPropertyAnyCase(element, out var data, "data") &&
                TryFindGameServerElement(data, out serverElement))
            {
                return true;
            }

            foreach (var property in element.EnumerateObject())
            {
                if (TryFindGameServerElement(property.Value, out serverElement))
                    return true;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                if (TryFindGameServerElement(item, out serverElement))
                    return true;
            }
        }

        serverElement = default;
        return false;
    }

    private void ObserveNetworkRegionLatencies(string body, string source)
    {
        if (!body.Contains("\"members\"", StringComparison.OrdinalIgnoreCase) &&
            !body.Contains("serverMeasurements", StringComparison.OrdinalIgnoreCase))
            return;

        try
        {
            using var doc = JsonDocument.Parse(body);
            ObserveMatchmakingPlayerPings(doc.RootElement, source);

            if (!body.Contains("serverMeasurements", StringComparison.OrdinalIgnoreCase))
                return;

            if (!TryFindBestNetworkRegionLatency(doc.RootElement, out var best))
                return;

            bool changed = _bestNetworkRegionLatency is null
                || !string.Equals(_bestNetworkRegionLatency.Region, best.Region, StringComparison.OrdinalIgnoreCase)
                || _bestNetworkRegionLatency.LatencyMs != best.LatencyMs;

            _bestNetworkRegionLatency = best;

            if (!changed)
                return;

            RejoinFixDiagnostics.Info("network", $"Best MCC network region is {best.Region} ({best.LatencyMs} ms) via {source}.");
            OnNetworkRegionLatencyChanged?.Invoke(this, best);
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("network", $"Failed to parse MCC network region measurements from {source}: {ex.Message}");
        }
    }

    private void ObserveMatchmakingPlayerPings(JsonElement root, string source)
    {
        var observed = new List<MatchmakingPlayerPing>();
        FindMatchmakingPlayerPings(root, observed);

        if (observed.Count == 0)
            return;

        var nextPings = new Dictionary<string, MatchmakingPlayerPing>(StringComparer.OrdinalIgnoreCase);
        foreach (var ping in observed)
        {
            ping.Xuid = NormalizeXuid(ping.Xuid);
            if (string.IsNullOrWhiteSpace(ping.Xuid))
                continue;

            // An MPSD document commonly contains the same member in several nested
            // objects.  Some copies only contain measurements while another holds
            // identity/group data.  Merge every copy instead of letting the last,
            // often sparse, occurrence erase a gamertag or squad discovered earlier.
            if (nextPings.TryGetValue(ping.Xuid, out var observedEarlier))
                MergeMatchmakingPlayerPing(ping, observedEarlier);

            if (_matchmakingPlayerPings.TryGetValue(ping.Xuid, out var existing))
                MergeMatchmakingPlayerPing(ping, existing);

            nextPings[ping.Xuid] = ping;
        }

        PreserveRecentLobbyPingsForPartialObservation(nextPings, source);

        bool changed = _matchmakingPlayerPings.Count != nextPings.Count ||
            nextPings.Any(kvp =>
                !_matchmakingPlayerPings.TryGetValue(kvp.Key, out var existing) ||
                !string.Equals(existing.Region, kvp.Value.Region, StringComparison.OrdinalIgnoreCase) ||
                existing.LatencyMs != kvp.Value.LatencyMs ||
                !string.Equals(existing.Gamertag, kvp.Value.Gamertag, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(existing.Team, kvp.Value.Team, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(existing.SquadId, kvp.Value.SquadId, StringComparison.OrdinalIgnoreCase) ||
                existing.AverageGroupSkillPercentile != kvp.Value.AverageGroupSkillPercentile);

        _matchmakingPlayerPings.Clear();
        foreach (var kvp in nextPings)
            _matchmakingPlayerPings[kvp.Key] = kvp.Value;

        if (!changed)
            return;

        var snapshot = _matchmakingPlayerPings.Values
            .OrderBy(p => p.Gamertag)
            .ThenBy(p => p.Xuid)
            .ToList();

        RejoinFixDiagnostics.Info("network", $"Observed current lobby members for {snapshot.Count} players via {source}.");
        PersistMatchmakingPlayerPings(snapshot, source);
        OnMatchmakingPlayerPingsObserved?.Invoke(this, snapshot);
    }

    private static void MergeMatchmakingPlayerPing(
        MatchmakingPlayerPing target,
        MatchmakingPlayerPing fallback)
    {
        if (string.IsNullOrWhiteSpace(target.Region))
            target.Region = fallback.Region;
        if (target.LatencyMs <= 0)
            target.LatencyMs = fallback.LatencyMs;
        if (string.IsNullOrWhiteSpace(target.Gamertag))
            target.Gamertag = fallback.Gamertag;
        if (string.IsNullOrWhiteSpace(target.Team))
            target.Team = fallback.Team;
        if (string.IsNullOrWhiteSpace(target.SquadId))
            target.SquadId = fallback.SquadId;
        if (!target.AverageGroupSkillPercentile.HasValue)
            target.AverageGroupSkillPercentile = fallback.AverageGroupSkillPercentile;
    }

    private static string NormalizeXuid(string xuid)
    {
        if (string.IsNullOrWhiteSpace(xuid))
            return "";

        string trimmed = xuid.Trim();
        if (!trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            return trimmed;

        return ulong.TryParse(
                trimmed[2..],
                System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture,
                out ulong value)
            ? value.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : trimmed;
    }

    private void PreserveRecentLobbyPingsForPartialObservation(
        Dictionary<string, MatchmakingPlayerPing> nextPings,
        string source)
    {
        if (nextPings.Count == 0 || nextPings.Count >= _matchmakingPlayerPings.Count)
            return;

        bool looksPartialRefresh = source.Contains("squad", StringComparison.OrdinalIgnoreCase) ||
            nextPings.Keys.All(xuid => _matchmakingPlayerPings.ContainsKey(xuid));
        if (!looksPartialRefresh)
            return;

        var cutoff = DateTime.UtcNow - TimeSpan.FromMinutes(10);
        foreach (var existing in _matchmakingPlayerPings.Values)
        {
            if (nextPings.ContainsKey(existing.Xuid) || existing.ObservedAt < cutoff)
                continue;

            nextPings[existing.Xuid] = existing;
        }
    }

    private void ClearMatchmakingPlayerPings(string reason)
    {
        if (_matchmakingPlayerPings.Count == 0)
            return;

        if (ShouldPreserveLobbyPingsDuringActiveGame(reason))
        {
            RejoinFixDiagnostics.Info("network", $"Ignored current lobby clear during active game: {reason}.");
            return;
        }

        _matchmakingPlayerPings.Clear();
        RejoinFixDiagnostics.Info("network", $"Cleared current lobby members: {reason}.");
        PersistMatchmakingPlayerPings(Array.Empty<MatchmakingPlayerPing>(), reason);
        OnMatchmakingPlayerPingsObserved?.Invoke(this, Array.Empty<MatchmakingPlayerPing>());
    }

    private bool ShouldPreserveLobbyPingsDuringActiveGame(string reason)
    {
        if (!reason.Contains("unconnected squad", StringComparison.OrdinalIgnoreCase))
            return false;

        if (_currentObservedGameServerInfo is null ||
            string.IsNullOrWhiteSpace(_currentObservedGameServerInfo.IPv4Address) ||
            _currentObservedGameServerInfo.CachedAt < DateTime.UtcNow - TimeSpan.FromMinutes(30))
        {
            return false;
        }

        return _matchmakingPlayerPings.Values.Any(p =>
            p.ObservedAt >= DateTime.UtcNow - TimeSpan.FromMinutes(10));
    }

    private static void PersistMatchmakingPlayerPings(IReadOnlyList<MatchmakingPlayerPing> snapshot, string source)
    {
        try
        {
            RejoinFixPaths.EnsureRootDirectory();
            File.WriteAllText(
                RejoinFixPaths.LastMatchmakingPingsFile,
                JsonSerializer.Serialize(snapshot, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("network", $"Failed to persist matchmaking ping snapshot: {ex.Message}");
        }

        foreach (var ping in snapshot)
        {
            string name = string.IsNullOrWhiteSpace(ping.Gamertag)
                ? ShortXuid(ping.Xuid)
                : ping.Gamertag;
            RejoinFixDiagnostics.Info(
                "network",
                $"LOBBY {name} ({ShortXuid(ping.Xuid)}) team={FormatTeamForLog(ping.Team)} squad={FormatSquadForLog(ping.SquadId)} {FormatRegionForLog(ping.Region)} | {ping.DisplayPing} ms via {source}");
        }
    }

    private static string TruncateForLog(string value, int maxLength) =>
        string.IsNullOrEmpty(value) || value.Length <= maxLength
            ? value
            : value[..maxLength] + "...";

    private static string FormatTeamForLog(string team) =>
        string.IsNullOrWhiteSpace(team) ? "?" : team.Trim();

    private static string FormatSquadForLog(string squadId) =>
        string.IsNullOrWhiteSpace(squadId) ? "?" : ShortSquadId(squadId);

    private static string ShortSquadId(string squadId)
    {
        if (string.IsNullOrWhiteSpace(squadId))
            return "unknown";

        var trimmed = squadId.Trim();
        return trimmed.Length <= 8 ? trimmed : $"...{trimmed[^8..]}";
    }

    private static string ShortXuid(string xuid)
    {
        if (string.IsNullOrWhiteSpace(xuid))
            return "unknown";

        var trimmed = xuid.Trim();
        return trimmed.Length <= 8 ? trimmed : $"...{trimmed[^8..]}";
    }

    private static string FormatRegionForLog(string region)
    {
        if (string.IsNullOrWhiteSpace(region))
            return "Unknown";

        var known = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["WestUs"] = "West US",
            ["SouthCentralUs"] = "South Central US",
            ["CentralUs"] = "Central US",
            ["NorthCentralUs"] = "North Central US",
            ["EastUs"] = "East US",
            ["EastUs2"] = "East US 2",
            ["BrazilSouth"] = "Brazil South",
            ["NorthEurope"] = "North Europe",
            ["WestEurope"] = "West Europe",
            ["SoutheastAsia"] = "Southeast Asia",
            ["EastAsia"] = "East Asia",
            ["JapanWest"] = "Japan West",
            ["JapanEast"] = "Japan East",
            ["AustraliaSoutheast"] = "Australia Southeast",
            ["AustraliaEast"] = "Australia East"
        };

        if (known.TryGetValue(region.Trim(), out var label))
            return label;

        var spaced = System.Text.RegularExpressions.Regex.Replace(
            region.Trim(),
            "([a-z])([A-Z0-9])",
            "$1 $2");
        return spaced.ToUpperInvariant() == spaced
            ? spaced
            : System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(spaced);
    }

    private static void FindMatchmakingPlayerPings(JsonElement element, List<MatchmakingPlayerPing> pings)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (property.NameEquals("members") && property.Value.ValueKind == JsonValueKind.Object)
                {
                    foreach (var member in property.Value.EnumerateObject())
                    {
                        if (TryParseMatchmakingPlayerPing(member.Value, member.Name, out var ping))
                            pings.Add(ping);
                    }
                }

                FindMatchmakingPlayerPings(property.Value, pings);
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
                FindMatchmakingPlayerPings(item, pings);
        }
    }

    private static bool TryParseMatchmakingPlayerPing(JsonElement member, string memberKey, out MatchmakingPlayerPing ping)
    {
        ping = new MatchmakingPlayerPing();

        if (member.ValueKind != JsonValueKind.Object)
            return false;

        string xuid = FindStringPropertyRecursive(member, "xuid", "Xuid", "xboxUserId", "XboxUserId");
        if (string.IsNullOrWhiteSpace(xuid) && LooksLikeSessionMemberXuid(memberKey))
            xuid = memberKey;
        string gamertag = FindStringPropertyRecursive(member, "gamertag", "Gamertag", "gamerTag", "GamerTag");
        if (string.IsNullOrWhiteSpace(xuid) && string.IsNullOrWhiteSpace(gamertag))
            return false;

        TryFindBestNetworkRegionLatency(member, out var best);

        ping = new MatchmakingPlayerPing
        {
            Xuid = xuid,
            Gamertag = gamertag,
            Team = FindStringPropertyRecursive(member, "initialTeam", "InitialTeam", "team", "Team", "teamId", "TeamId"),
            SquadId = FindStringPropertyRecursive(member, "squadId", "SquadId", "squadID", "SquadID", "partySquadId", "PartySquadId", "groupId", "GroupId"),
            Region = best.Region,
            LatencyMs = best.LatencyMs,
            AverageGroupSkillPercentile = FindNullableDoublePropertyRecursive(
                member,
                "AverageGroupSkillPercentile",
                "averageGroupSkillPercentile"),
            ObservedAt = DateTime.UtcNow
        };

        return true;
    }

    private static bool LooksLikeSessionMemberXuid(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Equals("me", StringComparison.OrdinalIgnoreCase))
            return false;

        string trimmed = value.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            return trimmed.Length > 2 &&
                   trimmed[2..].All(Uri.IsHexDigit);

        return trimmed.Length >= 12 && trimmed.All(char.IsDigit);
    }

    private static bool TryFindBestNetworkRegionLatency(JsonElement element, out NetworkRegionLatency best)
    {
        NetworkRegionLatency? result = null;
        FindBestNetworkRegionLatency(element, ref result);
        best = result ?? new NetworkRegionLatency();
        return result is not null;
    }

    private static void FindBestNetworkRegionLatency(JsonElement element, ref NetworkRegionLatency? best)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (property.NameEquals("serverMeasurements") &&
                    property.Value.ValueKind == JsonValueKind.Object)
                {
                    foreach (var region in property.Value.EnumerateObject())
                    {
                        if (TryReadRegionLatency(region.Value, out int latencyMs) &&
                            latencyMs >= 0 &&
                            (best is null || latencyMs < best.LatencyMs))
                        {
                            best = new NetworkRegionLatency
                            {
                                Region = region.Name,
                                LatencyMs = latencyMs,
                                ObservedAt = DateTime.UtcNow
                            };
                        }
                    }
                }

                FindBestNetworkRegionLatency(property.Value, ref best);
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
                FindBestNetworkRegionLatency(item, ref best);
        }
    }

    private static bool TryReadRegionLatency(JsonElement element, out int latencyMs)
    {
        latencyMs = 0;

        if (element.ValueKind == JsonValueKind.Number)
            return element.TryGetInt32(out latencyMs);

        if (element.ValueKind == JsonValueKind.Object &&
            TryGetPropertyAnyCase(element, out var latency, "latency", "Latency", "latencyMs", "LatencyMs"))
        {
            if (latency.ValueKind == JsonValueKind.Number)
                return latency.TryGetInt32(out latencyMs);

            if (latency.ValueKind == JsonValueKind.String)
                return int.TryParse(latency.GetString(), out latencyMs);
        }

        return false;
    }

    private static string FindStringPropertyRecursive(JsonElement element, params string[] names)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (names.Any(name => string.Equals(name, property.Name, StringComparison.OrdinalIgnoreCase)) &&
                    property.Value.ValueKind == JsonValueKind.String)
                {
                    return property.Value.GetString() ?? "";
                }
                if (names.Any(name => string.Equals(name, property.Name, StringComparison.OrdinalIgnoreCase)) &&
                    property.Value.ValueKind == JsonValueKind.Number)
                {
                    return property.Value.GetRawText();
                }

                var nested = FindStringPropertyRecursive(property.Value, names);
                if (!string.IsNullOrWhiteSpace(nested))
                    return nested;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var nested = FindStringPropertyRecursive(item, names);
                if (!string.IsNullOrWhiteSpace(nested))
                    return nested;
            }
        }

        return "";
    }

    private static double? FindNullableDoublePropertyRecursive(JsonElement element, params string[] names)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (names.Any(name => string.Equals(name, property.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    if (property.Value.ValueKind == JsonValueKind.Number &&
                        property.Value.TryGetDouble(out double number))
                    {
                        return number;
                    }

                    if (property.Value.ValueKind == JsonValueKind.String &&
                        double.TryParse(
                            property.Value.GetString(),
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out double parsed))
                    {
                        return parsed;
                    }
                }

                var nested = FindNullableDoublePropertyRecursive(property.Value, names);
                if (nested.HasValue)
                    return nested;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var nested = FindNullableDoublePropertyRecursive(item, names);
                if (nested.HasValue)
                    return nested;
            }
        }

        return null;
    }

    private static bool TryGetPropertyAnyCase(JsonElement element, out JsonElement value, params string[] names)
    {
        foreach (var name in names)
        {
            if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out value))
                return true;
        }

        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (names.Any(name => string.Equals(name, property.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static string GetStringProperty(JsonElement element, params string[] names) =>
        TryGetPropertyAnyCase(element, out var value, names) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    private static int GetIntProperty(JsonElement element, params string[] names)
    {
        if (!TryGetPropertyAnyCase(element, out var value, names))
            return 0;

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int result))
            return result;

        return value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out result)
            ? result
            : 0;
    }

    /// <summary>
    /// Returns true when MCC is PUTting a session DOCUMENT (not a /members/ sub-resource)
    /// and has included an If-Match header — the condition that triggers MPSD 412 replies
    /// when the ETag went stale during MCC's 40-second startup sequence.
    /// </summary>
    private static bool IsSessionPutWithIfMatch(Titanium.Web.Proxy.Http.Request req)
    {
        if (req.Method != "PUT") return false;
        if (!req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase)) return false;
        var path = req.RequestUri.AbsolutePath;
        if (!path.Contains("/sessions/",      StringComparison.OrdinalIgnoreCase)) return false;
        if ( path.Contains("/members/",       StringComparison.OrdinalIgnoreCase)) return false; // member sub-resource, not session doc
        foreach (var h in req.Headers)
            if (h.Name.Equals("If-Match", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    /// <summary>
    /// Performs a fresh GET of the MPSD session URL, extracts the current ETag, and
    /// replaces the stale If-Match value in MCC's outgoing PUT.  All of this happens
    /// BEFORE Titanium forwards the PUT to MPSD — the game sees a 200 instead of 412.
    /// Logs a synthetic "GET[ETag↑]" capture entry so the intervention is visible in the UI.
    /// Best-effort: any failure leaves the original PUT headers untouched.
    /// </summary>
    private async Task RefreshETagAsync(SessionEventArgs e, Titanium.Web.Proxy.Http.Request req)
    {
        var sessionUrl = req.Url;

        // Stash the stale ETag for the log entry
        string oldEtag = "";
        foreach (var h in req.Headers)
            if (h.Name.Equals("If-Match", StringComparison.OrdinalIgnoreCase))
                { oldEtag = h.Value; break; }

        try
        {
            // Fresh GET — proxy-bypassing so we don't re-intercept ourselves
            using var getReq = new HttpRequestMessage(HttpMethod.Get, sessionUrl);
            foreach (var h in req.Headers)
                if (h.Name.StartsWith("x-",        StringComparison.OrdinalIgnoreCase) ||
                    h.Name.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    getReq.Headers.TryAddWithoutValidation(h.Name, h.Value);

            using var getResp = await _refreshClient.SendAsync(getReq);
            int getCode = (int)getResp.StatusCode;

            // Pull the ETag from the standard response header
            string freshEtag = getResp.Headers.ETag?.Tag ?? "";

            // Emit a synthetic capture entry so the user can see the intervention
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "GET[ETag↑]",
                Url          = sessionUrl,
                Host         = req.RequestUri.Host,
                Path         = req.RequestUri.PathAndQuery,
                RequestBody  = $"ETag refresh before PUT\nold If-Match: {oldEtag}",
                StatusCode   = getCode,
                ResponseBody = string.IsNullOrEmpty(freshEtag)
                    ? "[no ETag in response — PUT forwarded unmodified]"
                    : $"fresh ETag: {freshEtag}\ninjected into PUT If-Match header",
            });

            if (string.IsNullOrEmpty(freshEtag) || !getResp.IsSuccessStatusCode) return;

            // Swap the stale ETag for the current one in MCC's outgoing PUT
            req.Headers.RemoveHeader("If-Match");
            req.Headers.AddHeader("If-Match", freshEtag);
        }
        catch (Exception ex)
        {
            // Never break the proxy — log the failure and let the original PUT go through
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "GET[ETag↑]",
                Url          = sessionUrl,
                Host         = req.RequestUri.Host,
                Path         = req.RequestUri.PathAndQuery,
                RequestBody  = $"ETag refresh failed\nold If-Match: {oldEtag}",
                StatusCode   = 0,
                ResponseBody = ex.Message,
            });
        }
    }

    /// <summary>Decompresses gzip bytes for display. Falls back to UTF-8 decode if not gzip.</summary>
    private static string TryDecompressGzip(byte[] data, string? contentEncoding)
    {
        try
        {
            if (string.Equals(contentEncoding, "gzip", StringComparison.OrdinalIgnoreCase))
            {
                using var ms = new MemoryStream(data);
                using var gz = new GZipStream(ms, CompressionMode.Decompress);
                using var sr = new StreamReader(gz, System.Text.Encoding.UTF8);
                return sr.ReadToEnd();
            }
            return System.Text.Encoding.UTF8.GetString(data);
        }
        catch
        {
            return $"[{data.Length} bytes — decode failed]";
        }
    }

    /// <summary>
    /// Parses a /handles POST body and writes the session reference + auth headers to
    /// %LocalAppData%\HaloMCCToolbox\RejoinFix\last-handle.json. Called synchronously before the
    /// response arrives so the data survives a game crash. Best-effort; never throws.
    /// </summary>
    private void PersistHandleToDisk(string bodyJson, Dictionary<string, string> requestHeaders)
    {
        try
        {
            using var doc  = JsonDocument.Parse(bodyJson);
            var       root = doc.RootElement;
            if (!root.TryGetProperty("sessionRef", out var refEl)) return;

            var info = new SavedHandleInfo
            {
                Scid         = refEl.TryGetProperty("scid",         out var scidEl) ? scidEl.GetString() ?? "" : "",
                TemplateName = refEl.TryGetProperty("templateName", out var tmEl)   ? tmEl.GetString()   ?? "" : "",
                SessionName  = refEl.TryGetProperty("name",         out var nameEl) ? nameEl.GetString() ?? "" : "",
                SavedAt      = DateTime.UtcNow,
                RequestHeaders = new Dictionary<string, string>(requestHeaders, StringComparer.OrdinalIgnoreCase),
            };

            if (string.IsNullOrEmpty(info.Scid) || string.IsNullOrEmpty(info.SessionName)) return;

            if (info.TemplateName.Equals("cascadesquadsession", StringComparison.OrdinalIgnoreCase))
            {
                RejoinSquadState? observedSquad = _lastSquadState;
                if (observedSquad is not null &&
                    string.Equals(observedSquad.SessionName, info.SessionName, StringComparison.OrdinalIgnoreCase))
                {
                    info.ObservedSquadMemberCount = observedSquad.MemberCount;
                    info.ObservedSquadSessionName = observedSquad.SessionName;
                }
            }

            if (_pendingCrashRestore &&
                ShouldKeepExistingSquadHandleDuringRestore(_lastSquadHandle, info))
            {
                RejoinFixDiagnostics.Warn(
                    "restore",
                    $"Kept saved party squad handle {_lastSquadHandle!.SessionShort}; ignored weaker post-crash handle {info.SessionShort} [{info.Mode.ToDisplayLabel()}].");
                return;
            }

            var dir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox", "RejoinFix");
            Directory.CreateDirectory(dir);
            File.WriteAllText(
                Path.Combine(dir, "last-handle.json"),
                JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true }));
            if (info.TemplateName.Equals("cascadesquadsession", StringComparison.OrdinalIgnoreCase))
                _lastSquadHandle = info;

            string modeSuffix = info.Mode == RejoinSessionMode.Unknown
                ? ""
                : $" [{info.Mode.ToDisplayLabel()}]";
            RejoinFixDiagnostics.Info("capture", $"Saved activity handle for {info.TemplateName}/{info.SessionShort}{modeSuffix}.");
        }
        catch
        {
            RejoinFixDiagnostics.Warn("capture", "Failed to save activity handle.");
        }
    }

    /// <summary>
    /// Parses a CascadeMatchmaking session URL and writes the session reference + auth
    /// headers + connection GUID to %LocalAppData%\HaloMCCToolbox\RejoinFix\last-match-session.json.
    /// Called when MCC PUTs to a matchmaking session (joining a match).  Best-effort; never throws.
    /// </summary>
    private void PersistMatchSessionToDisk(string url, Dictionary<string, string> requestHeaders, string requestBody = "")
    {
        try
        {
            // URL: https://sessiondirectory.xboxlive.com/serviceconfigs/{scid}
            //      /sessionTemplates/CascadeMatchmaking/sessions/{name}
            var uri  = new Uri(url);
            var segs = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);

            // Expected segments: serviceconfigs / {scid} / sessionTemplates / CascadeMatchmaking / sessions / {name}
            string scid         = "";
            string templateName = "";
            string sessionName  = "";
            for (int i = 0; i < segs.Length; i++)
            {
                if (segs[i].Equals("serviceconfigs", StringComparison.OrdinalIgnoreCase) && i + 1 < segs.Length)
                    scid = segs[i + 1];
                if (segs[i].Equals("sessionTemplates", StringComparison.OrdinalIgnoreCase) && i + 1 < segs.Length)
                    templateName = segs[i + 1];
                if (segs[i].Equals("sessions", StringComparison.OrdinalIgnoreCase) && i + 1 < segs.Length)
                    sessionName = segs[i + 1];
            }

            if (string.IsNullOrEmpty(scid) || string.IsNullOrEmpty(sessionName)) return;

            // Extract connection GUID from request body if present
            // Format: {"members":{"me":{"properties":{"system":{"active":true,"connection":"<GUID>"}}}}}
            string connectionGuid = "";
            if (!string.IsNullOrEmpty(requestBody))
            {
                try
                {
                    using var doc = JsonDocument.Parse(requestBody);
                    var root = doc.RootElement;
                    if (root.TryGetProperty("members", out var members) &&
                        members.TryGetProperty("me", out var me) &&
                        me.TryGetProperty("properties", out var props) &&
                        props.TryGetProperty("system", out var sys) &&
                        sys.TryGetProperty("connection", out var conn))
                    {
                        connectionGuid = conn.GetString() ?? "";
                        Debug.WriteLine($"[SAVE-Match] Captured connection GUID: {connectionGuid}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[SAVE-Match] Failed to extract connection GUID: {ex.Message}");
                }
            }

            var info = new SavedHandleInfo
            {
                Scid           = scid,
                TemplateName   = templateName,
                SessionName    = sessionName,
                SavedAt        = DateTime.UtcNow,
                ConnectionGuid = connectionGuid,
                RequestHeaders = new Dictionary<string, string>(requestHeaders, StringComparer.OrdinalIgnoreCase),
            };

            if (_lastSquadState is not null)
            {
                info.ObservedSquadMemberCount = _lastSquadState.MemberCount;
                info.ObservedSquadSessionName = _lastSquadState.SessionName;
            }

            if (_lastMatchSession is not null &&
                string.Equals(_lastMatchSession.TemplateName, info.TemplateName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_lastMatchSession.SessionName, info.SessionName, StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(info.ConnectionGuid) &&
                    !string.IsNullOrWhiteSpace(_lastMatchSession.ConnectionGuid))
                {
                    info.ConnectionGuid = _lastMatchSession.ConnectionGuid;
                }

                if (string.IsNullOrWhiteSpace(info.PlayerXuid) &&
                    !string.IsNullOrWhiteSpace(_lastMatchSession.PlayerXuid))
                {
                    info.PlayerXuid = _lastMatchSession.PlayerXuid;
                }

                bool preserveExistingSquadContext =
                    info.ObservedSquadMemberCount < _lastMatchSession.ObservedSquadMemberCount;

                if (preserveExistingSquadContext)
                {
                    info.ObservedSquadMemberCount = _lastMatchSession.ObservedSquadMemberCount;
                }

                if ((string.IsNullOrWhiteSpace(info.ObservedSquadSessionName) ||
                     preserveExistingSquadContext) &&
                    !string.IsNullOrWhiteSpace(_lastMatchSession.ObservedSquadSessionName))
                {
                    info.ObservedSquadSessionName = _lastMatchSession.ObservedSquadSessionName;
                }
            }

            if (_pendingCrashRestore &&
                _lastMatchSession is not null)
            {
                if (!string.IsNullOrWhiteSpace(info.ConnectionGuid) &&
                    !string.Equals(_lastMatchSession.ConnectionGuid, info.ConnectionGuid, StringComparison.OrdinalIgnoreCase))
                {
                    _lastMatchSession.ConnectionGuid = info.ConnectionGuid;
                    PersistSavedMatchSessionSnapshot(_lastMatchSession);
                    RejoinFixDiagnostics.Info("guid", $"Upgraded saved match session with replacement connection GUID from post-crash match touch: {info.ConnectionGuid}");
                }

                RejoinFixDiagnostics.Warn(
                    "restore",
                    $"Kept saved match session {_lastMatchSession.TemplateName}/{_lastMatchSession.SessionShort}; ignored post-crash match capture {info.TemplateName}/{info.SessionShort} (guid={info.ConnectionGuid}).");

                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "KEEP[Match]",
                    Url          = url,
                    Host         = "sessiondirectory.xboxlive.com",
                    Path         = new Uri(url).AbsolutePath,
                    RequestBody  = "Crash restore is armed; preserved saved match context and only accepted replacement GUID",
                    StatusCode   = 0,
                    ResponseBody = $"Preserved {_lastMatchSession.TemplateName}/{_lastMatchSession.SessionShort}",
                });
                return;
            }

            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox", "RejoinFix");
            Directory.CreateDirectory(dir);
            File.WriteAllText(
                Path.Combine(dir, "last-match-session.json"),
                JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true }));
            RejoinFixDiagnostics.Info("capture", $"Saved match session {info.TemplateName}/{info.SessionShort} (guid={info.ConnectionGuid}).");

            _lastMatchSession = info;

            // Diagnostic: emit a synthetic entry so the capture log shows this fired
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "SAVE[Match]",
                Url          = url,
                Host         = "sessiondirectory.xboxlive.com",
                Path         = new Uri(url).AbsolutePath,
                RequestBody  = $"template={templateName}  session={sessionName}",
                StatusCode   = 0,
                ResponseBody = $"Wrote last-match-session.json\n_lastMatchSession set",
            });

            OnMatchSessionSaved?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("capture", $"Failed to save match session: {ex.Message}");
            // Log the failure so we can diagnose why the save didn't work
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "SAVE[Match]",
                Url          = url,
                Host         = "sessiondirectory.xboxlive.com",
                Path         = "ERROR",
                RequestBody  = ex.GetType().Name,
                StatusCode   = 0,
                ResponseBody = ex.Message,
            });
        }
    }

    private static bool ShouldKeepExistingMatchSessionDuringRestore(SavedHandleInfo existing, SavedHandleInfo incoming)
    {
        if ((DateTime.UtcNow - existing.SavedAt).TotalMinutes > 30)
            return false;

        if (IsSavedMatchSessionMoreComplete(existing, incoming))
            return true;

        bool existingHasConnection = !string.IsNullOrWhiteSpace(existing.ConnectionGuid);
        bool incomingHasConnection = !string.IsNullOrWhiteSpace(incoming.ConnectionGuid);

        return existingHasConnection && !incomingHasConnection;
    }

    private static bool IsSavedMatchSessionMoreComplete(SavedHandleInfo left, SavedHandleInfo right) =>
        GetSavedMatchSessionCompletenessScore(left) > GetSavedMatchSessionCompletenessScore(right);

    private static bool IsBetterRestoreSource(SavedHandleInfo candidate, SavedHandleInfo current)
    {
        if (candidate.ObservedSquadMemberCount > current.ObservedSquadMemberCount)
            return true;

        if (candidate.ObservedSquadMemberCount < current.ObservedSquadMemberCount)
            return false;

        return IsSavedMatchSessionMoreComplete(candidate, current);
    }

    private static int GetSavedMatchSessionCompletenessScore(SavedHandleInfo info)
    {
        int score = 0;

        if (!string.IsNullOrWhiteSpace(info.Scid)) score++;
        if (!string.IsNullOrWhiteSpace(info.TemplateName)) score++;
        if (!string.IsNullOrWhiteSpace(info.SessionName)) score++;
        if (info.ObservedSquadMemberCount > 1) score += 5;
        else if (info.ObservedSquadMemberCount == 1) score++;
        if (!string.IsNullOrWhiteSpace(info.ConnectionGuid)) score += 4;
        if (!string.IsNullOrWhiteSpace(info.PlayerXuid)) score += 3;
        if (!string.IsNullOrWhiteSpace(info.SubscriptionId)) score += 2;
        if (info.RequestHeaders.ContainsKey("Authorization")) score += 2;
        if (info.RequestHeaders.ContainsKey("Signature")) score++;
        if (info.RequestHeaders.ContainsKey("If-Match")) score++;

        return score;
    }

    private static bool ShouldKeepExistingSquadHandleDuringRestore(SavedHandleInfo? existing, SavedHandleInfo incoming)
    {
        if (existing is null)
            return false;

        if (!IsSquadHandle(existing) || !IsSquadHandle(incoming))
            return false;

        if ((DateTime.UtcNow - existing.SavedAt).TotalMinutes > 30)
            return false;

        return existing.IsPartySquad && !incoming.IsPartySquad;
    }

    private static bool IsSquadHandle(SavedHandleInfo info) =>
        info.TemplateName.Equals("cascadesquadsession", StringComparison.OrdinalIgnoreCase);

    private void ObserveSquadSessionDocument(string url, string body, string source)
    {
        RejoinSquadState? nextState = null;
        int connectedCount = 0;
        bool observedMemberConnectionState = false;

        try
        {
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;

            int memberCount = 0;
            int acceptedCount = 0;
            int activeCount = 0;

            if (root.TryGetProperty("membersInfo", out var membersInfo))
            {
                if (membersInfo.TryGetProperty("count", out var countEl) && countEl.TryGetInt32(out var parsedCount))
                    memberCount = parsedCount;
                if (membersInfo.TryGetProperty("accepted", out var acceptedEl) && acceptedEl.TryGetInt32(out var parsedAccepted))
                    acceptedCount = parsedAccepted;
                if (membersInfo.TryGetProperty("active", out var activeEl) && activeEl.TryGetInt32(out var parsedActive))
                    activeCount = parsedActive;
            }

            if (memberCount == 0 &&
                root.TryGetProperty("members", out var members) &&
                members.ValueKind == JsonValueKind.Object)
            {
                memberCount = members.EnumerateObject().Count();
            }

            if (root.TryGetProperty("members", out var connectionMembers) &&
                connectionMembers.ValueKind == JsonValueKind.Object)
            {
                observedMemberConnectionState = true;
                connectedCount = connectionMembers.EnumerateObject()
                    .Count(member => !string.IsNullOrWhiteSpace(
                        FindStringPropertyRecursive(member.Value, "connection")));
            }

            if (memberCount <= 0)
                return;

            string sessionName = TryGetSessionNameFromUrl(url);
            if (string.IsNullOrWhiteSpace(sessionName))
                return;

            nextState = new RejoinSquadState
            {
                SessionName = sessionName,
                SavedAt = DateTime.UtcNow,
                MemberCount = memberCount,
                AcceptedCount = acceptedCount,
                ActiveCount = activeCount,
            };
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("squad", $"Failed to parse squad session document from {source}: {ex.Message}");
            return;
        }

        ArmCrashRestoreIfRestartSquadWriteDetected(nextState, source, observedMemberConnectionState, connectedCount);

        bool changed = _lastSquadState is null
            || !string.Equals(_lastSquadState.SessionName, nextState.SessionName, StringComparison.OrdinalIgnoreCase)
            || _lastSquadState.MemberCount != nextState.MemberCount
            || _lastSquadState.AcceptedCount != nextState.AcceptedCount
            || _lastSquadState.ActiveCount != nextState.ActiveCount;

        ArmCrashRestoreIfRestartSquadDetected(_lastSquadState, nextState, source);

        if (_pendingCrashRestore && ShouldKeepExistingSquadStateDuringRestore(_lastSquadState, nextState))
        {
            RejoinFixDiagnostics.Warn(
                "restore",
                $"Kept saved party squad state {_lastSquadState!.SessionName[..Math.Min(13, _lastSquadState.SessionName.Length)]}...; ignored weaker post-crash squad {nextState.SessionName[..Math.Min(13, nextState.SessionName.Length)]}... [{nextState.Mode.ToDisplayLabel()}].");
            return;
        }

        _lastSquadState = nextState;

        if (observedMemberConnectionState && connectedCount == 0)
            ClearMatchmakingPlayerPings($"observed unconnected squad via {source}");

        try
        {
            RejoinFixPaths.EnsureRootDirectory();
            File.WriteAllText(
                RejoinFixPaths.LastSquadStateFile,
                JsonSerializer.Serialize(nextState, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("squad", $"Failed to persist squad state: {ex.Message}");
        }

        if (changed)
        {
            RejoinFixDiagnostics.Info(
                "squad",
                $"Observed {nextState.Mode.ToDisplayLabel()} squad state for {nextState.SessionName[..Math.Min(13, nextState.SessionName.Length)]}… ({nextState.MemberCount} members) via {source}.");
            OnRejoinContextChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private void ArmCrashRestoreIfRestartSquadWriteDetected(
        RejoinSquadState incoming,
        string source,
        bool observedMemberConnectionState,
        int connectedCount)
    {
        if (_pendingCrashRestore ||
            _lastMatchSession is null ||
            !source.Equals("request-put", StringComparison.OrdinalIgnoreCase) ||
            !observedMemberConnectionState ||
            connectedCount != 0)
        {
            return;
        }

        if ((DateTime.UtcNow - _lastMatchSession.SavedAt).TotalMinutes > CRASH_RESTORE_TIMEOUT_MINUTES)
            return;

        if (_lastSquadState is not null &&
            string.Equals(_lastSquadState.SessionName, incoming.SessionName, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        SetPendingCrashRestore(_lastMatchSession);
        RejoinFixDiagnostics.Warn(
            "restore",
            $"Detected restart squad write {incoming.SessionName[..Math.Min(13, incoming.SessionName.Length)]}... [{incoming.Mode.ToDisplayLabel()}] with no live connection via {source}; armed crash restore for {_lastMatchSession.TemplateName}/{_lastMatchSession.SessionShort}.");
    }

    private void ArmCrashRestoreIfRestartSquadDetected(
        RejoinSquadState? existing,
        RejoinSquadState incoming,
        string source)
    {
        if (_pendingCrashRestore ||
            _lastMatchSession is null ||
            existing is null)
        {
            return;
        }

        if ((DateTime.UtcNow - _lastMatchSession.SavedAt).TotalMinutes > CRASH_RESTORE_TIMEOUT_MINUTES)
            return;

        bool newSquadAfterRecentMatch =
            existing.SavedAt <= _lastMatchSession.SavedAt &&
            incoming.SavedAt > _lastMatchSession.SavedAt &&
            !string.Equals(existing.SessionName, incoming.SessionName, StringComparison.OrdinalIgnoreCase);

        bool partyCollapsedToSolo =
            existing.Mode == RejoinSessionMode.Party &&
            incoming.Mode == RejoinSessionMode.Solo &&
            !string.Equals(existing.SessionName, incoming.SessionName, StringComparison.OrdinalIgnoreCase);

        if ((!newSquadAfterRecentMatch && !partyCollapsedToSolo) ||
            string.Equals(existing.SessionName, incoming.SessionName, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        SetPendingCrashRestore(_lastMatchSession);
        RejoinFixDiagnostics.Warn(
            "restore",
            $"Inferred missed MCC restart from squad transition {existing.SessionName[..Math.Min(13, existing.SessionName.Length)]}... [{existing.Mode.ToDisplayLabel()}] -> {incoming.SessionName[..Math.Min(13, incoming.SessionName.Length)]}... [{incoming.Mode.ToDisplayLabel()}] via {source}; armed crash restore for {_lastMatchSession.TemplateName}/{_lastMatchSession.SessionShort}.");
    }

    private static bool ShouldKeepExistingSquadStateDuringRestore(RejoinSquadState? existing, RejoinSquadState incoming)
    {
        if (existing is null)
            return false;

        if ((DateTime.UtcNow - existing.SavedAt).TotalMinutes > 30)
            return false;

        return existing.Mode == RejoinSessionMode.Party && incoming.Mode != RejoinSessionMode.Party;
    }

    private static string TryGetSessionNameFromUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var segs = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < segs.Length; i++)
            {
                if (segs[i].Equals("sessions", StringComparison.OrdinalIgnoreCase) && i + 1 < segs.Length)
                    return segs[i + 1];
            }
        }
        catch
        {
        }

        return "";
    }

    /// <summary>Clears the in-memory match session (called when user clicks CLEAR).</summary>
    public void ClearSavedMatchSession()
    {
        _lastMatchSession = null;
        _pendingCrashRestore = false;
        _pendingCrashRestoreStartedAt = DateTime.MinValue;  // Reset timeout clock
        _jitHandleDone = false;
        _jitPutDone = false;
        _blockMatchLeave = false;
        _cachedInjectedMatchBody = null;
        _cachedInjectedMatchEtag = "";
        ClearGhostSessionMode();  // Also clear ghost mode when clearing saved session
        RejoinFixDiagnostics.Info("capture", "Cleared saved match-session state.");
        OnCrashRestorePendingChanged?.Invoke(this, false);
    }

    /// <summary>
    /// Updates the live player identity from traffic. If MCC switches to a different
    /// account while the proxy stays running, clear stale rejoin state immediately so
    /// we never inject the old account into the new account's session flow.
    /// </summary>
    private void ObservePlayerXuid(string xuid)
    {
        if (string.IsNullOrWhiteSpace(xuid))
            return;

        bool changed = !string.Equals(_playerXuid, xuid, StringComparison.Ordinal);

        if (!string.IsNullOrEmpty(_playerXuid) &&
            !string.Equals(_playerXuid, xuid, StringComparison.Ordinal))
        {
            string previousXuid = _playerXuid;
            _lastSquadState = null;
            ClearSavedMatchSession();
            ClearGameServerRedirection();
            DeletePersistedRejoinState();
            RejoinFixDiagnostics.Warn("identity", $"Detected account switch ({previousXuid} -> {xuid}); cleared stale rejoin state.");

            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method = "CLEAR[Identity]",
                Url = $"xuid://{xuid}",
                Host = "sessiondirectory.xboxlive.com",
                Path = "/sessions?xuid=changed",
                RequestBody = $"Detected account switch from {previousXuid} to {xuid}",
                StatusCode = 0,
                ResponseBody = "Cleared saved rejoin session, ghost mode, cached injection bodies, and persisted rejoin artifacts.",
            });
        }

        _playerXuid = xuid;

        if (changed)
        {
            OnRejoinContextChanged?.Invoke(this, EventArgs.Empty);
            OnPlayerIdentityChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private static void DeletePersistedRejoinState()
    {
        static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                RejoinFixDiagnostics.Warn("identity", $"Failed to delete stale rejoin artifact '{Path.GetFileName(path)}': {ex.Message}");
            }
        }

        TryDelete(RejoinFixPaths.LastHandleFile);
        TryDelete(RejoinFixPaths.LastMatchSessionFile);
        TryDelete(RejoinFixPaths.LastSquadStateFile);
        TryDelete(RejoinFixPaths.LastGameServerFile);
    }

    /// <summary>Sets the in-memory match session from an externally loaded SavedHandleInfo.</summary>
    public void SetSavedMatchSession(SavedHandleInfo? info) => _lastMatchSession = info;

    /// <summary>Manually triggers the block-match-leave mechanism for testing/debugging.</summary>
    public void ForceBlockMatchLeave()
    {
        _blockMatchLeave = true;
        Debug.WriteLine("[MANUAL] Block match leave forced via UI button");
    }

    // ── Shared JIT PUT /members/me helper ─────────────────────────────────────
    // Used by both Phase B (OnBeforeRequestAsync, on CascadeMatchmaking GET) and
    // the session discovery injection (OnBeforeResponseAsync).
    //
    // KEY FIX: sends "active":true + a connection GUID so MPSD treats the member
    // as Active.  connectionRequiredForActiveMembers=true means you MUST provide a
    // connection UUID alongside active:true (400 without it).  The connection UUID
    // is normally an RTA WebSocket ID; we generate a placeholder so MPSD accepts
    // the PUT.  MPSD validates connections via 2-3 minute heartbeats, NOT instantly,
    // and timeouts only evaluate on the next read/write — giving MCC ample time to
    // establish its real WebSocket and replace the placeholder.
    private async Task<(int code, string body)> JitPutMembersMe(
        SavedHandleInfo match,
        Dictionary<string, string> freshHeaders,
        string logMethod)
    {
        int code = 0;
        string body = "";
        try
        {
            // GET for current ETag
            using var getReq = new HttpRequestMessage(HttpMethod.Get, match.SessionUrl);
            foreach (var (k, v) in freshHeaders)
                if (k.StartsWith("x-", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    getReq.Headers.TryAddWithoutValidation(k, v);
            using var getResp = await _refreshClient.SendAsync(getReq);
            string etag = getResp.Headers.ETag?.Tag ?? "";

            if (getResp.IsSuccessStatusCode && !string.IsNullOrEmpty(etag))
            {
                // PUT /members/me — active:true + connection GUID (required by
                // connectionRequiredForActiveMembers capability)
                string connGuid = GetCurrentGhostConnectionGuid();
                int restoreMemberCount = GetObservedSquadMemberCountForRestore();
                using var putReq = new HttpRequestMessage(HttpMethod.Put, match.SessionUrl);
                putReq.Content = new StringContent(
                    "{\"members\":{\"me\":{\"properties\":{\"system\":{\"active\":true,\"connection\":\"" + connGuid + "\"},\"custom\":{\"membercount\":" + restoreMemberCount + "}}}}}",
                    System.Text.Encoding.UTF8, "application/json");
                putReq.Headers.TryAddWithoutValidation("If-Match", etag);
                foreach (var (k, v) in freshHeaders)
                    if (k.StartsWith("x-", StringComparison.OrdinalIgnoreCase) ||
                        k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                        putReq.Headers.TryAddWithoutValidation(k, v);
                using var putResp = await _refreshClient.SendAsync(putReq);
                code = (int)putResp.StatusCode;
                body = await putResp.Content.ReadAsStringAsync();
            }
            else
            {
                code = (int)getResp.StatusCode;
                body = $"ETag fetch failed (status={(int)getResp.StatusCode}, etag={etag})";
            }
        }
        catch (Exception ex) { body = ex.Message; }

        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
        {
            Method       = logMethod,
            Url          = match.SessionUrl,
            Host         = "sessiondirectory.xboxlive.com",
            Path         = new Uri(match.SessionUrl).AbsolutePath,
            RequestBody  = $"PUT /members/me with active:true + connection GUID",
            StatusCode   = code,
            ResponseBody = body.Length > 300 ? body[..300] : body,
        });

        return (code, body);
    }

    // ── Response intercept ────────────────────────────────────────────────────
    //
    // Only watched domains reach here (BeforeTunnelConnectRequest sets DecryptSsl=false
    // for everything else, so non-watched hosts are opaque TCP tunnels and never fire
    // BeforeRequest/BeforeResponse).
    //
    // Only read GET response bodies — PUT/POST/DELETE responses are small
    // acknowledgments we don't need, and buffering them adds unnecessary latency.
    //
    // Within GETs: skip changeNumber= long-polls (held open for minutes by the server).
    private async Task OnBeforeResponseAsync(object sender, SessionEventArgs e)
    {
        // Check if crash restore timeout has expired
        CheckAndClearPendingCrashRestoreTimeout();

        if (e.UserData is not ProxyCaptureEntry entry) return;

        var resp = e.HttpClient.Response;
        entry.StatusCode = resp.StatusCode;

        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var h in resp.Headers)
            headers[h.Name] = h.Value;
        entry.ResponseHeaders = headers;

        // Refresh gamertag from X-Xbl-Debug header so account switches don't leave
        // us with a stale identity while the proxy remains running.
        if (headers.TryGetValue("X-Xbl-Debug", out var xblDebug) &&
            xblDebug.Contains("gamertag=", StringComparison.OrdinalIgnoreCase))
        {
            int gtIdx = xblDebug.IndexOf("gamertag=", StringComparison.OrdinalIgnoreCase) + 9;
            int gtEnd = xblDebug.IndexOf(';', gtIdx);
            string gamertag = gtEnd < 0 ? xblDebug[gtIdx..].Trim() : xblDebug[gtIdx..gtEnd].Trim();
            if (!string.Equals(_playerGamertag, gamertag, StringComparison.Ordinal))
            {
                _playerGamertag = gamertag;
                OnPlayerIdentityChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        var ct = resp.ContentType ?? "";
        bool isJson = ct.Contains("json", StringComparison.OrdinalIgnoreCase)
                   || ct.Contains("text", StringComparison.OrdinalIgnoreCase)
                   || ct.Contains("xml",  StringComparison.OrdinalIgnoreCase);

        // SmartMatch ticket creation is a known finite response. Some MCC/Xbox
        // responses omit a useful Content-Type, so this must not depend on isJson.
        bool smartMatchBodyRead = false;
        if (entry.Method == "POST" &&
            entry.Host.EndsWith("smartmatch.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            resp.HasBody)
        {
            var smartMatchBody = await e.GetResponseBodyAsString();
            entry.ResponseBody = smartMatchBody;
            e.SetResponseBodyString(smartMatchBody);
            smartMatchBodyRead = true;
            ObserveSmartMatchResponse(entry, smartMatchBody);
        }

        // ── Session discovery injection (PASSIVE MODE) ─────────────────────
        // When MCC restarts after a crash it queries GET /sessions?xuid=... .
        //
        // PASSIVE: If MPSD already returns the match session (player still active
        // on server), pass through the real response unmodified.  This lets MCC
        // see the original pre-crash member state — no fake connection GUIDs.
        //
        // FALLBACK: If MPSD returns empty (player was removed after heartbeat
        // timeout + inactiveRemovalTimeout:0), inject the match session so MCC
        // at least discovers it.  INJECT[Member] + FAKE[MatchPut] will handle
        // the rest downstream.
        //
        // MULTI-QUEUE: Check if all active queues are blocked. If so, don't accept match.
        bool injected = false;
        if (_lastMatchSession is not null &&
            entry.Method == "GET" &&
            entry.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            entry.Path.Contains("/sessions?", StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody)
        {
            var body = await e.GetResponseBodyAsString();
            var match = _lastMatchSession;


            // Always extract XUID from the query string (needed by INJECT[Member] downstream)
            int xuidIdx = entry.Url.IndexOf("xuid=", StringComparison.OrdinalIgnoreCase);
            if (xuidIdx >= 0)
            {
                xuidIdx += 5;
                int xuidEnd = entry.Url.IndexOf('&', xuidIdx);
                string xuid = xuidEnd < 0 ? entry.Url[xuidIdx..] : entry.Url[xuidIdx..xuidEnd];
                ObservePlayerXuid(xuid);
            }

            bool matchAlreadyInResults = body.Contains(match.SessionName, StringComparison.OrdinalIgnoreCase);
            bool isEmpty = body.Contains("\"results\":[]") || body.Contains("\"results\": []");

            if (_pendingCrashRestore && matchAlreadyInResults && body.Contains(_playerXuid))
            {
                // PASSIVE: Match session is already in MPSD discovery results AND player is in members.
                // Player is still an active member on the server — let MCC see the real response.
                // CRITICAL: Do NOT clear _pendingCrashRestore here! The flag must stay set
                // so that RequestParty and other subsequent rejoin requests can still be
                // redirected/injected. It will be cleared by ClearSavedMatchSession() or timeout.

                e.SetResponseBodyString(body);
                entry.ResponseBody = body;

                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "PASS[Discovery]",
                    Url          = entry.Url,
                    Host         = entry.Host,
                    Path         = entry.Path,
                    RequestBody  = "Match session already in MPSD discovery — player confirmed as member (passive mode)",
                    StatusCode   = 200,
                    ResponseBody = body.Length > 500 ? body[..500] + "…" : body,
                });
            }
            else if (_pendingCrashRestore && matchAlreadyInResults && !body.Contains(_playerXuid))
            {
                // CRITICAL FIX: Match in discovery but player was removed from session.
                // Need to inject player back into discovery so MCC can rejoin.
                // This fixes the case where player is removed during crash window but match is still active.
                Debug.WriteLine($"[CRASH-RESTORE] Player removed from match during crash window — injecting back into discovery");

                // Find and parse the match session from discovery results
                try
                {
                    using var doc = JsonDocument.Parse(body);
                    var root = doc.RootElement;

                    if (root.TryGetProperty("results", out var results) && results.ValueKind == JsonValueKind.Array)
                    {
                        // Find our match session in results
                        foreach (var result in results.EnumerateArray())
                        {
                            if (result.TryGetProperty("sessionRef", out var sessionRef) &&
                                sessionRef.TryGetProperty("name", out var nameEl) &&
                                nameEl.GetString() == match.SessionName)
                            {
                                // Found our match - inject xuid by finding the session name and prepending xuid before the object
                                var sessionNameQuoted = "\"name\":\"" + match.SessionName + "\"";
                                var xuidField = "\"xuid\":\"" + _playerXuid + "\",";

                                // Find the result object containing this session name and inject xuid after the opening brace
                                int idx = body.IndexOf(sessionNameQuoted);
                                if (idx > 0)
                                {
                                    // Find the opening brace of this result object (search backwards from sessionName)
                                    int braceIdx = body.LastIndexOf('{', idx);
                                    if (braceIdx > 0 && !body[braceIdx..idx].Contains("\"xuid\""))
                                    {
                                        // Insert xuid right after the opening brace
                                        var injectedBody = body.Insert(braceIdx + 1, xuidField);
                                        e.SetResponseBodyString(injectedBody);
                                        entry.ResponseBody = injectedBody;
                                        injected = true;
                                    }
                                }
                                else
                                {
                                    injected = true;  // Mark as attempted even if not found
                                }

                                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                                {
                                    Method       = "INJECT[Discovery]",
                                    Url          = entry.Url,
                                    Host         = entry.Host,
                                    Path         = entry.Path,
                                    RequestBody  = "Player removed from match during crash — re-injecting into discovery results",
                                    StatusCode   = 200,
                                    ResponseBody = entry.ResponseBody.Length > 500 ? entry.ResponseBody[..500] + "…" : entry.ResponseBody,
                                });
                                return;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[CRASH-RESTORE] Failed to inject player into discovery: {ex.Message}");
                }

                // Fallback: just pass through if injection fails
                e.SetResponseBodyString(body);
                entry.ResponseBody = body;
            }
            else if (!matchAlreadyInResults && (_pendingCrashRestore || isEmpty) &&
                     (DateTime.UtcNow - match.SavedAt).TotalMinutes <= 30)
            {
                // FALLBACK: Player was removed from session (heartbeat expired).
                // Inject the match session ref so MCC can still discover it.
                // INJECT[Member] will add the player to the GET response downstream.
                string xuid = _playerXuid;
                var injectedBody =
                    "{\"results\":[{\"xuid\":\"" + xuid +
                    "\",\"startTime\":\"" + match.SavedAt.ToString("O") +
                    "\",\"sessionRef\":{\"scid\":\"" + match.Scid +
                    "\",\"templateName\":\"" + match.TemplateName +
                    "\",\"name\":\"" + match.SessionName +
                    "\"}}]}";

                bool wasCrashRestore = _pendingCrashRestore;
                // CRITICAL FIX: Do NOT clear _pendingCrashRestore here!
                // The flag must stay set throughout the rejoin window so that subsequent requests
                // (like RequestParty, member PUT, etc.) can still be properly intercepted.
                // It will be cleared by ClearSavedMatchSession() when user clicks CLEAR, or
                // by a timeout if the rejoin doesn't complete in time.
                // _pendingCrashRestore = false;  // ← REMOVED

                e.SetResponseBodyString(injectedBody);
                entry.ResponseBody = injectedBody;
                injected = true;

                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "INJECT",
                    Url          = entry.Url,
                    Host         = entry.Host,
                    Path         = entry.Path,
                    RequestBody  = wasCrashRestore
                        ? "FALLBACK injection (player removed from session) — replaced MPSD results"
                        : "Session discovery was empty — injected saved matchmaking session",
                    StatusCode   = 200,
                    ResponseBody = injectedBody,
                });
            }
            else
            {
                // Stale match session or no injection needed — pass through unmodified
                e.SetResponseBodyString(body);
                entry.ResponseBody = body;
            }
        }

        // ── Member injection into CascadeMatchmaking session GET ────────────
        // After a crash, the player is removed from the session (inactiveRemovalTimeout: 0)
        // and joinRestriction:"local" + userAuthorizationStyle prevent us from re-adding
        // via PUT /members/me (403).  Instead, when MCC GETs the match session, we modify
        // the response to include our player in the members list so MCC sees itself as a
        // member and can properly connect to the game server.
        bool memberInjected = false;
        if (!injected && _lastMatchSession is not null &&
            !string.IsNullOrEmpty(_playerXuid) &&
            entry.Method == "GET" &&
            entry.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            entry.Path.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
            entry.Path.Contains(_lastMatchSession.SessionName, StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody)
        {
            var body = await e.GetResponseBodyAsString();
            PersistMatchmakingSessionDocument(entry.Url, body);
            // Only inject if our player is NOT in the members list
            if (!body.Contains(_playerXuid))
            {
                try
                {
                    var doc = System.Text.Json.JsonDocument.Parse(body);
                    var root = doc.RootElement;

                    // ── Parse membersInfo ────────────────────────────────────
                    int nextIdx = 0;   // next available member index (our new slot)
                    int count = 0;
                    int accepted = 0;
                    int active = 0;
                    if (root.TryGetProperty("membersInfo", out var mi))
                    {
                        if (mi.TryGetProperty("next", out var n)) nextIdx = n.GetInt32();
                        if (mi.TryGetProperty("count", out var c)) count = c.GetInt32();
                        if (mi.TryGetProperty("accepted", out var a)) accepted = a.GetInt32();
                        if (mi.TryGetProperty("active", out var ac)) active = ac.GetInt32();
                    }
                    if (nextIdx == 0) nextIdx = count + 1;

                    int newNext = nextIdx + 1; // new membersInfo.next after insertion

                    // ── Build member entry matching REAL MPSD format ─────────
                    // Real members have: next, joinTime, constants, properties,
                    // gamertag (root-level read-only), activeTitleId (root-level read-only).
                    // The linked-list "next" field on our member points to newNext (end sentinel).
                    // The PREVIOUS last member's "next" already == nextIdx, so it
                    // naturally chains into our new member without modification.
                    string connGuid = GetCurrentGhostConnectionGuid();
                    int restoreMemberCount = GetObservedSquadMemberCountForRestore();
                    string gt = !string.IsNullOrEmpty(_playerGamertag)
                        ? ",\"gamertag\":\"" + _playerGamertag + "\""
                        : "";
                    string memberJson =
                        "\"" + nextIdx + "\":{" +
                        "\"next\":" + newNext + "," +
                        "\"joinTime\":\"" + _lastMatchSession.SavedAt.ToString("O") + "\"," +
                        "\"constants\":{\"system\":{\"xuid\":\"" + _playerXuid + "\",\"index\":" + nextIdx + "}}," +
                        "\"properties\":{\"system\":{" +
                        "\"active\":true," +
                        "\"connection\":\"" + connGuid + "\"" +
                        "},\"custom\":{\"membercount\":" + restoreMemberCount + "}}" +
                        gt +
                        ",\"activeTitleId\":\"1144039928\"}";

                    // ── Splice into the members object ──────────────────────
                    int membersIdx = body.IndexOf("\"members\":{", StringComparison.Ordinal);
                    if (membersIdx >= 0)
                    {
                        // Find the matching closing brace for the members object
                        int braceStart = body.IndexOf('{', membersIdx + 10);
                        int depth = 1;
                        int pos = braceStart + 1;
                        while (pos < body.Length && depth > 0)
                        {
                            if (body[pos] == '{') depth++;
                            else if (body[pos] == '}') depth--;
                            if (depth > 0) pos++;
                        }
                        // pos = closing } of the members object.  Insert before it.
                        string modified = body[..pos] + "," + memberJson + body[pos..];

                        // ── Update membersInfo ONLY (surgical, not global regex) ──
                        // The old code used global regex which also changed member
                        // "next" fields, breaking the linked-list traversal.
                        // Now we isolate the membersInfo {...} block and only
                        // replace values within it.
                        int miStart = modified.IndexOf("\"membersInfo\":", StringComparison.Ordinal);
                        if (miStart >= 0)
                        {
                            int miBrace = modified.IndexOf('{', miStart);
                            int miEnd = modified.IndexOf('}', miBrace) + 1; // membersInfo is flat
                            string miSection = modified[miStart..miEnd];
                            string updatedMi = miSection
                                .Replace("\"next\":" + nextIdx,     "\"next\":" + newNext)
                                .Replace("\"count\":" + count,       "\"count\":" + (count + 1))
                                .Replace("\"accepted\":" + accepted, "\"accepted\":" + (accepted + 1))
                                .Replace("\"active\":" + active,     "\"active\":" + (active + 1));
                            modified = modified[..miStart] + updatedMi + modified[miEnd..];
                        }

                        e.SetResponseBodyString(modified);
                        entry.ResponseBody = modified;
                        memberInjected = true;

                        // Cache the injected body + ETag for faking subsequent PUT responses.
                        // When MCC PUTs to this session, MPSD returns 403 (not a real member).
                        // We intercept that 403 and return this cached body as a fake 200.
                        _cachedInjectedMatchBody = modified;
                        string respEtag = "";
                        if (entry.ResponseHeaders.TryGetValue("ETag", out var et))
                            respEtag = et;
                        _cachedInjectedMatchEtag = respEtag;

                        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                        {
                            Method       = "INJECT[Member]",
                            Url          = entry.Url,
                            Host         = entry.Host,
                            Path         = entry.Path,
                            RequestBody  = $"Added xuid={_playerXuid} as member {nextIdx} (conn={connGuid[..8]}…) gt={_playerGamertag}",
                            StatusCode   = 200,
                            ResponseBody = memberJson.Length > 400 ? memberJson[..400] : memberJson,
                        });
                    }
                }
                catch (Exception ex)
                {
                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method = "INJECT[Member]", Url = entry.Url, Host = "diag",
                        Path = "", StatusCode = 0, ResponseBody = "Parse error: " + ex.Message,
                    });
                    // Fall through — don't modify the response on error
                }
            }

            if (!memberInjected)
            {
                // Player already in members or injection failed — pass through
                e.SetResponseBodyString(body);
                entry.ResponseBody = body;
            }
            injected = memberInjected; // prevent double-read in normal capture
        }

        // ── Fake PUT response for CascadeMatchmaking during crash restore ──
        // When MCC finds itself in the injected member list, it tries to PUT
        // to the match session to update its member state. MPSD returns 403
        // (joinRestriction:"local" + not a real member). We intercept this
        // and return the cached INJECT[Member] body as a fake 200 so MCC
        // believes the PUT succeeded and proceeds to connect to the game server.
        if (!injected && !memberInjected &&
            _cachedInjectedMatchBody is not null &&
            _lastMatchSession is not null &&
            entry.Method == "PUT" &&
            entry.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            entry.Path.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
            entry.Path.Contains(_lastMatchSession.SessionName, StringComparison.OrdinalIgnoreCase) &&
            (resp.StatusCode == 403 || resp.StatusCode == 404 || resp.StatusCode == 409))
        {
            int originalStatus = resp.StatusCode;

            // Replace the error with the cached session body
            e.SetResponseBodyString(_cachedInjectedMatchBody);
            resp.StatusCode = 200;
            entry.StatusCode = 200;
            entry.ResponseBody = "[FAKE 200] Returned cached INJECT[Member] body (" + _cachedInjectedMatchBody.Length + " bytes)";

            // Set proper response headers so MCC treats this as a real session document
            if (!string.IsNullOrEmpty(_cachedInjectedMatchEtag))
            {
                resp.Headers.RemoveHeader("ETag");
                resp.Headers.AddHeader("ETag", _cachedInjectedMatchEtag);
            }
            resp.Headers.RemoveHeader("Content-Type");
            resp.Headers.AddHeader("Content-Type", "application/json");

            injected = true;
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "FAKE[MatchPut]",
                Url          = entry.Url,
                Host         = entry.Host,
                Path         = entry.Path,
                RequestBody  = $"Intercepted {originalStatus}→200 on PUT to CascadeMatchmaking",
                StatusCode   = 200,
                ResponseBody = $"Returned cached body ({_cachedInjectedMatchBody.Length} bytes) etag={_cachedInjectedMatchEtag}",
            });
        }

        // ── Game Server Observation (PlayFab response interception) ─────────────
        // Learn the active dedicated server from PlayFab responses. RequestParty is
        // the normal path, but MCC builds have used nearby PlayFab endpoints too, so
        // key off the server fields rather than only one URL shape.
        if (!injected && !memberInjected &&
            entry.Host.Contains("playfabapi.com", StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody)
        {
            var body = await e.GetResponseBodyAsString();

            try
            {
                if (TryParsePlayFabGameServer(body, out var serverInfo))
                {
                    if (!_pendingCrashRestore && !string.IsNullOrWhiteSpace(serverInfo.IPv4Address))
                    {
                        bool changedServer = !string.Equals(
                            _currentObservedGameServerInfo?.IPv4Address,
                            serverInfo.IPv4Address,
                            StringComparison.OrdinalIgnoreCase);

                        SetCurrentGameServer(serverInfo);
                        _gameServerRedirectionActive = false;  // Not active yet; only activate on crash restore

                        entry.ResponseBody = $"{(changedServer ? "UPDATED" : "REFRESHED")} game server: {serverInfo.ServerShort}";

                        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                        {
                            Method = "CACHE[GameServer]",
                            Url = entry.Url,
                            Host = entry.Host,
                            Path = entry.Path,
                            RequestBody = "Current game server info observed from latest PlayFab response",
                            StatusCode = 200,
                            ResponseBody = $"Region: {serverInfo.Region} | Server: {serverInfo.IPv4Address}:{serverInfo.Ports.FirstOrDefault()?.Num ?? 0} | ServerId: {ShortServerId(serverInfo.ServerId)}",
                        });
                        return;  // Skip normal body capture, we already have the data
                    }
                    // On restart: if we have cached server info, redirect to it
                    else if (_pendingCrashRestore && !string.IsNullOrWhiteSpace(serverInfo.IPv4Address))
                    {
                        bool changedServer = !string.Equals(
                            _currentObservedGameServerInfo?.IPv4Address,
                            serverInfo.IPv4Address,
                            StringComparison.OrdinalIgnoreCase);

                        SetCurrentGameServer(serverInfo);
                        _gameServerRedirectionActive = false;
                        // CRITICAL FIX: Do NOT clear _pendingCrashRestore here!
                        // MCC might make multiple RequestParty calls during the rejoin sequence,
                        // and ALL of them need to be redirected to the cached server.
                        // The flag will be cleared by SetPendingCrashRestore when MCC confirms
                        // successful rejoin or when the user clears the rejoin state.
                        // _pendingCrashRestore = false;  // ← REMOVED - was causing subsequent calls to bypass redirect

                        entry.ResponseBody = body;
                        e.SetResponseBodyString(body);

                        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                        {
                            Method = "LIVE[GameServer]",
                            Url = entry.Url,
                            Host = entry.Host,
                            Path = entry.Path,
                            RequestBody = "Crash restore active; accepted current PlayFab server assignment",
                            StatusCode = 200,
                            ResponseBody = $"{(changedServer ? "UPDATED" : "REFRESHED")} live server: {serverInfo.ServerShort}",
                        });
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't break the response
                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method = "ERROR[GameServer]",
                    Url = entry.Url,
                    Host = entry.Host,
                    Path = entry.Path,
                    RequestBody = "Failed to parse PlayFab game server response",
                    StatusCode = 500,
                    ResponseBody = ex.Message,
                });
            }

            // Still capture the body for normal logging
            entry.ResponseBody = body;
            e.SetResponseBodyString(body);
        }

        // ── Normal body capture ────────────────────────────────────────────
        if (entry.Method == "GET" &&
            entry.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            entry.Path.Contains("/CascadeSquadSession/sessions/", StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody &&
            ShouldReadBody(entry.Url))
        {
            var squadBody = await e.GetResponseBodyAsString();
            ObserveNetworkRegionLatencies(squadBody, "squad-response");
            ObserveSquadSessionDocument(entry.Url, squadBody, "response-get");
            e.SetResponseBodyString(squadBody);
            if (string.IsNullOrEmpty(entry.ResponseBody))
                entry.ResponseBody = squadBody;
        }

        if (!smartMatchBodyRead && !injected && !memberInjected && isJson && resp.HasBody && (entry.Method == "GET" || entry.Method == "POST") && ShouldReadBody(entry.Url))
        {
            var body = await e.GetResponseBodyAsString();
            ObserveNetworkRegionLatencies(body, "response-body");
            PersistMatchmakingSessionDocument(entry.Url, body);
            entry.ResponseBody = body;
            // Must re-set or the connection will fail (body stream is consumed)
            e.SetResponseBodyString(body);
        }

        OnRequestCaptured?.Invoke(this, entry);
    }

    private static void PersistMatchmakingSessionDocument(string url, string body)
    {
        if (string.IsNullOrWhiteSpace(body) ||
            !url.Contains("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase))
            return;

        string? path = null;
        string label = "";
        if (url.Contains("/CascadeMatchTicketSession/sessions/", StringComparison.OrdinalIgnoreCase))
        {
            path = RejoinFixPaths.LastMatchTicketSessionDocumentFile;
            label = "match ticket session";
        }
        else if (url.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase))
        {
            path = RejoinFixPaths.LastMatchmakingSessionDocumentFile;
            label = "matchmaking session";
        }

        if (path is null)
            return;

        try
        {
            // Validate before replacing the last known-good capture, then preserve
            // the exact service document rather than reshaping it.
            using var _ = JsonDocument.Parse(body);
            RejoinFixPaths.EnsureRootDirectory();
            File.WriteAllText(path, body);
            RejoinFixDiagnostics.Info("capture", $"Saved latest {label} document for passive analysis.");
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("capture", $"Failed to persist {label} document: {ex.Message}");
        }
    }

    /// <summary>
    /// Returns true only for URL patterns where we know the response is a finite,
    /// well-formed JSON document worth capturing.  Everything else is let through
    /// without body interception to avoid blocking game-critical connections.
    /// </summary>
    private static bool ShouldReadBody(string url)
    {

        // Never read bodies for services known to use streaming, long-polling,
        // or certificate-pinned connections — these would block the game thread.
        foreach (var b in _bypassHosts)
            if (url.Contains(b, StringComparison.OrdinalIgnoreCase))
                return false;

        // Halo Waypoint API — service records, lobby data (always finite JSON)
        if (url.Contains("halowaypoint.com", StringComparison.OrdinalIgnoreCase))
            return true;

        // PlayFab — matchmaking / telemetry (finite JSON)
        if (url.Contains("playfabapi.com", StringComparison.OrdinalIgnoreCase))
            return true;

        if (url.Contains("smartmatch.xboxlive.com", StringComparison.OrdinalIgnoreCase))
            return true;

        // Xbox Live activity/rejoin handles — POST response contains the new handle ID;
        // GET response contains the handle document with sessionRef.
        if (url.Contains("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            url.Contains("/handles", StringComparison.OrdinalIgnoreCase))
            return true;

        // Xbox Live session DOCUMENT (specific session by ID).
        // Xbox Live long-polls add ?changeNumber=N to hold the connection open for
        // minutes waiting for state changes — reading those blocks the game thread.
        // All OTHER query params (include=, version=, …) return finite JSON immediately.
        if (url.Contains("/serviceconfigs/", StringComparison.OrdinalIgnoreCase) &&
            url.Contains("/sessions/",       StringComparison.OrdinalIgnoreCase))
        {
            // changeNumber = long-poll subscription → skip.  Everything else is fine.
            return !url.Contains("changeNumber=", StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    // ── Domain filter ─────────────────────────────────────────────────────────
    private static bool IsDomainWatched(string host)
    {
        foreach (var d in _watchedDomains)
            if (host.EndsWith(d, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private void ObserveSmartMatchResponse(ProxyCaptureEntry entry, string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("waitTime", out var waitElement) ||
                !waitElement.TryGetInt32(out int waitSeconds) || waitSeconds < 0)
                return;

            string ticketId = doc.RootElement.TryGetProperty("ticketId", out var ticketElement)
                ? ticketElement.GetString() ?? ""
                : "";
            int giveUpSeconds = 120;
            try
            {
                using var requestDoc = JsonDocument.Parse(entry.RequestBody);
                if (requestDoc.RootElement.TryGetProperty("giveUpDuration", out var giveUpElement) &&
                    giveUpElement.TryGetInt32(out int parsedGiveUp) && parsedGiveUp > 0)
                    giveUpSeconds = parsedGiveUp;
            }
            catch { }

            string hopperName = "";
            var pathSegments = entry.Path.Split('/', StringSplitOptions.RemoveEmptyEntries);
            int hopperIndex = Array.FindIndex(pathSegments, x => x.Equals("hoppers", StringComparison.OrdinalIgnoreCase));
            if (hopperIndex >= 0 && hopperIndex + 1 < pathSegments.Length)
                hopperName = Uri.UnescapeDataString(pathSegments[hopperIndex + 1]);

            var estimate = new SmartMatchWaitEstimate(ticketId, hopperName, waitSeconds, giveUpSeconds, DateTimeOffset.UtcNow);
            OnSmartMatchWaitEstimateChanged?.Invoke(this, estimate);

            RejoinFixPaths.EnsureRootDirectory();
            File.WriteAllText(RejoinFixPaths.LastSmartMatchTicketFile, JsonSerializer.Serialize(new
            {
                capturedAt = estimate.CapturedAtUtc,
                stage = "response",
                entry.Method,
                entry.Url,
                entry.Host,
                entry.Path,
                entry.StatusCode,
                entry.RequestBody,
                responseBody = body
            }, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("smartmatch", $"Failed to parse wait estimate: {ex.Message}");
        }
    }

    public async Task<HopperPopulationResult> GetHopperStatisticsAsync(string hopperName, CancellationToken cancellationToken = default)
    {
        Dictionary<string, string> headers;
        string scid;
        lock (_smartMatchAuthLock)
        {
            headers = new Dictionary<string, string>(_smartMatchRequestHeaders, StringComparer.OrdinalIgnoreCase);
            scid = _smartMatchServiceConfigId;
        }

        if (string.IsNullOrWhiteSpace(scid) || headers.Count == 0)
            return new HopperPopulationResult(hopperName, null, null, "Start a matchmaking search to authorize population data.");

        string url = $"https://smartmatch.xboxlive.com/serviceconfigs/{Uri.EscapeDataString(scid)}/hoppers/{Uri.EscapeDataString(hopperName)}/stats";
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            foreach (var (name, value) in headers)
            {
                if (name.Equals("Host", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Content-Length", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Content-Type", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Signature", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("If-Match", StringComparison.OrdinalIgnoreCase))
                    continue;
                request.Headers.TryAddWithoutValidation(name, value);
            }
            request.Headers.Remove("X-Xbl-Contract-Version");
            request.Headers.TryAddWithoutValidation("X-Xbl-Contract-Version", "103");

            using var response = await _refreshClient.SendAsync(request, cancellationToken);
            string body = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                string detail = body.Length > 180 ? body[..180] + "…" : body;
                return new HopperPopulationResult(
                    hopperName,
                    null,
                    null,
                    string.IsNullOrWhiteSpace(detail)
                        ? $"HTTP {(int)response.StatusCode}"
                        : $"HTTP {(int)response.StatusCode}: {detail}");
            }

            using var doc = JsonDocument.Parse(body);
            int? waitTime = doc.RootElement.TryGetProperty("waitTime", out var wait) && wait.TryGetInt32(out int parsedWait)
                ? parsedWait : null;
            int? population = doc.RootElement.TryGetProperty("population", out var pop) && pop.TryGetInt32(out int parsedPopulation)
                ? parsedPopulation : null;
            return new HopperPopulationResult(hopperName, waitTime, population, "");
        }
        catch (Exception ex)
        {
            return new HopperPopulationResult(hopperName, null, null, ex.Message);
        }
    }

    private static bool ShouldDecryptTunnel(string host, int processId)
    {
        if (!IsDomainWatched(host))
            return false;

        if (IsKnownXboxShellProcess(processId))
        {
            RejoinFixDiagnostics.Info("proxy", $"Bypassed TLS interception for Xbox shell process on {host}.");
            return false;
        }

        return true;
    }

    private static bool IsKnownXboxShellProcess(int processId)
    {
        if (processId <= 0)
            return false;

        try
        {
            using var process = Process.GetProcessById(processId);
            string processName = process.ProcessName;
            foreach (var name in _xboxShellProcessNames)
            {
                if (processName.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch
        {
            // If process lookup fails, keep the old host-based behavior so rejoin capture still works.
        }

        return false;
    }

    // ── WinINet (no admin) ────────────────────────────────────────────────────
    private void SetWinINetProxy(string proxyAddress)
    {
        const string key = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";
        using var reg = Registry.CurrentUser.OpenSubKey(key, writable: true);
        if (reg is null) return;

        _savedProxyEnable   = (int)(reg.GetValue("ProxyEnable")   ?? 0);
        _savedProxyServer   = (string)(reg.GetValue("ProxyServer") ?? "");
        _savedProxyOverride = (string)(reg.GetValue("ProxyOverride") ?? "");

        reg.SetValue("ProxyEnable",   1,                                  RegistryValueKind.DWord);
        reg.SetValue("ProxyServer",   proxyAddress,                       RegistryValueKind.String);
        reg.SetValue("ProxyOverride", string.Join(';', _systemProxyBypassHosts), RegistryValueKind.String);

        InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
        InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH,          IntPtr.Zero, 0);
    }

    private void RestoreWinINetProxy()
    {
        const string key = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";
        using var reg = Registry.CurrentUser.OpenSubKey(key, writable: true);
        if (reg is null) return;

        reg.SetValue("ProxyEnable",   _savedProxyEnable,   RegistryValueKind.DWord);
        reg.SetValue("ProxyServer",   _savedProxyServer,   RegistryValueKind.String);
        reg.SetValue("ProxyOverride", _savedProxyOverride, RegistryValueKind.String);

        InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
        InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH,          IntPtr.Zero, 0);
    }

    // ── WinHTTP (admin / UAC required — needed for Halo MCC) ─────────────────
    private async Task TrySetWinHttpProxyAsync()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName        = "netsh",
                Arguments       = "winhttp import proxy source=ie",
                Verb            = "runas",   // triggers UAC elevation prompt
                UseShellExecute = true,
                CreateNoWindow  = true,
                WindowStyle     = ProcessWindowStyle.Hidden,
            };
            var p = Process.Start(psi);
            await Task.Run(() => p?.WaitForExit(8000));
        }
        catch
        {
            // User cancelled UAC or insufficient rights — surface the manual fallback
            RejoinFixDiagnostics.Warn("proxy", "WinHTTP proxy update needs elevation; manual netsh command may be required.");
            WinHttpManualSetRequired?.Invoke(this,
                "netsh winhttp import proxy source=ie");
        }
    }

    private static void TryResetWinHttpProxy()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName        = "netsh",
                Arguments       = "winhttp reset proxy",
                Verb            = "runas",
                UseShellExecute = true,
                CreateNoWindow  = true,
                WindowStyle     = ProcessWindowStyle.Hidden,
            });
        }
        catch { /* best-effort */ }
    }

    // ── Ghost Session Handling ────────────────────────────────────────────────
    // When enabled, fake MPSD responses to make MCC think it's still in the
    // session while we sync with real MPSD in the background.

    /// <summary>
    /// Check if a request is for the ghosted match session.
    /// CascadeSquadSession stays service-authoritative during restart.
    /// </summary>
    private bool IsRequestForGhostSession(Titanium.Web.Proxy.Http.Request req)
    {
        if (_ghostSession is null) return false;

        string url = req.RequestUri.AbsolutePath.ToLowerInvariant();
        string sessionName = _ghostSession.SessionName.ToLowerInvariant();
        // Match CascadeMatchmaking with same session name
        bool isMatchSession = url.Contains("/cascadematchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
                              url.Contains(sessionName, StringComparison.OrdinalIgnoreCase);

        return isMatchSession;
    }

    /// <summary>
    /// Handle ghost session requests: return fake responses that make MCC think it's
    /// still in the saved match session, while real sync happens in the background.
    /// </summary>
    private bool HandleGhostSessionRequest(Titanium.Web.Proxy.Http.Request req, SessionEventArgs e)
    {
        if (_ghostSession is null) return false;

        string method = (req.Method ?? "").ToUpperInvariant();
        string url = req.RequestUri?.AbsolutePath ?? "";
        string sessionName = _ghostSession.SessionName.ToLowerInvariant();

        // Determine if this is the saved match session
        bool isMatchSession = url.Contains("/cascadematchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
                              url.Contains(sessionName, StringComparison.OrdinalIgnoreCase);
        bool isSquadSession = false;

        // ── Match Session Interception ────────────────────────────────────
        // Block destructive writes only; real match reads must pass through.
        if (isMatchSession)
        {
            // GET match session — return fake "you're active"
            // DELETE /members/me from match session — block leave
            if (method == "DELETE" && url.Contains("/members/me", StringComparison.OrdinalIgnoreCase))
            {
                e.Ok("{}");
                LogGhostRequest("GHOST[Match-Leave]", req.Url, "204", "Blocked match leave");
                return true;
            }

            // PUT /members/me in match session — accept locally (sync in background)
            if (method == "PUT" && url.Contains("/members/me", StringComparison.OrdinalIgnoreCase))
            {
                e.Ok("{}");
                LogGhostRequest("GHOST[Match-PUT]", req.Url, "204", "Blocked, sync in background");
                return true;
            }
        }

        // ── Squad Session Interception ────────────────────────────────────
        // CORRECTED: Fake GETs but LET mutations through to MPSD!
        // Critical insight: MCC needs to actually UPDATE squad state for rejoin to work
        // Blocking mutations was TOO aggressive and prevented rejoin prep
        if (isSquadSession)
        {
            // GET squad session — return fake "squad is valid"
            if (method == "GET" && !url.Contains("/members/"))
            {
                string currentGuid = GetCurrentGhostConnectionGuid();
                string fakeSquadBody = $$"""
{
  "contractVersion": 1,
  "state": "active",
  "members": {
    "me": {
      "gamertag": "Player",
      "xuid": "{{_playerXuid}}",
      "active": true,
      "properties": {
        "system": {
          "active": true,
          "connection": "{{currentGuid}}"
        }
      }
    }
  }
}
""";
                e.Ok(fakeSquadBody);
                LogGhostRequest("GHOST[Squad-Get]", req.Url, "200", "Fake squad document");
                return true;
            }

            // GET /members/me in squad session
            if (method == "GET" && url.Contains("/members/me", StringComparison.OrdinalIgnoreCase))
            {
                string fakeMemberBody = GenerateFakeMemberDocument();
                e.Ok(fakeMemberBody);
                LogGhostRequest("GHOST[Squad-Member]", req.Url, "200", "Fake member (active=true)");
                return true;
            }

            // CRITICAL FIX: Let mutations through to real MPSD!
            // MCC needs to actually modify squad state (PUT /members/me, etc)
            // for rejoin to work. Don't block or fake - pass through.
            return false;
        }

        // ── Match Session (CascadeMatchmaking) Special Handling ──────────────
        // For match session, block mutations (leave, etc) but let GETs through
        if (isMatchSession)
        {
            // DELETE /members/me from match session — block leave
            if (method == "DELETE" && url.Contains("/members/me", StringComparison.OrdinalIgnoreCase))
            {
                e.Ok("{}");
                LogGhostRequest("GHOST[Match-Leave]", req.Url, "204", "Blocked match leave");
                return true;
            }

            // Let all other match requests through to MPSD
            // (including GETs and PUTs)
            return false;
        }

        // Not a ghost request, let it through
        return false;
    }

    /// <summary>Generate a fake session document that makes MCC think it's in an active session.</summary>
    private string GenerateFakeSessionDocument()
    {
        if (_ghostSession is null) return "{}";
        string currentGuid = GetCurrentGhostConnectionGuid();

        // Real session doc structure (minimal valid response)
        return $$"""
{
  "contractVersion": 1,
  "sessionRef": {
    "scid": "{{_ghostSession.Scid}}",
    "templateName": "{{_ghostSession.TemplateName}}",
    "name": "{{_ghostSession.SessionName}}"
  },
  "state": "active",
  "createdAt": "2026-03-08T00:00:00Z",
  "members": {
    "me": {
      "gamertag": "Player",
      "xuid": "{{_playerXuid}}",
      "roleTypes": [],
      "properties": {
        "system": {
          "active": true,
          "connection": "{{currentGuid}}",
          "joinTime": "2026-03-08T00:00:00Z"
        },
        "custom": {}
      }
    }
  },
  "constants": {
    "system": {
      "version": 1,
      "maxMembers": 12
    }
  }
}
""";
    }

    /// <summary>Generate a fake member document showing player is active.</summary>
    private string GenerateFakeMemberDocument()
    {
        string currentGuid = GetCurrentGhostConnectionGuid();
        return $$"""
{
  "gamertag": "Player",
  "xuid": "{{_playerXuid}}",
  "roleTypes": [],
  "properties": {
    "system": {
      "active": true,
      "connection": "{{currentGuid}}",
      "joinTime": "2026-03-08T00:00:00Z"
    },
    "custom": {}
  }
}
""";
    }

    private void LogGhostRequest(string method, string url, string statusCode, string notes)
    {
        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
        {
            Method       = method,
            Url          = url,
            Host         = "sessiondirectory.xboxlive.com",
            Path         = new Uri(url).AbsolutePath,
            RequestBody  = "GHOST MODE: " + notes,
            StatusCode   = int.Parse(statusCode),
            ResponseBody = "[Faked response]",
        });
    }

    private string GetCurrentGhostConnectionGuid()
    {
        if (_ghostSession is not null && !string.IsNullOrWhiteSpace(_ghostSession.ConnectionGuid))
            return _ghostSession.ConnectionGuid;

        if (_lastMatchSession is not null && !string.IsNullOrWhiteSpace(_lastMatchSession.ConnectionGuid))
            return _lastMatchSession.ConnectionGuid;

        return PlaceholderConnectionGuid;
    }

    private int GetObservedSquadMemberCountForRestore()
    {
        if (_ghostSession?.ObservedSquadMemberCount > 0)
            return _ghostSession.ObservedSquadMemberCount;

        if (_lastMatchSession?.ObservedSquadMemberCount > 0)
            return _lastMatchSession.ObservedSquadMemberCount;

        if (_lastSquadState?.MemberCount > 0)
            return _lastSquadState.MemberCount;

        return 1;
    }

    private void TryUpgradeGhostSessionConnectionGuid(string requestBody)
    {
        if (_ghostSession is null || string.IsNullOrWhiteSpace(requestBody))
            return;

        string newGuid = ExtractConnectionGuid(requestBody);
        if (string.IsNullOrWhiteSpace(newGuid))
            return;

        bool changed = !string.Equals(_ghostSession.ConnectionGuid, newGuid, StringComparison.OrdinalIgnoreCase);
        _ghostSession.ConnectionGuid = newGuid;

        if (_lastMatchSession is not null &&
            string.Equals(_lastMatchSession.SessionName, _ghostSession.SessionName, StringComparison.OrdinalIgnoreCase))
        {
            _lastMatchSession.ConnectionGuid = newGuid;
            PersistSavedMatchSessionSnapshot(_lastMatchSession);
        }

        _ghostSessionGuidUpgraded =
            string.IsNullOrWhiteSpace(_ghostSessionOriginalConnectionGuid) ||
            !string.Equals(_ghostSessionOriginalConnectionGuid, newGuid, StringComparison.OrdinalIgnoreCase);

        if (changed)
        {
            RejoinFixDiagnostics.Info("guid", $"Captured replacement connection GUID: {newGuid}");
            Debug.WriteLine($"[GUID] Captured NEW connection GUID: {newGuid}");
        }
    }

    private static string ExtractConnectionGuid(string requestBody)
    {
        if (string.IsNullOrWhiteSpace(requestBody))
            return "";

        try
        {
            using var doc = JsonDocument.Parse(requestBody);
            if (doc.RootElement.TryGetProperty("members", out var members) &&
                members.TryGetProperty("me", out var me) &&
                me.TryGetProperty("properties", out var properties) &&
                properties.TryGetProperty("system", out var system) &&
                system.TryGetProperty("connection", out var connection))
            {
                return connection.GetString() ?? "";
            }
        }
        catch
        {
            // Fall back to regex for partially-formed bodies.
        }

        var match = System.Text.RegularExpressions.Regex.Match(
            requestBody,
            @"""connection"":""([^""]+)""",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        return match.Success ? match.Groups[1].Value : "";
    }

    private static void PersistSavedMatchSessionSnapshot(SavedHandleInfo info)
    {
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox", "RejoinFix");
            Directory.CreateDirectory(dir);
            File.WriteAllText(
                Path.Combine(dir, "last-match-session.json"),
                JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("guid", $"Failed to persist upgraded connection GUID: {ex.Message}");
            Debug.WriteLine($"[GUID] Failed to persist upgraded connection GUID: {ex.Message}");
        }
    }

    /// <summary>
    /// Background task: automatically sync the ghost session with real MPSD.
    /// Once sync succeeds, disable ghost mode.
    /// </summary>
    private async Task AutoSyncGhostSessionAsync()
    {
        if (_ghostSession is null) return;

        try
        {
            // Wait 2 seconds before starting sync — let MCC settle after restart
            await Task.Delay(2000);

            // Give the restarted MCC a short window to publish its replacement connection GUID
            // before we touch the live MPSD session with stale pre-crash state.
            for (int i = 0; i < 40 && _ghostSession is not null; i++)
            {
                if (_ghostSessionGuidUpgraded || string.IsNullOrWhiteSpace(_ghostSessionOriginalConnectionGuid))
                    break;

                await Task.Delay(250);
            }

            // Attempt GET to check if session is alive
            using var getReq = new HttpRequestMessage(HttpMethod.Get, _ghostSession.SessionUrl);
            foreach (var (k, v) in _ghostSession.RequestHeaders)
                if (k.StartsWith("x-", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    getReq.Headers.TryAddWithoutValidation(k, v);

            using var getResp = await new HttpClient(new HttpClientHandler { UseProxy = false })
                .SendAsync(getReq);

            if (!getResp.IsSuccessStatusCode) return;  // Session dead, keep ghost mode

            string etag = getResp.Headers.ETag?.Tag ?? "";
            if (string.IsNullOrEmpty(etag)) return;  // No ETag, can't proceed

            // Session is alive! Now PUT /members/me to re-add player as active
            // CRITICAL: Include the connection GUID from when the player originally joined!
            // The game server validates rejoin attempts using this GUID. If we omit or change it,
            // the game server will reject the rejoin with "connection interrupted" error.
            string currentGuid = GetCurrentGhostConnectionGuid();
            string connectionField = string.IsNullOrEmpty(currentGuid)
                ? ""
                : $",\"connection\":\"{currentGuid}\"";
            int restoreMemberCount = GetObservedSquadMemberCountForRestore();

            var putBody = $$"""
{
  "members": {
    "me": {
      "properties": {
        "system": {
          "active": true{{connectionField}}
        },
        "custom": {
          "membercount": {{restoreMemberCount}}
        }
      }
    }
  }
}
""";

            using var putReq = new HttpRequestMessage(HttpMethod.Put, _ghostSession.SessionUrl);
            putReq.Content = new StringContent(putBody, System.Text.Encoding.UTF8, "application/json");
            putReq.Headers.TryAddWithoutValidation("If-Match", etag);
            foreach (var (k, v) in _ghostSession.RequestHeaders)
                if (k.StartsWith("x-", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    putReq.Headers.TryAddWithoutValidation(k, v);

            using var putResp = await new HttpClient(new HttpClientHandler { UseProxy = false })
                .SendAsync(putReq);

            if (putResp.IsSuccessStatusCode || putResp.StatusCode == System.Net.HttpStatusCode.NoContent)
            {
                // Sync succeeded! Keep ghost mode ON permanently - MCC needs it during rejoin
                _ghostSessionSyncSuccess = true;
                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "GHOST[AutoSync-Permanent]",
                    Url          = _ghostSession.SessionUrl,
                    Host         = "sessiondirectory.xboxlive.com",
                    Path         = new Uri(_ghostSession.SessionUrl).AbsolutePath,
                    RequestBody  = "Background automatic sync: GET + PUT succeeded",
                    StatusCode   = (int)putResp.StatusCode,
                    ResponseBody = "PERMANENT: Ghost mode stays ON to support rejoin flow",
                });

                // DO NOT disable ghost mode - MCC needs it to be active during the rejoin window
                // Disabling it too early breaks the rejoin process
                // _ghostSessionMode = false;  ← KEEP THIS COMMENTED
            }
        }
        catch (Exception ex)
        {
            // Sync failed, keep ghost mode active for retry
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "GHOST[SyncError]",
                Url          = _ghostSession?.SessionUrl ?? "unknown",
                Host         = "sessiondirectory.xboxlive.com",
                Path         = "",
                RequestBody  = "Background sync failed (will retry on next request)",
                StatusCode   = 0,
                ResponseBody = ex.Message,
            });
        }
    }

    // ── WinINet P/Invoke ──────────────────────────────────────────────────────
    [DllImport("wininet.dll", SetLastError = true)]
    private static extern bool InternetSetOption(
        IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);

    private const int INTERNET_OPTION_REFRESH          = 37;
    private const int INTERNET_OPTION_SETTINGS_CHANGED = 39;

    /// <summary>
    /// Reconstructs a PlayFab RequestParty response using cached game server info.
    /// This is used to redirect rejoin attempts to the original game server instead
    /// of letting PlayFab assign a new one after restart.
    /// </summary>
    private string ConstructRequestPartyResponse(GameServerInfo serverInfo)
    {
        var portsJson = string.Join(",", serverInfo.Ports.Select(p =>
            $"{{\"Name\":\"{EscapeJson(p.Name)}\",\"Num\":{p.Num},\"Protocol\":\"{EscapeJson(p.Protocol)}\"}}"));

        var json = $@"{{
  ""code"": 200,
  ""status"": ""OK"",
  ""data"": {{
    ""PartyId"": ""{EscapeJson(serverInfo.PartyId)}"",
    ""ServerId"": ""{EscapeJson(serverInfo.ServerId)}"",
    ""VmId"": ""{EscapeJson(serverInfo.VmId)}"",
    ""IPV4Address"": ""{EscapeJson(serverInfo.IPv4Address)}"",
    ""FQDN"": ""{EscapeJson(serverInfo.FQDN)}"",
    ""Ports"": [{portsJson}],
    ""Region"": ""{EscapeJson(serverInfo.Region)}"",
    ""State"": ""{EscapeJson(serverInfo.State)}"",
    ""ConnectedPlayers"": [],
    ""DTLSCertificateSHA2Thumbprint"": ""{EscapeJson(serverInfo.DTLSCertificateSHA2Thumbprint)}"",
    ""BuildId"": ""{EscapeJson(serverInfo.BuildId)}""
  }}
}}";
        return json;
    }

    /// <summary>Escapes special JSON characters in strings.</summary>
    private static string EscapeJson(string input)
    {
        if (string.IsNullOrEmpty(input)) return "";
        return input
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
    }

    // ── IDisposable ───────────────────────────────────────────────────────────
    public void Dispose() => Stop();
}

public sealed record SmartMatchWaitEstimate(
    string TicketId,
    string HopperName,
    int WaitSeconds,
    int GiveUpSeconds,
    DateTimeOffset CapturedAtUtc);

public sealed record HopperPopulationResult(
    string HopperName,
    int? WaitSeconds,
    int? Population,
    string Error);




