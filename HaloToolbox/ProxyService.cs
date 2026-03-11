using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Win32;
using Titanium.Web.Proxy;
using Titanium.Web.Proxy.EventArguments;
using Titanium.Web.Proxy.Models;

namespace HaloToolbox;

public class ProxyService : IDisposable
{
    // ── Configuration ─────────────────────────────────────────────────────────
    public int Port { get; set; } = 8888;

    // ── State ─────────────────────────────────────────────────────────────────
    public bool IsRunning { get; private set; }

    // Raised on the thread-pool; callers must marshal to the UI thread
    public event EventHandler<ProxyCaptureEntry>? OnRequestCaptured;

    // Raised when the WinHTTP elevation (UAC) is declined — provides the manual command
    public event EventHandler<string>? WinHttpManualSetRequired;

    // Raised when a CascadeMatchmaking session is persisted to disk
    public event EventHandler? OnMatchSessionSaved;

    // In-memory copy of the saved matchmaking session — used for session discovery injection
    private SavedHandleInfo? _lastMatchSession;

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

    // Game server redirection: cache the original game server info from RequestParty,
    // and redirect subsequent requests to use the same server (prevents PlayFab from
    // assigning a different server after restart, which breaks rejoin).
    private GameServerInfo? _cachedGameServerInfo = null;
    private bool _gameServerRedirectionActive = false;

    public bool IsGameServerRedirectionActive => _gameServerRedirectionActive;

    /// <summary>Clears cached game server info and disables redirection.</summary>
    public void ClearGameServerRedirection()
    {
        _cachedGameServerInfo = null;
        _gameServerRedirectionActive = false;
        // Also clear the persisted file
        try { File.Delete(GetGameServerCacheFile()); } catch { /* ignore */ }
    }

    /// <summary>Get the path to the cached game server file.</summary>
    private static string GetGameServerCacheFile() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                     "HaloIntel", "last-game-server.json");

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
            Debug.WriteLine($"Failed to load persisted game server: {ex.Message}");
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

        // CRITICAL FIX: Restore the persisted game server so we can redirect RequestParty
        // on rejoin. The server info is in memory normally, but we need to reload it from
        // disk in case MCC crashed and restarted before the proxy detected it.
        if (_cachedGameServerInfo == null)
        {
            _cachedGameServerInfo = LoadPersistedGameServer();
            if (_cachedGameServerInfo is not null)
            {
                Debug.WriteLine($"[RESTORE] Loaded persisted game server: {_cachedGameServerInfo.IPv4Address}:{_cachedGameServerInfo.Ports.FirstOrDefault()?.Num ?? 0}");
            }
        }

        // RACE CONDITION FIX: Accept matchSession parameter from caller (MainWindow)
        // to ensure we don't rely on _lastMatchSession timing. If caller provides it
        // and it's null locally, set it. This handles initialization race conditions.
        if (matchSession is not null && _lastMatchSession is null)
        {
            _lastMatchSession = matchSession;
            Debug.WriteLine($"[RESTORE] Match session restored from parameter: {matchSession.TemplateName}/{matchSession.SessionShort}");
        }

        // Enable aggressive ghost session mode
        if (_lastMatchSession is not null)
        {
            _ghostSessionMode = true;
            _ghostSession = _lastMatchSession;
            _ghostSessionSyncSuccess = false;

            // Start background sync immediately
            _ghostSessionSyncTask = AutoSyncGhostSessionAsync();
        }
    }

    public void ClearGhostSessionMode()
    {
        _ghostSessionMode = false;
        _ghostSession = null;
        _ghostSessionSyncSuccess = false;
    }

    public bool IsGhostSessionActive() => _ghostSessionMode;

    /// <summary>Check if crash restore timeout has expired and clear if needed.</summary>
    private void CheckAndClearPendingCrashRestoreTimeout()
    {
        if (_pendingCrashRestore && _pendingCrashRestoreStartedAt != DateTime.MinValue)
        {
            if ((DateTime.UtcNow - _pendingCrashRestoreStartedAt).TotalMinutes > CRASH_RESTORE_TIMEOUT_MINUTES)
            {
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
    // list, so only these three hosts are SSL-intercepted.  Everything else
    // (presence-heartbeat, userpresence, smartmatch, auth, …) becomes a plain TCP
    // tunnel with zero TLS overhead and zero interference with Xbox Live session state.
    private static readonly string[] _watchedDomains =
    [
        "halowaypoint.com",              // Halo Waypoint API + Spartan token extraction
        "sessiondirectory.xboxlive.com", // Xbox Live MPSD — session documents with skill data
        "playfabapi.com",                // PlayFab matchmaking / telemetry
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
        "HaloIntel", "proxy-root.pfx");

    // ── Start / Stop ──────────────────────────────────────────────────────────
    public async Task StartAsync()
    {
        if (IsRunning) return;

        Directory.CreateDirectory(Path.GetDirectoryName(CertStorePath)!);

        _server = new ProxyServer();
        _server.CertificateManager.PfxFilePath = CertStorePath;
        _server.CertificateManager.PfxPassword = "halointel-proxy";

        // Install CA into CurrentUser\Root — shows a one-time Windows trust dialog,
        // but does NOT require admin / UAC elevation.
        // Suppress: obsolete warning is informational; sync form still works correctly.
#pragma warning disable CS0618
        _server.CertificateManager.EnsureRootCertificate(
            userTrustRootCertificate:    true,
            machineTrustRootCertificate: false,
            trustRootCertificateAsAdmin: false);
#pragma warning restore CS0618

        _server.BeforeRequest  += OnBeforeRequestAsync;
        _server.BeforeResponse += OnBeforeResponseAsync;

        _endpoint = new ExplicitProxyEndPoint(IPAddress.Loopback, Port, decryptSsl: true);
        // BeforeTunnelConnectRequest lives on the endpoint, not the server —
        // must be registered after the endpoint is created.
        _endpoint.BeforeTunnelConnectRequest += OnBeforeTunnelConnectRequest;
        _server.AddEndPoint(_endpoint);
        await _server.StartAsync();

        // WinINet proxy (no admin required)
        SetWinINetProxy($"127.0.0.1:{Port}");

        // WinHTTP proxy — Halo MCC uses WinHTTP, not WinINet (admin / UAC required)
        await TrySetWinHttpProxyAsync();

        IsRunning = true;
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
    }

    // ── Tunnel-connect filter ─────────────────────────────────────────────────
    //
    // With decryptSsl:true on the endpoint, Titanium would MITM every HTTPS
    // connection by default — including presence heartbeats, auth, smartmatch —
    // adding TLS handshake overhead that disrupts Xbox Live session state.
    //
    // Here we opt non-watched domains OUT of SSL decryption: they become plain
    // TCP tunnels (zero overhead).  Only the three watched hosts are intercepted.
    private Task OnBeforeTunnelConnectRequest(object sender, TunnelConnectSessionEventArgs e)
    {
        e.DecryptSsl = IsDomainWatched(e.HttpClient.Request.RequestUri.Host);
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

        var entry = new ProxyCaptureEntry
        {
            Method         = req.Method,
            Url            = req.Url,
            Host           = req.RequestUri.Host,
            Path           = req.RequestUri.PathAndQuery,
            RequestHeaders = headers,
            RequestBody    = "",
        };

        // Safely capture request body for /handles endpoints (rejoin handle observation)
        if (req.HasBody && ShouldCaptureRequestBody(req))
        {
            var rawBytes = await e.GetRequestBody();
            e.SetRequestBody(rawBytes);  // restore exact bytes — no Content-Encoding change
            headers.TryGetValue("content-encoding", out var ce);
            entry.RequestBody = TryDecompressGzip(rawBytes, ce);

            // Persist to disk IMMEDIATELY — before awaiting the response — so a game
            // crash between the request and response doesn't lose the game session ref.
            PersistHandleToDisk(entry.RequestBody, headers);
        }

        // ── Block match leave during crash restore ────────────────────────────
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
        if (_blockMatchLeave && _lastMatchSession is not null &&
            req.Method == "PUT" &&
            req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.AbsolutePath.Contains(_lastMatchSession.SessionName, StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrEmpty(entry.RequestBody) &&
            entry.RequestBody.Contains("\"me\":null"))
        {
            // Short-circuit: fake 200 to MCC, leave never reaches MPSD
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

    /// <summary>Returns true for POST/PUT to /handles OR /sessions/ on the session directory.
    /// Captures PUT request bodies to session URLs so we can see what MCC writes
    /// (e.g., the 23-byte match session touch, squad session properties, etc.).</summary>
    private static bool ShouldCaptureRequestBody(Titanium.Web.Proxy.Http.Request req) =>
        (req.Method == "POST" || req.Method == "PUT") &&
        req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
        (req.RequestUri.AbsolutePath.Contains("/handles", StringComparison.OrdinalIgnoreCase) ||
         req.RequestUri.AbsolutePath.Contains("/sessions/", StringComparison.OrdinalIgnoreCase));

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
    /// %LocalAppData%\HaloIntel\last-handle.json. Called synchronously before the
    /// response arrives so the data survives a game crash. Best-effort; never throws.
    /// </summary>
    private static void PersistHandleToDisk(string bodyJson, Dictionary<string, string> requestHeaders)
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

            var dir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox");
            Directory.CreateDirectory(dir);
            File.WriteAllText(
                Path.Combine(dir, "last-handle.json"),
                JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { /* best-effort — never break the proxy */ }
    }

    /// <summary>
    /// Parses a CascadeMatchmaking session URL and writes the session reference + auth
    /// headers + connection GUID to %LocalAppData%\HaloMCCToolbox\last-match-session.json.
    /// Called when MCC PUTs to a matchmaking session (joining a match).  Best-effort; never throws.
    /// </summary>
    private void PersistMatchSessionToDisk(string url, Dictionary<string, string> requestHeaders, string requestBody = "")
    {
        try
        {
            // CRITICAL FIX: Don't persist closed sessions. The server may send:
            //   1. {"properties":{"system":{"closed":true}}} (direct property update)
            //   2. {"members":{"me":{"properties":{"system":{"closed":true}}}}} (member status)
            if (!string.IsNullOrEmpty(requestBody))
            {
                try
                {
                    using var doc = JsonDocument.Parse(requestBody);
                    var root = doc.RootElement;
                    bool isClosed = false;

                    // Check direct structure: {"properties":{"system":{"closed":true}}}
                    if (root.TryGetProperty("properties", out var props) &&
                        props.TryGetProperty("system", out var sys) &&
                        sys.TryGetProperty("closed", out var closed) &&
                        closed.GetBoolean())
                    {
                        isClosed = true;
                    }
                    // Check member structure: {"members":{"me":{"properties":{"system":{"closed":true}}}}}
                    else if (root.TryGetProperty("members", out var members) &&
                             members.TryGetProperty("me", out var me) &&
                             me.TryGetProperty("properties", out var meProps) &&
                             meProps.TryGetProperty("system", out var meSys) &&
                             meSys.TryGetProperty("closed", out var meClosed) &&
                             meClosed.GetBoolean())
                    {
                        isClosed = true;
                    }

                    if (isClosed)
                    {
                        Debug.WriteLine("[SAVE-Match] Skipping persistence of closed session");
                        return;  // Don't persist closed sessions
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[SAVE-Match] Failed to check closed state: {ex.Message}");
                }
            }

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

            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox");
            Directory.CreateDirectory(dir);
            File.WriteAllText(
                Path.Combine(dir, "last-match-session.json"),
                JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true }));

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
    }

    /// <summary>Sets the in-memory match session from an externally loaded SavedHandleInfo.
    /// If a new/different session is being loaded, clears the old ghost session to prevent
    /// rejoin attempts on stale sessions.</summary>
    public void SetSavedMatchSession(SavedHandleInfo? info)
    {
        // If we're loading a DIFFERENT session, clear ghost mode from the old one
        if (info is not null && _lastMatchSession is not null &&
            info.SessionName != _lastMatchSession.SessionName)
        {
            Debug.WriteLine($"[SESSION] New session detected ({info.SessionName}), clearing old ghost session");
            ClearGhostSessionMode();
        }
        _lastMatchSession = info;
    }

    /// <summary>Force-enables match-leave blocking (for manual UI trigger in rejoin testing).</summary>
    public void ForceBlockMatchLeave()
    {
        _blockMatchLeave = true;
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
                string connGuid = Guid.NewGuid().ToString();
                using var putReq = new HttpRequestMessage(HttpMethod.Put, match.SessionUrl);
                putReq.Content = new StringContent(
                    "{\"members\":{\"me\":{\"properties\":{\"system\":{\"active\":true,\"connection\":\"" + connGuid + "\"},\"custom\":{}}}}}",
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

        // Capture gamertag from X-Xbl-Debug header (format: "user0=...; gamertag=Foo; privilege=...")
        if (string.IsNullOrEmpty(_playerGamertag) &&
            headers.TryGetValue("X-Xbl-Debug", out var xblDebug) &&
            xblDebug.Contains("gamertag=", StringComparison.OrdinalIgnoreCase))
        {
            int gtIdx = xblDebug.IndexOf("gamertag=", StringComparison.OrdinalIgnoreCase) + 9;
            int gtEnd = xblDebug.IndexOf(';', gtIdx);
            _playerGamertag = gtEnd < 0 ? xblDebug[gtIdx..].Trim() : xblDebug[gtIdx..gtEnd].Trim();
        }

        var ct = resp.ContentType ?? "";
        bool isJson = ct.Contains("json", StringComparison.OrdinalIgnoreCase)
                   || ct.Contains("text", StringComparison.OrdinalIgnoreCase)
                   || ct.Contains("xml",  StringComparison.OrdinalIgnoreCase);

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
                if (!string.IsNullOrEmpty(xuid)) _playerXuid = xuid;
            }

            bool matchAlreadyInResults = body.Contains(match.SessionName, StringComparison.OrdinalIgnoreCase);
            bool isEmpty = body.Contains("\"results\":[]") || body.Contains("\"results\": []");

            if (_pendingCrashRestore && matchAlreadyInResults)
            {
                // PASSIVE: Match session is already in MPSD discovery results.
                // Player is still an active member on the server — let MCC see
                // the REAL response with the original pre-crash member state.
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
                    RequestBody  = "Match session already in MPSD discovery — passing through real response (passive mode)",
                    StatusCode   = 200,
                    ResponseBody = body.Length > 500 ? body[..500] + "…" : body,
                });
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
                    string connGuid = Guid.NewGuid().ToString();
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
                        "},\"custom\":{}}" +
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

        // ── Game Server Redirection (RequestParty response interception) ────────
        // Cache game server info from PlayFab RequestParty on initial match,
        // then redirect subsequent RequestParty calls to use the cached server.
        // This prevents PlayFab from assigning a different server after restart,
        // which would break rejoin since the new server doesn't have the player.
        if (!injected && !memberInjected &&
            entry.Host.Contains("playfabapi.com", StringComparison.OrdinalIgnoreCase) &&
            entry.Path.Contains("Party/RequestParty", StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody)
        {
            var body = await e.GetResponseBodyAsString();

            try
            {
                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;

                // Extract game server info from response
                if (root.TryGetProperty("data", out var data))
                {
                    var serverInfo = new GameServerInfo
                    {
                        PartyId = data.TryGetProperty("PartyId", out var pid) ? pid.GetString() ?? "" : "",
                        ServerId = data.TryGetProperty("ServerId", out var sid) ? sid.GetString() ?? "" : "",
                        VmId = data.TryGetProperty("VmId", out var vid) ? vid.GetString() ?? "" : "",
                        IPv4Address = data.TryGetProperty("IPV4Address", out var ip) ? ip.GetString() ?? "" : "",
                        FQDN = data.TryGetProperty("FQDN", out var fqdn) ? fqdn.GetString() ?? "" : "",
                        Region = data.TryGetProperty("Region", out var region) ? region.GetString() ?? "" : "",
                        State = data.TryGetProperty("State", out var state) ? state.GetString() ?? "" : "",
                        BuildId = data.TryGetProperty("BuildId", out var bid) ? bid.GetString() ?? "" : "",
                        DTLSCertificateSHA2Thumbprint = data.TryGetProperty("DTLSCertificateSHA2Thumbprint", out var cert) ? cert.GetString() ?? "" : "",
                        CachedAt = DateTime.UtcNow
                    };

                    // Parse ports
                    if (data.TryGetProperty("Ports", out var ports) && ports.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var port in ports.EnumerateArray())
                        {
                            serverInfo.Ports.Add(new GameServerPort
                            {
                                Name = port.TryGetProperty("Name", out var pname) ? pname.GetString() ?? "" : "",
                                Num = port.TryGetProperty("Num", out var pnum) ? pnum.GetInt32() : 0,
                                Protocol = port.TryGetProperty("Protocol", out var pproto) ? pproto.GetString() ?? "" : ""
                            });
                        }
                    }

                    // On first match: cache the server info
                    if (_cachedGameServerInfo == null && !_gameServerRedirectionActive)
                    {
                        _cachedGameServerInfo = serverInfo;
                        _gameServerRedirectionActive = false;  // Not active yet; only activate on crash restore

                        // PERSISTENCE: Save the game server to disk so it survives proxy restart/crash
                        PersistGameServerToDisk(serverInfo);

                        entry.ResponseBody = $"CACHED game server: {serverInfo.ServerShort}";

                        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                        {
                            Method = "CACHE[GameServer]",
                            Url = entry.Url,
                            Host = entry.Host,
                            Path = entry.Path,
                            RequestBody = "Game server info cached from RequestParty response",
                            StatusCode = 200,
                            ResponseBody = $"Server: {serverInfo.IPv4Address}:{serverInfo.Ports.FirstOrDefault()?.Num ?? 0} | ServerId: {serverInfo.ServerId[..13]}…",
                        });
                        return;  // Skip normal body capture, we already have the data
                    }
                    // On restart: if we have cached server info, redirect to it
                    else if (_cachedGameServerInfo is not null && _pendingCrashRestore)
                    {
                        _gameServerRedirectionActive = true;
                        // CRITICAL FIX: Do NOT clear _pendingCrashRestore here!
                        // MCC might make multiple RequestParty calls during the rejoin sequence,
                        // and ALL of them need to be redirected to the cached server.
                        // The flag will be cleared by SetPendingCrashRestore when MCC confirms
                        // successful rejoin or when the user clears the rejoin state.
                        // _pendingCrashRestore = false;  // ← REMOVED - was causing subsequent calls to bypass redirect

                        // Reconstruct the response with the cached server info
                        var redirectedBody = ConstructRequestPartyResponse(_cachedGameServerInfo);
                        e.SetResponseBodyString(redirectedBody);
                        entry.ResponseBody = redirectedBody;
                        resp.StatusCode = 200;
                        entry.StatusCode = 200;

                        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                        {
                            Method = "REDIRECT[GameServer]",
                            Url = entry.Url,
                            Host = entry.Host,
                            Path = entry.Path,
                            RequestBody = $"Redirecting to cached server instead of new assignment",
                            StatusCode = 200,
                            ResponseBody = $"Cached: {_cachedGameServerInfo.IPv4Address}:{_cachedGameServerInfo.Ports.FirstOrDefault()?.Num ?? 0}",
                        });
                        return;  // Skip normal body capture
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
                    RequestBody = "Failed to parse RequestParty response",
                    StatusCode = 500,
                    ResponseBody = ex.Message,
                });
            }

            // Still capture the body for normal logging
            entry.ResponseBody = body;
            e.SetResponseBodyString(body);
        }

        // ── Normal body capture ────────────────────────────────────────────
        if (!injected && !memberInjected && isJson && resp.HasBody && (entry.Method == "GET" || entry.Method == "POST") && ShouldReadBody(entry.Url))
        {
            var body = await e.GetResponseBodyAsString();
            entry.ResponseBody = body;
            // Must re-set or the connection will fail (body stream is consumed)
            e.SetResponseBodyString(body);
        }

        OnRequestCaptured?.Invoke(this, entry);
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
        reg.SetValue("ProxyOverride", "localhost;127.0.0.1;<local>",      RegistryValueKind.String);

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
    /// Check if a request is for the ghost session. AGGRESSIVE: catch ANY session with same SCID
    /// to prevent MCC from discovering it's not in rejoinable state.
    /// </summary>
    private bool IsRequestForGhostSession(Titanium.Web.Proxy.Http.Request req)
    {
        if (_ghostSession is null) return false;

        string url = req.RequestUri.AbsolutePath.ToLowerInvariant();
        string sessionName = _ghostSession.SessionName.ToLowerInvariant();
        string scid = _ghostSession.Scid.ToLowerInvariant();

        // Match CascadeMatchmaking with same session name
        bool isMatchSession = url.Contains("/cascadematchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
                              url.Contains(sessionName, StringComparison.OrdinalIgnoreCase);

        // Match cascadesquadsession with same SCID
        bool isSquadSession = url.Contains("/cascadesquadsession/sessions/", StringComparison.OrdinalIgnoreCase) &&
                              url.Contains(scid, StringComparison.OrdinalIgnoreCase);

        // AGGRESSIVE: Also catch ANY other session template with same SCID
        // (MCC might query other session templates during rejoin flow)
        bool isAnySCIDSession = url.Contains("/serviceconfigs/" + scid + "/sessionTemplates/", StringComparison.OrdinalIgnoreCase) &&
                                url.Contains("/sessions/");

        return isMatchSession || isSquadSession || isAnySCIDSession;
    }

    /// <summary>
    /// Handle ghost session requests: return fake responses that make MCC think it's
    /// still in the session, while real sync happens in the background.
    /// More aggressive: intercept both match and squad session, but let squad PUTs through.
    /// </summary>
    private bool HandleGhostSessionRequest(Titanium.Web.Proxy.Http.Request req, SessionEventArgs e)
    {
        if (_ghostSession is null) return false;

        string method = req.Method.ToUpperInvariant();
        string url = req.RequestUri.AbsolutePath;
        string sessionName = _ghostSession.SessionName.ToLowerInvariant();

        // Determine if this is match session or squad session
        bool isMatchSession = url.Contains("/cascadematchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
                              url.Contains(sessionName, StringComparison.OrdinalIgnoreCase);
        bool isSquadSession = url.Contains("/cascadesquadsession/sessions/", StringComparison.OrdinalIgnoreCase) &&
                              url.Contains(_ghostSession.Scid, StringComparison.OrdinalIgnoreCase);

        // ── Match Session Interception ────────────────────────────────────
        // Block leave, fake queries for match session
        if (isMatchSession)
        {
            // GET match session — return fake "you're active"
            if (method == "GET" && !url.Contains("/members/"))
            {
                string fakeSessionBody = GenerateFakeSessionDocument();
                e.Ok(fakeSessionBody);
                LogGhostRequest("GHOST[Match-Get]", req.Url, "200", "Fake match document");
                return true;
            }

            // GET /members/me in match session
            if (method == "GET" && url.Contains("/members/me", StringComparison.OrdinalIgnoreCase))
            {
                string fakeMemberBody = GenerateFakeMemberDocument();
                e.Ok(fakeMemberBody);
                LogGhostRequest("GHOST[Match-Member]", req.Url, "200", "Fake member (active=true)");
                return true;
            }

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
          "connection": "12345678-1234-1234-1234-123456789abc"
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
          "connection": "12345678-1234-1234-1234-123456789abc",
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
        return $$"""
{
  "gamertag": "Player",
  "xuid": "{{_playerXuid}}",
  "roleTypes": [],
  "properties": {
    "system": {
      "active": true,
      "connection": "12345678-1234-1234-1234-123456789abc",
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
            string connectionField = string.IsNullOrEmpty(_ghostSession.ConnectionGuid)
                ? ""
                : $",\"connection\":\"{_ghostSession.ConnectionGuid}\"";

            var putBody = $$"""
{
  "members": {
    "me": {
      "properties": {
        "system": {
          "active": true{{connectionField}}
        },
        "custom": {}
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
