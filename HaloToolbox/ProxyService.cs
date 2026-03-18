using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
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
    // ── Configuration ─────────────────────────────────────────────────────────
    public int Port { get; set; } = 8888;

    // ── State ─────────────────────────────────────────────────────────────────
    public bool IsRunning { get; private set; }

    // ── Events ────────────────────────────────────────────────────────────────

    // Raised on the thread-pool; callers must marshal to the UI thread.
    public event EventHandler<ProxyCaptureEntry>? OnRequestCaptured;

    // Raised when UAC is declined for WinHTTP setup — provides the manual command.
    public event EventHandler<string>? WinHttpManualSetRequired;

    // Raised when a CascadeMatchmaking session is newly captured.
    public event EventHandler? OnSessionCaptured;

    // ── Session capture + crash recovery ──────────────────────────────────────
    //
    // Normal play: capture only. Zero modifications to traffic.
    //
    // After a crash, within a 30s window:
    //
    //   1. Discovery injection:
    //      If MCC's /sessions? returns no match, inject the sessionRef so MCC
    //      can navigate to the session using its own live RTA connection.
    //
    //   2. Leave block (multi-shot, 30s window):
    //      After crash, MCC's RTA connection is dead. Membership validation fails,
    //      causing MCC to send PUT {"members":{"me":null}} to the session root URL
    //      (/cascadematchmaking/sessions/{uuid}) — possibly multiple times across
    //      different recovery branches.
    //      Block ALL such leaves for the captured session during the crash window.
    //      This forces MCC through its rejoin code path instead of queuing anew.

    private SavedHandleInfo? _capturedSession;
    private DateTime?        _crashDetectedAt;

    // Single window for both injection and leave blocking.
    // MCC takes ~34s from crash to RequestParty on rejoin — 60s gives safe margin.
    private const int CrashWindowSeconds = 60;

    // How many leaves we've blocked this recovery (diagnostic)
    private int _leaveBlockCount = 0;


    // Cache the RequestParty response during normal play so we can replay it
    // during crash recovery. Without this, MCC gets a NEW empty party instead
    // of the original match party, so the game connection fails.
    private string _cachedPartyResponse = "";

    // Background MPSD client — bypasses our own proxy.
    private static readonly HttpClient _mpsdBackgroundClient =
        new(new HttpClientHandler { UseProxy = false }) { Timeout = TimeSpan.FromSeconds(8) };

    private string _playerXuid     = "";
    private string _playerGamertag = "";

    // PlayFab entity token + host captured from MCC's own requests.
    // Used to make background RequestParty calls to prime the party cache
    // before a crash ever happens.
    private string _capturedEntityToken      = "";
    private string _capturedPlayFabHost      = "";
    private string _capturedPartyRequestBody = "";  // exact body MCC sends to /Party/RequestParty

    public SavedHandleInfo? CapturedSession => _capturedSession;

    /// <summary>
    /// Party cache status for display in the Rejoin Guard panel.
    /// </summary>
    public enum PartyCacheState { None, Priming, Stale, Ready }
    private PartyCacheState _partyCacheState  = PartyCacheState.None;
    private DateTime?       _partyCachedAt    = null;
    private string          _cachedPartyIp    = "";   // IP in _cachedPartyResponse
    private string          _livePartyIp      = "";   // IP MCC last received from PlayFab directly
    public  PartyCacheState PartyCache        => _partyCacheState;
    public  DateTime?       PartyCachedAt     => _partyCachedAt;
    public  string          CachedPartyIp     => _cachedPartyIp;
    public  string          LivePartyIp       => _livePartyIp;

    /// <summary>Raised when party cache state changes so the UI can refresh.</summary>
    public event EventHandler? OnPartyCacheChanged;

    public bool IsInCrashWindow =>
        _crashDetectedAt.HasValue &&
        (DateTime.UtcNow - _crashDetectedAt.Value).TotalSeconds <= CrashWindowSeconds;

    /// <summary>Called by Scanner when MCC exits. Arms discovery injection + leave block.</summary>
    public void SetCrashDetected()
    {
        _crashDetectedAt = DateTime.UtcNow;
        _leaveBlockCount = 0;
        Debug.WriteLine($"[CRASH] Detected at {_crashDetectedAt:HH:mm:ss} — 30s window started");

        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
        {
            Method       = "CRASH[Detected]",
            Url          = "",
            Host         = "local",
            Path         = "",
            RequestBody  = _capturedSession is not null
                ? $"session={_capturedSession.TemplateName}/{_capturedSession.SessionShort}"
                : "No session captured — cannot inject or block",
            StatusCode   = 0,
            ResponseBody = _capturedSession is not null
                ? "30s window: blocking all leave PUTs for captured session + injecting sessionRef if MPSD forgets us"
                : "Crash detected but no CascadeMatchmaking session was captured during play",
        });
    }

    /// <summary>Clears the captured session and all crash recovery state (e.g. user clicks CLEAR).</summary>
    public void ClearCapturedSession()
    {
        _capturedSession     = null;
        _crashDetectedAt     = null;
        _leaveBlockCount     = 0;
        // NOTE: _cachedPartyResponse and the disk file are intentionally NOT cleared here.
        // The party response must survive session clears so the NEXT crash can still redirect
        // RequestParty to the correct party server. It is overwritten by the next CAPTURE[Party].
    }

    // ── Proxy internals ───────────────────────────────────────────────────────
    private ProxyServer?           _server;
    private ExplicitProxyEndPoint? _endpoint;

    // WinINet originals — restored on Stop()
    private int    _savedProxyEnable;
    private string _savedProxyServer   = "";
    private string _savedProxyOverride = "";

    // ── Domain filter ─────────────────────────────────────────────────────────
    //
    // Only these three domains are SSL-intercepted. Everything else (presence
    // heartbeats, auth, smartmatch, …) becomes a plain TCP tunnel — zero overhead,
    // zero interference with Xbox Live session state.
    private static readonly string[] _watchedDomains =
    [
        "halowaypoint.com",              // Halo Waypoint API + Spartan token extraction
        "sessiondirectory.xboxlive.com", // Xbox Live MPSD — session documents + discovery
        "playfabapi.com",                // PlayFab matchmaking / telemetry
    ];

    // Hosts to skip body capture even within watched domains.
    private static readonly string[] _bypassHosts =
    [
        "banprocessor",  // MCC ban-check — may use unusual response format; skip body read
    ];

    // ── Certificate + persisted state paths ──────────────────────────────────
    private static string DataDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox");

    private static string CertStorePath        => Path.Combine(DataDir, "proxy-root.pfx");
    private static string PartyResponsePath    => Path.Combine(DataDir, "last-party-response.json");
    private static string PartyRequestBodyPath => Path.Combine(DataDir, "last-party-request-body.json");

    // ── Start / Stop ──────────────────────────────────────────────────────────
    public async Task StartAsync()
    {
        if (IsRunning) return;

        Directory.CreateDirectory(DataDir);
        TryLoadPartyResponseFromDisk();

        _server = new ProxyServer();
        _server.CertificateManager.PfxFilePath = CertStorePath;
        _server.CertificateManager.PfxPassword = "halointel-proxy";

        // Install CA into CurrentUser\Root — shows a one-time Windows trust dialog.
#pragma warning disable CS0618
        _server.CertificateManager.EnsureRootCertificate(
            userTrustRootCertificate:    true,
            machineTrustRootCertificate: false,
            trustRootCertificateAsAdmin: false);
#pragma warning restore CS0618

        _server.BeforeRequest  += OnBeforeRequestAsync;
        _server.BeforeResponse += OnBeforeResponseAsync;

        _endpoint = new ExplicitProxyEndPoint(IPAddress.Loopback, Port, decryptSsl: true);
        _endpoint.BeforeTunnelConnectRequest += OnBeforeTunnelConnectRequest;
        _server.AddEndPoint(_endpoint);
        await _server.StartAsync();

        SetWinINetProxy($"127.0.0.1:{Port}");
        await TrySetWinHttpProxyAsync();

        IsRunning = true;
    }

    public void Stop()
    {
        if (!IsRunning) return;

        RestoreWinINetProxy();
        TryResetWinHttpProxy();

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
    // Opt non-watched domains out of SSL decryption — they become plain TCP tunnels.
    private Task OnBeforeTunnelConnectRequest(object sender, TunnelConnectSessionEventArgs e)
    {
        e.DecryptSsl = IsDomainWatched(e.HttpClient.Request.RequestUri.Host);
        return Task.CompletedTask;
    }

    // ── Request intercept ─────────────────────────────────────────────────────
    //
    // We do NOT modify any outgoing requests. We only:
    //   1. Build the capture entry (for the log) from headers
    //   2. Read the body for /handles POST (display only)
    //   3. Extract the player XUID from discovery query strings
    //   4. Stash the entry on UserData for OnBeforeResponseAsync to complete
    private async Task OnBeforeRequestAsync(object sender, SessionEventArgs e)
    {
        var req = e.HttpClient.Request;
        if (!IsDomainWatched(req.RequestUri.Host)) return;

        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var h in req.Headers)
            headers[h.Name] = h.Value;

        var entry = new ProxyCaptureEntry
        {
            Method         = req.Method ?? "",
            Url            = req.Url,
            Host           = req.RequestUri.Host,
            Path           = req.RequestUri.PathAndQuery,
            RequestHeaders = headers,
            RequestBody    = "",
        };

        // Read body for /handles, /sessions PUT, and /members/me PUT — display + leave detection.
        if (req.HasBody && ShouldCaptureRequestBody(req))
        {
            var rawBytes = await e.GetRequestBody();
            e.SetRequestBody(rawBytes);  // restore exact bytes
            headers.TryGetValue("content-encoding", out var ce);
            entry.RequestBody = TryDecompressGzip(rawBytes, ce);
        }

        // ── Leave block (multi-shot, 30s crash window) ────────────────────────
        // After crash, MCC's RTA connection is dead. Membership validation fails,
        // causing MCC to send PUT {"members":{"me":null}} to the session root URL.
        // This can happen multiple times across different recovery branches.
        // Block ALL such leaves for the captured session during the crash window.
        //
        // The leave goes to the SESSION ROOT URL (not /members/me in the path):
        //   PUT /cascadematchmaking/sessions/{uuid}  body={"members":{"me":null}}
        // We match on: host + captured session name in path + leave body.
        if (!string.IsNullOrEmpty(entry.RequestBody) &&
            req.Method == "PUT" &&
            req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            _capturedSession is not null &&
            req.RequestUri.AbsolutePath.Contains(_capturedSession.SessionName, StringComparison.OrdinalIgnoreCase) &&
            IsInCrashWindow &&
            IsLeaveBody(entry.RequestBody))
        {
            _leaveBlockCount++;
            e.Ok("{}");
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "REJOIN[LeaveBlocked]",
                Url          = entry.Url,
                Host         = entry.Host,
                Path         = entry.Path,
                RequestBody  = entry.RequestBody,
                StatusCode   = 200,
                ResponseBody = $"Blocked leave #{_leaveBlockCount} (faked 200 {{}}) — preserving membership",
            });
            return;
        }

        // ── Connection update: grab new RTA connection GUID from squad PUT ──────
        // MCC creates the new squad session (with the live connection GUID) BEFORE
        // it sends the CascadeMatchmaking leave. Do NOT gate on _pendingConnectionUpdate —
        // instead, trigger on any squad PUT with a connection GUID during the crash window.
        // Auth header is taken directly from the squad PUT itself.
        if (IsInCrashWindow &&
            _capturedSession is not null &&
            req.Method == "PUT" &&
            req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.AbsolutePath.Contains("/cascadesquadsession/", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrEmpty(entry.RequestBody) &&
            TryExtractMemberConnection(entry.RequestBody, out var newConn, out var newSub))
        {
            headers.TryGetValue("Authorization", out var squadAuth);
            var authHeader   = squadAuth ?? "";
            var capturedSnap = _capturedSession;
            _ = Task.Run(() => UpdateMatchSessionConnectionAsync(capturedSnap, newConn, newSub, authHeader));
        }

        // Extract player XUID from session discovery URL (/sessions?xuid=...)
        if (req.RequestUri.PathAndQuery.Contains("/sessions?", StringComparison.OrdinalIgnoreCase))
        {
            int xuidIdx = req.Url.IndexOf("xuid=", StringComparison.OrdinalIgnoreCase);
            if (xuidIdx >= 0)
            {
                int end = req.Url.IndexOf('&', xuidIdx);
                _playerXuid = end < 0
                    ? req.Url[(xuidIdx + 5)..]
                    : req.Url[(xuidIdx + 5)..end];
            }
        }

        // Capture PlayFab entity token + host from any PlayFab request during active play.
        // Used to prime the party cache before a crash happens.
        if (req.RequestUri.Host.EndsWith("playfabapi.com", StringComparison.OrdinalIgnoreCase) &&
            headers.TryGetValue("X-EntityToken", out var entityToken) &&
            !string.IsNullOrEmpty(entityToken))
        {
            _capturedEntityToken = entityToken;
            _capturedPlayFabHost = req.RequestUri.Host;

            // Capture the exact request body MCC uses for RequestParty so we can
            // replay it verbatim in background prime calls. PlayFab requires fields
            // like "Version" or "Build" that an empty {} body omits.
            if (req.Method == "POST" &&
                req.RequestUri.AbsolutePath.Contains("/Party/RequestParty", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrEmpty(entry.RequestBody))
            {
                _capturedPartyRequestBody = entry.RequestBody;
                // Persist so the body survives proxy restarts — once captured, prime works forever.
                try { File.WriteAllText(PartyRequestBodyPath, entry.RequestBody, System.Text.Encoding.UTF8); }
                catch { /* best-effort */ }
            }
        }

        // ── PartyId swap: replace new post-crash UUID with saved original ─────
        // When MCC crashes, it loses party state and generates a new PartyId on
        // restart. PlayFab allocates a fresh empty server for the new ID instead
        // of returning the original match party. Replace MCC's new PartyId with
        // the one from our cached party response so PlayFab routes back to the
        // original party server where the match is still happening.
        if (IsInCrashWindow &&
            req.Method == "POST" &&
            req.RequestUri.Host.EndsWith("playfabapi.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.AbsolutePath.Contains("/Party/RequestParty", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrEmpty(entry.RequestBody) &&
            !string.IsNullOrEmpty(_cachedPartyResponse))
        {
            var cachedPartyId = TryExtractPartyId(_cachedPartyResponse);
            if (!string.IsNullOrEmpty(cachedPartyId))
            {
                try
                {
                    var bodyNode  = JsonNode.Parse(entry.RequestBody)?.AsObject();
                    var newPartyId = bodyNode?["PartyId"]?.GetValue<string>() ?? "";
                    if (bodyNode is not null &&
                        !string.Equals(newPartyId, cachedPartyId, StringComparison.OrdinalIgnoreCase))
                    {
                        bodyNode["PartyId"] = cachedPartyId;
                        var modifiedBody = bodyNode.ToJsonString();
                        e.SetRequestBody(System.Text.Encoding.UTF8.GetBytes(modifiedBody));
                        entry.RequestBody = modifiedBody;
                        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                        {
                            Method       = "REJOIN[PartyIdSwap]",
                            Url          = entry.Url,
                            Host         = entry.Host,
                            Path         = entry.Path,
                            RequestBody  = modifiedBody,
                            StatusCode   = 0,
                            ResponseBody = $"Swapped new PartyId {newPartyId[..Math.Min(8, newPartyId.Length)]}… → cached {cachedPartyId[..Math.Min(8, cachedPartyId.Length)]}… so PlayFab routes to original match party",
                        });
                    }
                }
                catch { /* best-effort — don't break the request on parse failure */ }
            }
        }

        // Stash for OnBeforeResponseAsync
        e.HttpClient.UserData = entry;
    }

    // ── Response intercept ────────────────────────────────────────────────────
    //
    // We do NOT modify responses during normal play.
    // We only:
    //   1. Capture the CascadeMatchmaking session from successful PUT responses
    //   2. Inject sessionRef into discovery response (ONLY within 30s crash window,
    //      ONLY if session is not already in MPSD results)
    //   3. Read + log response bodies for display
    private async Task OnBeforeResponseAsync(object sender, SessionEventArgs e)
    {
        var req  = e.HttpClient.Request;
        var resp = e.HttpClient.Response;
        if (!IsDomainWatched(req.RequestUri.Host)) return;

        var entry = e.HttpClient.UserData as ProxyCaptureEntry ?? new ProxyCaptureEntry
        {
            Method = req.Method ?? "",
            Url    = req.Url,
            Host   = req.RequestUri.Host,
            Path   = req.RequestUri.PathAndQuery,
        };
        entry.StatusCode = resp.StatusCode;

        // Capture gamertag from debug response header
        if (resp.Headers.GetFirstHeader("X-Xbl-Debug") is { } dbgH &&
            dbgH.Value.Contains("gt=", StringComparison.OrdinalIgnoreCase))
        {
            int gtIdx = dbgH.Value.IndexOf("gt=", StringComparison.OrdinalIgnoreCase);
            if (gtIdx >= 0)
            {
                int end = dbgH.Value.IndexOf(';', gtIdx);
                _playerGamertag = end < 0
                    ? dbgH.Value[(gtIdx + 3)..]
                    : dbgH.Value[(gtIdx + 3)..end];
            }
        }

        // ── Session capture ────────────────────────────────────────────────
        // When MCC successfully PUTs to CascadeMatchmaking (not /members/me),
        // it's actively joining a match — capture the session.
        if (req.Method == "PUT" &&
            req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.PathAndQuery.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
            !req.RequestUri.PathAndQuery.Contains("/members/", StringComparison.OrdinalIgnoreCase) &&
            (resp.StatusCode == 200 || resp.StatusCode == 201))
        {
            CaptureSessionFromUrl(req.Url);
        }

        // ── Discovery injection ───────────────────────────────────────────
        // Fallback only: if MCC's session discovery comes back empty (MPSD
        // removed us after crash), inject just the sessionRef so MCC can still
        // navigate to the session and use its own live RTA connection to rejoin.
        bool injected = false;
        if (_capturedSession is not null &&
            IsInCrashWindow &&
            req.Method == "GET" &&
            req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.PathAndQuery.Contains("/sessions?", StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody)
        {
            var body = await e.GetResponseBodyAsString();
            bool sessionInResults = body.Contains(_capturedSession.SessionName, StringComparison.OrdinalIgnoreCase);

            if (sessionInResults)
            {
                // Session still alive in MPSD — but the raw response may include OTHER
                // sessions (e.g. a cascadesquadsession with "closed":true) that cause MCC
                // to skip the rejoin prompt and start fresh matchmaking instead.
                // Fix: return the same minimal injected format as the Injected path —
                // only the CascadeMatchmaking sessionRef, nothing else.
                var filteredBody =
                    "{\"results\":[{\"xuid\":\"" + _playerXuid +
                    "\",\"startTime\":\"" + _capturedSession.SavedAt.ToString("O") +
                    "\",\"sessionRef\":{\"scid\":\"" + _capturedSession.Scid +
                    "\",\"templateName\":\"" + _capturedSession.TemplateName +
                    "\",\"name\":\"" + _capturedSession.SessionName +
                    "\"}}]}";

                e.SetResponseBodyString(filteredBody);
                entry.ResponseBody = filteredBody;

                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "DISCOVERY[Found]",
                    Url          = entry.Url,
                    Host         = entry.Host,
                    Path         = entry.Path,
                    StatusCode   = 200,
                    ResponseBody = $"Session still in MPSD — filtered to match-only sessionRef (strips closed squad) for {_capturedSession.TemplateName}/{_capturedSession.SessionShort}",
                });
            }
            else
            {
                // Session not in MPSD results — inject just the sessionRef so
                // MCC can discover the URL. MCC's own native rejoin flow handles auth + RTA.
                var injectedBody =
                    "{\"results\":[{\"xuid\":\"" + _playerXuid +
                    "\",\"startTime\":\"" + _capturedSession.SavedAt.ToString("O") +
                    "\",\"sessionRef\":{\"scid\":\"" + _capturedSession.Scid +
                    "\",\"templateName\":\"" + _capturedSession.TemplateName +
                    "\",\"name\":\"" + _capturedSession.SessionName +
                    "\"}}]}";

                e.SetResponseBodyString(injectedBody);
                entry.ResponseBody = injectedBody;
                injected = true;

                double elapsed = (DateTime.UtcNow - _crashDetectedAt!.Value).TotalSeconds;
                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "DISCOVERY[Injected]",
                    Url          = entry.Url,
                    Host         = entry.Host,
                    Path         = entry.Path,
                    StatusCode   = 200,
                    ResponseBody = $"Session not in MPSD — injected sessionRef for {_capturedSession.TemplateName}/{_capturedSession.SessionShort}  ({elapsed:F1}s after crash)",
                });
            }
        }

        // ── CascadeMatchmaking member injection ───────────────────────────
        // After a crash, MPSD may evict the player before we can update the
        // connection GUID (inactiveRemovalTimeout=0). MCC's GET on the session
        // returns 7 members (player absent) → MCC sees no rejoin button.
        // Inject the player's XUID back into the GET response so MCC shows
        // the rejoin button. The actual MPSD state is fixed ~200ms later by
        // UpdateMatchSessionConnectionAsync when the squad PUT arrives.

        // DEBUG: Log why member injection may or may not fire
        bool debugMemberInjection = !injected && _capturedSession is not null && !string.IsNullOrEmpty(_playerXuid) &&
                                    req.Method == "GET" && req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase);
        if (debugMemberInjection)
        {
            Debug.WriteLine($"[INJECT-DEBUG] CascadeMatchmaking GET detected. IsInCrashWindow={IsInCrashWindow}, _crashDetectedAt={_crashDetectedAt}, elapsed={(DateTime.UtcNow - (_crashDetectedAt ?? DateTime.UtcNow)).TotalSeconds}s");
        }

        if (!injected &&
            _capturedSession is not null &&
            IsInCrashWindow &&
            !string.IsNullOrEmpty(_playerXuid) &&
            req.Method == "GET" &&
            req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.AbsolutePath.Contains(_capturedSession.SessionName, StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.AbsolutePath.Contains("/CascadeMatchmaking/", StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody)
        {
            var sessionBody = await e.GetResponseBodyAsString();
            e.SetResponseBodyString(sessionBody);
            entry.ResponseBody = sessionBody;
            injected = true;

            if (!sessionBody.Contains(_playerXuid, StringComparison.OrdinalIgnoreCase))
            {
                var patched = InjectPlayerMember(sessionBody, _playerXuid);
                if (patched is not null)
                {
                    e.SetResponseBodyString(patched);
                    entry.ResponseBody = patched;

                    double elapsed = (DateTime.UtcNow - _crashDetectedAt!.Value).TotalSeconds;
                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method       = "INJECT[Member]",
                        Url          = entry.Url,
                        Host         = entry.Host,
                        Path         = entry.Path,
                        StatusCode   = 200,
                        ResponseBody = $"Injected XUID {_playerXuid} into CascadeMatchmaking GET ({elapsed:F1}s after crash) — MCC should show rejoin button",
                    });
                }
            }
        }

        // ── RequestParty cache + redirect ─────────────────────────────────
        // During normal play: cache the full response so we can replay it.
        // During crash recovery: replace the response with the cached one so
        // MCC connects to the original match party instead of a new empty one.
        if (req.Method == "POST" &&
            req.RequestUri.Host.EndsWith("playfabapi.com", StringComparison.OrdinalIgnoreCase) &&
            req.RequestUri.AbsolutePath.Contains("/Party/RequestParty", StringComparison.OrdinalIgnoreCase) &&
            resp.StatusCode == 200 &&
            resp.HasBody)
        {
            var partyBody = await e.GetResponseBodyAsString();
            e.SetResponseBodyString(partyBody);
            entry.ResponseBody = partyBody;

            if (_crashDetectedAt.HasValue && !string.IsNullOrEmpty(_cachedPartyResponse))
            {
                // During crash recovery: only redirect if PlayFab is sending MCC to a DIFFERENT
                // party server than the one in the cache. If PlayFab returns the correct server,
                // pass the REAL response through — it contains fresh party credentials that the
                // party server will accept. Redirecting a matching response would give MCC stale
                // credentials that the party server may reject after the DTLS connection dropped.
                var realIp   = TryExtractPartyIp(partyBody);
                var cachedIp = TryExtractPartyIp(_cachedPartyResponse);
                bool serversDiffer = !string.IsNullOrEmpty(realIp) &&
                                     !string.IsNullOrEmpty(cachedIp) &&
                                     !string.Equals(realIp, cachedIp, StringComparison.OrdinalIgnoreCase);

                if (serversDiffer)
                {
                    // PlayFab returned the wrong server — redirect to the original party server.
                    _livePartyIp = realIp;   // record what PlayFab sent before we overwrote it
                    e.SetResponseBodyString(_cachedPartyResponse);
                    entry.ResponseBody = _cachedPartyResponse;
                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method       = "REJOIN[PartyRedirect]",
                        Url          = entry.Url,
                        Host         = entry.Host,
                        Path         = entry.Path,
                        StatusCode   = 200,
                        ResponseBody = $"PlayFab→{realIp} WRONG — redirected to cached {cachedIp}: {_cachedPartyResponse[..Math.Min(100, _cachedPartyResponse.Length)]}",
                    });
                }
                else
                {
                    // PlayFab already returned the correct server — pass real response through
                    // so MCC gets fresh party credentials. Update the cache with the fresh response.
                    _cachedPartyResponse = partyBody;
                    _cachedPartyIp       = realIp;
                    _livePartyIp         = realIp;
                    SetPartyState(PartyCacheState.Ready);
                    try { File.WriteAllText(PartyResponsePath, partyBody, System.Text.Encoding.UTF8); }
                    catch { /* best-effort */ }
                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method       = "REJOIN[PartyPassthrough]",
                        Url          = entry.Url,
                        Host         = entry.Host,
                        Path         = entry.Path,
                        StatusCode   = 200,
                        ResponseBody = $"PlayFab→{realIp} matches cached — passing fresh response through (updated cache): {partyBody[..Math.Min(100, partyBody.Length)]}",
                    });
                }
            }
            else
            {
                // Cache: save this party response regardless of whether we're in a crash window.
                // RequestParty only fires during the rejoin flow (never during passive play),
                // so we must capture it whenever we see it — crash or not — so it's available
                // for the NEXT crash recovery via disk.
                //
                // Staleness check: skip caching if the party server has been idle for >1 hour
                // with zero connected players — PlayFab sometimes returns recycled stale
                // assignments that will reject DTLS connections from MCC.
                if (IsStaleParty(partyBody))
                {
                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method       = "PARTY[Stale]",
                        Url          = entry.Url,
                        Host         = entry.Host,
                        Path         = entry.Path,
                        StatusCode   = 200,
                        ResponseBody = $"Skipped stale party (LastStateTransitionTime >1h old): {partyBody[..Math.Min(120, partyBody.Length)]}",
                    });
                }
                else
                {
                    _cachedPartyResponse = partyBody;
                    _cachedPartyIp       = TryExtractPartyIp(partyBody);
                    _livePartyIp         = _cachedPartyIp;   // MCC's actual call — live = cached
                    SetPartyState(PartyCacheState.Ready);
                    try { File.WriteAllText(PartyResponsePath, partyBody, System.Text.Encoding.UTF8); }
                    catch { /* best-effort */ }

                    // Option B fallback: if we have the response but missed the request body,
                    // construct a body from the BuildId in the response so future primes work.
                    if (string.IsNullOrEmpty(_capturedPartyRequestBody))
                    {
                        var buildId = TryExtractBuildId(partyBody);
                        if (!string.IsNullOrEmpty(buildId))
                        {
                            _capturedPartyRequestBody = $"{{\"Build\":\"{buildId}\"}}";
                            try { File.WriteAllText(PartyRequestBodyPath, _capturedPartyRequestBody, System.Text.Encoding.UTF8); }
                            catch { /* best-effort */ }
                            Debug.WriteLine($"[PARTY] Constructed RequestParty body from BuildId: {buildId}");
                        }
                    }

                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method       = "CAPTURE[Party]",
                        Url          = entry.Url,
                        Host         = entry.Host,
                        Path         = entry.Path,
                        StatusCode   = 200,
                        ResponseBody = $"Cached party for crash recovery: {partyBody[..Math.Min(120, partyBody.Length)]}",
                    });
                }
            }

            OnRequestCaptured?.Invoke(this, entry);
            return;
        }

        // ── Body capture for display ───────────────────────────────────────
        if (!injected &&
            resp.HasBody &&
            ShouldReadBody(entry.Url) &&
            IsJsonContent(resp.ContentType) &&
            (req.Method == "GET" || req.Method == "POST"))
        {
            var body = await e.GetResponseBodyAsString();
            entry.ResponseBody = body;
            e.SetResponseBodyString(body);
        }

        OnRequestCaptured?.Invoke(this, entry);
    }

    // ── Session capture helper ────────────────────────────────────────────────
    private void CaptureSessionFromUrl(string url)
    {
        try
        {
            var    uri  = new Uri(url);
            var    segs = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            string scid = "", templateName = "", sessionName = "";

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

            bool isNew = _capturedSession?.SessionName != sessionName;

            // During crash window, MCC cleans up old session residue from previous runs.
            // Ignore any session that differs from the crash session — do not overwrite
            // _capturedSession or clear _crashDetectedAt until the window expires.
            if (isNew && IsInCrashWindow)
            {
                Debug.WriteLine($"[CAPTURE] Ignoring session change to {sessionName[..Math.Min(13, sessionName.Length)]}… during crash window (protecting {_capturedSession?.SessionName?[..Math.Min(13, _capturedSession.SessionName.Length)] ?? "null"})");
                return;
            }

            _capturedSession = new SavedHandleInfo
            {
                Scid         = scid,
                TemplateName = templateName,
                SessionName  = sessionName,
                SavedAt      = DateTime.UtcNow,
            };

            if (isNew)
            {
                // New session = MCC found a different match. Crash recovery is over.
                // CRITICAL FIX: Do NOT clear _crashDetectedAt here. The 60s crash window
                // must stay open throughout the entire rejoin process so party redirect
                // works (line 595 requires _crashDetectedAt.HasValue). The window will
                // expire naturally on its own schedule.

                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "CAPTURE[Session]",
                    Url          = url,
                    Host         = "sessiondirectory.xboxlive.com",
                    Path         = uri.AbsolutePath,
                    StatusCode   = 0,
                    ResponseBody = $"Captured {templateName}/{sessionName[..Math.Min(13, sessionName.Length)]}…",
                });
                OnSessionCaptured?.Invoke(this, EventArgs.Empty);

                // Auto-prime the party cache: fire a background RequestParty call using
                // MCC's current entity token so any subsequent crash has a valid party
                // cached without requiring a manual rejoin first.
                // Delay 5s to let PlayFab associate the entity with the new match.
                _ = Task.Run(async () =>
                {
                    await Task.Delay(5000);
                    await PrimeCacheAsync(auto: true);
                });
            }
        }
        catch { /* best-effort */ }
    }

    // ── Connection update helpers ─────────────────────────────────────────────

    /// <summary>
    /// Parses members.me.properties.system.connection and .subscription.id
    /// from a squad session PUT body. Returns false if not present.
    /// </summary>
    private static bool TryExtractMemberConnection(string body, out string connGuid, out string subId)
    {
        connGuid = "";
        subId    = "";
        try
        {
            using var doc  = JsonDocument.Parse(body);
            var       root = doc.RootElement;
            if (root.TryGetProperty("members", out var members)   &&
                members.TryGetProperty("me",    out var me)        &&
                me.TryGetProperty("properties", out var props)     &&
                props.TryGetProperty("system",  out var sys)       &&
                sys.TryGetProperty("connection", out var conn))
            {
                connGuid = conn.GetString() ?? "";
                if (sys.TryGetProperty("subscription", out var sub) &&
                    sub.TryGetProperty("id", out var subEl))
                    subId = subEl.GetString() ?? "";
                return !string.IsNullOrEmpty(connGuid);
            }
        }
        catch { /* malformed body — skip */ }
        return false;
    }

    /// <summary>
    /// Fires a background MPSD PUT to update the player's member connection in the
    /// captured CascadeMatchmaking session, replacing the stale dead RTA connection
    /// with the new live one MCC just established.
    /// </summary>
    private async Task UpdateMatchSessionConnectionAsync(
        SavedHandleInfo session, string connGuid, string subId, string authHeader)
    {
        var url = $"https://sessiondirectory.xboxlive.com/serviceconfigs/{session.Scid}" +
                  $"/sessionTemplates/{session.TemplateName}/sessions/{session.SessionName}";

        var bodyStr =
            "{\"members\":{\"me\":{\"properties\":{\"system\":{" +
            "\"active\":true," +
            $"\"connection\":\"{connGuid}\"," +
            $"\"subscription\":{{\"id\":\"{subId}\",\"changeTypes\":[\"everything\"]}}" +
            "}}}}}";

        int  finalStatus  = 0;
        string finalBody  = "";
        bool succeeded    = false;

        for (int attempt = 1; attempt <= 3 && !succeeded; attempt++)
        {
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Put, url)
                {
                    Content = new StringContent(bodyStr, System.Text.Encoding.UTF8, "application/json"),
                };
                if (!string.IsNullOrEmpty(authHeader))
                    request.Headers.TryAddWithoutValidation("Authorization", authHeader);
                request.Headers.TryAddWithoutValidation("X-Xbl-Contract-Version", "107");
                request.Headers.TryAddWithoutValidation("If-Match", "*");

                var resp  = await _mpsdBackgroundClient.SendAsync(request);
                finalBody = await resp.Content.ReadAsStringAsync();
                finalStatus = (int)resp.StatusCode;

                if (finalStatus < 300)
                    succeeded = true;
                else if (attempt < 3)
                    await Task.Delay(150);
            }
            catch (Exception ex)
            {
                finalBody = $"Exception: {ex.Message}";
                if (attempt < 3)
                    await Task.Delay(150);
            }
        }

        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
        {
            Method       = "REJOIN[ConnectionUpdate]",
            Url          = url,
            Host         = "sessiondirectory.xboxlive.com",
            Path         = new Uri(url).AbsolutePath,
            RequestBody  = $"connection={connGuid[..Math.Min(8, connGuid.Length)]}... sub={subId[..Math.Min(8, subId.Length)]}...",
            StatusCode   = finalStatus,
            ResponseBody = succeeded
                ? $"Connection updated ({finalStatus}) — CascadeMatchmaking membership now has live RTA GUID"
                : $"Update failed after 3 attempts ({finalStatus}): {finalBody[..Math.Min(300, finalBody.Length)]}",
        });
    }

    // ── Party helpers ─────────────────────────────────────────────────────────

    /// <summary>
    /// Extracts the PartyId from a PlayFab RequestParty response body (data.PartyId).
    /// Returns empty string if not found or on parse failure.
    /// </summary>
    private static string TryExtractPartyId(string partyJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(partyJson);
            if (doc.RootElement.TryGetProperty("data", out var data) &&
                data.TryGetProperty("PartyId", out var partyId))
                return partyId.GetString() ?? "";
        }
        catch { }
        return "";
    }

    /// <summary>
    /// Extracts IPV4Address from a PlayFab RequestParty response body.
    /// Returns empty string if not found or on parse failure.
    /// </summary>
    private static string TryExtractPartyIp(string partyJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(partyJson);
            if (doc.RootElement.TryGetProperty("data", out var data) &&
                data.TryGetProperty("IPV4Address", out var ip))
                return ip.GetString() ?? "";
        }
        catch { /* malformed — skip */ }
        return "";
    }

    /// <summary>
    /// Returns true if the party response is stale and should not be cached.
    /// Stale = LastStateTransitionTime is >1 hour ago.
    /// NOTE: ConnectedPlayers is always empty in MCC captures regardless of match state —
    /// it is NOT a reliable staleness indicator and must not be used.
    /// </summary>
    private static bool IsStaleParty(string partyJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(partyJson);
            if (!doc.RootElement.TryGetProperty("data", out var data)) return false;

            // Only check LastStateTransitionTime — ConnectedPlayers is always [] in MCC
            if (data.TryGetProperty("LastStateTransitionTime", out var lstt) &&
                DateTime.TryParse(lstt.GetString(), null,
                    System.Globalization.DateTimeStyles.RoundtripKind, out var transitionTime))
            {
                return (DateTime.UtcNow - transitionTime) > TimeSpan.FromHours(1);
            }
        }
        catch { /* malformed — treat as non-stale to be safe */ }
        return false;
    }

    // ── Party cache priming ───────────────────────────────────────────────────

    /// <summary>
    /// Makes a background RequestParty call using MCC's captured entity token
    /// to prime the party cache before any crash occurs.
    /// Called automatically 5s after a new session is captured, and available
    /// as a public method for the UI "Prime Party Cache" button.
    /// </summary>
    public async Task PrimeCacheAsync(bool auto = false)
    {
        if (string.IsNullOrEmpty(_capturedEntityToken) || string.IsNullOrEmpty(_capturedPlayFabHost))
        {
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = auto ? "PARTY[PrimeSkipped]" : "PARTY[PrimeFailed]",
                Url          = "",
                Host         = "local",
                Path         = "",
                StatusCode   = 0,
                ResponseBody = "No entity token captured yet — cannot prime party cache",
            });
            return;
        }

        if (string.IsNullOrEmpty(_capturedPartyRequestBody))
        {
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = auto ? "PARTY[PrimeSkipped]" : "PARTY[PrimeFailed]",
                Url          = "",
                Host         = "local",
                Path         = "",
                StatusCode   = 0,
                ResponseBody = "No RequestParty body captured yet — MCC hasn't called RequestParty this session. Play until matchmaking then retry.",
            });
            return;
        }

        SetPartyState(PartyCacheState.Priming);

        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Post,
                $"https://{_capturedPlayFabHost}/Party/RequestParty");
            req.Headers.Add("X-EntityToken", _capturedEntityToken);
            req.Headers.Add("Accept", "application/json");
            // Use MCC's own request body verbatim — PlayFab requires Version/Build fields
            // that an empty {} body omits, causing a 400 BadRequest.
            // Also swap PartyId to the cached original so PlayFab routes to the real match party.
            var primeBody = string.IsNullOrEmpty(_capturedPartyRequestBody) ? "{}" : _capturedPartyRequestBody;
            if (!string.IsNullOrEmpty(_cachedPartyResponse))
            {
                var cachedPartyId = TryExtractPartyId(_cachedPartyResponse);
                if (!string.IsNullOrEmpty(cachedPartyId))
                {
                    try
                    {
                        var bodyNode = JsonNode.Parse(primeBody)?.AsObject();
                        if (bodyNode is not null)
                        {
                            bodyNode["PartyId"] = cachedPartyId;
                            primeBody = bodyNode.ToJsonString();
                        }
                    }
                    catch { /* best-effort */ }
                }
            }
            req.Content = new StringContent(primeBody, System.Text.Encoding.UTF8, "application/json");

            var resp = await _mpsdBackgroundClient.SendAsync(req);
            var body = await resp.Content.ReadAsStringAsync();

            if ((int)resp.StatusCode == 200 && !string.IsNullOrEmpty(body))
            {
                if (IsStaleParty(body))
                {
                    SetPartyState(PartyCacheState.Stale);
                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method       = "PARTY[PrimeStale]",
                        Url          = $"https://{_capturedPlayFabHost}/Party/RequestParty",
                        Host         = _capturedPlayFabHost,
                        Path         = "/Party/RequestParty",
                        StatusCode   = 200,
                        ResponseBody = $"Background RequestParty returned stale party — not caching: {body[..Math.Min(120, body.Length)]}",
                    });
                }
                else
                {
                    _cachedPartyResponse = body;
                    _cachedPartyIp       = TryExtractPartyIp(body);
                    try { File.WriteAllText(PartyResponsePath, body, System.Text.Encoding.UTF8); }
                    catch { /* best-effort */ }
                    SetPartyState(PartyCacheState.Ready);
                    OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                    {
                        Method       = auto ? "PARTY[AutoPrimed]" : "PARTY[ManualPrimed]",
                        Url          = $"https://{_capturedPlayFabHost}/Party/RequestParty",
                        Host         = _capturedPlayFabHost,
                        Path         = "/Party/RequestParty",
                        StatusCode   = 200,
                        ResponseBody = $"Party cache primed — crash recovery ready: {body[..Math.Min(120, body.Length)]}",
                    });
                }
            }
            else
            {
                SetPartyState(PartyCacheState.None);
                OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                {
                    Method       = "PARTY[PrimeFailed]",
                    Url          = $"https://{_capturedPlayFabHost}/Party/RequestParty",
                    Host         = _capturedPlayFabHost,
                    Path         = "/Party/RequestParty",
                    StatusCode   = (int)resp.StatusCode,
                    ResponseBody = $"Background RequestParty failed ({(int)resp.StatusCode}): {body[..Math.Min(200, body.Length)]}",
                });
            }
        }
        catch (Exception ex)
        {
            SetPartyState(PartyCacheState.None);
            OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
            {
                Method       = "PARTY[PrimeFailed]",
                Url          = $"https://{_capturedPlayFabHost}/Party/RequestParty",
                Host         = _capturedPlayFabHost,
                Path         = "/Party/RequestParty",
                StatusCode   = 0,
                ResponseBody = $"Background RequestParty exception: {ex.Message}",
            });
        }
    }

    private void SetPartyState(PartyCacheState state)
    {
        _partyCacheState = state;
        if (state == PartyCacheState.Ready)
            _partyCachedAt = DateTime.Now;
        else if (state == PartyCacheState.None)
            _partyCachedAt = null;
        // Priming and Stale leave the timestamp alone so the last-good time stays visible.
        OnPartyCacheChanged?.Invoke(this, EventArgs.Empty);
    }

    // ── Party persistence ─────────────────────────────────────────────────────

    /// <summary>
    /// Loads the persisted RequestParty response from disk into memory.
    /// Called at proxy start so crash recovery can redirect even if RequestParty
    /// wasn't observed in the current proxy session (e.g. fired during matchmaking
    /// before the proxy captured it, or in a previous MCC session).
    /// </summary>
    private void TryLoadPartyResponseFromDisk()
    {
        // Load persisted party response.
        try
        {
            if (File.Exists(PartyResponsePath))
            {
                var saved = File.ReadAllText(PartyResponsePath, System.Text.Encoding.UTF8);
                if (!string.IsNullOrWhiteSpace(saved))
                {
                    if (IsStaleParty(saved))
                    {
                        Debug.WriteLine("[PARTY] Disk party is stale (LastStateTransitionTime >1h old) — discarding");
                        try { File.Delete(PartyResponsePath); } catch { /* best-effort */ }
                    }
                    else
                    {
                        _cachedPartyResponse = saved;
                        _cachedPartyIp       = TryExtractPartyIp(saved);
                        _partyCacheState     = PartyCacheState.Ready;
                        _partyCachedAt       = File.GetLastWriteTime(PartyResponsePath);
                        Debug.WriteLine($"[PARTY] Loaded saved party response from disk ({saved.Length} chars)");
                        OnRequestCaptured?.Invoke(this, new ProxyCaptureEntry
                        {
                            Method       = "PARTY[LoadedFromDisk]",
                            Url          = PartyResponsePath,
                            Host         = "local",
                            Path         = PartyResponsePath,
                            StatusCode   = 0,
                            ResponseBody = $"Party response loaded from disk ({saved.Length} chars) — ready for crash recovery redirect",
                        });
                    }
                }
            }
        }
        catch { /* best-effort */ }

        // Load persisted RequestParty request body.
        // Once captured from any MCC session, this body works forever across restarts
        // because the Version/Build fields don't change between matches (only game updates).
        try
        {
            if (File.Exists(PartyRequestBodyPath))
            {
                var body = File.ReadAllText(PartyRequestBodyPath, System.Text.Encoding.UTF8);
                if (!string.IsNullOrWhiteSpace(body))
                {
                    _capturedPartyRequestBody = body;
                    Debug.WriteLine($"[PARTY] Loaded saved RequestParty request body from disk ({body.Length} chars)");
                }
            }
        }
        catch { /* best-effort */ }

        // Option B: if we still have no request body, try to construct one from the
        // BuildId embedded in the cached party response. PlayFab accepts {"Build":"<id>"}
        // as a valid RequestParty body when the title uses build-based party routing.
        if (string.IsNullOrEmpty(_capturedPartyRequestBody) && !string.IsNullOrEmpty(_cachedPartyResponse))
        {
            var buildId = TryExtractBuildId(_cachedPartyResponse);
            if (!string.IsNullOrEmpty(buildId))
            {
                _capturedPartyRequestBody = $"{{\"Build\":\"{buildId}\"}}";
                Debug.WriteLine($"[PARTY] Constructed RequestParty body from cached BuildId: {buildId}");
            }
        }
    }

    /// <summary>Extracts the BuildId field from a PlayFab RequestParty response body.</summary>
    private static string TryExtractBuildId(string partyJson)
    {
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(partyJson);
            if (doc.RootElement.TryGetProperty("data", out var data) &&
                data.TryGetProperty("BuildId", out var buildId))
                return buildId.GetString() ?? "";
        }
        catch { }
        return "";
    }

    // ── Member injection helper ───────────────────────────────────────────────

    /// <summary>
    /// Injects the player's XUID into a CascadeMatchmaking GET response body
    /// when they have been evicted (MPSD inactiveRemovalTimeout=0 + dead RTA).
    /// Adds a minimal active member entry at the next slot index so MCC sees
    /// the player and shows the rejoin button. Returns null on parse failure.
    /// </summary>
    private static string? InjectPlayerMember(string sessionJson, string playerXuid)
    {
        try
        {
            var root = JsonNode.Parse(sessionJson)?.AsObject();
            if (root is null) return null;

            var members = root["members"]?.AsObject();
            if (members is null) return null;

            // Determine the next available slot index.
            var membersInfo = root["membersInfo"]?.AsObject();
            int nextIndex   = 0;
            try   { nextIndex = membersInfo?["next"]?.GetValue<int>() ?? 0; }
            catch { /* fall through to scan */ }

            if (nextIndex == 0)
            {
                foreach (var kvp in members)
                    if (int.TryParse(kvp.Key, out int idx) && idx >= nextIndex)
                        nextIndex = idx + 1;
            }

            // Minimal member entry — just enough for MCC to recognise the player.
            // Connection GUID will be updated by UpdateMatchSessionConnectionAsync.
            members[nextIndex.ToString()] = new JsonObject
            {
                ["xuid"]       = playerXuid,
                ["properties"] = new JsonObject
                {
                    ["system"] = new JsonObject { ["active"] = true }
                },
            };

            if (membersInfo is not null)
            {
                try { membersInfo["count"]    = membersInfo["count"]?.GetValue<int>()    + 1 ?? 1; } catch { }
                try { membersInfo["next"]     = membersInfo["next"]?.GetValue<int>()     + 1 ?? nextIndex + 1; } catch { }
                try { membersInfo["accepted"] = membersInfo["accepted"]?.GetValue<int>() + 1 ?? 1; } catch { }
                try { membersInfo["active"]   = membersInfo["active"]?.GetValue<int>()   + 1 ?? 1; } catch { }
            }

            return root.ToJsonString();
        }
        catch { return null; }
    }

    // ── URL helpers ───────────────────────────────────────────────────────────

    private static bool ShouldCaptureRequestBody(Titanium.Web.Proxy.Http.Request req)
    {
        string path = req.RequestUri.AbsolutePath;
        // /handles — rejoin handle POSTs
        if (path.Contains("/handles", StringComparison.OrdinalIgnoreCase)) return true;
        // PUT to any session URL — includes /members/me (needed for leave detection)
        if (req.Method == "PUT" && path.Contains("/sessions/", StringComparison.OrdinalIgnoreCase)) return true;
        // PlayFab /Party/RequestParty POST — capture body so we can replay it in background prime calls.
        // Without this, _capturedPartyRequestBody is never populated and all prime attempts fail with 400.
        if (req.Method == "POST" && path.Contains("/Party/RequestParty", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    /// <summary>Returns true if the body is a MPSD leave request: {"members":{"me":null}}.</summary>
    private static bool IsLeaveBody(string body) =>
        body.Contains("\"me\":null",  StringComparison.OrdinalIgnoreCase) ||
        body.Contains("\"me\": null", StringComparison.OrdinalIgnoreCase);

    private static bool ShouldReadBody(string url)
    {
        foreach (var b in _bypassHosts)
            if (url.Contains(b, StringComparison.OrdinalIgnoreCase)) return false;

        if (url.Contains("halowaypoint.com", StringComparison.OrdinalIgnoreCase)) return true;
        if (url.Contains("playfabapi.com",   StringComparison.OrdinalIgnoreCase)) return true;

        if (url.Contains("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
            url.Contains("/handles", StringComparison.OrdinalIgnoreCase))
            return true;

        if (url.Contains("/serviceconfigs/", StringComparison.OrdinalIgnoreCase) &&
            url.Contains("/sessions/",       StringComparison.OrdinalIgnoreCase))
            return !url.Contains("changeNumber=", StringComparison.OrdinalIgnoreCase);

        return false;
    }

    private static bool IsJsonContent(string? contentType) =>
        !string.IsNullOrEmpty(contentType) &&
        (contentType.Contains("json", StringComparison.OrdinalIgnoreCase) ||
         contentType.Contains("text", StringComparison.OrdinalIgnoreCase) ||
         contentType.Contains("xml",  StringComparison.OrdinalIgnoreCase));

    private static bool IsDomainWatched(string host)
    {
        foreach (var d in _watchedDomains)
            if (host.EndsWith(d, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

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

    // ── WinINet (no admin required) ───────────────────────────────────────────
    private void SetWinINetProxy(string proxyAddress)
    {
        const string key = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";
        using var reg = Registry.CurrentUser.OpenSubKey(key, writable: true);
        if (reg is null) return;

        _savedProxyEnable   = (int)(reg.GetValue("ProxyEnable")   ?? 0);
        _savedProxyServer   = (string)(reg.GetValue("ProxyServer") ?? "");
        _savedProxyOverride = (string)(reg.GetValue("ProxyOverride") ?? "");

        reg.SetValue("ProxyEnable",   1,                                 RegistryValueKind.DWord);
        reg.SetValue("ProxyServer",   proxyAddress,                      RegistryValueKind.String);
        reg.SetValue("ProxyOverride", "localhost;127.0.0.1;<local>",     RegistryValueKind.String);

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
                Verb            = "runas",
                UseShellExecute = true,
                CreateNoWindow  = true,
                WindowStyle     = ProcessWindowStyle.Hidden,
            };
            var p = Process.Start(psi);
            await Task.Run(() => p?.WaitForExit(8000));
        }
        catch
        {
            WinHttpManualSetRequired?.Invoke(this, "netsh winhttp import proxy source=ie");
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

    // ── WinINet P/Invoke ──────────────────────────────────────────────────────
    [DllImport("wininet.dll", SetLastError = true)]
    private static extern bool InternetSetOption(
        IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);

    private const int INTERNET_OPTION_REFRESH          = 37;
    private const int INTERNET_OPTION_SETTINGS_CHANGED = 39;

    // ── IDisposable ───────────────────────────────────────────────────────────
    public void Dispose() => Stop();
}
