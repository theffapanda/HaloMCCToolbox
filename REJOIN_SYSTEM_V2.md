# Rejoin System V2: Complete Technical Documentation
## HaloMCCToolbox Crash Recovery & Session Rejoin

**Status: WORKING (Solo)** | **Date: 2026-03-17** | **Build: net8.0-windows**

---

## Table of Contents

1. [System Overview](#1-system-overview)
2. [The Problem](#2-the-problem)
3. [Architecture](#3-architecture)
4. [The Recovery Timeline](#4-the-recovery-timeline)
5. [Component Deep Dive](#5-component-deep-dive)
   - 5.1 [MCC Watcher (Crash Detection)](#51-mcc-watcher-crash-detection)
   - 5.2 [Block Match Leave](#52-block-match-leave)
   - 5.3 [Session Persistence](#53-session-persistence)
   - 5.4 [Ghost Session Mode](#54-ghost-session-mode)
   - 5.5 [AutoSync One-Shot](#55-autosync-one-shot)
   - 5.6 [GUID Upgrade (FIX #16)](#56-guid-upgrade-fix-16)
   - 5.7 [Session Discovery Injection](#57-session-discovery-injection)
   - 5.8 [Squad Handle Override](#58-squad-handle-override)
   - 5.9 [Game Server Redirect](#59-game-server-redirect)
   - 5.10 [ETag Refresh](#510-etag-refresh)
6. [MPSD Session Constants](#6-mpsd-session-constants)
7. [Critical Bug History](#7-critical-bug-history)
8. [File Reference](#8-file-reference)
9. [Capture Analysis: The Working Run](#9-capture-analysis-the-working-run)

---

## 1. System Overview

The Rejoin System allows a player to re-enter a Halo MCC match after an unexpected exit (crash, kill, disconnect). It works by intercepting Xbox MPSD (Multiplayer Session Directory) traffic through an HTTP proxy, preserving session state across MCC restarts, and faking/injecting responses so MCC believes it never left the match.

The system comprises **ten coordinated interventions** across two source files:

| # | Component | File | Purpose |
|---|-----------|------|---------|
| 1 | MCC Watcher | `Scanner.xaml.cs` | Detects MCC crash via process monitoring |
| 2 | Block Match Leave | `ProxyService.cs` | Intercepts leave requests, fakes success |
| 3 | Session Persistence | Both | Saves session/handle/server to disk |
| 4 | Ghost Session Mode | `ProxyService.cs` | Fakes MPSD responses during recovery |
| 5 | AutoSync One-Shot | `ProxyService.cs` | Re-adds player to MPSD at +2s |
| 6 | GUID Upgrade | `ProxyService.cs` | Replaces dead connection GUID with live one |
| 7 | Session Discovery Injection | `ProxyService.cs` | Fakes session list so MCC sees the match |
| 8 | Squad Handle Override | `ProxyService.cs` | Points activity handle to saved squad |
| 9 | Game Server Redirect | `ProxyService.cs` | Sends MCC to the correct game server |
| 10 | ETag Refresh | `ProxyService.cs` | Fixes 412 conflicts from stale ETags |

---

## 2. The Problem

### Why MCC Can't Rejoin On Its Own

When MCC exits (crash or close), several things happen that destroy the session:

1. **Immediate Leave Request**: MCC sends `PUT {"members":{"me":null}}` to the CascadeMatchmaking session, explicitly removing itself
2. **RTA WebSocket Dies**: The Real-Time Activity WebSocket to `rta.xboxlive.com` disconnects, and MPSD's heartbeat will detect this within ~30 seconds
3. **No Session Memory**: MCC starts fresh on restart with no knowledge of the previous match
4. **New Connection GUID**: MCC generates a new RTA WebSocket connection, getting a new GUID that doesn't match the saved session membership
5. **`inactiveRemovalTimeout: 0`**: CascadeMatchmaking sessions have ZERO tolerance for inactive members. The instant MPSD detects a dead connection and any request touches the session, the member is removed permanently.

### The Race Against Time

```
+0s   MCC crashes. RTA WebSocket dies.
+0s   MCC sends leave request (blocked by us).
+2s   Our AutoSync PUT re-confirms player as active (saved GUID).
+17s  MCC restarts. New RTA WebSocket established (new GUID).
+17s  GUID Upgrade fires: PUT new GUID to CascadeMatchmaking.
+30s  MPSD heartbeat would detect dead WebSocket...
      BUT new GUID is live. No timeout fires.
+33s  User clicks rejoin. Ghost mode serves fake session data.
+34s  RequestParty returns game server. Player connects.
      Session membership preserved. Rejoin succeeds.
```

Without the GUID Upgrade (FIX #16), the old GUID's dead WebSocket is detected at ~30s, and `inactiveRemovalTimeout: 0` removes the player on the next session access.

---

## 3. Architecture

```
                    MCC (Halo MCC)
                         |
                    [HTTP Proxy]
                    Port 8888
                         |
              +----------+----------+
              |                     |
    sessiondirectory       ee38.playfabapi.com
     .xboxlive.com          (PlayFab APIs)
              |                     |
         MPSD Sessions        RequestParty
         (PUT/GET/POST)       (Game Server)
```

**Key Files:**
- `ProxyService.cs` (~2330 lines) - All proxy interception logic
- `Scanner.xaml.cs` (~1405 lines) - UI, crash detection, session loading
- `Models.cs` - Data models (`SavedHandleInfo`, `GameServerInfo`, etc.)

**Proxy Library:** Unobtanium.Web.Proxy v0.1.5 (fork of Titanium.Web.Proxy)

**Persisted Files** (in `%LOCALAPPDATA%\HaloMCCToolbox\`):
- `last-handle.json` - Squad session activity handle
- `last-match-session.json` - CascadeMatchmaking session info + connection GUID
- `last-game-server.json` - Game server IP/port from RequestParty

---

## 4. The Recovery Timeline

This is the exact sequence from the working capture (`HaloCapture-20260317-081807CLOSE`):

```
TIME        EVENT                         DETAIL
08:16:14    SAVE[Match]                   Match session persisted to disk during normal gameplay
08:17:02    RESTORE[Start]                MCC crash detected by watcher (PID gone)
08:17:02    AUTO-BLOCK[Armed]             Block-match-leave armed for next leave request
08:17:02    SetPendingCrashRestore()      Ghost mode enabled, AutoSync scheduled
08:17:04    GHOST[AutoSync-OneShot]       GET+PUT to MPSD with SAVED GUID (200 OK)
08:17:19    MCC restarts                  New squad PUT captured with new connection GUID
08:17:19    UPGRADE[GUID-Success]         PUT new live GUID to CascadeMatchmaking session
08:17:19    GHOST[Match-Get]              Fake match document served to MCC
08:17:19    GHOST[Squad-Get]              Fake squad document served to MCC
08:17:22    OVERRIDE[SquadHandle]         Handle POST redirected to saved squad UUID
08:17:33    MCC-REJOIN-CLICK              User clicks rejoin prompt in MCC
08:17:34    RequestParty                  Game server obtained
08:17:34+   OVERRIDE[SquadHandle] x9      Continued handle overrides during rejoin
08:18:03    GHOST[Match-Get]              Continued ghost responses
08:18:05    Session PUTs                  MCC establishing new squad for the match
```

---

## 5. Component Deep Dive

### 5.1 MCC Watcher (Crash Detection)

**File:** `Scanner.xaml.cs`, lines 1313-1402
**Interval:** 1 second (changed from 5s to catch fast restarts)

The watcher polls for the MCC process and detects two scenarios:
1. **Normal exit**: Process disappears between ticks
2. **Fast restart**: Process PID changes while appearing continuously "running"

```csharp
// Scanner.xaml.cs - MccWatcher_Tick
private void MccWatcher_Tick(object? sender, EventArgs e)
{
    bool running = Process.GetProcessesByName("MCC-Win64-Shipping").Length > 0
                || Process.GetProcessesByName("MCC-Win64-Shipping-EAC").Length > 0;

    // CRITICAL: Track PID to detect fast restarts within the 1s poll window
    int currentMccPid = 0;
    if (running)
    {
        var mccProcess = Process.GetProcessesByName("MCC-Win64-Shipping")
            .Concat(Process.GetProcessesByName("MCC-Win64-Shipping-EAC"))
            .FirstOrDefault();
        if (mccProcess is not null)
            currentMccPid = mccProcess.Id;
    }

    // If PID changed while MCC appeared "running", it's a fast restart
    if (running && _mccWasRunning && _lastMccPid > 0
        && currentMccPid != _lastMccPid && !_restoreInProgress)
    {
        _mccWasRunning = false;  // Force the exit handler path
    }

    // MCC just exited - arm all recovery systems
    if (_mccWasRunning && !running && !_restoreInProgress &&
        (_savedHandle is not null || _savedMatchSession is not null))
    {
        _proxy.ForceBlockMatchLeave();           // Arm leave blocker
        if (_savedMatchSession is not null)
            _proxy.SetPendingCrashRestore(_savedMatchSession);  // Arm ghost mode + AutoSync
        if (_savedHandle is not null)
            _proxy.SetSavedSquadHandle(_savedHandle);           // Arm handle override
    }

    _mccWasRunning = running;
    if (running) _lastMccPid = currentMccPid;
}
```

**Why 1 second?** At 5 seconds, if the user quit and relaunched MCC in <5s, the watcher never saw `running=false`. The `_lastMccPid` check is defense-in-depth: even if the watcher sees MCC as continuously "running," a PID change reveals the restart.

---

### 5.2 Block Match Leave

**File:** `ProxyService.cs`, lines ~455-490

When MCC exits, it sends `PUT {"members":{"me":null}}` to the CascadeMatchmaking session to remove itself. This is intercepted and short-circuited with a fake 200 OK.

```csharp
// ProxyService.cs - OnBeforeRequest
if (_blockMatchLeave && _lastMatchSession is not null &&
    req.Method == "PUT" &&
    req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", ...) &&
    req.RequestUri.AbsolutePath.Contains(
        _lastMatchSession.SessionName, StringComparison.OrdinalIgnoreCase))
{
    // Check if the body is a leave request: {"members":{"me":null}}
    if (!string.IsNullOrEmpty(entry.RequestBody) &&
        entry.RequestBody.Contains("\"me\":null"))
    {
        e.Ok("{}");  // Fake success - leave NEVER reaches MPSD
        _blockMatchLeave = false;  // One-shot: only block the first leave
        return;
    }
}
```

**One-shot design:** The block fires once and disarms. This prevents blocking legitimate leave requests later (e.g., when the user actually wants to quit the match after a successful rejoin).

---

### 5.3 Session Persistence

**File:** `ProxyService.cs` (PersistHandleToDisk, PersistMatchSessionToDisk, PersistGameServerToDisk)
**File:** `Scanner.xaml.cs` (LoadSavedHandle, LoadSavedMatchSession)

Three types of data are persisted to `%LOCALAPPDATA%\HaloMCCToolbox\`:

#### Match Session (`last-match-session.json`)
Captured from CascadeMatchmaking PUT responses. Contains:
- `scid`, `templateName`, `sessionName` - session coordinates
- `connectionGuid` - the player's RTA connection GUID (critical for AutoSync)
- `subscriptionId` - RTA subscription ID
- `requestHeaders` - auth headers for background requests

**Guard:** During crash recovery (`_pendingCrashRestore`), persistence is blocked to prevent MCC's new session from overwriting the saved one:
```csharp
// ProxyService.cs - PersistMatchSessionToDisk
if (_pendingCrashRestore)
{
    Debug.WriteLine("[SAVE-Match] Skipping persistence during crash restore");
    return;
}
```

#### Squad Handle (`last-handle.json`)
Captured from `/handles` POST requests. Contains the cascadesquadsession reference that links the player's activity to the match.

**Guard:** Same `_pendingCrashRestore` guard prevents overwrite during recovery.

#### Game Server (`last-game-server.json`)
Captured from PlayFab `RequestParty` responses. Contains:
- `ipv4Address`, `ports` - actual game server endpoint
- `serverId`, `vmId`, `region` - PlayFab allocation identifiers

---

### 5.4 Ghost Session Mode

**File:** `ProxyService.cs`, `TryHandleGhostMode` method

When crash recovery is active, ALL requests to the saved CascadeMatchmaking and cascadesquadsession are intercepted BEFORE they reach MPSD. The proxy returns fake session documents that make MCC believe:
- The match session exists and is active
- The player is a member of the session
- The session has valid properties

```csharp
// ProxyService.cs - TryHandleGhostMode (simplified)
private bool TryHandleGhostMode(SessionEventArgs e, HttpWebClient.Request req, ...)
{
    if (!_ghostSessionMode || _ghostSession is null) return false;

    string url = req.Url.ToLowerInvariant();
    string method = req.Method;

    // Match session GET -> return fake session document
    if (url.Contains(_ghostSession.SessionName.ToLowerInvariant()) &&
        url.Contains("cascadematchmaking"))
    {
        if (method == "GET" && !url.Contains("/members/"))
        {
            string fakeSessionBody = GenerateFakeSessionDocument();
            e.Ok(fakeSessionBody);

            // CRITICAL: Cache the fake body so game server redirect condition is met
            _cachedInjectedMatchBody = fakeSessionBody;
            _cachedInjectedMatchEtag = $"\"{Guid.NewGuid():N}\"";
            return true;
        }
    }

    // Squad session GET -> return fake squad document
    if (url.Contains("cascadesquadsession"))
    {
        if (method == "GET")
        {
            string fakeSquadBody = GenerateFakeSquadDocument();
            e.Ok(fakeSquadBody);
            return true;
        }
    }

    return false;
}
```

#### Fake Session Document Generation

The fake document includes the player as an active member with the **saved connection GUID**:

```csharp
// ProxyService.cs - GenerateFakeSessionDocument
private string GenerateFakeSessionDocument()
{
    string connectionGuid = string.IsNullOrEmpty(_ghostSession.ConnectionGuid)
        ? Guid.NewGuid().ToString()
        : _ghostSession.ConnectionGuid;  // SAVED GUID from before crash

    return $$"""
    {
      "membersInfo": { "first": 0, "next": 1, "count": 1, "accepted": 1 },
      "constants": { ... },
      "properties": {
        "system": {
          "matchmaking": { "serverConnectionString": "west us" }
        }
      },
      "members": {
        "0": {
          "constants": { "system": { "xuid": "{{_playerXuid}}" } },
          "properties": {
            "system": {
              "active": true,
              "connection": "{{connectionGuid}}",
              "subscription": {
                "id": "{{subscriptionId}}",
                "changeTypes": ["everything"]
              }
            }
          }
        }
      }
    }
    """;
}
```

**Why ghost mode stays ON permanently:** Ghost mode is never disabled during the recovery window. If it were turned off after AutoSync, MCC's next real MPSD request could trigger `inactiveRemovalTimeout: 0` and remove the player. Ghost mode ensures NO real MPSD access occurs from MCC during recovery.

---

### 5.5 AutoSync One-Shot

**File:** `ProxyService.cs`, lines 2005-2165, method `AutoSyncGhostSessionAsync`

Fires once at +2 seconds after crash detection. Sends a GET+PUT directly to MPSD (bypassing the proxy) to re-confirm the player as an active session member.

```csharp
// ProxyService.cs - AutoSyncGhostSessionAsync
private async Task AutoSyncGhostSessionAsync()
{
    if (_ghostSession is null) return;

    await Task.Delay(2000);  // Wait 2s for MCC to settle

    var authHeaders = _lastCapturedAuthHeaders ?? _ghostSession.RequestHeaders;

    // GET: Verify session is alive and get current ETag
    using var getReq = new HttpRequestMessage(HttpMethod.Get, _ghostSession.SessionUrl);
    foreach (var (k, v) in authHeaders)
        if ((k.StartsWith("x-", ...) || k.Equals("Authorization", ...)) &&
            !k.Equals("Signature", ...))  // Skip Signature - per-request crypto, can't replay
            getReq.Headers.TryAddWithoutValidation(k, v);

    using var getResp = await new HttpClient(
        new HttpClientHandler { UseProxy = false }).SendAsync(getReq);

    if (!getResp.IsSuccessStatusCode) return;  // Session dead
    string etag = getResp.Headers.ETag?.Tag ?? "";

    // PUT: Re-add player with SAVED connection GUID
    string connectionGuid = _ghostSession.ConnectionGuid ?? "";
    string connectionField = string.IsNullOrEmpty(connectionGuid)
        ? "" : $",\"connection\":\"{connectionGuid}\"";

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
    putReq.Content = new StringContent(putBody, Encoding.UTF8, "application/json");
    putReq.Headers.TryAddWithoutValidation("If-Match", etag);
    // ... same auth headers, skip Signature ...

    using var putResp = await new HttpClient(
        new HttpClientHandler { UseProxy = false }).SendAsync(putReq);
    // Log success/failure
}
```

**Why one-shot, not a loop?** A loop caused the very problem it was trying to prevent. Each PUT to MPSD is a "read or write request" that triggers `inactiveRemovalTimeout: 0` processing. At ~30s when the dead WebSocket is detected, the loop's own PUT triggered the timeout and removed the player. One-shot at +2s succeeds because the WebSocket is still within the heartbeat window, then stops accessing MPSD entirely.

**Why `UseProxy = false`?** The PUT goes directly to MPSD, bypassing our own proxy. If it went through the proxy, ghost mode would intercept it and return a fake response instead of actually reaching MPSD.

**Why skip `Signature` header?** The `Signature` header is a per-request cryptographic proof. It can't be replayed from a previous request. MPSD accepts requests without it (the `Authorization` header provides authentication).

---

### 5.6 GUID Upgrade (FIX #16)

**File:** `ProxyService.cs`, lines ~537-570 (trigger), lines ~2167-2270 (method)

**This is the critical fix that made solo rejoin work.**

#### The Problem

The AutoSync one-shot at +2s uses the **saved** connection GUID (from before the crash). This GUID points to a dead RTA WebSocket. At ~30s, MPSD's heartbeat detects the dead WebSocket and marks the member inactive. The next session access triggers `inactiveRemovalTimeout: 0` and permanently removes the player.

#### The Solution

When MCC restarts, it establishes a **new RTA WebSocket** and gets a new connection GUID. This new GUID appears in MCC's first `cascadesquadsession` PUT (creating its new squad session). We capture this GUID and immediately fire a PUT to the CascadeMatchmaking session to replace the dead GUID with the live one.

#### Trigger (in OnBeforeRequest, GUID capture block)

```csharp
// ProxyService.cs - FIX #16 trigger, inside the GUID capture block
// Fires when: crash recovery is active, new GUID captured from squad PUT,
// and it differs from the saved GUID
if (_pendingCrashRestore && !_guidUpgradeDone &&
    !string.IsNullOrEmpty(_capturedConnectionGuid) &&
    _ghostSession is not null &&
    req.RequestUri.AbsolutePath.Contains(
        "/cascadesquadsession/", StringComparison.OrdinalIgnoreCase))
{
    string savedGuid = _ghostSession.ConnectionGuid ?? "";
    if (!string.Equals(_capturedConnectionGuid, savedGuid,
        StringComparison.OrdinalIgnoreCase))
    {
        _guidUpgradeDone = true;  // One-shot

        // Capture fresh auth headers from THIS request (MCC's new valid auth)
        var upgradeHeaders = new Dictionary<string, string>(...);
        foreach (var hdr in req.Headers)
        {
            if ((hdr.Name.StartsWith("x-", ...) ||
                 hdr.Name.Equals("Authorization", ...)) &&
                !hdr.Name.Equals("Signature", ...))
                upgradeHeaders[hdr.Name] = hdr.Value;
        }

        string newGuid = _capturedConnectionGuid;
        string newSubId = _capturedSubscriptionId;
        string sessionUrl = _ghostSession.SessionUrl;

        _ = Task.Run(() => GuidUpgradeAsync(
            sessionUrl, newGuid, newSubId, upgradeHeaders));
    }
}
```

#### The Upgrade PUT

```csharp
// ProxyService.cs - GuidUpgradeAsync
private async Task GuidUpgradeAsync(string sessionUrl, string newGuid,
    string newSubId, Dictionary<string, string> authHeaders)
{
    // GET current ETag
    using var getReq = new HttpRequestMessage(HttpMethod.Get, sessionUrl);
    foreach (var (k, v) in authHeaders)
        getReq.Headers.TryAddWithoutValidation(k, v);

    using var getResp = await new HttpClient(
        new HttpClientHandler { UseProxy = false }).SendAsync(getReq);
    if (!getResp.IsSuccessStatusCode) return;

    string etag = getResp.Headers.ETag?.Tag ?? "";

    // Build PUT with NEW live GUID + subscription
    string subscriptionBlock = string.IsNullOrEmpty(newSubId)
        ? ""
        : $",\"subscription\":{{\"id\":\"{newSubId}\"," +
          $"\"changeTypes\":[\"everything\"]}}";

    var putBody = $$"""
    {
      "members": {
        "me": {
          "properties": {
            "system": {
              "active": true,
              "connection": "{{newGuid}}"{{subscriptionBlock}}
            },
            "custom": {}
          }
        }
      }
    }
    """;

    using var putReq = new HttpRequestMessage(HttpMethod.Put, sessionUrl);
    putReq.Content = new StringContent(putBody, Encoding.UTF8, "application/json");
    putReq.Headers.TryAddWithoutValidation("If-Match", etag);
    // ... auth headers ...

    using var putResp = await new HttpClient(
        new HttpClientHandler { UseProxy = false }).SendAsync(putReq);
    // Log UPGRADE[GUID-Success] or UPGRADE[GUID-PUT-Fail]
}
```

#### Why This Works

```
BEFORE FIX #16:
+0s   Crash. Old GUID (54c52880...) -> dead WebSocket
+2s   AutoSync PUT with old GUID. Player active BUT connection dead.
+30s  Heartbeat detects dead WebSocket for GUID 54c52880.
+30s  Member marked inactive. inactiveRemovalTimeout=0 arms.
+34s  Any session access -> player REMOVED permanently.

AFTER FIX #16:
+0s   Crash. Old GUID (54c52880...) -> dead WebSocket
+2s   AutoSync PUT with old GUID. Player active, connection dead.
+17s  MCC restarts. New GUID (85903ee9...) -> LIVE WebSocket.
+17s  UPGRADE PUT: Replace 54c52880 with 85903ee9 on CascadeMatchmaking.
+30s  Heartbeat checks GUID 85903ee9 -> WebSocket is ALIVE.
+30s  No inactive marking. No timeout. Player stays active.
+33s  Rejoin click. Player connects. SUCCESS.
```

The GUID Upgrade is the **bridge** between the one-shot AutoSync (which buys time with the old GUID) and MCC's new RTA WebSocket (which provides a permanent live connection). Without it, there's a ~13-second gap where the session has a dead GUID and is vulnerable to timeout removal.

---

### 5.7 Session Discovery Injection

**File:** `ProxyService.cs`, INJECT block

When MCC restarts and queries MPSD for its active sessions (`GET /sessions?xuid=...&private=true&inactive=true`), the real response may not include the match (if the player was briefly removed). The proxy injects a fake result containing the saved session.

```csharp
// ProxyService.cs - OnAfterResponse, session discovery
if (entry.Path.Contains("/sessions?") && entry.Path.Contains("xuid="))
{
    if (_lastMatchSession is not null && _pendingCrashRestore)
    {
        // Build fake session list with the saved match session
        var fakeResult = new {
            results = new[] {
                new {
                    xuid = _playerXuid,
                    startTime = _lastMatchSession.SavedAt,
                    sessionRef = new {
                        scid = _lastMatchSession.Scid,
                        templateName = _lastMatchSession.TemplateName,
                        name = _lastMatchSession.SessionName
                    }
                }
            }
        };
        e.SetResponseBodyString(JsonSerializer.Serialize(fakeResult));
        // Log INJECT
    }
}
```

This makes MCC's session discovery find the saved match, triggering the rejoin prompt.

---

### 5.8 Squad Handle Override

**File:** `ProxyService.cs`, lines ~575-620

When MCC restarts, it creates a NEW squad session and POSTs a new activity handle pointing to it. The override intercepts these handle POSTs and replaces the new squad UUID with the saved one, so the activity handle points to the original match's squad.

```csharp
// ProxyService.cs - OnBeforeRequest, handle override
if (_pendingCrashRestore && _savedSquadHandle is not null &&
    req.Method == "POST" && !string.IsNullOrEmpty(entry.RequestBody) &&
    entry.RequestBody.Contains("cascadesquadsession", ...))
{
    string savedSessionName = _savedSquadHandle.SessionName;
    if (!string.IsNullOrEmpty(savedSessionName))
    {
        using var doc = JsonDocument.Parse(entry.RequestBody);
        if (root.TryGetProperty("sessionRef", out var sessionRef) &&
            sessionRef.TryGetProperty("name", out var nameEl))
        {
            string originalName = nameEl.GetString() ?? "";
            if (!string.Equals(originalName, savedSessionName, ...))
            {
                // Replace the UUID in the body
                string newBody = entry.RequestBody.Replace(
                    originalName, savedSessionName);
                e.SetRequestBody(Encoding.UTF8.GetBytes(newBody));
                // Log OVERRIDE[SquadHandle]
            }
        }
    }
}
```

**Fields:**
```csharp
private SavedHandleInfo? _savedSquadHandle;  // Set by Scanner via SetSavedSquadHandle()
```

---

### 5.9 Game Server Redirect

**File:** `ProxyService.cs`, lines ~1525-1580

When MCC calls PlayFab's `RequestParty` to get a game server, the proxy can redirect this to the original cached server instead of accepting a new allocation.

```csharp
// ProxyService.cs - OnAfterResponse, RequestParty handling
if (entry.Path.Contains("Party/RequestParty", ...))
{
    // Parse server info from response
    var serverInfo = ParseGameServerInfo(responseBody);

    // REDIRECT PATH: During crash recovery, use cached server
    if (_cachedGameServerInfo is not null &&
        (_pendingCrashRestore || _cachedInjectedMatchBody is not null))
    {
        _gameServerRedirectionActive = true;
        var redirectedBody = ConstructRequestPartyResponse(_cachedGameServerInfo);
        e.SetResponseBodyString(redirectedBody);
        // Log REDIRECT[GameServer]
        return;
    }

    // CACHE PATH: Normal gameplay, first RequestParty
    else if (_cachedGameServerInfo == null && !_gameServerRedirectionActive)
    {
        _cachedGameServerInfo = serverInfo;
        PersistGameServerToDisk(serverInfo);
        // Log CACHE[GameServer]
    }
}
```

**Disk persistence:** The game server is saved to `last-game-server.json` so it survives proxy restarts. On crash restore, `SetPendingCrashRestore()` loads it from disk if not already in memory.

**Current limitation:** For some match types, the game server assignment doesn't flow through `RequestParty` during normal gameplay (it comes through the Xbox GDK networking layer). In these cases, the server can't be cached, and the rejoin `RequestParty` returns a new allocation. The GUID Upgrade (FIX #16) keeps the player in MPSD regardless, allowing the game server connection to validate.

---

### 5.10 ETag Refresh

**File:** `ProxyService.cs`, ETag refresh block (~line 580+)

MPSD uses ETags for optimistic concurrency. When MCC PUTs a session with a stale `If-Match` header, MPSD returns 412 Precondition Failed. The proxy detects this pattern and silently refreshes the ETag:

```csharp
// ProxyService.cs - OnBeforeRequest, ETag refresh
// When a session PUT has an If-Match header, do a quick GET to get the
// current ETag and swap it in before the PUT reaches MPSD.
if (req.Method == "PUT" && req.Headers.Any(h => h.Name == "If-Match"))
{
    using var freshGet = new HttpRequestMessage(HttpMethod.Get, req.Url);
    // ... copy auth headers, skip Signature ...
    using var freshResp = await httpClient.SendAsync(freshGet);
    string freshEtag = freshResp.Headers.ETag?.Tag ?? "*";
    req.Headers.RemoveHeader("If-Match");
    req.Headers.AddHeader("If-Match", freshEtag);
}
```

---

## 6. MPSD Session Constants

These are the CascadeMatchmaking session constants that govern timeout behavior:

```json
{
  "inactiveRemovalTimeout": 0,
  "connectionRequiredForActiveMembers": true,
  "readyRemovalTimeout": 30000,
  "reservedRemovalTimeout": 30000,
  "sessionEmptyTimeout": 0,
  "maxMembersCount": 64
}
```

**Critical constants:**
- **`inactiveRemovalTimeout: 0`** - Members marked inactive are removed **immediately** on the next session access (not real-time; it's lazy evaluation). This is the primary adversary.
- **`connectionRequiredForActiveMembers: true`** - Active members MUST have a valid RTA connection. If the connection GUID points to a dead WebSocket, the member is marked inactive.

---

## 7. Critical Bug History

Each bug was discovered through capture analysis and fixed iteratively:

| # | Bug | Root Cause | Fix |
|---|-----|-----------|-----|
| 4 | Handle persistence overwrites saved session | `PersistHandleToDisk` ran during recovery | Added `_pendingCrashRestore` guard |
| 5 | Ghost mode disabled during recovery | `SetSavedMatchSession` cleared session on UUID mismatch | Guard `if (_pendingCrashRestore) return` |
| 6 | Ghost mode never activates | `_lastMatchSession is null` condition prevented parameter use | Removed `&& _lastMatchSession is null` |
| 7 | Leave request corrupts saved GUID | Race: leave PUT persists empty connection GUID | Added `"me":null` body check in persistence |
| 8 | Handle file missing connection data | `PersistHandleToDisk` didn't capture connection fields | Enriched with `_capturedConnectionGuid` etc. |
| 9 | Squad handle override missing | Handle POST used new squad UUID instead of saved | Added handle body replacement in OnBeforeRequest |
| 9a/9b | Saved sessions overwritten during recovery | `PersistHandleToDisk`/`PersistMatchSessionToDisk` unguarded | Added `_pendingCrashRestore` guards |
| 9c | AutoSync PUT missing connection GUID | PUT body didn't include connection field | Added connection GUID to PUT body |
| 10 | Handle override gated on stale flag | `_pendingCrashRestore` cleared before handle POST arrived | Removed flag guard from override |
| Watcher | Missed fast MCC restarts | 5s interval + no PID tracking | 1s interval + `_lastMccPid` tracking |
| Redirect | Game server redirect gated on `_pendingCrashRestore` | Condition too narrow | Added `\|\| _cachedInjectedMatchBody is not null` |
| Paths | `HaloIntel` vs `HaloMCCToolbox` | Renamed project, old paths remained | Updated all 3 path references |
| AutoSync | Loop PUTs triggered `inactiveRemovalTimeout` | Each PUT was an MPSD access that triggered timeout check | Reverted to one-shot (matching PortableCLEAN) |
| **16** | **Dead GUID causes removal at ~30s** | **Old GUID's WebSocket dies; heartbeat detects at ~30s** | **GUID Upgrade: PUT new live GUID at +17s** |

---

## 8. File Reference

### ProxyService.cs (~2330 lines)

**Key Fields:**
```csharp
private bool _pendingCrashRestore;              // Master recovery flag
private bool _blockMatchLeave;                  // One-shot leave blocker
private bool _ghostSessionMode;                 // Ghost mode active
private SavedHandleInfo? _ghostSession;         // Session data for ghost mode
private bool _guidUpgradeDone;                  // FIX #16 one-shot flag
private string _capturedConnectionGuid;         // New GUID from MCC restart
private string _capturedSubscriptionId;         // New subscription from MCC restart
private Dictionary<string, string>? _lastCapturedAuthHeaders;  // Fresh auth
private GameServerInfo? _cachedGameServerInfo;  // Cached game server for redirect
private SavedHandleInfo? _savedSquadHandle;     // Saved squad for handle override
private string? _cachedInjectedMatchBody;       // Cached fake session body
private SavedHandleInfo? _lastMatchSession;     // Current/saved match session
```

**Key Methods:**
| Method | Lines | Purpose |
|--------|-------|---------|
| `SetPendingCrashRestore()` | 164-207 | Arms all recovery systems |
| `AutoSyncGhostSessionAsync()` | 2005-2165 | One-shot GET+PUT at +2s |
| `GuidUpgradeAsync()` | 2167-2270 | Replaces dead GUID with live one |
| `TryHandleGhostMode()` | ~670-780 | Intercepts MPSD requests during recovery |
| `GenerateFakeSessionDocument()` | ~800-860 | Builds fake CascadeMatchmaking response |
| `OnBeforeRequest()` | ~380-650 | Leave blocking, GUID capture, handle override |
| `OnAfterResponse()` | ~1200-1600 | Game server redirect, session injection |
| `PersistMatchSessionToDisk()` | ~700-750 | Saves match session (guarded) |
| `PersistHandleToDisk()` | ~400-420 | Saves squad handle (guarded) |
| `ConstructRequestPartyResponse()` | ~2175-2220 | Builds fake RequestParty response |

### Scanner.xaml.cs (~1405 lines)

**Key Fields:**
```csharp
private SavedHandleInfo? _savedHandle;          // Last squad handle from disk
private SavedHandleInfo? _savedMatchSession;    // Last match session from disk
private bool _mccWasRunning;                    // Previous watcher tick state
private int _lastMccPid;                        // PID tracking for fast restarts
private bool _restoreInProgress;                // Debounce flag
```

**Key Methods:**
| Method | Lines | Purpose |
|--------|-------|---------|
| `MccWatcher_Tick()` | 1313-1402 | 1s poll for MCC process, triggers restore |
| `LoadSavedHandle()` | 970-996 | Reads last-handle.json from disk |
| `LoadSavedMatchSession()` | 998-1026 | Reads last-match-session.json from disk |
| `ClearRejoinData()` | 235-246 | Clears all saved state (disk + memory) |

---

## 9. Capture Analysis: The Working Run

**Capture:** `HaloCapture-20260317-081807CLOSE.zip` (501 files)

### What Fired Correctly

| Diagnostic Event | Count | Status |
|-----------------|-------|--------|
| `RESTORE[Start]` | 1 | Fired at 08:17:02 with correct session data |
| `AUTO-BLOCK[Armed]` | 1 | Armed immediately after restore |
| `GHOST[AutoSync-OneShot]` | 1 | GET+PUT succeeded at +2s (200 OK) |
| `UPGRADE[GUID-Success]` | 1 | New GUID PUT succeeded at +17s |
| `GHOST[Match-Get]` | 2 | Fake match docs served to MCC |
| `GHOST[Squad-Get]` | 7 | Fake squad docs served to MCC |
| `OVERRIDE[SquadHandle]` | 10 | All handle POSTs redirected to saved squad |
| `MCC-REJOIN-CLICK` | 1 | User successfully clicked rejoin at 08:17:33 |
| `CACHE[GameServer]` | 1 | Game server obtained from RequestParty |

### Key Timestamps

```
08:16:14.716  SAVE[Match]           CascadeMatchmaking/6b4fe654-29e0 saved
08:17:02.155  RESTORE[Start]        handle=c5a5f8a2  match=6b4fe654
08:17:02.159  AUTO-BLOCK[Armed]     Block armed for next leave request
08:17:04.872  AutoSync-OneShot      conn=54c52880... (saved) -> 200 OK
08:17:19.088  GHOST[Match-Get]      Fake match document served
08:17:19.302  Squad PUT captured    NEW GUID: 85903ee9-2958-4397-...
08:17:19.xxx  UPGRADE[GUID-Success] GUID upgraded on CascadeMatchmaking
08:17:22.589  OVERRIDE[SquadHandle] 50369981 -> c5a5f8a2 (saved)
08:17:33.661  MCC-REJOIN-CLICK      User clicked rejoin
08:17:34.790  CACHE[GameServer]     145.190.8.32:30656
```

### Session Data

**Saved Match Session:**
- Template: `CascadeMatchmaking`
- Session: `6b4fe654-29e0-4c39-9478-a3cdacc12d93`
- Connection GUID (saved): `54c52880-19c4-4c3d-8055-b5a16d024ff1`
- Connection GUID (upgraded): `85903ee9-2958-4397-9ad7-1a6844af0b35`

**Saved Squad Handle:**
- Template: `cascadesquadsession`
- Session: `c5a5f8a2-8cea-4c95-b26c-bb8f7ef9ca0b`

**Game Server:**
- IP: `145.190.8.32:30656`
- Region: EastUs
- Protocol: UDP (XRNM)

---

## Appendix: State Machine

```
                    NORMAL GAMEPLAY
                         |
                    [MCC Running]
                    Proxy captures:
                    - Match session -> disk
                    - Squad handle -> disk
                    - Game server -> disk
                    - Connection GUID -> memory
                         |
                    [MCC Crashes]
                         |
                    WATCHER DETECTS
                    (1s poll, PID check)
                         |
              +----------+----------+
              |          |          |
         ForceBlock   SetPending   SetSaved
         MatchLeave   CrashRestore SquadHandle
              |          |
              |    +-----+------+
              |    |            |
              | GhostMode   AutoSync
              | ENABLED     +2s OneShot
              |    |            |
              |    |     [PUT saved GUID]
              |    |            |
              |    |     [MCC Restarts]
              |    |     New squad PUT
              |    |     captured with
              |    |     NEW GUID
              |    |            |
              |    |     GUID UPGRADE
              |    |     PUT new GUID to
              |    |     CascadeMatchmaking
              |    |            |
              |    +-----+------+
              |          |
              |    [MCC Requests]
              |    Ghost mode intercepts
              |    Fake session docs
              |    Handle overrides
              |          |
              |    [Rejoin Click]
              |    RequestParty -> server
              |    MCC connects to game
              |          |
              |    [SUCCESS]
              |    Player back in match
              |          |
              +----------+
                         |
                    RECOVERY COMPLETE
```
