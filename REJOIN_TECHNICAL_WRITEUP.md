# Rejoin Guard: Technical Writeup
## Bridging the Gap Between Expected MPSD Behavior and Toolbox Reality

---

## Overview

The rejoin feature allows players to re-enter a match session after an unexpected game exit (crash, disconnect, kill, voluntary close). While MPSD (Xbox Live Session Directory) has native session persistence, several critical mechanisms have been disabled or removed from MCC. We had to implement workarounds to restore the feature.

---

## Expected MPSD Behavior (Design Intent)

### How Rejoin Should Work (Ideal Case)

1. **Player in Match**: MCC is a member of a CascadeMatchmaking session on MPSD
2. **Unexpected Exit**: MCC crashes or is closed
3. **Session Persists**: MPSD keeps the session alive and the player's membership active
4. **Natural Rejoin**: Player restarts MCC and queues for matchmaking
5. **MPSD Recognition**: MPSD sees the player is already in a session and shows rejoin prompt
6. **Automatic Restoration**: MCC re-enters the session seamlessly

### Key MPSD Assumptions
- **Heartbeats**: MCC sends periodic session heartbeats to keep its membership active
- **Leave Requests**: Player membership removal only happens when player explicitly quits match
- **Session Timeout**: Inactive members are removed after extended period (minutes+)
- **Stale Connection Detection**: MPSD can detect when MCC's connection is stale vs. player voluntarily left

---

## What's Actually Broken in MCC

### 1. **MCC Sends Immediate Leave on Exit**
**Problem**: When MCC closes (whether by user action, crash, or kill), it sends a `PUT` to `{match-session}/members/me` with body `{"members":{"me":null}}` IMMEDIATELY, before any graceful shutdown. This is treated as a voluntary leave.

**Expected**: Session should be marked inactive, player membership persists for timeout period. Player can rejoin by queuing.

**Actual**: Session treats it as explicit player departure. Player is removed. On restart, no session to rejoin.

**Our Intervention**:
```csharp
// ProxyService.cs, block-match-leave logic
if (_blockMatchLeave && _lastMatchSession is not null &&
    req.Method == "PUT" &&
    req.RequestUri.Host.EndsWith("sessiondirectory.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
    req.RequestUri.AbsolutePath.Contains(_lastMatchSession.SessionName, StringComparison.OrdinalIgnoreCase) &&
    !string.IsNullOrEmpty(entry.RequestBody) &&
    entry.RequestBody.Contains("\"me\":null"))
{
    // Short-circuit: fake 200 to MCC, leave never reaches MPSD
    e.Ok("{}");
    _blockMatchLeave = false;  // one-shot
}
```

**Why Necessary**: Without this, the first leave request succeeds (204 response), player is removed from session, rejoin is impossible. With this block, player stays in session on MPSD, rejoin prompt appears.

---

### 2. **No Automatic Exit Detection Mechanism**
**Problem**: There's no way for the toolbox to know when MCC has exited. The proxy only sees outbound requests; it doesn't monitor MCC process state.

**Expected**: MCC would implement heartbeat timeout detection internally, or MPSD would detect stale connections via TCP keepalive.

**Actual**: No heartbeat. MCC just sends a leave and dies. The proxy has no signal to arm the block mechanism.

**Our Intervention**:
```csharp
// Scanner.xaml.cs, MccWatcher_Tick
if (_mccWasRunning && !running && !_restoreInProgress &&
    (_savedHandle is not null || _savedMatchSession is not null))
{
    // Auto-arm block-match-leave: detect MCC exit and automatically block next leave
    _proxy.ForceBlockMatchLeave();
    AddDiag("AUTO-BLOCK[Armed]", "MCC exit detected — block armed for next leave request", "");

    if (_savedMatchSession is not null)
        _proxy.SetPendingCrashRestore(_savedMatchSession);

    _restoreInProgress = false;
}
```

**Timer**: Runs every 5 seconds. Detects when `Process.GetProcessesByName("MCC-Win64-Shipping")` transitions from existing → empty.

**Why Necessary**: Without automatic detection, the block only works when user manually clicks "Log Quit" button. If game crashes or closes naturally without clicking, the block isn't armed and the leave succeeds. Automatic detection makes it work in all exit scenarios.

---

### 3. **Session State Not Persisted Across Restart**
**Problem**: When MCC restarts, it has no memory of which match session it was in. The session data (session name, template, SCID, etc.) is lost.

**Expected**: MPSD would either remember the player's session memberships, or MCC would query MPSD to find outstanding sessions.

**Actual**: MCC has no query for "what sessions am I still a member of?" MCC starts fresh, calls session discovery endpoints.

**Our Intervention**:
```csharp
// Scanner.xaml.cs, AddCaptureEntry
// When proxy captures a PUT to the match session, reload and persist it
if (entry.Method == "PUT" &&
    entry.Url.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase) &&
    !string.IsNullOrEmpty(entry.RequestBody) &&
    entry.RequestBody.Contains("\"members\"") &&
    !entry.RequestBody.Contains("\"me\":null"))
{
    LoadSavedMatchSession();
}

// When proxy captures a POST to /handles (activity handle), reload and persist it
if (entry.Method == "POST" &&
    entry.Url.Contains("/handles", StringComparison.OrdinalIgnoreCase) &&
    !string.IsNullOrEmpty(entry.RequestBody))
{
    LoadSavedHandle();
}
```

Saved to disk:
- `%APPDATA%\HaloMCCToolbox\last-match-session.json` (session name, template, SCID, connection GUID)
- `%APPDATA%\HaloMCCToolbox\last-handle.json` (activity handle reference for session reconstruction)

**Why Necessary**: Without persisting session state, on MCC restart there's no way to reconstruct which session the player was in. MPSD has no way to push this info to a restarting MCC. We have to capture it from network requests and save it.

---

### 4. **No Connection GUID Preservation**
**Problem**: When MCC re-enters a match session, it must include its original connection GUID in the session member properties. This GUID ties the restarted MCC instance to the existing session membership (proving continuity).

**Expected**: MCC would extract its own connection GUID from somewhere and include it in rejoin requests.

**Actual**: MCC generates a new GUID on restart, so MPSD sees it as a different "connection" and rejects the re-entry as invalid.

**Our Intervention**:
```csharp
// SavedHandleInfo / SavedMatchSession models capture connectionGuid
public string ConnectionGuid { get; set; } = "";

// When restoring, ReaddToMatchSessionAsync includes it in the PUT body:
public class SessionMemberUpdate
{
    public Members Members { get; set; }
}

// With system properties containing the original connection GUID
```

**Why Necessary**: Without preserving the connection GUID, the re-entry PUT gets rejected (403 Forbidden) because MPSD sees it as an unauthorized connection attempt. The GUID proves the same player session is rejoining.

---

### 5. **Ghost Session Mode for JIT Injection**
**Problem**: After MCC restarts, MPSD no longer believes MCC is in the match session. If MCC queries session details, MPSD returns either 401 (unauthorized) or the player isn't in the members list. This causes MCC to think there's no session to rejoin.

**Expected**: MPSD would return a stale but valid session membership allowing re-entry.

**Actual**: MPSD returns 401/403 because the membership was deleted when the leave request was processed (before we blocked it, in old behavior).

**Our Intervention**:
```csharp
// ProxyService.cs, ghost session mode
if (_pendingCrashRestore is not null &&
    req.RequestUri.AbsolutePath.Contains(_pendingCrashRestore.SessionName))
{
    // Return fake/cached session data to MCC so it thinks it's still in the session
    // This allows MCC to re-add itself without 401 errors
    e.Ok(CachedSessionJson);
}
```

**Why Necessary**: Without ghost responses, MCC gets 401 on first session query post-restart, assumes it's not in the session, and proceeds with normal matchmaking without attempting to rejoin. Ghost mode makes MCC think the session is still valid and worth re-entering.

---

### 6. **ETag Conflict Handling (412 Errors)**
**Problem**: When the proxy re-enters MCC into a session via JIT injection, the session ETag (concurrency token) may have changed since the session was captured. MCC's PUT includes the old ETag, MPSD rejects with 412 Precondition Failed.

**Expected**: MCC would handle 412 gracefully, refetch the ETag, and retry.

**Actual**: MCC doesn't retry 412. It treats it as a hard failure. The rejoin attempt fails silently.

**Our Intervention**:
```csharp
// Scanner.xaml.cs, RefreshETagAsync
private async Task<string> RefreshETagAsync(string sessionUrl, Dictionary<string, string> headers)
{
    var freshReq = new HttpRequestMessage(HttpMethod.Get, sessionUrl);
    foreach (var h in headers)
        freshReq.Headers.Add(h.Key, h.Value);

    var freshRes = await _mpsdClient.SendAsync(freshReq);
    if (freshRes.Headers.TryGetValues("ETag", out var values))
        return values.First();

    return "*";  // wildcard ETag (bypass check)
}

// Called when 412 is encountered, refreshes ETag and retries PUT
```

**Why Necessary**: Without ETag refresh, the rejoin PUT fails with 412. The error is logged but the session is considered "failed to restore". With refresh, we get the current ETag and the PUT succeeds.

---

## Summary of Interventions

| Feature | MPSD Expected | MCC Reality | Our Fix |
|---------|---------------|-------------|---------|
| **Graceful Session Exit** | Player stays in session on voluntary close, timeout removes them | MCC sends immediate leave, removes player | **Block the leave request** |
| **Exit Detection** | MPSD detects stale connections, or MCC implements heartbeat timeout | No detection mechanism; proxy blind to MCC state | **Monitor MCC process, auto-arm block on exit** |
| **Session Memory** | MPSD tracks memberships; MCC can query outstanding sessions | MCC starts fresh, has no session memory | **Capture and persist session state to disk** |
| **Connection Continuity** | Connection GUID ties rejoining instance to original membership | MCC generates new GUID on restart | **Save connection GUID, include in rejoin requests** |
| **Session Availability** | Valid membership = accessible session | Post-leave, session is 401/403 to player | **Fake session responses in "ghost mode"** |
| **ETag Concurrency** | MPSD updates ETag on session changes | Old ETag causes 412 on rejoin attempt | **Refresh ETag on 412, retry** |

---

## Why This Feature Doesn't Work Without These Changes

Each intervention solves a specific failure point:

1. **Without block-match-leave**: Leave request succeeds, player removed, no session to rejoin → 404
2. **Without auto-detection**: Block only works for manual button clicks, crashes fail silently → 403
3. **Without session persistence**: Restarted MCC doesn't know which session to target → 404 / no attempt
4. **Without connection GUID**: Rejoin PUT rejected as unauthorized connection → 403
5. **Without ghost mode**: MCC sees 401 on session query, assumes no session to rejoin → skips rejoin
6. **Without ETag refresh**: Concurrent session changes cause 412, rejoin fails → 412 error loop

---

## Conclusion

The rejoin feature requires **six concurrent interventions** because MCC's architecture has diverged significantly from what MPSD expects. Rather than a simple "session persistence" feature that MPSD natively supports, we've had to build a complete crash-recovery system that:

- **Intercepts** network traffic to prevent session destruction
- **Monitors** process state to detect when intervention is needed
- **Persists** session data locally (MPSD doesn't provide re-discovery)
- **Reconstructs** session membership with proper continuity tokens
- **Fakes** session discovery responses during recovery
- **Handles** concurrency conflicts from stale cached state

This is fundamentally a **workaround for missing features in MCC's implementation**, not a feature that MPSD provides.
