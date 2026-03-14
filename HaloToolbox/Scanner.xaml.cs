using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace HaloToolbox;

public partial class Scanner : UserControl
{
    // Dedicated client for MPSD control requests — bypasses our own proxy to avoid loops
    private static readonly HttpClient _mpsdClient = new(new HttpClientHandler { UseProxy = false })
    {
        Timeout = TimeSpan.FromSeconds(5)
    };

    // ── Proxy ─────────────────────────────────────────────────────────────────
    private readonly ProxyService _proxy = new();
    private readonly ObservableCollection<ProxyCaptureEntry> _captureLog = new();
    private ProxyCaptureEntry? _selectedEntry;
    private ICollectionView?   _captureView;

    // ── Session Intel ──────────────────────────────────────────────────────────
    private readonly ObservableCollection<SessionMember> _sessionMembers = new();
    private string _lastSessionJson    = "";
    private string _lastSessionUrl     = "";
    private Dictionary<string, string> _lastSessionHeaders = new();
    private bool   _autoLeaveEnabled   = false;
    private readonly ObservableCollection<string> _sessionLog = new();

    // ── Hopper population query ────────────────────────────────────────────────
    private string                     _ticketScid        = "";
    private string                     _ticketHopper      = "";
    private string                     _ticketId          = "";
    private Dictionary<string, string> _ticketAuthHeaders = new();

    // ── Rejoin Guard / MCC Crash Restore ───────────────────────────────────────
    // NOTE: ProxyService saves to HaloMCCToolbox directory
    private static readonly string HandleFile =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                     "HaloMCCToolbox", "last-handle.json");
    private static readonly string MatchSessionFile =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                     "HaloMCCToolbox", "last-match-session.json");

    private SavedHandleInfo? _savedHandle;
    private SavedHandleInfo? _savedMatchSession;
    private readonly System.Windows.Threading.DispatcherTimer _mccWatcher =
        new() { Interval = TimeSpan.FromSeconds(1) };  // CRITICAL FIX: 5s → 1s (catch fast restarts)
    private bool _mccWasRunning = false;
    private int _lastMccPid = 0;  // CRITICAL FIX: Track PID to detect fast MCC restarts
    private bool _restoreInProgress = false;

    public Scanner()
    {
        InitializeComponent();

        _proxy.OnRequestCaptured += (_, entry) =>
            Dispatcher.InvokeAsync(() => AddCaptureEntry(entry));

        _proxy.WinHttpManualSetRequired += (_, _) =>
            Dispatcher.InvokeAsync(() => ProxyWinHttpNote.Visibility = Visibility.Visible);

        // CRITICAL: When proxy saves a match session to disk, reload it and sync to proxy
        // This ensures the in-memory ghost session is up-to-date for JIT injection on MCC restart
        _proxy.OnMatchSessionSaved += (_, _) =>
            Dispatcher.InvokeAsync(() => LoadSavedMatchSession());

        _captureView = CollectionViewSource.GetDefaultView(_captureLog);
        _captureView.Filter = o => FilterCapture((ProxyCaptureEntry)o);

        ProxyCaptureList.ItemsSource  = _captureView;
        SessionMemberList.ItemsSource = _sessionMembers;
        SessionLogList.ItemsSource    = _sessionLog;

        // Initialize MCC watcher for rejoin crash restore
        // CRITICAL: Run on UI thread — LoadSavedHandle/LoadSavedMatchSession call AddDiag
        // which updates the UI. Running on Task.Run (background thread) causes cross-thread
        // exceptions that prevent the watcher from starting.
        _ = Dispatcher.InvokeAsync(() =>
        {
            // Clear saved session data on every Toolbox startup. This prevents stale data from
            // a previous Xbox account (or previous session) from interfering when the user
            // launches the app fresh. The proxy will re-capture a new session when MCC makes
            // its first request with the current account.
            ClearRejoinData();

            LoadSavedHandle();
            LoadSavedMatchSession();
            _mccWatcher.Tick += MccWatcher_Tick;
            _mccWatcher.Start();
        }, System.Windows.Threading.DispatcherPriority.Normal);

        // NOTE: Proxy is NOT disposed on tab unload.
        // Disposal only happens when the entire application closes (handled in MainWindow).
        // This allows the proxy to remain active even when the Scanner tab is not selected.
    }

    // ── Proxy toggle ──────────────────────────────────────────────────────────
    private async void ProxyToggle_Click(object sender, RoutedEventArgs e)
    {
        if (_proxy.IsRunning)
        {
            _proxy.Stop();
            ProxyToggleBtn.Content      = "▶  START PROXY";
            ProxyStatusDot.Fill         = new SolidColorBrush(Color.FromRgb(0x48, 0x4F, 0x58));
            ProxyStatusText.Text        = "STOPPED";
            ProxyWinHttpNote.Visibility = Visibility.Collapsed;
        }
        else
        {
            await StartProxyAsync();
        }
    }

    /// <summary>
    /// Public method to start the proxy from external callers (e.g., Fix Rejoins button).
    /// Can be safely called from MainWindow or other components.
    /// </summary>
    public async Task<bool> StartProxyAsync()
    {
        if (_proxy.IsRunning) return true;  // Already running

        if (int.TryParse(ProxyPortBox.Text, out int port))
            _proxy.Port = port;

        ProxyToggleBtn.IsEnabled = false;
        ProxyStatusText.Text     = "STARTING…";
        try
        {
            await _proxy.StartAsync();
            ProxyToggleBtn.Content = "■  STOP PROXY";
            ProxyStatusDot.Fill    = new SolidColorBrush(Color.FromRgb(0x3F, 0xB9, 0x50));
            ProxyStatusText.Text   = $"LISTENING :{_proxy.Port}  —  Restart MCC";
            return true;
        }
        catch (Exception ex)
        {
            ProxyStatusText.Text = $"ERROR: {ex.Message}";
            ProxyStatusDot.Fill  = new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49));
            return false;
        }
        finally
        {
            ProxyToggleBtn.IsEnabled = true;
        }
    }

    /// <summary>
    /// Public method to dispose of the proxy. Called from MainWindow Closing event.
    /// </summary>
    public void DisposeProxy()
    {
        if (_proxy.IsRunning)
            _proxy.Stop();
        _proxy.Dispose();
    }

    /// <summary>
    /// Public property to check if proxy is running.
    /// </summary>
    public bool IsProxyRunning => _proxy.IsRunning;

    /// <summary>
    /// Updates the Fix Rejoins button in MainWindow based on proxy running state.
    /// Called after proxy starts/stops.
    /// </summary>
    public void UpdateFixRejoinsButtonState(Button btnFixRejoins)
    {
        if (_proxy.IsRunning)
            btnFixRejoins.Content = "◀  RUNNING";
        else
            btnFixRejoins.Content = "RUN FIX";
    }

    /// <summary>
    /// Adds a lightweight diagnostic capture entry (for debugging event chains).
    /// Skips during cold startup when proxy is not yet running and capture log is empty.
    /// </summary>
    private void AddDiag(string method, string body, string url = "")
    {
        if (_captureLog.Count == 0 && !_proxy.IsRunning) return;  // skip during cold startup
        AddCaptureEntry(new ProxyCaptureEntry
        {
            Method       = method,
            Url          = url,
            Host         = "diag",
            Path         = "",
            StatusCode   = 0,
            ResponseBody = body,
        });
    }

    /// <summary>
    /// Updates internal rejoin guard state based on current proxy status.
    /// Called after loading handles/sessions from disk to ensure state consistency.
    /// </summary>
    private void UpdateRejoinGuardUi()
    {
        bool hasAny = _savedHandle is not null || _savedMatchSession is not null;

        if (_savedHandle is null)
            RejoinSessionText.Text = "no session saved";
        else
            RejoinSessionText.Text = $"{_savedHandle.TemplateName} / {_savedHandle.SessionShort}   {_savedHandle.SavedAtStr}";

        if (_savedMatchSession is null)
            RejoinMatchText.Text = "";
        else
            RejoinMatchText.Text = $"MATCH  {_savedMatchSession.TemplateName} / {_savedMatchSession.SessionShort}   {_savedMatchSession.SavedAtStr}";

        // Show game server redirection status if active
        if (_proxy.IsGameServerRedirectionActive)
            RejoinStatusText.Text = "🎯 GAME SERVER REDIRECTED — connecting to original server…";
        // Show if ghost mode is active
        else if (_proxy.IsGhostSessionActive())
            RejoinStatusText.Text = "⚡ GHOST MODE — syncing with MPSD…";
        else
            RejoinStatusText.Text = hasAny ? "" : "";

        RejoinCheckBtn.IsEnabled   = hasAny;
        RejoinRestoreBtn.IsEnabled = hasAny;
    }

    /// <summary>
    /// Clears all saved rejoin session data and cleans up proxy state.
    /// Matches HaloIntel's RejoinClear_Click behavior.
    /// </summary>
    public void ClearRejoinData()
    {
        _savedHandle       = null;
        _savedMatchSession = null;
        _proxy.ClearSavedMatchSession();     // also clear proxy's in-memory copy (stops injection)
        _proxy.ClearGameServerRedirection(); // also clear game server redirection
        _proxy.ClearGhostSessionMode();      // clear any ghost mode state
        try { File.Delete(HandleFile); }       catch { /* ignore */ }
        try { File.Delete(MatchSessionFile); } catch { /* ignore */ }
        UpdateRejoinGuardUi();
        AddDiag("CLEAR[Rejoin]", "Cleared all saved rejoin data from disk and proxy");
    }

    // ── Capture entry management ──────────────────────────────────────────────
    private void AddCaptureEntry(ProxyCaptureEntry entry)
    {
        _captureLog.Insert(0, entry);
        if (_captureLog.Count > 500)
            _captureLog.RemoveAt(_captureLog.Count - 1);

        TryParseSessionEntry(entry);

        // Reload saved handle from disk when a /handles POST is captured —
        // PersistHandleToDisk already wrote it at request time, so the file is current.
        if (entry.Method == "POST" &&
            entry.Url.Contains("/handles", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(entry.RequestBody))
            LoadSavedHandle();

        // Same belt-and-suspenders reload for the matchmaking session file.
        // PersistMatchSessionToDisk fires during request handling; by the time
        // this capture entry arrives (response phase), the file is on disk.
        if (entry.Method == "PUT" &&
            entry.Url.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase))
            LoadSavedMatchSession();
    }

    private void ProxyCaptureList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ProxyCaptureList.SelectedItem is not ProxyCaptureEntry entry) return;
        _selectedEntry = entry;
        ShowEntryDetail(entry);
    }

    private void ShowEntryDetail(ProxyCaptureEntry entry)
    {
        ProxyDetailUrl.Text = $"{entry.Method}  {entry.Url}";

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(entry.RequestBody))
            parts.Add($"── REQUEST BODY ──\n{TryPrettyJson(entry.RequestBody)}");
        if (!string.IsNullOrWhiteSpace(entry.ResponseBody))
            parts.Add($"── RESPONSE BODY ──\n{TryPrettyJson(entry.ResponseBody)}");

        ProxyDetailBody.Text = string.Join("\n\n", parts);
    }

    private static string TryPrettyJson(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return raw;
        try
        {
            using var doc = JsonDocument.Parse(raw);
            return JsonSerializer.Serialize(doc.RootElement,
                new JsonSerializerOptions { WriteIndented = true });
        }
        catch { return raw; }
    }

    // ── Header bar button handlers ────────────────────────────────────────────
    private void ProxyCopyJson_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedEntry is null) return;
        Clipboard.SetText(TryPrettyJson(_selectedEntry.ResponseBody));
    }

    private void ProxyClear_Click(object sender, RoutedEventArgs e)
    {
        _captureLog.Clear();
        _selectedEntry       = null;
        ProxyDetailUrl.Text  = "Select an entry to view details";
        ProxyDetailBody.Text = "";
    }

    private void ProxySaveZip_Click(object sender, RoutedEventArgs e)
    {
        if (_captureLog.Count == 0)
        {
            MessageBox.Show("No captured packets to save.", "Save ZIP",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title    = "Save Captured Packets",
            Filter   = "ZIP files (*.zip)|*.zip",
            FileName = $"HaloCapture-{DateTime.Now:yyyyMMdd-HHmmss}.zip"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            var entries = _captureLog.ToList();
            using var fs  = new FileStream(dlg.FileName, FileMode.Create);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create);

            for (int i = 0; i < entries.Count; i++)
            {
                var entry = entries[i];
                var obj = new
                {
                    entry.Timestamp, entry.Method, entry.Url, entry.Host, entry.Path,
                    entry.StatusCode,
                    RequestHeaders  = entry.RequestHeaders,
                    entry.RequestBody,
                    ResponseHeaders = entry.ResponseHeaders,
                    entry.ResponseBody
                };

                var json     = JsonSerializer.Serialize(obj, new JsonSerializerOptions { WriteIndented = true });
                var name     = $"{i:D4}_{entry.Timestamp:HHmmss-fff}_{entry.Method}_{SanitizeFileName(entry.Host + entry.Path)}.json";
                var zipEntry = zip.CreateEntry(name, CompressionLevel.SmallestSize);
                using var writer = new StreamWriter(zipEntry.Open());
                writer.Write(json);
            }

            var summary = entries.Select((ent, idx) => new
                { Index = idx, ent.Timestamp, ent.Method, ent.StatusCode, ent.Url });
            var summaryJson  = JsonSerializer.Serialize(summary, new JsonSerializerOptions { WriteIndented = true });
            var summaryEntry = zip.CreateEntry("_index.json", CompressionLevel.SmallestSize);
            using (var sw = new StreamWriter(summaryEntry.Open()))
                sw.Write(summaryJson);

            MessageBox.Show($"Saved {entries.Count} packets to:\n{dlg.FileName}", "Save ZIP",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save ZIP:\n{ex.Message}", "Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sb      = new System.Text.StringBuilder(name.Length);
        foreach (var c in name)
            sb.Append(invalid.Contains(c) ? '_' : c);
        var result = sb.ToString();
        return result.Length > 80 ? result[..80] : result;
    }

    private void ProxyExtractToken_Click(object sender, RoutedEventArgs e)
    {
        foreach (var entry in _captureLog)
        {
            var token = entry.ExtractedSpartanToken;
            if (!string.IsNullOrEmpty(token))
            {
                Clipboard.SetText(token);
                ProxyStatusText.Text = "TOKEN COPIED TO CLIPBOARD";
                return;
            }
        }
        ProxyStatusText.Text = "NO TOKEN FOUND IN CAPTURED TRAFFIC";
    }

    private void ProxyCollapseToggle_Click(object sender, RoutedEventArgs e)
    {
        bool isCollapsed = ProxyPanelContent.Visibility == Visibility.Collapsed;
        ProxyPanelContent.Visibility = isCollapsed ? Visibility.Visible : Visibility.Collapsed;
        ProxyCollapseBtn.Content     = isCollapsed ? "▲" : "▼";
    }

    // ── Capture filter ────────────────────────────────────────────────────────
    private static readonly string[] _rejoinMethods =
    {
        "SAVE[", "LOAD[", "RESTORE[", "INJECT", "PASS[", "BLOCK[",
        "PUT[JIT", "POST[JIT", "PATCH[", "GET[ETag", "FAKE["
    };

    private bool FilterCapture(ProxyCaptureEntry e)
    {
        if (RejoinFilterToggle.IsChecked == true)
        {
            bool isRejoin = false;
            foreach (var prefix in _rejoinMethods)
                if (e.Method.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    || e.Method.Equals("INJECT", StringComparison.OrdinalIgnoreCase))
                { isRejoin = true; break; }
            if (!isRejoin) return false;
        }

        var text = ProxyFilterBox.Text.Trim();
        if (string.IsNullOrEmpty(text)) return true;
        return e.Host.Contains(text,   StringComparison.OrdinalIgnoreCase)
            || e.Path.Contains(text,   StringComparison.OrdinalIgnoreCase)
            || e.Method.Contains(text, StringComparison.OrdinalIgnoreCase);
    }

    private void ProxyFilterBox_TextChanged(object sender, TextChangedEventArgs e)
        => _captureView?.Refresh();

    private void RejoinFilterToggle_Changed(object sender, RoutedEventArgs e)
        => _captureView?.Refresh();

    // ── Session Intel ──────────────────────────────────────────────────────────

    /// <summary>
    /// Checks whether the captured entry is a /serviceconfigs/ session response,
    /// and if so, parses it to populate the Session Intel panel.
    /// </summary>
    private void TryParseSessionEntry(ProxyCaptureEntry entry)
    {
        var url = entry.Url;

        if (!url.Contains("/serviceconfigs/", StringComparison.OrdinalIgnoreCase) ||
            !url.Contains("/sessions/",       StringComparison.OrdinalIgnoreCase))
            return;

        if (string.IsNullOrWhiteSpace(entry.ResponseBody)) return;

        try
        {
            using var doc  = JsonDocument.Parse(entry.ResponseBody);
            var       root = doc.RootElement;

            // ── Matchmaking ticket session → store hopper info for population query ──
            if (root.TryGetProperty("servers", out var srvEl) &&
                srvEl.TryGetProperty("matchmaking", out var mmSrvEl) &&
                mmSrvEl.TryGetProperty("properties", out var mmPropsEl) &&
                mmPropsEl.TryGetProperty("system", out var mmSysEl) &&
                mmSysEl.TryGetProperty("status", out var mmStatusEl) &&
                mmStatusEl.GetString() == "searching" &&
                mmPropsEl.TryGetProperty("custom", out var mmCustEl))
            {
                string scid   = mmCustEl.TryGetProperty("ticketScid",       out var scidEl) ? scidEl.GetString()  ?? "" : "";
                string hopper = mmCustEl.TryGetProperty("ticketHopperName", out var hopEl)  ? hopEl.GetString()   ?? "" : "";
                string tid    = mmCustEl.TryGetProperty("ticketId",         out var tidEl)  ? tidEl.GetString()   ?? "" : "";
                if (!string.IsNullOrEmpty(scid) && !string.IsNullOrEmpty(hopper))
                {
                    _ticketScid        = scid;
                    _ticketHopper      = hopper;
                    _ticketId          = tid;
                    _ticketAuthHeaders = entry.RequestHeaders;
                    HopperNameText.Text   = hopper;
                    HopperResultText.Text = "";
                }
            }

            if (!root.TryGetProperty("members", out var membersEl) ||
                membersEl.ValueKind != JsonValueKind.Object)
                return;

            var rawMembers = new List<SessionMember>();

            foreach (var prop in membersEl.EnumerateObject())
            {
                var m      = prop.Value;
                var member = new SessionMember();

                if (m.TryGetProperty("gamertag", out var gtEl))
                    member.Gamertag = gtEl.GetString() ?? "";
                if (m.TryGetProperty("xuid", out var xuidEl))
                    member.Xuid = xuidEl.GetString() ?? "";

                if (m.TryGetProperty("constants", out var constsEl) &&
                    constsEl.TryGetProperty("custom", out var ccEl))
                {
                    // Path 1: constants.custom.matchmakingResult (CascadeMatchmaking format)
                    if (ccEl.TryGetProperty("matchmakingResult", out var mmEl))
                    {
                        if (string.IsNullOrEmpty(member.InitialTeam) &&
                            mmEl.TryGetProperty("initialTeam", out var itEl))
                            member.InitialTeam = itEl.GetString() ?? "";

                        if (mmEl.TryGetProperty("ticketAttrs", out var attrsEl))
                        {
                            ReadSkillPcts(attrsEl, member);
                            if (string.IsNullOrEmpty(member.PartyId) &&
                                attrsEl.TryGetProperty("PartyID", out var partyIdEl))
                                member.PartyId = partyIdEl.ValueKind == JsonValueKind.String
                                    ? partyIdEl.GetString() ?? ""
                                    : partyIdEl.GetRawText();
                        }
                        if (member.AvgSkillPct == 0)
                            ReadSkillPcts(mmEl, member);
                    }

                    // Path 2: directly under constants.custom
                    if (member.AvgSkillPct == 0)
                        ReadSkillPcts(ccEl, member);

                    if (string.IsNullOrEmpty(member.InitialTeam) &&
                        ccEl.TryGetProperty("initialTeam", out var ccItEl))
                        member.InitialTeam = ccItEl.GetString() ?? "";
                    if (string.IsNullOrEmpty(member.PartyId) &&
                        ccEl.TryGetProperty("PartyID", out var ccPidEl))
                        member.PartyId = ccPidEl.ValueKind == JsonValueKind.String
                            ? ccPidEl.GetString() ?? ""
                            : ccPidEl.GetRawText();
                }

                // Fallback: team in properties.system.team
                if (string.IsNullOrEmpty(member.InitialTeam) &&
                    m.TryGetProperty("properties", out var propsEl))
                {
                    if (propsEl.TryGetProperty("system", out var sysEl))
                    {
                        if (sysEl.TryGetProperty("team",        out var tEl)) member.InitialTeam = tEl.GetString() ?? "";
                        if (sysEl.TryGetProperty("initialTeam", out var iEl)) member.InitialTeam = iEl.GetString() ?? "";
                    }
                    if (string.IsNullOrEmpty(member.InitialTeam) &&
                        propsEl.TryGetProperty("custom", out var pcEl) &&
                        pcEl.TryGetProperty("initialTeam", out var pcItEl))
                        member.InitialTeam = pcItEl.GetString() ?? "";
                }

                if (string.IsNullOrEmpty(member.InitialTeam) &&
                    m.TryGetProperty("initialTeam", out var directTeamEl))
                    member.InitialTeam = directTeamEl.GetString() ?? "";

                // Path 3: skill directly on member element
                if (member.AvgSkillPct == 0)
                    ReadSkillPcts(m, member);

                if (!string.IsNullOrEmpty(member.Gamertag))
                    rawMembers.Add(member);
            }

            if (rawMembers.Count == 0) return;

            // Don't replace rich skill data with a skill-less session document.
            bool hasSkillData = rawMembers.Any(m => m.AvgSkillPct > 0);
            if (!hasSkillData && _sessionMembers.Count > 0)
            {
                // Game-lobby session: update InitialTeam in-place, preserve all skill data
                var byGt      = _sessionMembers.ToDictionary(m => m.Gamertag, StringComparer.OrdinalIgnoreCase);
                bool anyUpdate = false;
                foreach (var incoming in rawMembers)
                {
                    if (!string.IsNullOrEmpty(incoming.InitialTeam) &&
                        byGt.TryGetValue(incoming.Gamertag, out var existing) &&
                        existing.InitialTeam != incoming.InitialTeam)
                    {
                        existing.InitialTeam = incoming.InitialTeam;
                        anyUpdate = true;
                    }
                }
                if (anyUpdate)
                {
                    var sorted = _sessionMembers
                        .OrderBy(m => m.InitialTeam, StringComparer.Ordinal)
                        .ThenByDescending(m => m.AvgSkillPct)
                        .ToList();
                    _sessionMembers.Clear();
                    foreach (var m in sorted) _sessionMembers.Add(m);
                }
                return;
            }

            _lastSessionJson    = entry.ResponseBody;
            _lastSessionUrl     = entry.Url;
            _lastSessionHeaders = entry.RequestHeaders;
            AssignPartyInfo(rawMembers);
            RunSessionDetection(rawMembers, entry);

            rawMembers.Sort((a, b) =>
            {
                int tc = string.Compare(a.InitialTeam, b.InitialTeam, StringComparison.Ordinal);
                return tc != 0 ? tc : b.AvgSkillPct.CompareTo(a.AvgSkillPct);
            });

            _sessionMembers.Clear();
            foreach (var sm in rawMembers)
                _sessionMembers.Add(sm);

            SessionHopperText.Text       = ExtractSessionTemplate(entry.Url);
        }
        catch { /* swallow parse errors — malformed JSON / unexpected shape */ }
    }

    private static void ReadSkillPcts(JsonElement el, SessionMember member)
    {
        // JSON values are 0–1 scale; model uses 0–100
        if (el.TryGetProperty("MinGroupSkillPercentile",     out var minEl) && minEl.ValueKind == JsonValueKind.Number)
            member.MinSkillPct = minEl.GetDouble() * 100.0;
        if (el.TryGetProperty("MaxGroupSkillPercentile",     out var maxEl) && maxEl.ValueKind == JsonValueKind.Number)
            member.MaxSkillPct = maxEl.GetDouble() * 100.0;
        if (el.TryGetProperty("AverageGroupSkillPercentile", out var avgEl) && avgEl.ValueKind == JsonValueKind.Number)
            member.AvgSkillPct = avgEl.GetDouble() * 100.0;
    }

    private static void AssignPartyInfo(List<SessionMember> members)
    {
        var groups = members
            .Where(m => !string.IsNullOrEmpty(m.PartyId))
            .GroupBy(m => m.PartyId)
            .Where(g => g.Count() > 1)
            .OrderByDescending(g => g.Count())
            .ToList();

        int idx = 0;
        foreach (var grp in groups)
        {
            foreach (var m in grp)
            {
                m.PartyIndex = idx;
                m.PartySize  = grp.Count();
            }
            idx++;
        }
        // Solo players remain PartyIndex = -1, PartySize = 1
    }

    private static string ExtractSessionTemplate(string url)
    {
        var marker = "/sessiontemplates/";
        int i = url.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (i < 0) return "";
        var rest = url[(i + marker.Length)..];
        int end  = rest.IndexOf('/');
        return end >= 0 ? rest[..end] : rest;
    }

    // ── Auto-leave toggle ─────────────────────────────────────────────────────
    private async void AutoLeaveToggle_Checked(object sender, RoutedEventArgs e)
    {
        _autoLeaveEnabled = true;

        // When toggled ON, immediately leave any active ticket session
        if (!string.IsNullOrEmpty(_lastSessionUrl))
        {
            AddLog("Auto-Leave enabled: leaving active ticket");
            await LeaveSessionAsync(_lastSessionUrl, _lastSessionHeaders);
        }
        else
        {
            AddLog("Auto-Leave enabled: no active ticket to leave");
        }
    }

    private void AutoLeaveToggle_Unchecked(object sender, RoutedEventArgs e)
    {
        _autoLeaveEnabled = false;
        AddLog("Auto-Leave disabled");
    }

    // ── Session log helpers ───────────────────────────────────────────────────
    private void AddLog(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        _sessionLog.Add(line);
        SessionLogScroll.ScrollToEnd();
        SessionLogBorder.Visibility = Visibility.Visible;
    }

    private int ReadThreshold()
    {
        if (int.TryParse(StackThresholdBox.Text.Trim(), out int t) && t >= 1)
            return t;
        return 4;
    }

    private HashSet<string> ReadAvoidList() =>
        new(AvoidListBox.Text
                .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => l.Length > 0),
            StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Runs avoid-list and stack-size detection against rawMembers.
    /// Logs results, updates the alert banner, and fires auto-leave if enabled.
    /// </summary>
    private void RunSessionDetection(List<SessionMember> rawMembers, ProxyCaptureEntry? entry)
    {
        int    threshold = ReadThreshold();
        var    avoidList = ReadAvoidList();
        string myGt      = MyGamertagBox.Text.Trim();

        AddLog($"Session: {rawMembers.Count} players" +
               (string.IsNullOrEmpty(myGt) ? "" : $"  |  me: {myGt}"));

        // ── Avoid list check ─────────────────────────────────────────────────
        bool avoidTriggered = false;
        if (avoidList.Count > 0)
        {
            var hits = rawMembers.Where(m => avoidList.Contains(m.Gamertag)).ToList();
            if (hits.Any())
            {
                avoidTriggered = true;
                AddLog($"Avoid list hit: {string.Join(", ", hits.Select(m => m.Gamertag))}");
            }
        }

        // ── Stack check (everyone except user's own party) ───────────────────
        IEnumerable<SessionMember> checkPool;
        if (!string.IsNullOrEmpty(myGt))
        {
            var me = rawMembers.FirstOrDefault(m =>
                m.Gamertag.Equals(myGt, StringComparison.OrdinalIgnoreCase));
            if (me != null && me.PartySize > 1)
                checkPool = rawMembers.Where(m => m.PartyIndex != me.PartyIndex);
            else
                checkPool = rawMembers.Where(m =>
                    !m.Gamertag.Equals(myGt, StringComparison.OrdinalIgnoreCase));
        }
        else
            checkPool = rawMembers;

        var bigGroup = checkPool
            .Where(m => m.PartySize >= threshold)
            .GroupBy(m => m.PartyIndex)
            .OrderByDescending(g => g.Count())
            .FirstOrDefault();

        bool stackTriggered = bigGroup is not null;
        if (stackTriggered)
            AddLog($"Stack detected: {bigGroup!.Count()}× — {string.Join(", ", bigGroup!.Select(m => m.Gamertag))}");

        // ── Update banner and optionally leave ───────────────────────────────
        bool shouldAlert = avoidTriggered || stackTriggered;
        if (shouldAlert)
        {
            var parts = new List<string>();
            if (stackTriggered)
                parts.Add($"⚠  {bigGroup!.Count()}-STACK: {string.Join(", ", bigGroup!.Select(m => m.Gamertag))}");
            if (avoidTriggered)
            {
                var hitNames = rawMembers
                    .Where(m => avoidList.Contains(m.Gamertag))
                    .Select(m => m.Gamertag);
                parts.Add($"⛔  AVOID: {string.Join(", ", hitNames)}");
            }

            SessionAlertText.Text         = string.Join("  |  ", parts);
            SessionAlertBanner.Visibility = Visibility.Visible;
            System.Media.SystemSounds.Exclamation.Play();

            // Avoid list ALWAYS triggers leave (explicit user intent)
            if (avoidTriggered && entry is not null)
            {
                AddLog("Avoid list triggered: LEAVING");
                _ = LeaveSessionAsync(entry.Url, entry.RequestHeaders);
            }
            // Stack detection only leaves if Auto-Leave is enabled
            else if (stackTriggered && _autoLeaveEnabled && entry is not null)
            {
                AddLog("Stack detected: LEAVING (Auto-Leave enabled)");
                _ = LeaveSessionAsync(entry.Url, entry.RequestHeaders);
            }
            else if (stackTriggered && !_autoLeaveEnabled)
            {
                AddLog("Stack detected: OFF (toggle Auto-Leave to auto-leave stacks)");
            }
        }
        else
        {
            SessionAlertBanner.Visibility = Visibility.Collapsed;
            AddLog("No threats detected");
        }
    }

    /// <summary>
    /// Sends DELETE /members/me to the MPSD session, causing the game to detect
    /// a session loss and abort the loading sequence. Best-effort.
    /// </summary>
    private async Task LeaveSessionAsync(string sessionUrl, Dictionary<string, string> requestHeaders)
    {
        var    uri       = new Uri(sessionUrl);
        string baseUrl   = $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath.TrimEnd('/')}";
        string deleteUrl = baseUrl + "/members/me";

        var statusCode = 0;
        var body       = "";
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Delete, deleteUrl);
            foreach (var (k, v) in requestHeaders)
                if (k.StartsWith("x-",        StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    req.Headers.TryAddWithoutValidation(k, v);

            using var resp = await _mpsdClient.SendAsync(req);
            statusCode = (int)resp.StatusCode;
            body       = await resp.Content.ReadAsStringAsync();
        }
        catch (Exception ex) { body = ex.Message; }

        // Show result in capture log
        AddCaptureEntry(new ProxyCaptureEntry
        {
            Method       = "DELETE",
            Url          = deleteUrl,
            Host         = new Uri(deleteUrl).Host,
            Path         = new Uri(deleteUrl).AbsolutePath,
            StatusCode   = statusCode,
            ResponseBody = body
        });

        // ── Cancel SmartMatch ticket (stops re-queue) ─────────────────────────
        if (!string.IsNullOrEmpty(_ticketScid) &&
            !string.IsNullOrEmpty(_ticketHopper) &&
            !string.IsNullOrEmpty(_ticketId))
        {
            var ticketUrl = $"https://smartmatch.xboxlive.com/serviceconfigs/{_ticketScid}" +
                            $"/hoppers/{_ticketHopper}/tickets/{_ticketId}";
            var tCode = 0;
            var tBody = "";
            try
            {
                using var treq = new HttpRequestMessage(HttpMethod.Delete, ticketUrl);
                foreach (var (k, v) in requestHeaders)
                    if (k.Equals("x-xbl-contract-version", StringComparison.OrdinalIgnoreCase))
                        continue;
                    else if (k.StartsWith("x-",        StringComparison.OrdinalIgnoreCase) ||
                             k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                        treq.Headers.TryAddWithoutValidation(k, v);
                treq.Headers.TryAddWithoutValidation("x-xbl-contract-version", "103");
                using var tresp = await _mpsdClient.SendAsync(treq);
                tCode = (int)tresp.StatusCode;
                tBody = await tresp.Content.ReadAsStringAsync();
            }
            catch (Exception tex) { tBody = tex.Message; }

            AddCaptureEntry(new ProxyCaptureEntry
            {
                Method       = "DELETE",
                Url          = ticketUrl,
                Host         = "smartmatch.xboxlive.com",
                Path         = $"/serviceconfigs/{_ticketScid}/hoppers/{_ticketHopper}/tickets/{_ticketId}",
                StatusCode   = tCode,
                ResponseBody = tBody
            });
        }
    }

    // ── Session Intel button handlers ─────────────────────────────────────────
    private void SessionClear_Click(object sender, RoutedEventArgs e)
    {
        _sessionMembers.Clear();
        _sessionLog.Clear();
        SessionHopperText.Text        = "";
        SessionAlertBanner.Visibility = Visibility.Collapsed;
        SessionAlertText.Text         = "";
        SessionLogBorder.Visibility   = Visibility.Collapsed;
    }

    private void SessionCopyRaw_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(_lastSessionJson))
            Clipboard.SetText(_lastSessionJson);
    }

    private void TestLeave_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastSessionUrl))
        {
            MessageBox.Show(
                "No active ticket session.\n\nStart the proxy, queue for a match, and wait for the Session Intel panel to appear.",
                "No Session", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        AddLog("Manual leave triggered");
        _ = LeaveSessionAsync(_lastSessionUrl, _lastSessionHeaders);
    }

    // ── Hopper population query ────────────────────────────────────────────────
    private async void QueryHopper_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_ticketScid) || string.IsNullOrEmpty(_ticketHopper))
        {
            HopperResultText.Text = "No ticket session captured yet";
            return;
        }

        HopperResultText.Text    = "Querying…";
        QueryHopperBtn.IsEnabled = false;

        var url = $"https://smartmatch.xboxlive.com/serviceconfigs/{_ticketScid}" +
                  $"/hoppers/{_ticketHopper}";
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.TryAddWithoutValidation("Accept", "application/json");
            foreach (var (k, v) in _ticketAuthHeaders)
                req.Headers.TryAddWithoutValidation(k, v);

            using var resp = await _mpsdClient.SendAsync(req);
            var       body = await resp.Content.ReadAsStringAsync();

            AddCaptureEntry(new ProxyCaptureEntry
            {
                Method       = "GET",
                Url          = url,
                Host         = "smartmatch.xboxlive.com",
                Path         = $"/serviceconfigs/{_ticketScid}/hoppers/{_ticketHopper}",
                StatusCode   = (int)resp.StatusCode,
                ResponseBody = body
            });

            if (!resp.IsSuccessStatusCode)
            {
                HopperResultText.Text = $"HTTP {(int)resp.StatusCode}";
                return;
            }

            using var doc   = JsonDocument.Parse(body);
            var       root2 = doc.RootElement;
            var       parts = new List<string>();

            if (root2.TryGetProperty("waitTime",             out var wtEl))  parts.Add($"~{wtEl.GetInt32()}s wait");
            if (root2.TryGetProperty("population",            out var popEl)) parts.Add($"{popEl.GetInt32()} players");
            else if (root2.TryGetProperty("playersSearching", out var psEl))  parts.Add($"{psEl.GetInt32()} searching");

            HopperResultText.Text = parts.Count > 0
                ? string.Join("  |  ", parts)
                : "OK — click entry to inspect JSON";
        }
        catch (Exception ex)
        {
            HopperResultText.Text = $"Error: {ex.Message}";
        }
        finally
        {
            QueryHopperBtn.IsEnabled = true;
        }
    }

    // ── Rejoin Guard: Load/Save Handle and Match Session ────────────────────────
    private void LoadSavedHandle()
    {
        try
        {
            if (!File.Exists(HandleFile))
            {
                AddDiag("LOAD[Handle]", "file not found", HandleFile);
                return;
            }
            var json = File.ReadAllText(HandleFile);
            _savedHandle = JsonSerializer.Deserialize<SavedHandleInfo>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            var h = _savedHandle;
            AddDiag("LOAD[Handle]",
                h is not null ? $"OK  {h.TemplateName}/{h.SessionShort}  saved={h.SavedAtStr}" : "deserialized to null",
                HandleFile);
        }
        catch (Exception ex)
        {
            AddDiag("LOAD[Handle]", $"ERROR: {ex.GetType().Name}: {ex.Message}", HandleFile);
        }
        finally { UpdateRejoinGuardUi(); }
    }

    private void LoadSavedMatchSession()
    {
        try
        {
            if (!File.Exists(MatchSessionFile))
            {
                AddDiag("LOAD[Match]", "file not found", MatchSessionFile);
                return;
            }
            var json = File.ReadAllText(MatchSessionFile);
            _savedMatchSession = JsonSerializer.Deserialize<SavedHandleInfo>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            // Sync to proxy so session discovery injection works
            _proxy.SetSavedMatchSession(_savedMatchSession);

            var m = _savedMatchSession;
            AddDiag("LOAD[Match]",
                m is not null ? $"OK  {m.TemplateName}/{m.SessionShort}  saved={m.SavedAtStr}" : "deserialized to null",
                MatchSessionFile);
        }
        catch (Exception ex)
        {
            AddDiag("LOAD[Match]", $"ERROR: {ex.GetType().Name}: {ex.Message}", MatchSessionFile);
        }
        finally { UpdateRejoinGuardUi(); }
    }

    // ── Rejoin Guard: Button Handlers ─────────────────────────────────────────────
    private async void RejoinCheck_Click(object sender, RoutedEventArgs e)
    {
        if (_savedHandle is null && _savedMatchSession is null) return;
        RejoinCheckBtn.IsEnabled   = false;
        RejoinRestoreBtn.IsEnabled = false;
        RejoinStatusText.Text      = "checking…";
        try
        {
            if (_savedMatchSession is not null)
                await CheckSessionAsync(_savedMatchSession);
            else if (_savedHandle is not null)
                await CheckSessionAsync(_savedHandle);
        }
        finally
        {
            bool hasAny = _savedHandle is not null || _savedMatchSession is not null;
            RejoinCheckBtn.IsEnabled   = hasAny;
            RejoinRestoreBtn.IsEnabled = hasAny;
        }
    }

    private async void RejoinRestore_Click(object sender, RoutedEventArgs e)
    {
        if (_savedHandle is null && _savedMatchSession is null) return;
        RejoinCheckBtn.IsEnabled   = false;
        RejoinRestoreBtn.IsEnabled = false;
        RejoinStatusText.Text      = "restoring…";
        try
        {
            // Re-add to match session first (PUT /members/me)
            if (_savedMatchSession is not null)
                await ReaddToMatchSessionAsync(_savedMatchSession);
            // Then POST activity handles
            if (_savedMatchSession is not null)
                await RestoreRejoinAsync(_savedMatchSession);
            if (_savedHandle is not null)
                await RestoreRejoinAsync(_savedHandle);
        }
        finally
        {
            bool hasAny = _savedHandle is not null || _savedMatchSession is not null;
            RejoinCheckBtn.IsEnabled   = hasAny;
            RejoinRestoreBtn.IsEnabled = hasAny;
        }
    }

    private void RejoinClear_Click(object sender, RoutedEventArgs e)
    {
        ClearRejoinData();
    }

    /// <summary>Log a timestamp marker for when user closes MCC (debugging rejoin flow).</summary>
    private void LogQuit_Click(object sender, RoutedEventArgs e)
    {
        var entry = new ProxyCaptureEntry
        {
            Timestamp = DateTime.Now,
            Method = "DEBUG[UserEvent]",
            Url = "user-action",
            Host = "user-action",
            Path = "/MCC-QUIT",
            RequestBody = "User marked: Closing MCC",
            StatusCode = 0,
            ResponseBody = $"Time: {DateTime.Now:HH:mm:ss.fff}"
        };
        AddCaptureEntry(entry);

        // Force close MCC to simulate crash
        // NOTE: The MCC watcher will automatically detect the exit and arm the block
        var mccProcesses = Process.GetProcessesByName("MCC-Win64-Shipping")
            .Concat(Process.GetProcessesByName("MCC-Win64-Shipping-EAC"))
            .ToList();

        if (mccProcesses.Count == 0)
        {
            AddDiag("MCC-FORCE-CLOSE", "No MCC processes found (may need admin)", "");
        }
        else
        {
            foreach (var proc in mccProcesses)
            {
                try
                {
                    proc.Kill();
                    AddDiag("MCC-FORCE-CLOSE", $"Killed PID {proc.Id} ({proc.ProcessName})", "");
                }
                catch (Exception ex)
                {
                    AddDiag("MCC-FORCE-CLOSE-FAIL", $"Failed to kill PID {proc.Id}: {ex.GetType().Name}", ex.Message);
                }
            }
        }
    }

    /// <summary>Log a timestamp marker for when user clicks rejoin (debugging rejoin flow).</summary>
    private void LogRejoin_Click(object sender, RoutedEventArgs e)
    {
        var entry = new ProxyCaptureEntry
        {
            Timestamp = DateTime.Now,
            Method = "DEBUG[UserEvent]",
            Url = "user-action",
            Host = "user-action",
            Path = "/MCC-REJOIN-CLICK",
            RequestBody = "User marked: Clicking rejoin",
            StatusCode = 0,
            ResponseBody = $"Time: {DateTime.Now:HH:mm:ss.fff}"
        };
        AddCaptureEntry(entry);
    }

    // ── Rejoin Guard: Session Check & Restore ──────────────────────────────────────
    private async Task CheckSessionAsync(SavedHandleInfo handle)
    {
        int    code = 0;
        string body = "";
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, handle.SessionUrl);
            foreach (var (k, v) in handle.RequestHeaders)
                if (k.StartsWith("x-",        StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    req.Headers.TryAddWithoutValidation(k, v);

            using var resp = await _mpsdClient.SendAsync(req);
            code = (int)resp.StatusCode;
            body = await resp.Content.ReadAsStringAsync();
        }
        catch (Exception ex) { body = ex.Message; }

        AddCaptureEntry(new ProxyCaptureEntry
        {
            Method       = "GET",
            Url          = handle.SessionUrl,
            Host         = "sessiondirectory.xboxlive.com",
            Path         = new Uri(handle.SessionUrl).AbsolutePath,
            StatusCode   = code,
            ResponseBody = body,
        });

        var (text, color) = code switch
        {
            200        => ("✓ session alive",                                   Color.FromRgb(0x3F, 0xB9, 0x50)),
            401 or 403 => ("✗ auth expired — restart proxy to refresh",         Color.FromRgb(0xF8, 0x51, 0x49)),
            404        => ("✗ session dead (expired after crash)",               Color.FromRgb(0xF8, 0x51, 0x49)),
            0          => ("✗ request failed — check proxy / network",           Color.FromRgb(0xF8, 0x51, 0x49)),
            _          => ($"HTTP {code}",                                       Color.FromRgb(0xD2, 0x99, 0x22)),
        };
        RejoinStatusText.Text       = text;
        RejoinStatusText.Foreground = new SolidColorBrush(color);
    }

    /// <summary>
    /// Re-adds the player to the matchmaking session by PUTting /members/me.
    /// This ensures session discovery returns the match on MCC relaunch.
    /// cascadesquadsession requires a live WebSocket (connectionRequiredForActiveMembers),
    /// but CascadeMatchmaking may not — so we attempt the PUT there only.
    /// </summary>
    private async Task ReaddToMatchSessionAsync(SavedHandleInfo match)
    {
        // Step 1: GET the session to verify it's alive and grab current ETag
        string etag = "";
        int getCode = 0;
        try
        {
            using var getReq = new HttpRequestMessage(HttpMethod.Get, match.SessionUrl);
            foreach (var (k, v) in match.RequestHeaders)
                if (k.StartsWith("x-", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    getReq.Headers.TryAddWithoutValidation(k, v);

            using var getResp = await _mpsdClient.SendAsync(getReq);
            getCode = (int)getResp.StatusCode;
            etag = getResp.Headers.ETag?.Tag ?? "";
        }
        catch { return; }

        AddCaptureEntry(new ProxyCaptureEntry
        {
            Method       = "GET[ETag]",
            Url          = match.SessionUrl,
            Host         = "sessiondirectory.xboxlive.com",
            Path         = new Uri(match.SessionUrl).AbsolutePath,
            RequestBody  = "Pre-PUT ETag fetch for member re-add",
            StatusCode   = getCode,
            ResponseBody = string.IsNullOrEmpty(etag) ? "[no ETag]" : $"ETag: {etag}",
        });

        if (getCode != 200 || string.IsNullOrEmpty(etag)) return;

        // Step 2: PUT to session URL with members/me to re-add ourselves
        // IMPORTANT: Do NOT change the connection GUID! The game server has the original GUID
        // and validates rejoin attempts against it. Only set active:true without a new GUID.
        var putBody = "{\"members\":{\"me\":{\"properties\":{\"system\":{\"active\":true},\"custom\":{}}}}}";

        int putCode = 0;
        string putResp = "";
        try
        {
            using var putReq = new HttpRequestMessage(HttpMethod.Put, match.SessionUrl);
            putReq.Content = new StringContent(putBody, System.Text.Encoding.UTF8, "application/json");
            putReq.Headers.TryAddWithoutValidation("If-Match", etag);
            foreach (var (k, v) in match.RequestHeaders)
                if (k.StartsWith("x-", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    putReq.Headers.TryAddWithoutValidation(k, v);

            using var resp = await _mpsdClient.SendAsync(putReq);
            putCode = (int)resp.StatusCode;
            putResp = await resp.Content.ReadAsStringAsync();
        }
        catch (Exception ex) { putResp = ex.Message; }

        AddCaptureEntry(new ProxyCaptureEntry
        {
            Method       = "PUT[REJOIN]",
            Url          = match.SessionUrl,
            Host         = "sessiondirectory.xboxlive.com",
            Path         = new Uri(match.SessionUrl).AbsolutePath,
            RequestBody  = putBody,
            StatusCode   = putCode,
            ResponseBody = putResp,
        });

        var (text, color) = putCode switch
        {
            200 or 204 => ("✓ re-added to match session",                          Color.FromRgb(0x3F, 0xB9, 0x50)),
            412        => ("✗ ETag conflict — session changed",                     Color.FromRgb(0xD2, 0x99, 0x22)),
            404        => ("✗ match session expired",                               Color.FromRgb(0xF8, 0x51, 0x49)),
            401 or 403 => ("✗ auth expired",                                        Color.FromRgb(0xF8, 0x51, 0x49)),
            _          => ($"PUT /members/me → HTTP {putCode}",                     Color.FromRgb(0xD2, 0x99, 0x22)),
        };
        RejoinStatusText.Text       = text;
        RejoinStatusText.Foreground = new SolidColorBrush(color);
    }

    private async Task RestoreRejoinAsync(SavedHandleInfo handle)
    {
        // POST /handles — re-create the activity handle so MCC can find the session on relaunch.
        var handleBody =
            $$"""{"type":"activity","sessionRef":{"scid":"{{handle.Scid}}","templateName":"{{handle.TemplateName}}","name":"{{handle.SessionName}}"},"version":1}""";
        int    handleCode = 0;
        string handleResp = "";
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Post,
                "https://sessiondirectory.xboxlive.com/handles");
            req.Content = new StringContent(handleBody, System.Text.Encoding.UTF8, "application/json");
            foreach (var (k, v) in handle.RequestHeaders)
                if (k.StartsWith("x-",        StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                    req.Headers.TryAddWithoutValidation(k, v);
            using var resp = await _mpsdClient.SendAsync(req);
            handleCode = (int)resp.StatusCode;
            handleResp = await resp.Content.ReadAsStringAsync();
        }
        catch (Exception ex) { handleResp = ex.Message; }

        AddCaptureEntry(new ProxyCaptureEntry
        {
            Method       = "POST",
            Url          = "https://sessiondirectory.xboxlive.com/handles",
            Host         = "sessiondirectory.xboxlive.com",
            Path         = "/handles",
            RequestBody  = handleBody,
            StatusCode   = handleCode,
            ResponseBody = handleResp,
        });

        var (text, color) = handleCode switch
        {
            200 or 201 => ("✓ handle restored — launch MCC and accept rejoin",  Color.FromRgb(0x3F, 0xB9, 0x50)),
            404        => ("✗ session expired — cannot restore",                 Color.FromRgb(0xF8, 0x51, 0x49)),
            401 or 403 => ("✗ auth expired — restart proxy to refresh",          Color.FromRgb(0xF8, 0x51, 0x49)),
            0          => ("✗ request failed — check proxy / network",           Color.FromRgb(0xF8, 0x51, 0x49)),
            _          => ($"HTTP {handleCode}",                                 Color.FromRgb(0xD2, 0x99, 0x22)),
        };
        RejoinStatusText.Text       = text;
        RejoinStatusText.Foreground = new SolidColorBrush(color);
    }

    // ── Rejoin Guard: MCC Process Monitor ───────────────────────────────────────
    private int _watcherTickCount = 0;

    private void MccWatcher_Tick(object? sender, EventArgs e)
    {
        try
        {
            _watcherTickCount++;

            // DEFENSIVE LOAD: Ensure session is always fresh before crash detection.
            // If MCC crashes before a PUT is captured, _savedMatchSession could be null
            // even though the session file exists on disk. This ensures ghost mode can
            // activate by guaranteeing the session is loaded when needed.
            // Only attempt if the file actually exists — avoids spamming LOAD[Match] "file not
            // found" every tick when there is no session (e.g. after CLEAR[Rejoin]).
            if (_savedMatchSession is null && File.Exists(MatchSessionFile))
                LoadSavedMatchSession();

            bool running = Process.GetProcessesByName("MCC-Win64-Shipping").Length > 0
                        || Process.GetProcessesByName("MCC-Win64-Shipping-EAC").Length > 0;

            // CRITICAL FIX: Track MCC process ID to detect fast restarts
            // If MCC restarts within the watcher interval (now 1s), Process.GetProcessesByName
            // might show it as "running" even though it's a NEW process with a different PID.
            // This detects that scenario and triggers crash restore.
            int currentMccPid = 0;
            if (running)
            {
                var mccProcess = Process.GetProcessesByName("MCC-Win64-Shipping")
                    .Concat(Process.GetProcessesByName("MCC-Win64-Shipping-EAC"))
                    .FirstOrDefault();
                if (mccProcess is not null)
                    currentMccPid = mccProcess.Id;
            }

            // Detect if MCC restarted (PID changed) while appearing "running"
            if (running && _mccWasRunning && _lastMccPid > 0 && currentMccPid != _lastMccPid && !_restoreInProgress)
            {
                AddDiag("WATCHER[FastRestart]", $"MCC PID changed ({_lastMccPid} → {currentMccPid}) — treating as crash", "");
                _mccWasRunning = false;  // Force the exit handler to trigger
            }

            if (running)
                _lastMccPid = currentMccPid;  // Update PID tracker

            // Heartbeat: Log every 10 ticks so we can see watcher is alive and what state it sees
            if (_watcherTickCount % 10 == 0)
                AddDiag("WATCHER[Heartbeat]", $"tick={_watcherTickCount}  running={running}  wasRunning={_mccWasRunning}  session={(_savedMatchSession?.SessionShort ?? "null")}  inProgress={_restoreInProgress}", "");

            if (_mccWasRunning && !running && !_restoreInProgress &&
                (_savedHandle is not null || _savedMatchSession is not null))
            {
                // MCC just exited — arm the proxy for just-in-time injection on restart
                _mccWasRunning      = false;
                _restoreInProgress  = true;

                // Diagnostic: log what state we have at restore time
                AddDiag("RESTORE[Start]",
                    $"handle={(_savedHandle is not null ? _savedHandle.TemplateName + "/" + _savedHandle.SessionShort : "null")}  " +
                    $"match={(_savedMatchSession is not null ? _savedMatchSession.TemplateName + "/" + _savedMatchSession.SessionShort : "null")}");

                // Auto-arm block-match-leave: when MCC exits (whether clicked or crashed),
                // arm the proxy to block the next leave request, keeping the session alive
                _proxy.ForceBlockMatchLeave();
                AddDiag("AUTO-BLOCK[Armed]", "MCC exit detected — block armed for next leave request", "");

                // Arm the proxy: on MCC's first sessiondirectory request, the proxy will:
                //   1. PUT /members/me to the match session (JIT, with MCC's fresh auth)
                //   2. POST activity handle for the match session
                //   3. Force-replace session discovery results (if MCC does discovery)
                if (_savedMatchSession is not null)
                    _proxy.SetPendingCrashRestore(_savedMatchSession);

                _restoreInProgress = false;
            }
            else if (running && !_mccWasRunning)
            {
                // MCC just launched — clear the restore debounce
                _restoreInProgress = false;
            }

            _mccWasRunning = running;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[ERROR] MccWatcher_Tick: {ex.GetType().Name}: {ex.Message}");
        }
    }
}
