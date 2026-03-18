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
    // HttpClient for out-of-band MPSD/SmartMatch requests (auto-leave, hopper query)
    // UseProxy=false to avoid routing through our own proxy.
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

    // ── Rejoin Guard / MCC Crash Detection ────────────────────────────────────
    private readonly System.Windows.Threading.DispatcherTimer _mccWatcher =
        new() { Interval = TimeSpan.FromMilliseconds(250) };
    private bool _mccWasRunning    = false;
    private int  _lastMccPid       = 0;   // Track PID to detect fast MCC restarts
    private bool _restoreInProgress = false;

    public Scanner()
    {
        InitializeComponent();

        _proxy.OnRequestCaptured += (_, entry) =>
            Dispatcher.InvokeAsync(() => AddCaptureEntry(entry));

        _proxy.WinHttpManualSetRequired += (_, _) =>
            Dispatcher.InvokeAsync(() => ProxyWinHttpNote.Visibility = Visibility.Visible);

        _proxy.OnSessionCaptured += (_, _) =>
            Dispatcher.InvokeAsync(() => UpdateRejoinGuardUi());

        _proxy.OnPartyCacheChanged += (_, _) =>
            Dispatcher.InvokeAsync(() => UpdateRejoinGuardUi());

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
            UpdateRejoinGuardUi();
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
    /// Called to refresh the Rejoin Guard panel from proxy state.
    /// </summary>
    private void UpdateRejoinGuardUi()
    {
        var captured = _proxy.CapturedSession;

        if (captured is null)
        {
            RejoinSessionText.Text = "no session captured";
            RejoinMatchText.Text   = "";
        }
        else
        {
            RejoinSessionText.Text = $"{captured.TemplateName} / {captured.SessionShort}   {captured.SavedAtStr}";
            RejoinMatchText.Text   = _proxy.IsInCrashWindow
                ? $"CRASH DETECTED — {CrashWindowSeconds}s injection window active"
                : "";
        }

        RejoinStatusText.Text      = "";
        RejoinCheckBtn.IsEnabled   = false;
        RejoinRestoreBtn.IsEnabled = false;

        var cachedAt  = _proxy.PartyCachedAt;
        var tsStr     = cachedAt.HasValue ? $"  {cachedAt.Value:HH:mm:ss}" : "";

        // Determine if the party predates the current session capture.
        // If it does, warn: the cached party might be from a different match.
        bool partyPredatesSession = cachedAt.HasValue &&
                                    captured is not null &&
                                    cachedAt.Value < captured.SavedAt.ToLocalTime();

        switch (_proxy.PartyCache)
        {
            case ProxyService.PartyCacheState.None:
                RejoinPartyText.Text       = "party: not cached";
                RejoinPartyText.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0x7D, 0x85, 0x90));
                break;
            case ProxyService.PartyCacheState.Priming:
                RejoinPartyText.Text       = "party: priming…";
                RejoinPartyText.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0xFF, 0xC0, 0x30));
                break;
            case ProxyService.PartyCacheState.Stale:
                RejoinPartyText.Text       = $"party: stale (not cached){tsStr}";
                RejoinPartyText.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0xFF, 0x7B, 0x25));
                break;
            case ProxyService.PartyCacheState.Ready:
            {
                var cachedIp = _proxy.CachedPartyIp;
                var liveIp   = _proxy.LivePartyIp;
                bool ipsKnown = !string.IsNullOrEmpty(cachedIp) && !string.IsNullOrEmpty(liveIp);
                bool ipMatch  = ipsKnown && string.Equals(cachedIp, liveIp, StringComparison.OrdinalIgnoreCase);

                if (ipsKnown && !ipMatch)
                {
                    // We know both — they differ. Cache will redirect on crash.
                    RejoinPartyText.Text       = $"party: cached={cachedIp}  live={liveIp}{tsStr}  (redirect on crash)";
                    RejoinPartyText.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0xFF, 0xC0, 0x30));
                }
                else if (ipsKnown)
                {
                    // Both known, match — ideal.
                    RejoinPartyText.Text       = $"party: ready ✓  {cachedIp}{tsStr}";
                    RejoinPartyText.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x3F, 0xB9, 0x50));
                }
                else if (partyPredatesSession)
                {
                    var lag    = captured!.SavedAt.ToLocalTime() - cachedAt!.Value;
                    var lagStr = lag.TotalMinutes >= 1
                        ? $"{(int)lag.TotalMinutes}m before session"
                        : $"{(int)lag.TotalSeconds}s before session";
                    RejoinPartyText.Text       = $"party: old ⚠  {cachedIp}  {cachedAt!.Value:HH:mm:ss}  ({lagStr})";
                    RejoinPartyText.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0xFF, 0x7B, 0x25));
                }
                else
                {
                    RejoinPartyText.Text       = $"party: ready ✓  {cachedIp}{tsStr}";
                    RejoinPartyText.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0x3F, 0xB9, 0x50));
                }
                break;
            }
        }
    }

    private const int CrashWindowSeconds = 60;

    /// <summary>Clears the captured session and crash window from proxy state.</summary>
    public void ClearRejoinData()
    {
        _proxy.ClearCapturedSession();
        UpdateRejoinGuardUi();
        AddDiag("CLEAR[Rejoin]", "Cleared captured session from proxy");
    }

    // ── Capture entry management ──────────────────────────────────────────────
    private void AddCaptureEntry(ProxyCaptureEntry entry)
    {
        _captureLog.Insert(0, entry);
        if (_captureLog.Count > 500)
            _captureLog.RemoveAt(_captureLog.Count - 1);

        TryParseSessionEntry(entry);

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
        "CAPTURE[", "CRASH[", "DISCOVERY[", "WATCHER[",
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

    // ── Rejoin Guard: Button Handlers ────────────────────────────────────────────
    private void RejoinCheck_Click(object sender, RoutedEventArgs e)   { /* proxy handles discovery automatically */ }
    private void RejoinRestore_Click(object sender, RoutedEventArgs e) { /* proxy handles discovery automatically */ }
    private void RejoinClear_Click(object sender, RoutedEventArgs e)   { ClearRejoinData(); }

    private async void PrimeParty_Click(object sender, RoutedEventArgs e)
    {
        PrimePartyBtn.IsEnabled = false;
        PrimePartyBtn.Content   = "⚡ PRIMING…";
        try
        {
            await _proxy.PrimeCacheAsync(auto: false);
            UpdateRejoinGuardUi();
        }
        finally
        {
            PrimePartyBtn.IsEnabled = true;
            PrimePartyBtn.Content   = "⚡ PRIME PARTY";
        }
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

    // ── AI Capture Analysis ───────────────────────────────────────────────────

    private static readonly string ApiKeyPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox", "claude-api-key.txt");

    private static readonly HttpClient _claudeClient =
        new(new HttpClientHandler { UseProxy = false }) { Timeout = TimeSpan.FromSeconds(60) };

    private const string ClaudeSystemPrompt =
        "You are analyzing network captures from Halo: MCC's multiplayer session system. " +
        "The app is a MITM proxy that intercepts HTTPS traffic between Halo MCC and Xbox Live / PlayFab.\n\n" +
        "Key concepts:\n" +
        "- MPSD (sessiondirectory.xboxlive.com): Xbox Live Multiplayer Session Directory. Manages session membership.\n" +
        "- connectionRequiredForActiveMembers=true: members must have a live RTA WebSocket connection GUID or MPSD evicts them.\n" +
        "- RTA connection GUID: a UUID tied to the player's live Xbox Live WebSocket. Becomes stale/dead after a crash.\n" +
        "- CascadeMatchmaking: the MPSD session template for active matches (8 players).\n" +
        "- cascadesquadsession: the squad-level session (1-4 players, links to the match).\n" +
        "- CAPTURE[Session]: proxy saved a CascadeMatchmaking session during normal play.\n" +
        "- CRASH[Detected]: MCC process exited — crash recovery window armed (30s).\n" +
        "- DISCOVERY[Found/Injected]: proxy handling /sessions? discovery response.\n" +
        "- REJOIN[LeaveBlocked]: proxy blocked a PUT {members:{me:null}} leave during recovery.\n" +
        "- REJOIN[ConnectionUpdate]: proxy fired background PUT to update stale RTA connection GUID in MPSD.\n" +
        "- RequestParty (PlayFab): MCC requests the game server party. Validates MPSD membership.\n\n" +
        "Analyze the capture entries provided. Focus on: whether the rejoin flow succeeded or failed, " +
        "what HTTP status codes indicate, timing of key events, and what should be fixed or tried next. " +
        "Be specific — cite entry methods, status codes, and timestamps from the data.";

    private async void Analyze_Click(object sender, RoutedEventArgs e)
    {
        AnalyzeBtn.IsEnabled = false;
        AnalyzeBtn.Content   = "⏳ ANALYZING...";

        try
        {
            var apiKey = await GetOrPromptApiKeyAsync();
            if (string.IsNullOrEmpty(apiKey)) return;

            // Collect entries: from last CRASH[Detected] forward, or last 75 if no crash
            var allEntries = _captureLog.ToList();
            int crashIdx   = -1;
            for (int i = allEntries.Count - 1; i >= 0; i--)
            {
                if (allEntries[i].Method.StartsWith("CRASH[", StringComparison.OrdinalIgnoreCase))
                { crashIdx = i; break; }
            }

            var subset = crashIdx >= 0
                ? allEntries.Skip(crashIdx).ToList()
                : allEntries.TakeLast(75).ToList();

            // Serialize entries compactly
            var entrySummaries = subset.Select(en => new
            {
                t      = en.Timestamp.ToString("HH:mm:ss.fff"),
                method = en.Method,
                status = en.StatusCode,
                path   = en.Path.Length > 120 ? en.Path[..120] + "…" : en.Path,
                req    = en.RequestBody.Length  > 300 ? en.RequestBody[..300]  + "…" : en.RequestBody,
                resp   = en.ResponseBody.Length > 300 ? en.ResponseBody[..300] + "…" : en.ResponseBody,
            });

            var captureJson = JsonSerializer.Serialize(entrySummaries,
                new JsonSerializerOptions { WriteIndented = false });

            var userMessage = $"Here are {subset.Count} capture entries" +
                (crashIdx >= 0 ? " starting from the last crash detection" : " (last 75)") +
                $":\n\n{captureJson}";

            var analysis = await SendToClaudeAsync(apiKey, userMessage);
            ShowAnalysisWindow(analysis, subset.Count);
        }
        catch (Exception ex)
        {
            ShowAnalysisWindow($"Error: {ex.Message}", 0);
        }
        finally
        {
            AnalyzeBtn.IsEnabled = true;
            AnalyzeBtn.Content   = "🤖 ANALYZE";
        }
    }

    private async Task<string> GetOrPromptApiKeyAsync()
    {
        if (File.Exists(ApiKeyPath))
        {
            var stored = (await File.ReadAllTextAsync(ApiKeyPath)).Trim();
            if (!string.IsNullOrEmpty(stored)) return stored;
        }

        // Prompt for key
        var win = new Window
        {
            Title           = "Claude API Key",
            Width           = 480, Height = 140,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner           = Window.GetWindow(this),
            ResizeMode      = ResizeMode.NoResize,
            Background      = new SolidColorBrush(Color.FromRgb(0x14, 0x18, 0x20)),
        };
        var box = new TextBox
        {
            Margin      = new Thickness(12, 16, 12, 8),
            FontFamily  = new System.Windows.Media.FontFamily("Consolas"),
            Background  = new SolidColorBrush(Color.FromRgb(0x0A, 0x0C, 0x10)),
            Foreground  = new SolidColorBrush(Color.FromRgb(0xC8, 0xD8, 0xE8)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x2A, 0x32, 0x38)),
            Padding     = new Thickness(6, 4, 6, 4),
            Text        = "",
        };
        var btn = new Button
        {
            Content             = "Save & Continue",
            Margin              = new Thickness(12, 0, 12, 12),
            HorizontalAlignment = HorizontalAlignment.Right,
            Padding             = new Thickness(12, 4, 12, 4),
        };
        string result = "";
        btn.Click += (_, _) => { result = box.Text.Trim(); win.Close(); };
        var panel = new StackPanel();
        panel.Children.Add(new TextBlock
        {
            Text       = "Enter your Anthropic API key (saved locally, never transmitted elsewhere):",
            Foreground = new SolidColorBrush(Color.FromRgb(0x7D, 0x85, 0x90)),
            FontFamily = new System.Windows.Media.FontFamily("Consolas"),
            FontSize   = 10,
            Margin     = new Thickness(12, 12, 12, 0),
        });
        panel.Children.Add(box);
        panel.Children.Add(btn);
        win.Content = panel;
        win.ShowDialog();

        if (string.IsNullOrEmpty(result)) return "";
        Directory.CreateDirectory(Path.GetDirectoryName(ApiKeyPath)!);
        await File.WriteAllTextAsync(ApiKeyPath, result);
        return result;
    }

    private static async Task<string> SendToClaudeAsync(string apiKey, string userMessage)
    {
        var payload = JsonSerializer.Serialize(new
        {
            model      = "claude-opus-4-6",
            max_tokens = 2048,
            system     = ClaudeSystemPrompt,
            messages   = new[] { new { role = "user", content = userMessage } },
        });

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages")
        {
            Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json"),
        };
        req.Headers.TryAddWithoutValidation("x-api-key",          apiKey);
        req.Headers.TryAddWithoutValidation("anthropic-version",  "2023-06-01");

        var resp     = await _claudeClient.SendAsync(req);
        var respBody = await resp.Content.ReadAsStringAsync();

        if (!resp.IsSuccessStatusCode)
            return $"API error {(int)resp.StatusCode}: {respBody}";

        using var doc   = JsonDocument.Parse(respBody);
        var       root  = doc.RootElement;
        if (root.TryGetProperty("content", out var content) &&
            content.GetArrayLength() > 0 &&
            content[0].TryGetProperty("text", out var text))
            return text.GetString() ?? "(empty response)";

        return "(could not parse response)";
    }

    private void ShowAnalysisWindow(string analysis, int entryCount)
    {
        var win = new Window
        {
            Title                 = $"Claude Analysis — {entryCount} entries",
            Width                 = 800, Height = 600,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner                 = Window.GetWindow(this),
            Background            = new SolidColorBrush(Color.FromRgb(0x0A, 0x0C, 0x10)),
        };

        var copyBtn = new Button
        {
            Content             = "Copy",
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin              = new Thickness(8),
            Padding             = new Thickness(12, 4, 12, 4),
        };
        copyBtn.Click += (_, _) => Clipboard.SetText(analysis);

        var textBox = new TextBox
        {
            Text            = analysis,
            IsReadOnly      = true,
            TextWrapping    = TextWrapping.Wrap,
            AcceptsReturn   = true,
            FontFamily      = new System.Windows.Media.FontFamily("Consolas"),
            FontSize        = 11,
            Background      = new SolidColorBrush(Color.FromRgb(0x0A, 0x0C, 0x10)),
            Foreground      = new SolidColorBrush(Color.FromRgb(0xC8, 0xD8, 0xE8)),
            BorderThickness = new Thickness(0),
            Padding         = new Thickness(10),
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };

        var dock = new DockPanel();
        DockPanel.SetDock(copyBtn, Dock.Bottom);
        dock.Children.Add(copyBtn);
        dock.Children.Add(new ScrollViewer
        {
            Content = textBox,
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        });
        win.Content = dock;
        win.Show();
    }

    // ── MCC Process Monitor ───────────────────────────────────────────────────
    private int _watcherTickCount = 0;

    private void MccWatcher_Tick(object? sender, EventArgs e)
    {
        try
        {
            _watcherTickCount++;

            bool running = Process.GetProcessesByName("MCC-Win64-Shipping").Length > 0
                        || Process.GetProcessesByName("MCC-Win64-Shipping-EAC").Length > 0;

            // Track PID to detect fast restarts: if MCC restarts within the 1s poll window,
            // GetProcessesByName still returns "running" but with a different PID.
            int currentMccPid = 0;
            if (running)
            {
                var mccProcess = Process.GetProcessesByName("MCC-Win64-Shipping")
                    .Concat(Process.GetProcessesByName("MCC-Win64-Shipping-EAC"))
                    .FirstOrDefault();
                if (mccProcess is not null)
                    currentMccPid = mccProcess.Id;
            }

            if (running && _mccWasRunning && _lastMccPid > 0 && currentMccPid != _lastMccPid && !_restoreInProgress)
            {
                AddDiag("WATCHER[FastRestart]", $"MCC PID changed ({_lastMccPid} → {currentMccPid}) — treating as crash");
                _mccWasRunning = false;  // force the exit handler below to trigger
            }

            if (running)
                _lastMccPid = currentMccPid;

            // Heartbeat every 10 ticks
            if (_watcherTickCount % 10 == 0)
                AddDiag("WATCHER[Heartbeat]", $"tick={_watcherTickCount}  running={running}  wasRunning={_mccWasRunning}  captured={(_proxy.CapturedSession?.SessionShort ?? "null")}  crashWindow={_proxy.IsInCrashWindow}");

            if (_mccWasRunning && !running && !_restoreInProgress)
            {
                _mccWasRunning     = false;
                _restoreInProgress = true;

                var captured = _proxy.CapturedSession;
                AddDiag("CRASH[Detected]",
                    $"MCC exited  captured={captured?.TemplateName + "/" + (captured?.SessionShort ?? "null")}");

                // Open the 30s discovery-injection window.
                // The proxy will inject sessionRef into /sessions? if MPSD no longer has us.
                _proxy.SetCrashDetected();
                UpdateRejoinGuardUi();

                _restoreInProgress = false;
            }
            else if (running && !_mccWasRunning)
            {
                _restoreInProgress = false;
                UpdateRejoinGuardUi();
            }

            _mccWasRunning = running;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[ERROR] MccWatcher_Tick: {ex.GetType().Name}: {ex.Message}");
        }
    }

}

