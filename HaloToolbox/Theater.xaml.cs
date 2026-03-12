using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace HaloToolbox;

public partial class Theater : UserControl
{
    // ── Static config ──────────────────────────────────────────────────────────

    private static readonly string LocalLow = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "AppData", "LocalLow");

    private static readonly string TheaterRoot =
        Path.Combine(LocalLow, "MCC", "Temporary", "UserContent");

    private static readonly string BackupRoot = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox", "TheaterBackups");

    private static readonly string CustomNamesFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox", "theater-names.json");

    // Ordered game keys (determines CboGame indices 1-5)
    private static readonly string[] GameKeys =
        ["Halo2A", "Halo3", "Halo3ODST", "Halo4", "HaloReach"];

    private static readonly Dictionary<string, (string DisplayName, Color Accent)> GameInfo = new()
    {
        ["Halo2A"]    = ("Halo 2: Anniv.", Color.FromRgb(0x39, 0xD0, 0xC8)),
        ["Halo3"]     = ("Halo 3",         Color.FromRgb(0x58, 0xA6, 0xFF)),
        ["Halo3ODST"] = ("Halo 3: ODST",   Color.FromRgb(0xD2, 0x99, 0x22)),
        ["Halo4"]     = ("Halo 4",         Color.FromRgb(0xF8, 0x51, 0x49)),
        ["HaloReach"] = ("Halo: Reach",    Color.FromRgb(0xBC, 0x8C, 0xF9)),
    };

    // ── Map name resolution (Halo 3) ───────────────────────────────────────────
    // Reversed from MainWindow.MapToTheaterPrefix. Longer prefixes ordered first
    // where ambiguity exists (e.g. asq_chillou must precede asq_chill_).
    private static readonly (string Prefix, string DisplayName)[] _h3Prefixes =
    [
        ("asq_chillou", "Cold Storage"),   // Cold Storage — MUST come before asq_chill_
        ("asq_chill_",  "Narrows"),
        ("asq_armory_", "Rat's Nest"),
        ("asq_docks_",  "Longshore"),
        ("asq_shrine_", "Sandtrap"),
        ("asq_sidewin", "Avalanche"),
        ("asq_descent", "Assembly"),
        ("asq_lockout", "Blackout"),
        ("asq_fortres", "Citadel"),
        ("asq_constru", "Construct"),
        ("asq_s3d_edg", "Edge"),
        ("asq_salvati", "Epitaph"),
        ("asq_warehou", "Foundry"),
        ("asq_ghostto", "Ghost Town"),
        ("asq_guardia", "Guardian"),
        ("asq_midship", "Heretic"),
        ("asq_deadloc", "High Ground"),
        ("asq_s3d_tur", "Icebox"),
        ("asq_isolati", "Isolation"),
        ("asq_zanziba", "Last Resort"),
        ("asq_spaceca", "Orbital"),
        ("asq_sandbox", "Sandbox"),
        ("asq_snowbou", "Snowbound"),
        ("asq_bunkerw", "Standoff"),
        ("asq_cyberde", "The Pit"),
        ("asq_riverwo", "Valhalla"),
        ("asq_s3d_wat", "Waterfall"),
    ];

    /// <summary>
    /// Returns a human-readable map name for a theater .mov file.
    /// Halo 3: exact prefix lookup from known asq_* naming convention.
    /// Other games: best-effort cleanup (strip engine prefixes, title-case).
    /// </summary>
    private static string ResolveMapName(string gameKey, string fileName)
    {
        var noExt = Path.GetFileNameWithoutExtension(fileName);

        if (gameKey == "Halo3")
        {
            foreach (var (prefix, displayName) in _h3Prefixes)
                if (noExt.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return displayName;
        }

        return CleanMapFileName(noExt);
    }

    /// <summary>
    /// Generic cleanup: strip engine prefixes, level-number prefixes (Reach),
    /// trailing hash/timestamp segments, then title-case.
    /// </summary>
    private static string CleanMapFileName(string noExt)
    {
        var s = noExt;

        // Strip known engine prefixes
        foreach (var prefix in new[] { "asq_", "dlc_", "mp_", "ffa_", "coop_", "ms_" })
            if (s.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            { s = s[prefix.Length..]; break; }

        // Strip leading level-number prefix (Reach: "20_sword_slayer" → "sword_slayer")
        if (s.Length > 3 && char.IsDigit(s[0]) && char.IsDigit(s[1]) && s[2] == '_')
            s = s[3..];

        // Split on underscores; stop at first segment that looks like a hash/timestamp
        // (6+ chars of only hex digits — covers both decimal timestamps and hex hashes)
        var meaningful = new List<string>();
        foreach (var part in s.Split('_', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part.Length >= 6 && part.All(c => char.IsDigit(c) ||
                (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')))
                break; // hit a hash/timestamp — stop

            meaningful.Add(char.ToUpperInvariant(part[0]) + part[1..].ToLowerInvariant());
        }

        return meaningful.Count > 0 ? string.Join(" ", meaningful) : noExt;
    }

    // ── Instance state ─────────────────────────────────────────────────────────

    private readonly ObservableCollection<TheaterClip> _clips = new();
    private readonly ICollectionView _view;
    private readonly List<FileSystemWatcher> _watchers = new();
    private readonly Dictionary<string, string> _customNames = new(); // "{GameKey}:{FileName}" → name

    private string _filterGame   = "";
    private string _filterSearch = "";

    // ── Constructor ────────────────────────────────────────────────────────────

    public Theater()
    {
        InitializeComponent();

        // Initialize view first so SelectionChanged handlers can safely reference it
        _view = CollectionViewSource.GetDefaultView(_clips);
        _view.Filter = FilterClip;
        ClipList.ItemsSource = _view;

        // Game filter combo
        CboGame.Items.Add("All Games");
        foreach (var key in GameKeys)
            CboGame.Items.Add(GameInfo[key].DisplayName);
        CboGame.SelectedIndex = 0;

        // Sort combo
        CboSort.Items.Add("Newest First");
        CboSort.Items.Add("Oldest First");
        CboSort.Items.Add("Map A–Z");
        CboSort.SelectedIndex = 0;

        TxtBackupPath.Text = $"BACKUP: {BackupRoot}";
    }

    // ── Lifecycle ──────────────────────────────────────────────────────────────

    private void Theater_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            Directory.CreateDirectory(BackupRoot);
            LoadCustomNames();
            InitialScan();
            InitializeWatchers();
        }
        catch (Exception ex)
        {
            try { TxtStatus.Text = $"LOAD ERROR: {ex.GetType().Name}: {ex.Message}"; }
            catch { System.Windows.MessageBox.Show($"Theater load error:\n{ex}", "Theater Error"); }
        }
    }

    private void Theater_Unloaded(object sender, RoutedEventArgs e)
    {
        foreach (var w in _watchers) { w.EnableRaisingEvents = false; w.Dispose(); }
        _watchers.Clear();
    }

    // ── Custom names persistence ───────────────────────────────────────────────

    private void LoadCustomNames()
    {
        _customNames.Clear();
        if (!File.Exists(CustomNamesFile)) return;
        try
        {
            var json = File.ReadAllText(CustomNamesFile);
            var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
            if (dict is null) return;
            foreach (var (k, v) in dict)
                if (!string.IsNullOrWhiteSpace(v)) _customNames[k] = v;
        }
        catch { /* corrupt file — start fresh */ }
    }

    private void SaveCustomNames()
    {
        try
        {
            var json = JsonSerializer.Serialize(_customNames,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(CustomNamesFile, json);
        }
        catch { /* non-critical */ }
    }

    private string CustomNamesKey(TheaterClip clip) => $"{clip.GameKey}:{clip.FileName}";

    // ── Initial scan ───────────────────────────────────────────────────────────

    private void InitialScan()
    {
        _clips.Clear();

        // Pass 1: clips present in source folders
        foreach (var key in GameKeys)
        {
            try
            {
                var sourceFolder = Path.Combine(TheaterRoot, key, "Movie");
                if (!Directory.Exists(sourceFolder)) continue;

                var backupFolder = Path.Combine(BackupRoot, key);
                Directory.CreateDirectory(backupFolder);

                foreach (var file in Directory.EnumerateFiles(sourceFolder, "*.mov"))
                {
                    try
                    {
                        var fi   = new FileInfo(file);
                        var clip = BuildClip(key, fi, sourcePresent: true);
                        _clips.Add(clip);
                        BackupIfNeeded(clip);
                    }
                    catch { /* skip unreadable file */ }
                }
            }
            catch { /* skip inaccessible game folder */ }
        }

        // Pass 2: clips only in backup (source deleted or game not installed)
        foreach (var key in GameKeys)
        {
            try
            {
                var backupFolder = Path.Combine(BackupRoot, key);
                if (!Directory.Exists(backupFolder)) continue;

                foreach (var file in Directory.EnumerateFiles(backupFolder, "*.mov"))
                {
                    try
                    {
                        var fi = new FileInfo(file);
                        if (_clips.Any(c => c.GameKey == key && c.FileName == fi.Name)) continue;

                        var clip = BuildClipFromBackup(key, fi);
                        _clips.Add(clip);
                    }
                    catch { /* skip unreadable backup file */ }
                }
            }
            catch { /* skip inaccessible backup folder */ }
        }

        ApplySort();
        RebuildByGameMenu();
        UpdateStatus();
    }

    private TheaterClip BuildClip(string key, FileInfo fi, bool sourcePresent)
    {
        var backupPath = Path.Combine(BackupRoot, key, fi.Name);
        var (displayName, accent) = GameInfo[key];
        var clip = new TheaterClip
        {
            Game           = displayName,
            GameKey        = key,
            FileName       = fi.Name,
            MapName        = Path.GetFileNameWithoutExtension(fi.Name),
            MapDisplayName = ResolveMapName(key, fi.Name),
            FileSizeBytes  = fi.Length,
            RecordedAt     = fi.LastWriteTime,
            SourcePath     = fi.FullName,
            BackupPath     = backupPath,
            IsBackedUp     = File.Exists(backupPath),
            SourcePresent  = sourcePresent,
            GameBrush      = new SolidColorBrush(accent),
        };
        ApplyCustomName(clip);
        return clip;
    }

    private TheaterClip BuildClipFromBackup(string key, FileInfo backupFi)
    {
        var sourcePath = Path.Combine(TheaterRoot, key, "Movie", backupFi.Name);
        var (displayName, accent) = GameInfo[key];
        var clip = new TheaterClip
        {
            Game           = displayName,
            GameKey        = key,
            FileName       = backupFi.Name,
            MapName        = Path.GetFileNameWithoutExtension(backupFi.Name),
            MapDisplayName = ResolveMapName(key, backupFi.Name),
            FileSizeBytes  = backupFi.Length,
            RecordedAt     = backupFi.LastWriteTime,
            SourcePath     = sourcePath,
            BackupPath     = backupFi.FullName,
            IsBackedUp     = true,
            SourcePresent  = false,
            GameBrush      = new SolidColorBrush(accent),
        };
        ApplyCustomName(clip);
        return clip;
    }

    private void ApplyCustomName(TheaterClip clip)
    {
        if (_customNames.TryGetValue(CustomNamesKey(clip), out var name))
            clip.CustomName = name;
    }

    private static void BackupIfNeeded(TheaterClip clip)
    {
        if (File.Exists(clip.BackupPath)) { clip.IsBackedUp = true; return; }
        try
        {
            File.Copy(clip.SourcePath, clip.BackupPath, overwrite: false);
            clip.IsBackedUp = true;
        }
        catch { /* non-critical; will retry on next scan */ }
    }

    // ── FileSystemWatcher ──────────────────────────────────────────────────────

    private void InitializeWatchers()
    {
        foreach (var w in _watchers) { w.EnableRaisingEvents = false; w.Dispose(); }
        _watchers.Clear();

        foreach (var key in GameKeys)
        {
            var folder = Path.Combine(TheaterRoot, key, "Movie");
            if (!Directory.Exists(folder)) continue;

            var w = new FileSystemWatcher(folder, "*.mov")
            {
                NotifyFilter        = NotifyFilters.FileName | NotifyFilters.LastWrite,
                EnableRaisingEvents = true,
            };
            var capturedKey = key;
            w.Created += (_, e) => OnFileCreated(capturedKey, e.FullPath);
            w.Deleted += (_, e) => OnFileDeleted(capturedKey, e.Name ?? "");
            _watchers.Add(w);
        }
    }

    private async void OnFileCreated(string gameKey, string fullPath)
    {
        await Task.Delay(600); // let MCC finish writing

        await Dispatcher.InvokeAsync(() =>
        {
            var fi = new FileInfo(fullPath);
            if (!fi.Exists) return;

            // DEDUP — if already tracked, this was a restore write; just mark source present
            var existing = _clips.FirstOrDefault(c => c.GameKey == gameKey && c.FileName == fi.Name);
            if (existing is not null)
            {
                existing.SourcePresent = true;
                UpdateStatus();
                return;
            }

            // Genuinely new clip
            var clip = BuildClip(gameKey, fi, sourcePresent: true);
            _clips.Add(clip);
            BackupIfNeeded(clip);
            ApplySort();
            RebuildByGameMenu();
            UpdateStatus($"New clip: {clip.DisplayName} ({clip.FileSizeStr})");
        });
    }

    private void OnFileDeleted(string gameKey, string fileName)
    {
        Dispatcher.InvokeAsync(() =>
        {
            var clip = _clips.FirstOrDefault(c => c.GameKey == gameKey && c.FileName == fileName);
            if (clip is not null)
            {
                clip.SourcePresent = false;
                UpdateStatus();
            }
        });
    }

    // ── Row interaction — click toggles selection, double-click renames ──────────

    private void Row_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        // Skip on double-click (ClickCount == 2) — rename is handled in MouseLeftButtonDown
        if (e.ClickCount >= 2) return;
        if ((sender as FrameworkElement)?.DataContext is TheaterClip clip)
            clip.IsSelected = !clip.IsSelected;
    }

    private void Row_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount != 2) return;
        if ((sender as FrameworkElement)?.DataContext is not TheaterClip clip) return;

        // Cancel any other active rename first
        foreach (var c in _clips.Where(c => c.IsRenaming && c != clip))
            c.IsRenaming = false;

        clip.IsRenaming = true;
        e.Handled = true; // prevent selection toggle on the Up event
    }

    private void RenameBox_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if ((bool)e.NewValue && sender is TextBox tb)
            Dispatcher.BeginInvoke(() => { tb.Focus(); tb.SelectAll(); });
    }

    private void RenameBox_KeyDown(object sender, KeyEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not TheaterClip clip) return;

        if (e.Key == Key.Return || e.Key == Key.Enter)
        {
            CommitRename(clip, ((TextBox)sender).Text);
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            clip.IsRenaming = false;
            e.Handled = true;
        }
    }

    private void RenameBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is TheaterClip clip && clip.IsRenaming)
            CommitRename(clip, ((TextBox)sender).Text);
    }

    private void CommitRename(TheaterClip clip, string text)
    {
        if (!clip.IsRenaming) return; // guard against double-commit
        clip.CustomName = string.IsNullOrWhiteSpace(text) ? null : text.Trim();
        clip.IsRenaming = false;

        var key = CustomNamesKey(clip);
        if (clip.CustomName is null)
            _customNames.Remove(key);
        else
            _customNames[key] = clip.CustomName;

        SaveCustomNames();
    }

    // ── Restore ────────────────────────────────────────────────────────────────

    private void BtnRestore_Click(object sender, RoutedEventArgs e)
    {
        var btn = (Button)sender;
        btn.ContextMenu.PlacementTarget = btn;
        btn.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        btn.ContextMenu.IsOpen = true;
    }

    private void MniRestoreSelected_Click(object sender, RoutedEventArgs e)
    {
        var selected = _clips.Where(c => c.IsSelected).ToList();
        if (selected.Count == 0) { UpdateStatus("No clips selected — click rows to select them."); return; }
        ConfirmAndRestore(selected, $"{selected.Count} selected clip(s)");
    }

    private void MniRestoreAll_Click(object sender, RoutedEventArgs e)
    {
        var all = _clips.ToList();
        if (all.Count == 0) return;
        ConfirmAndRestore(all, $"all {all.Count} clip(s)");
    }

    private void MniRestoreByGame_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi && mi.Tag is string gameKey)
        {
            var forGame = _clips.Where(c => c.GameKey == gameKey).ToList();
            ConfirmAndRestore(forGame, $"{forGame.Count} {GameInfo[gameKey].DisplayName} clip(s)");
        }
    }

    private void MniRestoreOne_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi && mi.Tag is TheaterClip clip)
            ConfirmAndRestore([clip], $"\"{clip.DisplayName}\"");
    }

    private void ConfirmAndRestore(List<TheaterClip> clips, string description)
    {
        if (clips.Count == 0) return;

        var result = MessageBox.Show(
            $"Restore {description} to their original MCC theater folders?\n\nExisting source files will be overwritten.",
            "Confirm Restore",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes) return;

        int count = ExecuteRestore(clips);
        foreach (var c in clips) c.IsSelected = false;
        UpdateStatus($"Restored {count} clip(s) to source folders.");
    }

    /// <summary>
    /// Copies backup files back to MCC source folders.
    /// FileSystemWatcher.Created fires for each copy; OnFileCreated deduplicates on
    /// (GameKey, FileName) so no infinite backup loop occurs.
    /// </summary>
    private static int ExecuteRestore(IEnumerable<TheaterClip> clips)
    {
        int count = 0;
        foreach (var clip in clips)
        {
            if (!File.Exists(clip.BackupPath)) continue;
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(clip.SourcePath)!);
                File.Copy(clip.BackupPath, clip.SourcePath, overwrite: true);
                clip.SourcePresent = true;
                count++;
            }
            catch { /* skip errored file */ }
        }
        return count;
    }

    // ── Row context menu ───────────────────────────────────────────────────────

    private void MniOpenSource_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi && mi.Tag is TheaterClip clip)
        {
            var dir = Path.GetDirectoryName(clip.SourcePath) ?? "";
            if (Directory.Exists(dir))
                Process.Start("explorer.exe", dir);
            else
                UpdateStatus($"Source folder not found: {dir}");
        }
    }

    private void MniOpenBackup_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi && mi.Tag is TheaterClip clip)
        {
            var dir = Path.GetDirectoryName(clip.BackupPath) ?? "";
            Directory.CreateDirectory(dir);
            Process.Start("explorer.exe", dir);
        }
    }

    private void MniCopyPath_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi && mi.Tag is TheaterClip clip)
        {
            Clipboard.SetText(clip.BackupPath);
            UpdateStatus($"Copied: {clip.BackupPath}");
        }
    }

    private void MniDeleteClip_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi || mi.Tag is not TheaterClip clip) return;

        var result = MessageBox.Show(
            $"Delete \"{clip.DisplayName}\"?\n\nThis will remove the backup copy. Source file (if present) will also be deleted.",
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            // Delete backup
            if (File.Exists(clip.BackupPath))
                File.Delete(clip.BackupPath);

            // Delete source if present
            if (clip.SourcePresent && File.Exists(clip.SourcePath))
                File.Delete(clip.SourcePath);

            // Remove from collection
            _clips.Remove(clip);
            UpdateStatus($"Deleted: {clip.DisplayName}");
        }
        catch (Exception ex)
        {
            UpdateStatus($"Delete failed: {ex.Message}");
        }
    }

    // ── Filter & sort ──────────────────────────────────────────────────────────

    private bool FilterClip(object obj)
    {
        if (obj is not TheaterClip clip) return false;

        if (!string.IsNullOrEmpty(_filterGame) && clip.GameKey != _filterGame)
            return false;

        if (!string.IsNullOrEmpty(_filterSearch) &&
            !clip.DisplayName.Contains(_filterSearch, StringComparison.OrdinalIgnoreCase) &&
            !clip.FileName.Contains(_filterSearch, StringComparison.OrdinalIgnoreCase))
            return false;

        return true;
    }

    private void ApplySort()
    {
        _view.SortDescriptions.Clear();
        var sort = CboSort?.SelectedIndex switch
        {
            1 => new SortDescription(nameof(TheaterClip.RecordedAt), ListSortDirection.Ascending),
            2 => new SortDescription(nameof(TheaterClip.MapDisplayName), ListSortDirection.Ascending),
            _ => new SortDescription(nameof(TheaterClip.RecordedAt), ListSortDirection.Descending),
        };
        _view.SortDescriptions.Add(sort);
    }

    private void RebuildByGameMenu()
    {
        if (MniByGameParent is null) return;
        MniByGameParent.Items.Clear();
        foreach (var key in GameKeys)
        {
            if (!_clips.Any(c => c.GameKey == key)) continue;
            var mi = new MenuItem
            {
                Header     = GameInfo[key].DisplayName,
                Tag        = key,
                FontFamily = new FontFamily("Consolas"),
                FontSize   = 11,
            };
            mi.Click += MniRestoreByGame_Click;
            MniByGameParent.Items.Add(mi);
        }
        MniByGameParent.IsEnabled = MniByGameParent.Items.Count > 0;
    }

    // ── Status / empty state ───────────────────────────────────────────────────

    private void UpdateStatus(string? extra = null)
    {
        if (TxtStatus is null || TxtEmpty is null) return;

        int watching = _watchers.Count;
        int total    = _clips.Count;
        var last     = _clips.OrderByDescending(c => c.RecordedAt).FirstOrDefault()?.DisplayName ?? "—";

        TxtStatus.Text = $"WATCHING: {watching} FOLDERS  ●  {total} CLIPS  ●  LAST: {last}"
                       + (extra is not null ? $"  ●  {extra}" : "");

        TxtEmpty.Visibility = total == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    // ── UI event handlers ──────────────────────────────────────────────────────

    private void BtnScan_Click(object sender, RoutedEventArgs e) => InitialScan();

    private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(BackupRoot);
        Process.Start("explorer.exe", BackupRoot);
    }

    private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
    {
        bool anyUnselected = _clips.Any(c => !c.IsSelected);
        foreach (var c in _clips) c.IsSelected = anyUnselected;
        BtnSelectAll.Content = anyUnselected ? "DESELECT ALL" : "SELECT ALL";
    }

    private void CboGame_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _filterGame = CboGame.SelectedIndex <= 0
            ? ""
            : GameKeys[CboGame.SelectedIndex - 1];
        _view?.Refresh();
    }

    private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        _filterSearch = TxtSearch.Text.Trim();
        _view?.Refresh();
    }

    private void CboSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        ApplySort();
    }
}
