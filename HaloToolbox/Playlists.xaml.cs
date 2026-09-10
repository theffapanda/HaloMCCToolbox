using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Linq;

namespace HaloToolbox;

public partial class Playlists : UserControl
{
    private const string RotationDataPath = "Data\\playlist-rotations.csv";

    private readonly ObservableCollection<PlaylistViewGroup> _visibleGroups = new();
    private readonly List<PlaylistViewSummary> _allPlaylists = new();
    private readonly HashSet<string> _selectedTagIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _selectedGames = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _expandedGroupIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<PlaylistRotationRecord> _allRotationRecords = new();
    private readonly ObservableCollection<PlaylistRotationViewRow> _confirmedRotationRows = new();
    private bool _loaded;
    private PlaylistMode _mode = PlaylistMode.Social;
    private PlaylistSubview _subview = PlaylistSubview.LiveComposer;
    private string _mccInstallationPath = App.DefaultMccPath;

    private static readonly string[] GameOrder =
    [
        "Halo: Reach",
        "Halo CE",
        "Halo 2",
        "Halo 2A",
        "Halo 3",
        "Halo 3: ODST",
        "Halo 4"
    ];

    private static readonly (string Prefix, string DisplayName)[] H3Prefixes =
    [
        ("asq_chillou", "Cold Storage"),
        ("asq_chill_", "Narrows"),
        ("asq_armory_", "Rat's Nest"),
        ("asq_docks_", "Longshore"),
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
        ("s3d_turf", "Icebox"),
        ("s3d_tur", "Icebox"),
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

    private static readonly Dictionary<string, (string MapId, string DisplayName)[]> KnownMapFiles = new()
    {
        ["halo2a"] =
        [
            ("ca_coagulation", "Bloodline"),
            ("ca_lockout", "Lockdown"),
            ("ca_sanctuary", "Shrine"),
            ("ca_zanzibar", "Stonetown"),
            ("ca_warlock", "Warlord"),
            ("ca_ascension", "Zenith"),
            ("ca_forge_skybox01", "Awash"),
            ("ca_forge_skybox02", "Nebula"),
            ("ca_forge_skybox03", "Skyward"),
            ("ca_relic", "Remnant"),
        ],
        ["halo3odst"] =
        [
            ("sc150", "Kikowani Station"),
            ("sc140", "NMPD HQ"),
            ("sc130", "ONI Alpha Site"),
            ("sc120", "Kizingo Boulevard"),
            ("sc110", "Uplift Reserve"),
            ("sc100", "Tayari Plaza"),
            ("h100", "Mombasa Streets"),
            ("l300", "Coastal Highway"),
            ("l200", "Data Hive"),
            ("c200", "Epilogue"),
            ("c100", "Prepare To Drop"),
        ],
        ["halo4"] =
        [
            ("ca_deadlycrossing", "Monolith"),
            ("ca_forge_island", "Forge Island"),
            ("ca_forge_bonanza", "Impact"),
            ("ca_forge_erosion", "Erosion"),
            ("ca_blood_cavern", "Abandon"),
            ("ca_blood_crash", "Exile"),
            ("ca_gore_valley", "Longbow"),
            ("ca_spiderweb", "Daybreak"),
            ("ca_highrise", "Perdition"),
            ("ca_dropoff", "Vertigo"),
            ("ca_creeper", "Pitfall"),
            ("ca_rattler", "Skyline"),
            ("ca_redoubt", "Vortex"),
            ("ca_warhouse", "Adrift"),
            ("ca_forge_ravine", "Ravine"),
            ("ca_canyon", "Meltdown"),
            ("ca_tower", "Solace"),
            ("ca_basin", "Outcast"),
            ("ca_port", "Landfall"),
            ("wraparound", "Haven"),
            ("z05_cliffside", "Complex"),
            ("z11_valhalla", "Ragnarok"),
            ("zd_02_grind", "Harvest"),
            ("dlc_dejewel", "Shatter"),
            ("dlc_dejunkyard", "Wreckage"),
            ("dlc_forge_island", "Forge Island"),
            ("ff87_chopperbowl", "Quarry"),
            ("ff86_sniperally", "Sniper Alley"),
            ("ff90_fortsw", "Fortress"),
            ("ff84_temple", "The Refuge"),
            ("ff81_scurve", "The Cauldron"),
            ("ff81_courtyard", "The Gate"),
            ("ff91_complex", "Galileo Base"),
            ("ff92_valhalla", "Two Giants"),
            ("ff151_mezzanine", "Control"),
            ("ff153_caverns", "Warrens"),
            ("ff152_vortex", "Cyclone"),
            ("ff155_breach", "Harvester"),
            ("ff154_hillside", "Apex"),
            ("dlc01_factory", "Lockup"),
            ("dlc01_engine", "Infinity"),
        ],
        ["haloreach"] =
        [
            ("20_sword_slayer", "Sword Base"),
            ("45_launch_station", "Countdown"),
            ("50_panopticon", "Boardwalk"),
            ("52_ivory_tower", "Reflection"),
            ("70_boneyard", "Boneyard"),
            ("45_aftship", "Zealot"),
            ("35_island", "Spire"),
            ("30_settlement", "Powerhouse"),
            ("forge_halo", "Forge World"),
            ("dlc_slayer", "Anchor 9"),
            ("dlc_invasion", "Breakpoint"),
            ("dlc_medium", "Tempest"),
            ("trainingpreserve", "Highlands"),
            ("condemned", "Condemned"),
            ("cex_beavercreek", "Battle Canyon"),
            ("cex_headlong", "Breakneck"),
            ("cex_hangemhigh", "High Noon"),
            ("cex_damnation", "Penance"),
            ("cex_timberland", "Ridgeline"),
            ("cex_prisoner", "Solitary"),
            ("ff50_park", "Beachhead"),
            ("ff45_corvette", "Corvette"),
            ("ff20_cortyard", "Courtyard"),
            ("ff60_icecave", "Glacier"),
            ("ff70_holdout", "Holdout"),
            ("ff60_airview", "Outpost"),
            ("ff10_prototype", "Overlook"),
            ("ff30_waterfront", "Waterfront"),
            ("ff_unearthed", "Unearthed"),
            ("cex_ff_halo", "Installation 04"),
        ],
    };

    public Playlists()
    {
        InitializeComponent();
        GroupList.ItemsSource = _visibleGroups;
        GridConfirmedRotations.ItemsSource = _confirmedRotationRows;
    }

    private void Playlists_Loaded(object sender, RoutedEventArgs e)
    {
        if (_loaded) return;
        _loaded = true;

        CboScheduleFilter.ItemsSource = Enum.GetValues<RotationScheduleFilter>();
        CboScheduleFilter.SelectedItem = RotationScheduleFilter.All;
        DpRotationDate.SelectedDate = DateTime.Today;

        LoadPlaylists();
        LoadRotationSchedule();
        SetSubview(PlaylistSubview.LiveComposer);
    }

    private void BtnRefresh_Click(object sender, RoutedEventArgs e)
    {
        if (_subview == PlaylistSubview.LiveComposer)
            LoadPlaylists();
        else
            LoadRotationSchedule();
    }

    public void SetMccInstallationPath(string path)
    {
        var normalizedPath = string.IsNullOrWhiteSpace(path) ? App.DefaultMccPath : path.Trim();
        if (_mccInstallationPath.Equals(normalizedPath, StringComparison.OrdinalIgnoreCase))
            return;

        _mccInstallationPath = normalizedPath;
        if (_loaded && _subview == PlaylistSubview.LiveComposer)
            LoadPlaylists();
    }

    private void BtnLiveComposer_Click(object sender, RoutedEventArgs e) => SetSubview(PlaylistSubview.LiveComposer);
    private void BtnRotationSchedule_Click(object sender, RoutedEventArgs e) => SetSubview(PlaylistSubview.RotationSchedule);
    private void BtnSocial_Click(object sender, RoutedEventArgs e) => SetMode(PlaylistMode.Social);
    private void BtnRanked_Click(object sender, RoutedEventArgs e) => SetMode(PlaylistMode.Ranked);
    private void CboSize_SelectionChanged(object sender, SelectionChangedEventArgs e) => RefreshPlaylistsForMode();
    private void CboPlaylist_SelectionChanged(object sender, SelectionChangedEventArgs e) => RefreshSelection();
    private void CboScheduleFilter_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyRotationFilters();
    private void CboScheduleSearch_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyRotationFilters();
    private void DpRotationDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e) => ApplyRotationFilters();

    private void GridConfirmedRotations_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (GridConfirmedRotations.SelectedItem is PlaylistRotationViewRow row)
            UpdateRotationSummary(row.Source);
    }

    private void SetSubview(PlaylistSubview subview)
    {
        _subview = subview;
        var liveActive = subview == PlaylistSubview.LiveComposer;

        LiveControlsPanel.Visibility = liveActive ? Visibility.Visible : Visibility.Collapsed;
        ScheduleControlsPanel.Visibility = liveActive ? Visibility.Collapsed : Visibility.Visible;
        LiveComposerView.Visibility = liveActive ? Visibility.Visible : Visibility.Collapsed;
        RotationScheduleView.Visibility = liveActive ? Visibility.Collapsed : Visibility.Visible;

        BtnRefresh.Content = liveActive ? "RELOAD XML" : "RELOAD SCHEDULE";
        TxtViewTitle.Text = liveActive ? "PLAYLISTS" : "PLAYLIST ROTATION SCHEDULE";
        TxtHeaderSub.Text = liveActive
            ? "Browse live hopper composition and playlist weights from MCC's XML"
            : "Browse weekly featured playlist rotations and current playlist picks";

        ApplyModeButton(BtnLiveComposer, liveActive);
        ApplyModeButton(BtnRotationSchedule, !liveActive);

        if (liveActive)
            RefreshSelection();
        else
            ApplyRotationFilters();
    }

    private void LoadPlaylists()
    {
        try
        {
            _allPlaylists.Clear();
            var playlistXmlPath = GetPlaylistXmlPath();

            if (!File.Exists(playlistXmlPath))
            {
                ShowEmpty($"Playlist XML not found at:{Environment.NewLine}{playlistXmlPath}");
                return;
            }

            var doc = XDocument.Load(playlistXmlPath);
            var mixTags = doc
                .Descendants("MixTag")
                .Select(x => new
                {
                    Id = (string?)x.Attribute("id") ?? "",
                    Name = (string?)x.Attribute("name") ?? ""
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.Id))
                .ToDictionary(x => x.Id, x => x.Name, StringComparer.OrdinalIgnoreCase);

            var playlists = doc
                .Descendants("Playlist")
                .Select(x => ParsePlaylist(x, mixTags))
                .Where(x => x is not null && !x.IsDebug)
                .Cast<PlaylistViewSummary>()
                .ToList();

            _allPlaylists.AddRange(playlists);
            SetMode(PlaylistMode.Social);
        }
        catch (Exception ex)
        {
            ShowEmpty("Failed to load playlist data." + Environment.NewLine + ex.Message);
        }
    }

    private string GetPlaylistXmlPath() =>
        Path.Combine(_mccInstallationPath, "data", "careerdb", "findgamehopperdb-v4.xml");

    private void LoadRotationSchedule()
    {
        try
        {
            _allRotationRecords.Clear();

            var dataPath = Path.Combine(AppContext.BaseDirectory, RotationDataPath);
            if (!File.Exists(dataPath))
            {
                TxtConfirmedSummary.Text = $"Rotation schedule data not found at {dataPath}";
                return;
            }

            foreach (var line in File.ReadAllLines(dataPath).Skip(1))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split('|');
                if (parts.Length < 6)
                    continue;

                if (!DateOnly.TryParseExact(parts[0], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                    continue;

                _allRotationRecords.Add(new PlaylistRotationRecord
                {
                    Date = date,
                    SocialPlaylist = NormalizeRotationCell(parts[1]),
                    Ranked4v4Playlist = NormalizeRotationCell(parts[2]),
                    Ranked2v2Playlist = NormalizeRotationCell(parts[3]),
                    Status = ParseRotationStatus(parts[4]),
                    SourceNote = NormalizeRotationCell(parts[5])
                });
            }

            _allRotationRecords.Sort((a, b) => a.Date.CompareTo(b.Date));
            PopulateRotationSearchValues();
            ApplyRotationFilters();
        }
        catch (Exception ex)
        {
            TxtConfirmedSummary.Text = "Failed to load rotation schedule." + Environment.NewLine + ex.Message;
        }
    }

    private void SetMode(PlaylistMode mode)
    {
        _mode = mode;
        ApplyModeButtonState();
        PopulateSizes();
        RefreshPlaylistsForMode();
    }

    private void ApplyModeButtonState()
    {
        ApplyModeButton(BtnSocial, _mode == PlaylistMode.Social);
        ApplyModeButton(BtnRanked, _mode == PlaylistMode.Ranked);
        TxtSelectionLabel.Text = _mode == PlaylistMode.Social ? "SOCIAL PLAYLIST" : "RANKED PLAYLIST";
        TxtCategoriesLabel.Text = _mode == PlaylistMode.Social ? "GAME CATEGORIES INCLUDED" : "GAMES INCLUDED";
    }

    private static void ApplyModeButton(Button button, bool active)
    {
        if (active)
        {
            button.BorderBrush = (Brush)Application.Current.Resources["AccentBrush"];
            button.Foreground = (Brush)Application.Current.Resources["AccentBrush"];
            button.Background = new SolidColorBrush(Color.FromArgb(0x18, 0x00, 0xC8, 0xFF));
        }
        else
        {
            button.BorderBrush = new SolidColorBrush(Color.FromRgb(0x2A, 0x32, 0x38));
            button.Foreground = new SolidColorBrush(Color.FromRgb(0x7D, 0x85, 0x90));
            button.Background = Brushes.Transparent;
        }
    }

    private void PopulateSizes()
    {
        var sizes = _allPlaylists
            .Where(MatchesCurrentMode)
            .Select(x => x.SizeLabel)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(SizeSortKey)
            .ToList();

        CboSize.ItemsSource = sizes;
        CboSize.SelectedIndex = sizes.Count > 0 ? 0 : -1;
    }

    private void RefreshPlaylistsForMode()
    {
        if (CboSize.SelectedItem is not string size)
        {
            CboPlaylist.ItemsSource = null;
            ShowEmpty("No playlists are available for the selected mode.");
            return;
        }

        var playlists = _allPlaylists
            .Where(MatchesCurrentMode)
            .Where(x => x.SizeLabel.Equals(size, StringComparison.OrdinalIgnoreCase))
            .OrderBy(x => x.SelectionName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        CboPlaylist.ItemsSource = playlists;
        CboPlaylist.DisplayMemberPath = nameof(PlaylistViewSummary.SelectionName);
        CboPlaylist.Text = "";
        CboPlaylist.SelectedIndex = playlists.Count > 0 ? 0 : -1;
        RefreshSelection();
    }

    private void RefreshSelection()
    {
        BuildTagButtons();
        RefreshVisibleGroups();
    }

    private void BuildTagButtons()
    {
        TagPanel.Children.Clear();
        GamePanel.Children.Clear();
        _selectedTagIds.Clear();
        _selectedGames.Clear();

        if (CboPlaylist.SelectedItem is not PlaylistViewSummary selected)
            return;

        foreach (var game in selected.Entries
                     .Select(x => x.GameDisplay)
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(GameSortKey)
                     .ThenBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            _selectedGames.Add(game);
            GamePanel.Children.Add(CreateGameToggle(game, true));
        }

        if (_mode == PlaylistMode.Social)
        {
            foreach (var group in selected.TagGroups.OrderBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase))
            {
                _selectedTagIds.Add(group.TagId);
                TagPanel.Children.Add(CreateTagToggle(group.DisplayName, group.TagId, true));
            }
            return;
        }

        foreach (var group in selected.TagGroups.OrderBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase))
            TagPanel.Children.Add(CreateStaticChip(group.DisplayName));
    }

    private Button CreateGameToggle(string label, bool isSelected)
    {
        var button = new Button
        {
            Content = label,
            Tag = label,
            Margin = new Thickness(0, 0, 8, 8),
            Padding = new Thickness(12, 6, 12, 6),
            FontFamily = new FontFamily("Consolas"),
            FontSize = 10,
            BorderThickness = new Thickness(1),
            Cursor = System.Windows.Input.Cursors.Hand
        };

        ApplyTagButtonState(button, isSelected);
        button.Click += GameButton_Click;
        return button;
    }

    private Button CreateTagToggle(string label, string tagId, bool isSelected)
    {
        var button = new Button
        {
            Content = label,
            Tag = tagId,
            Margin = new Thickness(0, 0, 8, 8),
            Padding = new Thickness(12, 6, 12, 6),
            FontFamily = new FontFamily("Consolas"),
            FontSize = 10,
            BorderThickness = new Thickness(1),
            Cursor = System.Windows.Input.Cursors.Hand
        };

        ApplyTagButtonState(button, isSelected);
        button.Click += TagButton_Click;
        return button;
    }

    private Border CreateStaticChip(string label)
    {
        return new Border
        {
            Margin = new Thickness(0, 0, 8, 8),
            Padding = new Thickness(12, 6, 12, 6),
            CornerRadius = new CornerRadius(4),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x2A, 0x32, 0x38)),
            Background = new SolidColorBrush(Color.FromArgb(0x10, 0xFF, 0xFF, 0xFF)),
            Child = new TextBlock
            {
                Text = label,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 10,
                Foreground = (Brush)Application.Current.Resources["TextBrush"]
            }
        };
    }

    private void TagButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string tagId) return;

        if (_selectedTagIds.Contains(tagId))
            _selectedTagIds.Remove(tagId);
        else
            _selectedTagIds.Add(tagId);

        ApplyTagButtonState(button, _selectedTagIds.Contains(tagId));
        RefreshVisibleGroups();
    }

    private void GameButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string game) return;

        if (_selectedGames.Contains(game))
            _selectedGames.Remove(game);
        else
            _selectedGames.Add(game);

        ApplyTagButtonState(button, _selectedGames.Contains(game));
        RefreshVisibleGroups();
    }

    private void GroupExpander_Expanded(object sender, RoutedEventArgs e)
    {
        if (sender is Expander { Tag: string tagId } && !string.IsNullOrWhiteSpace(tagId))
            _expandedGroupIds.Add(tagId);
    }

    private void GroupExpander_Collapsed(object sender, RoutedEventArgs e)
    {
        if (sender is Expander { Tag: string tagId } && !string.IsNullOrWhiteSpace(tagId))
            _expandedGroupIds.Remove(tagId);
    }

    private static void ApplyTagButtonState(Button button, bool selected)
    {
        if (selected)
        {
            button.Background = new SolidColorBrush(Color.FromArgb(0x18, 0x00, 0xC8, 0xFF));
            button.BorderBrush = (Brush)Application.Current.Resources["AccentBrush"];
            button.Foreground = (Brush)Application.Current.Resources["AccentBrush"];
        }
        else
        {
            button.Background = Brushes.Transparent;
            button.BorderBrush = new SolidColorBrush(Color.FromRgb(0x2A, 0x32, 0x38));
            button.Foreground = new SolidColorBrush(Color.FromRgb(0x7D, 0x85, 0x90));
        }
    }

    private void RefreshVisibleGroups()
    {
        _visibleGroups.Clear();

        if (CboPlaylist.SelectedItem is not PlaylistViewSummary selected)
        {
            ShowEmpty("Select a playlist to view its hopper entries.");
            return;
        }

        IEnumerable<PlaylistViewGroup> groups = selected.TagGroups;
        if (_mode == PlaylistMode.Social)
        {
            groups = groups
                .Where(x => _selectedTagIds.Count == 0 || _selectedTagIds.Contains(x.TagId))
                .Select(x =>
                {
                    var filteredEntries = x.Entries
                        .Where(e => _selectedGames.Contains(e.GameDisplay))
                        .ToList();

                    var filteredWeight = filteredEntries.Sum(e => e.Weight);
                    foreach (var entry in filteredEntries)
                        entry.WeightShare = filteredWeight <= 0 ? 0 : (double)entry.Weight / filteredWeight;

                    return new PlaylistViewGroup
                    {
                        TagId = x.TagId,
                        DisplayName = x.DisplayName,
                        Subtitle = x.Subtitle,
                        TotalWeight = filteredWeight,
                        IsExpanded = _expandedGroupIds.Contains(x.TagId),
                        Entries = filteredEntries
                    };
                })
                .Where(x => x.Entries.Count > 0);
        }
        else
        {
            groups = groups
                .Select(x =>
                {
                    var filteredEntries = x.Entries
                        .Where(e => _selectedGames.Contains(e.GameDisplay))
                        .ToList();

                    var filteredWeight = filteredEntries.Sum(e => e.Weight);
                    foreach (var entry in filteredEntries)
                        entry.WeightShare = filteredWeight <= 0 ? 0 : (double)entry.Weight / filteredWeight;

                    return new PlaylistViewGroup
                    {
                        TagId = x.TagId,
                        DisplayName = x.DisplayName,
                        Subtitle = x.Subtitle,
                        TotalWeight = filteredWeight,
                        IsExpanded = _expandedGroupIds.Contains(x.TagId),
                        Entries = filteredEntries
                    };
                })
                .Where(x => x.Entries.Count > 0);
        }

        var materialized = groups.OrderBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase).ToList();
        var totalWeight = materialized.Sum(x => x.TotalWeight);
        foreach (var group in materialized)
            group.TotalPlaylistWeight = totalWeight;
        foreach (var group in materialized)
            _visibleGroups.Add(group);

        TxtPlaylistName.Text = selected.SelectionName;
        TxtPlaylistSubtitle.Text = $"{selected.ModeLabel}  •  {selected.SizeLabel}  •  {selected.PlayerSummary}";
        TxtPlayerCount.Text = selected.EffectivePlayerCount;
        TxtPartySize.Text = selected.MaxPartySize;
        TxtTeamSize.Text = selected.TeamSizeDisplay;
        TxtEntryCount.Text = materialized.Sum(x => x.Entries.Count).ToString();
        TxtWeightTotal.Text = materialized.Sum(x => x.TotalWeight).ToString();

        if (_visibleGroups.Count == 0)
            ShowEmpty("No entries matched the current category selection.");
        else
            HideEmpty();
    }

    private void ApplyRotationFilters()
    {
        if (_allRotationRecords.Count == 0)
            return;

        var selectedDate = DateOnly.FromDateTime((DpRotationDate.SelectedDate ?? DateTime.Today).Date);
        var filter = CboScheduleFilter.SelectedItem is RotationScheduleFilter value ? value : RotationScheduleFilter.All;
        var selectedPlaylist = CboScheduleSearch.SelectedItem as string;
        if (string.Equals(selectedPlaylist, "All Playlists", StringComparison.OrdinalIgnoreCase))
            selectedPlaylist = "";

        IEnumerable<PlaylistRotationRecord> filtered = _allRotationRecords;
        if (!string.IsNullOrWhiteSpace(selectedPlaylist))
        {
            filtered = filtered.Where(x =>
                MatchesRotationSelection(x.SocialPlaylist, selectedPlaylist) ||
                MatchesRotationSelection(x.Ranked4v4Playlist, selectedPlaylist) ||
                MatchesRotationSelection(x.Ranked2v2Playlist, selectedPlaylist));
        }

        _confirmedRotationRows.Clear();

        foreach (var record in filtered)
        {
            var row = PlaylistRotationViewRow.FromRecord(record, selectedDate, filter);
            if (row is null)
                continue;

            _confirmedRotationRows.Add(row);
        }

        TxtConfirmedSummary.Text = $"{_confirmedRotationRows.Count} rotation weeks loaded  •  coverage {_allRotationRecords.First().Date:MMM d, yyyy} to {_allRotationRecords.Last().Date:MMM d, yyyy}";

        var selectedRecord = filtered
            .Where(x => x.Date <= selectedDate)
            .OrderByDescending(x => x.Date)
            .FirstOrDefault()
            ?? filtered.OrderBy(x => x.Date).FirstOrDefault()
            ?? _allRotationRecords.First();

        UpdateRotationSummary(selectedRecord);
        SelectCurrentRows(selectedRecord.Date);
    }

    private void SelectCurrentRows(DateOnly date)
    {
        var selectedRow = _confirmedRotationRows.FirstOrDefault(x => x.Source.Date == date);
        GridConfirmedRotations.SelectedItem = selectedRow;

        if (selectedRow is not null)
        {
            GridConfirmedRotations.UpdateLayout();
            GridConfirmedRotations.ScrollIntoView(selectedRow);
        }
    }

    private void UpdateRotationSummary(PlaylistRotationRecord selectedRecord)
    {
        TxtCurrentWeek.Text = selectedRecord.Date.ToString("MMM d, yyyy", CultureInfo.InvariantCulture);
        TxtNextReset.Text = $"Next reset: {selectedRecord.Date.AddDays(7).ToString("dddd, MMM d", CultureInfo.InvariantCulture)}";

        var today = DateOnly.FromDateTime(DateTime.Today);
        var activeRecord = _allRotationRecords
            .Where(x => x.Date <= today)
            .OrderByDescending(x => x.Date)
            .FirstOrDefault()
            ?? _allRotationRecords.First();

        var activeSocial = ResolveLaneValue(activeRecord, RotationScheduleFilter.Social, out var socialSource);
        var activeRanked4v4 = ResolveLaneValue(activeRecord, RotationScheduleFilter.Ranked4v4, out var ranked4v4Source);
        var activeRanked2v2 = ResolveLaneValue(activeRecord, RotationScheduleFilter.Ranked2v2, out var ranked2v2Source);

        TxtCurrentSocial.Text = DisplayRotationCell(activeSocial);
        TxtCurrentRanked4v4.Text = DisplayRotationCell(activeRanked4v4);
        TxtCurrentRanked2v2.Text = DisplayRotationCell(activeRanked2v2);
        TxtCurrentSocialMeta.Text = BuildPlaylistMetaRelative(activeSocial, RotationScheduleFilter.Social, activeRecord.Date, socialSource != activeRecord.Date);
        TxtCurrentRanked4v4Meta.Text = BuildPlaylistMetaRelative(activeRanked4v4, RotationScheduleFilter.Ranked4v4, activeRecord.Date, ranked4v4Source != activeRecord.Date);
        TxtCurrentRanked2v2Meta.Text = BuildPlaylistMetaRelative(activeRanked2v2, RotationScheduleFilter.Ranked2v2, activeRecord.Date, ranked2v2Source != activeRecord.Date);
        TxtScheduleHint.Text = $"Selected week: {selectedRecord.Date:MMM d, yyyy}" +
                               (string.IsNullOrWhiteSpace(selectedRecord.SourceNote) ? "" : $"  •  note: {selectedRecord.SourceNote}");
    }

    private string BuildPlaylistMeta(string? playlistName, RotationScheduleFilter lane, DateOnly referenceDate, bool isEstimatedFromHistory = false)
    {
        if (string.IsNullOrWhiteSpace(playlistName))
            return "Not listed this week";

        var matches = _allRotationRecords
            .Where(x => x.Date < referenceDate)
            .Where(x => GetLaneValue(x, lane)?.Equals(playlistName, StringComparison.OrdinalIgnoreCase) == true)
            .OrderByDescending(x => x.Date)
            .ToList();

        var latest = matches.FirstOrDefault();
        if (latest is null)
            return "No prior appearances";

        var weeksAgo = Math.Max(0, (referenceDate.DayNumber - latest.Date.DayNumber) / 7);
        return $"Last seen {latest.Date:MMM d, yyyy}  •  {weeksAgo} weeks ago";
    }

    private static bool MatchesRotationSelection(string? value, string selectedPlaylist) =>
        !string.IsNullOrWhiteSpace(value) && value.Equals(selectedPlaylist, StringComparison.OrdinalIgnoreCase);

    private string BuildPlaylistMetaRelative(string? playlistName, RotationScheduleFilter lane, DateOnly referenceDate, bool isEstimatedFromHistory = false)
    {
        if (string.IsNullOrWhiteSpace(playlistName))
            return "Not listed this week";

        var latest = _allRotationRecords
            .Where(x => x.Date < referenceDate)
            .Where(x => GetLaneValue(x, lane)?.Equals(playlistName, StringComparison.OrdinalIgnoreCase) == true)
            .OrderByDescending(x => x.Date)
            .FirstOrDefault();

        if (latest is null)
            return "No prior appearances";

        var weeksAgo = Math.Max(0, (referenceDate.DayNumber - latest.Date.DayNumber) / 7);
        return isEstimatedFromHistory
            ? $"Estimated from {latest.Date:MMM d, yyyy}  •  {weeksAgo} weeks ago"
            : $"Last seen {latest.Date:MMM d, yyyy}  •  {weeksAgo} weeks ago";
    }

    private void PopulateRotationSearchValues()
    {
        var values = _allRotationRecords
            .SelectMany(record => new[]
            {
                record.SocialPlaylist,
                record.Ranked4v4Playlist,
                record.Ranked2v2Playlist
            })
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList();

        values.Insert(0, "All Playlists");
        CboScheduleSearch.ItemsSource = values;
        CboScheduleSearch.SelectedIndex = 0;
    }

    private string ResolveLaneValue(PlaylistRotationRecord record, RotationScheduleFilter lane, out DateOnly sourceDate)
    {
        var directValue = GetLaneValue(record, lane);
        if (!string.IsNullOrWhiteSpace(directValue))
        {
            sourceDate = record.Date;
            return directValue;
        }

        var fallback = _allRotationRecords
            .Where(x => x.Date <= record.Date)
            .OrderByDescending(x => x.Date)
            .Select(x => new { x.Date, Value = GetLaneValue(x, lane) })
            .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.Value));

        if (fallback is not null)
        {
            sourceDate = fallback.Date;
            return fallback.Value!;
        }

        sourceDate = record.Date;
        return "";
    }

    private static string? GetLaneValue(PlaylistRotationRecord record, RotationScheduleFilter lane) => lane switch
    {
        RotationScheduleFilter.Social => record.SocialPlaylist,
        RotationScheduleFilter.Ranked4v4 => record.Ranked4v4Playlist,
        RotationScheduleFilter.Ranked2v2 => record.Ranked2v2Playlist,
        _ => null
    };

    private bool MatchesCurrentMode(PlaylistViewSummary playlist) =>
        _mode == PlaylistMode.Social ? !playlist.IsRanked : playlist.IsRanked;

    private void ShowEmpty(string message)
    {
        TxtEmpty.Text = message;
        EmptyState.Visibility = Visibility.Visible;
        TxtPlaylistName.Text = "No playlist selected";
        TxtPlaylistSubtitle.Text = "";
        TxtPlayerCount.Text = "-";
        TxtPartySize.Text = "-";
        TxtTeamSize.Text = "-";
        TxtEntryCount.Text = "0";
        TxtWeightTotal.Text = "0";
    }

    private void HideEmpty() => EmptyState.Visibility = Visibility.Collapsed;

    private static bool ParseBool(string? value) =>
        value?.Equals("true", StringComparison.OrdinalIgnoreCase) == true;

    private static int SizeSortKey(string value) => value switch
    {
        "FIREFIGHT" => 0,
        "1V1" => 1,
        "2V2" => 2,
        "4V4" => 4,
        "6V6" => 6,
        "8V8" => 8,
        "12-PLAYER FFA" => 12,
        _ => 99
    };

    private static int GameSortKey(string game)
    {
        var index = Array.FindIndex(GameOrder, x => x.Equals(game, StringComparison.OrdinalIgnoreCase));
        return index >= 0 ? index : int.MaxValue;
    }

    private static string NormalizeRotationCell(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "-")
            return "";
        return value.Trim();
    }

    private static string DisplayRotationCell(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "-" : value;

    private static PlaylistRotationStatus ParseRotationStatus(string value) =>
        value.Equals("estimated", StringComparison.OrdinalIgnoreCase)
            ? PlaylistRotationStatus.Estimated
            : PlaylistRotationStatus.Confirmed;

    private static PlaylistViewSummary? ParsePlaylist(XElement playlistElement, Dictionary<string, string> mixTags)
    {
        var id = (string?)playlistElement.Attribute("id") ?? "";
        if (string.IsNullOrWhiteSpace(id))
            return null;

        var rawName = (string?)playlistElement.Attribute("name") ?? "";
        var hopperName = (string?)playlistElement.Attribute("mpsdHopperName") ?? "";
        var statName = (string?)playlistElement.Attribute("mpsdHopperStatName") ?? "";
        var isMix = ParseBool((string?)playlistElement.Attribute("isMix"));
        var isRanked = ParseBool((string?)playlistElement.Attribute("isRanked"));
        var isDebug = ParseBool((string?)playlistElement.Attribute("isDebug"));
        var minPlayers = (string?)playlistElement.Attribute("minPlayer") ?? "";
        var maxPlayers = (string?)playlistElement.Attribute("maxPlayerCount") ?? "";
        var maxParty = (string?)playlistElement.Attribute("maxPartySize") ?? "";

        var teamSize = playlistElement
            .Descendants("Team")
            .Select(x => (string?)x.Attribute("max_size"))
            .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? "1";

        var entries = playlistElement.Elements("HopperEntry").Select(ParseEntry).ToList();
        var groups = entries
            .GroupBy(x => x.TagId, StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var totalWeight = group.Sum(x => x.Weight);
                foreach (var entry in group)
                    entry.WeightShare = totalWeight <= 0 ? 0 : (double)entry.Weight / totalWeight;

                var tagName = mixTags.TryGetValue(group.Key, out var display) ? display : "";
                return new PlaylistViewGroup
                {
                    TagId = group.Key,
                    DisplayName = BuildTagDisplayName(group.Key, tagName),
                    Subtitle = $"{group.First().CategoryDisplay}  •  {group.Count()} entries",
                    TotalWeight = totalWeight,
                    Entries = group.OrderByDescending(x => x.Weight).ThenBy(x => x.MapDisplay, StringComparer.OrdinalIgnoreCase).ToList()
                };
            })
            .ToList();

        return new PlaylistViewSummary
        {
            Id = id,
            RawName = rawName,
            HopperName = hopperName,
            StatName = statName,
            IsMix = isMix,
            IsRanked = isRanked,
            IsDebug = isDebug,
            MinPlayers = minPlayers,
            MaxPlayers = maxPlayers,
            MaxPartySize = maxParty,
            TeamSize = teamSize,
            TagGroups = groups,
            Entries = entries
        };
    }

    private static PlaylistViewEntry ParseEntry(XElement entry)
    {
        var mapId = (string?)entry.Attribute("map") ?? "";
        var gameVariant = (string?)entry.Attribute("game_variant") ?? "";
        var mapVariant = (string?)entry.Attribute("map_variant") ?? "";
        var tagId = (string?)entry.Attribute("tag") ?? "";
        var category = (string?)entry.Attribute("category") ?? "";
        var weight = int.TryParse((string?)entry.Attribute("weight"), out var parsedWeight) ? parsedWeight : 0;
        var gameKey = InferGameKey(mapId, gameVariant);

        var strippedMap = mapId.Replace("_map_id_", "", StringComparison.OrdinalIgnoreCase);
        strippedMap = strippedMap.Replace($"{gameKey}_", "", StringComparison.OrdinalIgnoreCase);

        return new PlaylistViewEntry
        {
            TagId = tagId,
            Category = category,
            Weight = weight,
            MapId = mapId,
            MapDisplay = ResolvePlaylistMapName(gameKey, strippedMap),
            MapVariant = mapVariant,
            GameVariant = gameVariant,
            GameKey = gameKey,
            TagDisplay = BuildTagDisplayName(tagId, "")
        };
    }

    private static string InferGameKey(string mapId, string gameVariant)
    {
        var source = $"{mapId} {gameVariant}".ToLowerInvariant();
        if (source.Contains("haloreach") || source.StartsWith("hr_")) return "haloreach";
        if (source.Contains("groundhog") || source.Contains("halo2a") || source.StartsWith("h2a_")) return "halo2a";
        if (source.Contains("halo3odst") || source.StartsWith("odst_")) return "halo3odst";
        if (source.Contains("halo4") || source.StartsWith("h4_")) return "halo4";
        if (source.Contains("halo3") || source.StartsWith("h3_")) return "halo3";
        if (source.Contains("halo2") || source.StartsWith("h2_")) return "halo2";
        if (source.Contains("halo1") || source.StartsWith("h1_")) return "halo1";
        return "unknown";
    }

    private static string BuildTagDisplayName(string tagId, string mixTagName)
    {
        var cleaned = BeautifyToken(tagId);
        if (tagId.Equals("brSlayer", StringComparison.OrdinalIgnoreCase)) return "BR Slayer";
        if (tagId.Equals("arSlayer", StringComparison.OrdinalIgnoreCase)) return "AR Slayer";
        if (tagId.Equals("flagBomb", StringComparison.OrdinalIgnoreCase)) return "Flag and Bomb";
        if (tagId.Equals("zoneControl", StringComparison.OrdinalIgnoreCase)) return "Zone Control";
        if (tagId.Equals("assetDenial", StringComparison.OrdinalIgnoreCase)) return "Asset Denial";
        if (tagId.Equals("actionsack", StringComparison.OrdinalIgnoreCase)) return "Action Sack";
        if (tagId.Equals("snipers", StringComparison.OrdinalIgnoreCase)) return "Snipers";
        if (tagId.Equals("infection", StringComparison.OrdinalIgnoreCase)) return "Infection";
        if (!string.IsNullOrWhiteSpace(mixTagName) && !mixTagName.StartsWith("$", StringComparison.Ordinal))
            return BeautifyToken(mixTagName);
        return cleaned;
    }

    private static string ResolvePlaylistMapName(string gameKey, string rawMapId)
    {
        if (gameKey.Equals("halo3", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var (prefix, displayName) in H3Prefixes)
                if (rawMapId.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return displayName;
        }

        if (TryResolveKnownMapName(gameKey, rawMapId, out var resolved))
            return resolved;

        return BeautifyToken(rawMapId);
    }

    private static bool TryResolveKnownMapName(string gameKey, string rawName, out string displayName)
    {
        displayName = "";
        if (!KnownMapFiles.TryGetValue(gameKey, out var knownMaps))
            return false;

        var candidates = BuildLookupCandidates(rawName);
        var bestMatch = knownMaps
            .Select(map => new
            {
                map.DisplayName,
                MapId = NormalizeLookupKey(map.MapId),
                Score = candidates
                    .Where(candidate => MatchesMapId(candidate, map.MapId))
                    .Select(candidate => NormalizeLookupKey(candidate).Length)
                    .DefaultIfEmpty(0)
                    .Max()
            })
            .Where(match => match.Score > 0)
            .OrderByDescending(match => match.Score)
            .ThenByDescending(match => match.MapId.Length)
            .FirstOrDefault();

        if (bestMatch is null)
            return false;

        displayName = bestMatch.DisplayName;
        return true;
    }

    private static List<string> BuildLookupCandidates(string rawName)
    {
        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddCandidate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            var trimmed = value.Trim('_', '-', ' ');
            if (!string.IsNullOrWhiteSpace(trimmed))
                candidates.Add(trimmed);
        }

        AddCandidate(rawName);
        var stripped = Regex.Replace(rawName, @"(?i)\bmglo[-_]*\d+\b", "_");
        AddCandidate(stripped);

        foreach (var prefix in new[] { "asq_", "dlc_", "mp_", "ffa_", "coop_", "ms_", "_map_id_" })
        {
            if (stripped.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                AddCandidate(stripped[prefix.Length..]);
        }

        foreach (var token in stripped.Split(['_', '-'], StringSplitOptions.RemoveEmptyEntries))
            AddCandidate(token);

        return candidates.ToList();
    }

    private static bool MatchesMapId(string candidate, string mapId)
    {
        var normalizedCandidate = NormalizeLookupKey(candidate);
        var normalizedMapId = NormalizeLookupKey(mapId);
        return normalizedCandidate.Equals(normalizedMapId, StringComparison.OrdinalIgnoreCase)
            || normalizedCandidate.EndsWith(normalizedMapId, StringComparison.OrdinalIgnoreCase)
            || normalizedMapId.EndsWith(normalizedCandidate, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeLookupKey(string value) =>
        Regex.Replace(value ?? "", @"[^a-z0-9]", "", RegexOptions.IgnoreCase).ToLowerInvariant();

    private static string BeautifyToken(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return "Unknown";

        var value = raw.Trim().Trim('$');
        value = value.Replace("GameCategory_", "", StringComparison.OrdinalIgnoreCase);
        value = value.Replace('_', ' ');
        value = Regex.Replace(value, "([a-z])([A-Z])", "$1 $2");
        value = Regex.Replace(value, @"\s+", " ").Trim();

        if (string.IsNullOrWhiteSpace(value))
            return "Unknown";

        var words = value.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(word =>
        {
            var upper = word.ToUpperInvariant();
            if (upper is "SWAT" or "BTB" or "MLG" or "FFA" or "BR" or "AR") return upper;
            return char.ToUpperInvariant(word[0]) + word[1..].ToLowerInvariant();
        });

        return string.Join(" ", words);
    }
}

public enum PlaylistMode
{
    Social,
    Ranked
}

public enum PlaylistSubview
{
    LiveComposer,
    RotationSchedule
}

public enum RotationScheduleFilter
{
    All,
    Social,
    Ranked4v4,
    Ranked2v2
}

public enum PlaylistRotationStatus
{
    Confirmed,
    Estimated
}

internal sealed class PlaylistRotationRecord
{
    public DateOnly Date { get; init; }
    public string SocialPlaylist { get; init; } = "";
    public string Ranked4v4Playlist { get; init; } = "";
    public string Ranked2v2Playlist { get; init; } = "";
    public PlaylistRotationStatus Status { get; init; }
    public string SourceNote { get; init; } = "";
}

internal sealed class PlaylistRotationViewRow
{
    public required PlaylistRotationRecord Source { get; init; }
    public required string DateDisplay { get; init; }
    public required string SocialDisplay { get; init; }
    public required string Ranked4v4Display { get; init; }
    public required string Ranked2v2Display { get; init; }
    public required string StatusDisplay { get; init; }
    public required Brush RowBackground { get; init; }
    public required FontWeight RowFontWeight { get; init; }

    public static PlaylistRotationViewRow? FromRecord(PlaylistRotationRecord record, DateOnly selectedDate, RotationScheduleFilter filter)
    {
        if (filter == RotationScheduleFilter.Social && string.IsNullOrWhiteSpace(record.SocialPlaylist)) return null;
        if (filter == RotationScheduleFilter.Ranked4v4 && string.IsNullOrWhiteSpace(record.Ranked4v4Playlist)) return null;
        if (filter == RotationScheduleFilter.Ranked2v2 && string.IsNullOrWhiteSpace(record.Ranked2v2Playlist)) return null;

        var isSelectedWeek = record.Date == selectedDate;
        var isCurrentWeek = record.Status == PlaylistRotationStatus.Confirmed &&
                            record.Date <= DateOnly.FromDateTime(DateTime.Today) &&
                            record.Date.AddDays(7) > DateOnly.FromDateTime(DateTime.Today);

        return new PlaylistRotationViewRow
        {
            Source = record,
            DateDisplay = record.Date.ToString("MMM d, yyyy", CultureInfo.InvariantCulture),
            SocialDisplay = DisplayValue(record.SocialPlaylist, filter, RotationScheduleFilter.Social),
            Ranked4v4Display = DisplayValue(record.Ranked4v4Playlist, filter, RotationScheduleFilter.Ranked4v4),
            Ranked2v2Display = DisplayValue(record.Ranked2v2Playlist, filter, RotationScheduleFilter.Ranked2v2),
            StatusDisplay = record.Status == PlaylistRotationStatus.Estimated ? "Estimated" : isCurrentWeek ? "Current" : "Confirmed",
            RowBackground = BuildBackground(record.Status, isCurrentWeek, isSelectedWeek),
            RowFontWeight = isCurrentWeek || isSelectedWeek ? FontWeights.SemiBold : FontWeights.Normal
        };
    }

    private static string DisplayValue(string value, RotationScheduleFilter activeFilter, RotationScheduleFilter lane) =>
        activeFilter != RotationScheduleFilter.All && activeFilter != lane ? "—" : string.IsNullOrWhiteSpace(value) ? "-" : value;

    private static Brush BuildBackground(PlaylistRotationStatus status, bool isCurrentWeek, bool isSelectedWeek)
    {
        if (isSelectedWeek)
            return new SolidColorBrush(Color.FromArgb(0x22, 0x00, 0xC8, 0xFF));
        if (isCurrentWeek)
            return new SolidColorBrush(Color.FromArgb(0x16, 0x00, 0xC8, 0xFF));
        if (status == PlaylistRotationStatus.Estimated)
            return new SolidColorBrush(Color.FromArgb(0x08, 0xFF, 0xFF, 0xFF));
        return Brushes.Transparent;
    }
}

internal sealed class PlaylistViewSummary
{
    public string Id { get; init; } = "";
    public string RawName { get; init; } = "";
    public string HopperName { get; init; } = "";
    public string StatName { get; init; } = "";
    public bool IsMix { get; init; }
    public bool IsRanked { get; init; }
    public bool IsDebug { get; init; }
    public string MinPlayers { get; init; } = "";
    public string MaxPlayers { get; init; } = "";
    public string MaxPartySize { get; init; } = "";
    public string TeamSize { get; init; } = "";
    public List<PlaylistViewGroup> TagGroups { get; init; } = new();
    public List<PlaylistViewEntry> Entries { get; init; } = new();

    public string ModeLabel => IsRanked ? "Ranked" : "Social";
    public string EffectivePlayerCount => !string.IsNullOrWhiteSpace(MaxPlayers) ? MaxPlayers : MinPlayers;
    public string PlayerSummary => $"{MinPlayers}-{MaxPlayers} players  •  party up to {MaxPartySize}";
    public string TeamSizeDisplay => TeamSize == "1" ? "FFA" : $"Teams of {TeamSize}";

    public string SizeLabel
    {
        get
        {
            if (Entries.Any(x => x.CategoryDisplay.Equals("Firefight", StringComparison.OrdinalIgnoreCase)))
                return "FIREFIGHT";

            if (EffectivePlayerCount == "2" && TeamSize == "1")
                return "1V1";

            if (IsRanked && int.TryParse(EffectivePlayerCount, out var rankedPlayers))
            {
                if (rankedPlayers <= 4) return "2V2";
                if (rankedPlayers <= 8) return "4V4";
                if (rankedPlayers <= 12) return "6V6";
                if (rankedPlayers <= 16) return "8V8";
            }

            if (EffectivePlayerCount == "4" && TeamSize == "2")
                return "2V2";

            if (EffectivePlayerCount == "8" && TeamSize == "4")
                return "4V4";

            if (EffectivePlayerCount == "12" && TeamSize == "1")
                return "12-PLAYER FFA";

            if (EffectivePlayerCount == "16" && TeamSize == "8")
                return "8V8";

            return $"{EffectivePlayerCount}-PLAYER";
        }
    }

    public string SelectionName
    {
        get
        {
            if (HopperName.Equals("Cascade12PFFAMix", StringComparison.OrdinalIgnoreCase))
                return "Infection";
            if (StatName.Equals("h4SquadBattle0", StringComparison.OrdinalIgnoreCase)
                || HopperName.Equals("Cascade12PTeamRanked", StringComparison.OrdinalIgnoreCase))
                return "Halo 4 Squad Battle";
            if (IsRanked && StatName.Equals("h3RankedTeamSlayer0", StringComparison.OrdinalIgnoreCase))
                return "H3 Team Slayer";
            if (IsRanked && StatName.Equals("h3HardcoreDoubles1", StringComparison.OrdinalIgnoreCase))
                return "H3 Hardcore Doubles";
            if (IsRanked && StatName.Equals("h3TeamDoubles0", StringComparison.OrdinalIgnoreCase))
                return "H3 Team Doubles";
            if (IsRanked && StatName.Equals("HCERankedDoubles1", StringComparison.OrdinalIgnoreCase))
                return "HCE Hardcore Doubles";
            if (IsRanked && StatName.Equals("h2aRankedHardcore1", StringComparison.OrdinalIgnoreCase))
                return "H2A Team Hardcore";
            if (IsRanked && StatName.Equals("hrRankedTeamSlayer0", StringComparison.OrdinalIgnoreCase))
                return "Halo Reach Team Slayer";
            if (IsRanked && StatName.Equals("h3RankedHardcore0", StringComparison.OrdinalIgnoreCase))
                return "H3 Team Hardcore";
            if (!IsRanked && HopperName.Equals("Cascade8PTeamMix", StringComparison.OrdinalIgnoreCase))
                return "Social 4v4 Composer";
            if (!IsRanked && HopperName.Equals("Cascade16PTeamMix", StringComparison.OrdinalIgnoreCase))
                return "Social 8v8 Composer";
            if (!IsRanked && StatName.Equals("FirefightDoubles", StringComparison.OrdinalIgnoreCase))
                return "Firefight Doubles";
            if (!IsRanked && StatName.Equals("hrFirefightDoubles0", StringComparison.OrdinalIgnoreCase))
                return "Halo Reach Firefight Doubles";
            if (!IsRanked && StatName.Equals("odstFirefightDoubles0", StringComparison.OrdinalIgnoreCase))
                return "ODST Firefight Doubles";
            if (!string.IsNullOrWhiteSpace(StatName))
                return BeautifyName(StatName);
            if (!string.IsNullOrWhiteSpace(HopperName))
                return BeautifyName(HopperName);
            if (!string.IsNullOrWhiteSpace(Id))
                return BeautifyName(Id);
            return BeautifyName(RawName);
        }
    }

    private static string BeautifyName(string raw)
    {
        var value = raw.Replace('_', ' ');
        value = Regex.Replace(value, "([a-z])([A-Z])", "$1 $2");
        value = Regex.Replace(value, @"\d+$", "");
        value = Regex.Replace(value, @"\s+", " ").Trim();
        return value;
    }

    public override string ToString() => SelectionName;
}

internal sealed class PlaylistViewGroup
{
    public string TagId { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Subtitle { get; init; } = "";
    public int TotalWeight { get; init; }
    public int TotalPlaylistWeight { get; set; }
    public bool IsExpanded { get; set; }
    public List<PlaylistViewEntry> Entries { get; init; } = new();
    public string EntrySummary => $"{Entries.Count} entries";
    public string HeaderSummary => $"{Subtitle}  •  {Entries.Count} entries";
    public string HeaderWeightDisplay => TotalPlaylistWeight <= 0
        ? TotalWeight.ToString()
        : $"{(double)TotalWeight / TotalPlaylistWeight * 100:F1}%  •  {TotalWeight}";
}

internal sealed class PlaylistViewEntry
{
    public string TagId { get; init; } = "";
    public string TagDisplay { get; init; } = "";
    public string Category { get; init; } = "";
    public int Weight { get; init; }
    public string MapId { get; init; } = "";
    public string MapDisplay { get; init; } = "";
    public string MapVariant { get; init; } = "";
    public string GameVariant { get; init; } = "";
    public string GameKey { get; init; } = "";
    public double WeightShare { get; set; }

    public string WeightShareDisplay => $"{WeightShare * 100:F1}%";
    public string MapVariantDisplay => string.IsNullOrWhiteSpace(MapVariant) ? "-" : MapVariant;
    public string CategoryDisplay => string.IsNullOrWhiteSpace(Category) ? "-" : Category.Replace("GameCategory_", "");

    public string GameDisplay => GameKey switch
    {
        "halo1" => "Halo CE",
        "halo2" => "Halo 2",
        "halo2a" => "Halo 2A",
        "halo3" => "Halo 3",
        "halo3odst" => "Halo 3: ODST",
        "halo4" => "Halo 4",
        "haloreach" => "Halo: Reach",
        _ => "Unknown"
    };

    public SolidColorBrush GameBrush => GameKey switch
    {
        "halo1" => new SolidColorBrush(Color.FromRgb(0x73, 0xD9, 0x8C)),
        "halo2" => new SolidColorBrush(Color.FromRgb(0xFF, 0xB8, 0x4D)),
        "halo2a" => new SolidColorBrush(Color.FromRgb(0x39, 0xD0, 0xC8)),
        "halo3" => new SolidColorBrush(Color.FromRgb(0x58, 0xA6, 0xFF)),
        "halo3odst" => new SolidColorBrush(Color.FromRgb(0xD2, 0x99, 0x22)),
        "halo4" => new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49)),
        "haloreach" => new SolidColorBrush(Color.FromRgb(0xBC, 0x8C, 0xF9)),
        _ => new SolidColorBrush(Color.FromRgb(0x7D, 0x85, 0x90))
    };
}
