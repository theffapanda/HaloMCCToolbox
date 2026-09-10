using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

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

    private static readonly string DownpatchIndexFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox", "downpatch-workspaces.json");

    private const string HaloMccAppId = "976730";
    private const string MccBaseDepotId = "976731";
    private const string Halo3MultiplayerDepotId = "976739";
    private static readonly DateTime CurrentBuildNoDownpatchCutoff =
        new(2025, 2, 28, 0, 0, 0, DateTimeKind.Local);

    private static readonly Regex SavedFilmDescriptionDateRegex = new(
        @"\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\b",
        RegexOptions.Compiled);

    private static readonly Regex SavedFilmMapPathRegex = new(
        @"halo3\\maps\\([A-Za-z0-9_]+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Dictionary<string, string> Halo3MapIds = new(StringComparer.OrdinalIgnoreCase)
    {
        ["construct"] = "Construct",
        ["salvation"] = "Epitaph",
        ["guardian"] = "Guardian",
        ["deadlock"] = "High Ground",
        ["isolation"] = "Isolation",
        ["zanzibar"] = "Last Resort",
        ["chill"] = "Narrows",
        ["shrine"] = "Sandtrap",
        ["snowbound"] = "Snowbound",
        ["cyberdyne"] = "The Pit",
        ["riverworld"] = "Valhalla",
        ["warehouse"] = "Foundry",
        ["armory"] = "Rat's Nest",
        ["bunkerworld"] = "Standoff",
        ["sidewinder"] = "Avalanche",
        ["lockout"] = "Blackout",
        ["ghosttown"] = "Ghost Town",
        ["chillout"] = "Cold Storage",
        ["descent"] = "Assembly",
        ["spacecamp"] = "Orbital",
        ["sandbox"] = "Sandbox",
        ["fortress"] = "Citadel",
        ["docks"] = "Longshore",
        ["midship"] = "Heretic",
        ["s3d_waterfall"] = "Waterfall",
        ["s3d_edge"] = "Edge",
        ["s3d_turf"] = "Icebox",
    };

    private static readonly DownpatchManifestPair[] Halo3DownpatchCatalog =
    [
        new("2020-05-13", new DateTime(2020, 5, 13, 3, 8, 38, DateTimeKind.Utc), "7589034107403041459", "3736036753335813549"),
        new("2020-05-22", new DateTime(2020, 5, 22, 2, 59, 32, DateTimeKind.Utc), "2682561640792648202", "176998430543082065"),
        new("2020-07-14", new DateTime(2020, 7, 14, 3, 0, 8, DateTimeKind.Utc), "608110741451978584", "6725604593536929481"),
        new("2020-07-29", new DateTime(2020, 7, 29, 16, 59, 19, DateTimeKind.Utc), "8826895997030691501", "5169039549295643489"),
        new("2020-09-22", new DateTime(2020, 9, 22, 3, 2, 1, DateTimeKind.Utc), "8147576778063649641", "6844316636715771668"),
        new("2020-10-14", new DateTime(2020, 10, 14, 17, 3, 53, DateTimeKind.Utc), "8085881832139938494", "7034980811403089544"),
        new("2020-10-19", new DateTime(2020, 10, 19, 23, 2, 28, DateTimeKind.Utc), "9050062892283056554", "8677987896651073422"),
        new("2020-11-17", new DateTime(2020, 11, 17, 3, 58, 48, DateTimeKind.Utc), "8804662812684949302", "7464644334442106984"),
        new("2020-12-16", new DateTime(2020, 12, 16, 17, 59, 43, DateTimeKind.Utc), "7917039139591885241", "5431621953899624022"),
        new("2021-01-27", new DateTime(2021, 1, 27, 17, 3, 18, DateTimeKind.Utc), "3916960619121697109", "3747952903602120628"),
        new("2021-04-07", new DateTime(2021, 4, 7, 16, 0, 12, DateTimeKind.Utc), "7717469693536934734", "4920557822834955354"),
        new("2021-04-28", new DateTime(2021, 4, 28, 17, 1, 6, DateTimeKind.Utc), "7350425281184071515", "8057158599503793264"),
        new("2021-06-23", new DateTime(2021, 6, 23, 17, 1, 3, DateTimeKind.Utc), "6099447580914881304", "4980497131152773649"),
        new("2021-06-25", new DateTime(2021, 6, 25, 17, 0, 59, DateTimeKind.Utc), "7397581775992835466", "6081273071651234854"),
        new("2021-07-22", new DateTime(2021, 7, 22, 17, 0, 7, DateTimeKind.Utc), "3600222516169950408", "6189276660200426150"),
        new("2021-10-13", new DateTime(2021, 10, 13, 16, 59, 53, DateTimeKind.Utc), "5283684822972828228", "7256534212836780722"),
        new("2021-10-18", new DateTime(2021, 10, 18, 17, 0, 26, DateTimeKind.Utc), "740379036428020971", "7546014327679828214"),
        new("2021-11-03", new DateTime(2021, 11, 3, 17, 0, 35, DateTimeKind.Utc), "4248105459810321792", "1131436886954838062"),
        new("2021-11-30", new DateTime(2021, 11, 30, 17, 0, 11, DateTimeKind.Utc), "1018723109337407870", "855678896425413970"),
        new("2022-04-11", new DateTime(2022, 4, 11, 16, 59, 43, DateTimeKind.Utc), "771679180145980166", "8489219392321953435"),
        new("2022-04-14", new DateTime(2022, 4, 14, 17, 11, 28, DateTimeKind.Utc), "4637054143923919713", "2208456434069436152"),
        new("2022-04-27", new DateTime(2022, 4, 27, 16, 59, 57, DateTimeKind.Utc), "5952135321245530570", "8766808327326615635"),
        new("2022-05-11", new DateTime(2022, 5, 11, 17, 0, 34, DateTimeKind.Utc), "2470982831463401783", "4446442061387975346"),
        new("2022-06-29", new DateTime(2022, 6, 29, 17, 0, 6, DateTimeKind.Utc), "5773691394673377371", "4112232820370899695"),
        new("2022-08-31", new DateTime(2022, 8, 31, 17, 0, 12, DateTimeKind.Utc), "3398001368072059567", "2472348697636137511"),
        new("2022-12-07", new DateTime(2022, 12, 7, 17, 0, 9, DateTimeKind.Utc), "3208837841069898068", "5664726381295792027"),
        new("2022-12-19", new DateTime(2022, 12, 19, 18, 0, 30, DateTimeKind.Utc), "4214904329398965160", "615611205645680843"),
        new("2023-04-05", new DateTime(2023, 4, 5, 17, 1, 1, DateTimeKind.Utc), "5133523842924247498", "2095204872053229091"),
        new("2023-07-12", new DateTime(2023, 7, 12, 17, 0, 16, DateTimeKind.Utc), "7734934274879970574", "3544043389882203801"),
        new("2023-07-20", new DateTime(2023, 7, 20, 17, 0, 58, DateTimeKind.Utc), "864876826350492019", "8906722332519243293"),
        new("2023-09-19", new DateTime(2023, 9, 19, 17, 0, 52, DateTimeKind.Utc), "6734576022768850860", "462579476231054134"),
        new("2024-02-14", new DateTime(2024, 2, 14, 18, 1, 18, DateTimeKind.Utc), "978825021605647054", "9155859885212250288"),
        new("2025-02-28", new DateTime(2025, 2, 28, 20, 51, 51, DateTimeKind.Utc), "5290165787340864952", "411871384308362488"),
        new("2025-09-17", new DateTime(2025, 9, 17, 17, 14, 36, DateTimeKind.Utc), "4956458579120161185", "3902112679418599911"),
    ];

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

    // Map file ids are taken from Halopedia's Map_file reference:
    // https://www.halopedia.org/Map_file
    private static readonly Dictionary<string, (string MapId, string DisplayName)[]> _knownMapFiles = new()
    {
        ["Halo2A"] =
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
        ["Halo3ODST"] =
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
        ["Halo4"] =
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
        ["HaloReach"] =
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

        if (TryResolveKnownMapName(gameKey, noExt, out var resolved))
            return resolved;

        return CleanMapFileName(noExt);
    }

    private static bool TryResolveKnownMapName(string gameKey, string noExt, out string displayName)
    {
        displayName = "";
        if (!_knownMapFiles.TryGetValue(gameKey, out var knownMaps))
            return false;

        var candidates = BuildLookupCandidates(noExt);
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

        foreach (var prefix in new[] { "asq_", "dlc_", "mp_", "ffa_", "coop_", "ms_" })
            if (stripped.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                AddCandidate(stripped[prefix.Length..]);

        var cleaned = CleanMapFileName(stripped);
        AddCandidate(cleaned);

        var tokens = ExtractMeaningfulTokens(stripped);
        for (int start = 0; start < tokens.Count; start++)
        {
            for (int length = 1; length <= tokens.Count - start; length++)
            {
                var slice = tokens.Skip(start).Take(length).ToArray();
                AddCandidate(string.Join("_", slice));
                AddCandidate(string.Concat(slice));
            }
        }

        return candidates.OrderByDescending(candidate => NormalizeLookupKey(candidate).Length).ToList();
    }

    private static List<string> ExtractMeaningfulTokens(string value)
    {
        var sanitized = Regex.Replace(value, @"(?i)\bmglo[-_]*\d+\b", "_");
        sanitized = Regex.Replace(sanitized, @"[^a-zA-Z0-9]+", "_");

        var tokens = new List<string>();
        foreach (var token in sanitized.Split('_', StringSplitOptions.RemoveEmptyEntries))
        {
            if (LooksLikeHashOrTimestamp(token))
                break;

            if (token.Equals("asq", StringComparison.OrdinalIgnoreCase))
                continue;

            tokens.Add(token);
        }

        return tokens;
    }

    private static bool MatchesMapId(string candidate, string mapId)
    {
        var candidateKey = NormalizeLookupKey(candidate);
        var mapKey = NormalizeLookupKey(mapId);

        if (candidateKey.Length == 0 || mapKey.Length == 0)
            return false;

        if (candidateKey.Equals(mapKey, StringComparison.Ordinal))
            return true;

        if (candidateKey.Length >= 5 && candidateKey.Contains(mapKey, StringComparison.Ordinal))
            return true;

        return candidateKey.Length >= 5 && mapKey.StartsWith(candidateKey, StringComparison.Ordinal);
    }

    private static string NormalizeLookupKey(string value)
        => new(value.Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());

    private static bool LooksLikeHashOrTimestamp(string value)
        => value.Length >= 6 && value.All(c => char.IsDigit(c) ||
            (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'));

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
            if (LooksLikeHashOrTimestamp(part))
                break; // hit a hash/timestamp — stop

            meaningful.Add(char.ToUpperInvariant(part[0]) + part[1..].ToLowerInvariant());
        }

        return meaningful.Count > 0 ? string.Join(" ", meaningful) : noExt;
    }

    // ── Instance state ─────────────────────────────────────────────────────────

    private readonly ObservableCollection<TheaterClip> _clips = new();
    private readonly ICollectionView _view;
    private readonly List<FileSystemWatcher> _watchers = new();
    private DownpatchWorkspaceIndex _downpatchIndex = new();
    private TheaterClip? _externalDownpatchClip;
    private bool _settingExternalDownpatchClip;
    private readonly System.Windows.Threading.DispatcherTimer _downpatchLogWatchTimer = new();
    private DownpatchWorkflowStep _downpatchWorkflowStep = DownpatchWorkflowStep.Idle;
    private DateTime _downpatchWatchStartedAt = DateTime.MinValue;
    private bool _downpatchStagingActive;
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
        CboDownpatchClip.ItemsSource = _clips;
        TxtDownpatchWorkspace.Text = App.LoadDownpatchWorkspacePath();
        LoadDownpatchIndex();
        _downpatchLogWatchTimer.Interval = TimeSpan.FromSeconds(2);
        _downpatchLogWatchTimer.Tick += DownpatchLogWatchTimer_Tick;
        RefreshDownpatchUi();
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
        UpdateSelectAllButton();
        if (CboDownpatchClip is not null && CboDownpatchClip.SelectedItem is null && _clips.Count > 0)
            CboDownpatchClip.SelectedItem = _clips.FirstOrDefault(c => c.GameKey == "Halo3") ?? _clips[0];
        RefreshDownpatchUi();
    }

    private TheaterClip BuildClip(string key, FileInfo fi, bool sourcePresent)
    {
        var backupPath = Path.Combine(BackupRoot, key, fi.Name);
        var (displayName, accent) = GameInfo[key];
        var filmDate = ReadSavedFilmDate(fi.FullName, fi.LastWriteTime);
        var clip = new TheaterClip
        {
            Game           = displayName,
            GameKey        = key,
            FileName       = fi.Name,
            MapName        = Path.GetFileNameWithoutExtension(fi.Name),
            MapDisplayName = ResolveMapName(key, fi.Name),
            FileSizeBytes  = fi.Length,
            RecordedAt     = filmDate.RecordedAt,
            RecordedAtSource = filmDate.Source,
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
        var filmDate = ReadSavedFilmDate(backupFi.FullName, backupFi.LastWriteTime);
        var clip = new TheaterClip
        {
            Game           = displayName,
            GameKey        = key,
            FileName       = backupFi.Name,
            MapName        = Path.GetFileNameWithoutExtension(backupFi.Name),
            MapDisplayName = ResolveMapName(key, backupFi.Name),
            FileSizeBytes  = backupFi.Length,
            RecordedAt     = filmDate.RecordedAt,
            RecordedAtSource = filmDate.Source,
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

    private static SavedFilmDate ReadSavedFilmDate(string path, DateTime fallback)
        => ReadSavedFilmHeader(path, fallback).Date;

    private static SavedFilmHeader ReadSavedFilmHeader(string path, DateTime fallback)
    {
        try
        {
            var fileInfo = new FileInfo(path);
            var buffer = new byte[Math.Min(4096, fileInfo.Length)];
            using (var stream = File.OpenRead(path))
            {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read < buffer.Length)
                    Array.Resize(ref buffer, read);
            }

            var headerText = Encoding.ASCII.GetString(buffer);
            string gameKey = headerText.Contains("halo 3 saved film", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Contains(@"halo3\maps\", StringComparison.OrdinalIgnoreCase)
                ? "Halo3"
                : "";
            string mapId = "";
            string mapDisplayName = "";
            var mapMatch = SavedFilmMapPathRegex.Match(headerText);
            if (mapMatch.Success)
            {
                mapId = mapMatch.Groups[1].Value.TrimEnd('\0');
                if (Halo3MapIds.TryGetValue(mapId, out var display))
                    mapDisplayName = display;
            }

            var match = SavedFilmDescriptionDateRegex.Match(headerText);
        if (match.Success &&
            DateTime.TryParseExact(
                $"{match.Groups[1].Value} {match.Groups[2].Value}, {match.Groups[3].Value} {match.Groups[4].Value}:{match.Groups[5].Value}:{match.Groups[6].Value}",
                "MMMM d, yyyy H:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out var embeddedDate))
        {
            embeddedDate = DateTime.SpecifyKind(embeddedDate, DateTimeKind.Local);
            return new SavedFilmHeader(
                new SavedFilmDate(embeddedDate, "Saved film header"),
                gameKey,
                    mapId,
                    mapDisplayName,
                    headerText);
            }

            if (buffer.Length >= 0x110)
            {
                uint unixSeconds = BitConverter.ToUInt32(buffer, 0x10C);
                if (unixSeconds > 1_500_000_000 && unixSeconds < 2_200_000_000)
                {
                    var local = DateTimeOffset.FromUnixTimeSeconds(unixSeconds).LocalDateTime;
                    return new SavedFilmHeader(
                        new SavedFilmDate(local, "Saved film timestamp"),
                        gameKey,
                        mapId,
                        mapDisplayName,
                        headerText);
                }
            }
        }
        catch
        {
            // Fall back to the filesystem date below.
        }

        return new SavedFilmHeader(new SavedFilmDate(fallback, "File modified"), "", "", "", "");
    }

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
        {
            clip.IsSelected = !clip.IsSelected;
            UpdateSelectAllButton();
        }
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
        UpdateSelectAllButton();
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
        ConfirmAndDelete([clip], $"\"{clip.DisplayName}\"");
    }

    private void ConfirmAndDelete(List<TheaterClip> clips, string description)
    {
        if (clips.Count == 0) return;

        var plural = clips.Count == 1 ? "clip" : "clips";
        var result = MessageBox.Show(
            $"Delete {description}?\n\nThis will remove the backup copy. Source files (if present) will also be deleted.",
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes) return;

        var (deleted, failed) = ExecuteDelete(clips);
        RebuildByGameMenu();
        UpdateStatus(failed == 0
            ? $"Deleted {deleted} {plural}."
            : $"Deleted {deleted} {plural}; {failed} failed.");
        UpdateSelectAllButton();
    }

    private (int Deleted, int Failed) ExecuteDelete(IEnumerable<TheaterClip> clips)
    {
        int deleted = 0;
        int failed = 0;

        foreach (var clip in clips.ToList())
        {
            try
            {
                if (File.Exists(clip.BackupPath))
                    File.Delete(clip.BackupPath);

                if (clip.SourcePresent && File.Exists(clip.SourcePath))
                    File.Delete(clip.SourcePath);

                _customNames.Remove(CustomNamesKey(clip));
                _clips.Remove(clip);
                deleted++;
            }
            catch
            {
                failed++;
            }
        }

        if (deleted > 0)
            SaveCustomNames();

        return (deleted, failed);
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

    private void LoadDownpatchIndex()
    {
        try
        {
            if (!File.Exists(DownpatchIndexFile))
            {
                _downpatchIndex = new DownpatchWorkspaceIndex();
                return;
            }

            _downpatchIndex = JsonSerializer.Deserialize<DownpatchWorkspaceIndex>(
                File.ReadAllText(DownpatchIndexFile)) ?? new DownpatchWorkspaceIndex();
        }
        catch
        {
            _downpatchIndex = new DownpatchWorkspaceIndex();
        }
    }

    private void SaveDownpatchIndex()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(DownpatchIndexFile)!);
            File.WriteAllText(
                DownpatchIndexFile,
                JsonSerializer.Serialize(_downpatchIndex, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private void RefreshDownpatchUi()
    {
        if (TxtDownpatchVersion is null)
            return;

        var clip = GetSelectedDownpatchClip();
        var workspace = TxtDownpatchWorkspace?.Text.Trim() ?? "";
        if (clip is null)
        {
            TxtDownpatchClipDetails.Text = "No clip selected.";
            TxtDownpatchDateDetails.Text = "";
            TxtDownpatchVersion.Text = "Select a Halo 3 theater clip.";
            TxtDownpatchReuse.Text = "";
            TxtDownpatchBaseCommand.Text = "";
            TxtDownpatchMultiplayerCommand.Text = "";
            TxtDownpatchNotes.Text = "";
            SetDownpatchLaunchCommand("");
            SetDownpatchPlaybackButtons(false);
            SetDownpatchDepotButtons(false);
            return;
        }

        TxtDownpatchClipDetails.Text =
            $"{clip.DisplayName}\n{clip.Game} / {clip.MapDisplayName}\n{clip.FileName}\n{clip.FileSizeStr}";
        TxtDownpatchDateDetails.Text =
            $"Detected date: {clip.RecordedAt:MMMM d, yyyy h:mm:ss tt}\nSource: {clip.RecordedAtSource}";

        if (!clip.GameKey.Equals("Halo3", StringComparison.OrdinalIgnoreCase))
        {
            TxtDownpatchVersion.Text = "Downpatch recovery is currently mapped for Halo 3 clips only.";
            TxtDownpatchReuse.Text = "";
            TxtDownpatchBaseCommand.Text = "";
            TxtDownpatchMultiplayerCommand.Text = "";
            TxtDownpatchNotes.Text = "Theater backup still works for this clip, but the Steam depot pairing has not been mapped for this game yet.";
            SetDownpatchLaunchCommand("");
            SetDownpatchPlaybackButtons(false);
            SetDownpatchDepotButtons(false);
            return;
        }

        if (ToUtc(clip.RecordedAt) >= ToUtc(CurrentBuildNoDownpatchCutoff))
        {
            TxtDownpatchVersion.Text = "Current MCC build should read this clip date.";
            TxtDownpatchReuse.Text = "No depot download is needed for Halo 3 films from February 28, 2025 or newer. Install a safe-named clip, then launch current MCC with EAC disabled.";
            TxtDownpatchBaseCommand.Text = "";
            TxtDownpatchMultiplayerCommand.Text = "";

            var currentNotes = new StringBuilder();
            currentNotes.AppendLine("The toolbox automatically installs a safe copy into:");
            currentNotes.AppendLine(Path.Combine(TheaterRoot, "Halo3", "Movie"));
            currentNotes.AppendLine();
            currentNotes.AppendLine("The safe copy keeps the original file and uses MCC's asq_<map>_<hash>_<hash>.mov naming pattern.");
            currentNotes.AppendLine("Then use LAUNCH NO-EAC to start the configured current MCC install directly.");
            TxtDownpatchNotes.Text = currentNotes.ToString();

            var currentMccRoot = App.LoadMccInstallationPath();
            SetDownpatchLaunchCommand(BuildNoEacLaunchCommand(currentMccRoot));
            SetDownpatchPlaybackButtons(!string.IsNullOrWhiteSpace(currentMccRoot));
            SetDownpatchDepotButtons(false);
            return;
        }

        var version = ResolveHalo3Version(clip.RecordedAt);
        if (version is null)
        {
            TxtDownpatchVersion.Text = "No built-in manifest pair covers this clip date.";
            TxtDownpatchReuse.Text = "Use manual manifests for now, or add this date to the catalog.";
            TxtDownpatchBaseCommand.Text = "";
            TxtDownpatchMultiplayerCommand.Text = "";
            TxtDownpatchNotes.Text = "The current built-in catalog starts at January 27, 2021.";
            SetDownpatchLaunchCommand("");
            SetDownpatchPlaybackButtons(false);
            SetDownpatchDepotButtons(false);
            return;
        }

        var next = NextHalo3Version(version);
        var windowText = next is null
            ? $"{version.EffectiveUtc:yyyy-MM-dd} and newer"
            : $"{version.EffectiveUtc:yyyy-MM-dd} through {next.EffectiveUtc.AddSeconds(-1):yyyy-MM-dd}";
        var suggestedFolder = BuildSuggestedDownpatchFolder(workspace, version, next);
        var tracked = FindTrackedDownpatch(clip.GameKey, version);
        bool trackedFolderExists = tracked is not null && Directory.Exists(tracked.FolderPath);
        bool suggestedExists = !string.IsNullOrWhiteSpace(suggestedFolder) && Directory.Exists(suggestedFolder);
        var launchFolder = trackedFolderExists ? tracked!.FolderPath : suggestedExists ? suggestedFolder : "";

        TxtDownpatchVersion.Text = $"Version window: {windowText}";
        TxtDownpatchReuse.Text = trackedFolderExists
            ? $"Reusable folder already tracked: {tracked!.FolderPath}"
            : suggestedExists
                ? $"Matching folder exists in workspace: {suggestedFolder}"
                : "No reusable downpatch folder is tracked for this version yet.";
        if (BtnCopyNextDownpatchCommand is not null)
        {
            BtnCopyNextDownpatchCommand.Content = _downpatchWorkflowStep switch
            {
                DownpatchWorkflowStep.BaseComplete => "COPY MULTIPLAYER COMMAND",
                DownpatchWorkflowStep.WatchingBase => "WATCHING BASE",
                DownpatchWorkflowStep.WatchingMultiplayer => "WATCHING MULTIPLAYER",
                DownpatchWorkflowStep.Staging => "STAGING DEPOTS",
                DownpatchWorkflowStep.Complete => "COPY BASE COMMAND",
                _ => "COPY BASE COMMAND"
            };
            BtnCopyNextDownpatchCommand.IsEnabled =
                _downpatchWorkflowStep is not DownpatchWorkflowStep.WatchingBase
                    and not DownpatchWorkflowStep.WatchingMultiplayer
                    and not DownpatchWorkflowStep.Staging;
        }
        bool hasCompleteManifestPair = !string.IsNullOrWhiteSpace(version.BaseManifestId) &&
            !string.IsNullOrWhiteSpace(version.MultiplayerManifestId);
        SetDownpatchDepotButtons(hasCompleteManifestPair);

        TxtDownpatchBaseCommand.Text = $"download_depot {HaloMccAppId} {MccBaseDepotId} {version.BaseManifestId}";
        TxtDownpatchMultiplayerCommand.Text = string.IsNullOrWhiteSpace(version.MultiplayerManifestId)
            ? "Missing H3 multiplayer manifest ID for this patch timestamp."
            : $"download_depot {HaloMccAppId} {Halo3MultiplayerDepotId} {version.MultiplayerManifestId}";

        var notes = new StringBuilder();
        notes.AppendLine(BuildDownpatchWorkflowInstruction());
        notes.AppendLine();
        notes.AppendLine($"Workspace target: {(string.IsNullOrWhiteSpace(suggestedFolder) ? "(choose a workspace folder)" : suggestedFolder)}");
        notes.AppendLine($"Base depot cache: steamapps\\content\\app_{HaloMccAppId}\\depot_{MccBaseDepotId}");
        notes.AppendLine($"Multiplayer depot cache: steamapps\\content\\app_{HaloMccAppId}\\depot_{Halo3MultiplayerDepotId}");
        notes.AppendLine();
        notes.AppendLine("The toolbox watches Steam's console log for each Depot download complete line.");
        notes.AppendLine("After both depots finish, files are staged and tracked automatically in the isolated workspace folder.");
        notes.AppendLine("Staging writes steam_appid.txt to both the folder root and MCC\\Binaries\\Win64.");
        notes.AppendLine("The toolbox automatically installs a safe theater copy when you drop or browse to a Halo 3 saved film.");
        if (!hasCompleteManifestPair)
            notes.AppendLine("This patch timestamp is known, but the matching H3 multiplayer manifest ID is not in the local catalog yet.");
        if (clip.RecordedAt >= new DateTime(2022, 8, 31))
            notes.AppendLine("This date may also need the Halo 3 campaign/DLC step in Steam before theater playback works.");
        TxtDownpatchNotes.Text = notes.ToString();

        SetDownpatchLaunchCommand(string.IsNullOrWhiteSpace(launchFolder)
            ? ""
            : BuildNoEacLaunchCommand(launchFolder));
        SetDownpatchPlaybackButtons(!string.IsNullOrWhiteSpace(launchFolder));
    }

    private string BuildDownpatchWorkflowInstruction()
    {
        return _downpatchWorkflowStep switch
        {
            DownpatchWorkflowStep.WatchingBase =>
                "Step 1: paste/enter the Base command in Steam Console. Watching for Base depot completion...",
            DownpatchWorkflowStep.BaseComplete =>
                "Step 2: Base is complete. Click COPY NEXT COMMAND, paste/enter the Multiplayer command in Steam Console.",
            DownpatchWorkflowStep.WatchingMultiplayer =>
                "Step 2: paste/enter the Multiplayer command in Steam Console. Watching for Multiplayer depot completion...",
            DownpatchWorkflowStep.Staging =>
                "Step 3: both depots are complete. Copying files into the isolated downpatch folder...",
            DownpatchWorkflowStep.Complete =>
                "Complete: both depots finished, and the downpatch folder has been staged/tracked.",
            _ =>
                "Step 1: click COPY NEXT COMMAND, paste/enter the Base command in Steam Console, then leave this tab open while it watches the log."
        };
    }

    private static DownpatchManifestPair? ResolveHalo3Version(DateTime clipDate)
    {
        var clipUtc = ToUtc(clipDate);
        return Halo3DownpatchCatalog
            .Where(v => v.EffectiveUtc <= clipUtc)
            .OrderByDescending(v => v.EffectiveUtc)
            .FirstOrDefault();
    }

    private static DateTime ToUtc(DateTime value)
        => value.Kind == DateTimeKind.Utc
            ? value
            : DateTime.SpecifyKind(value, value.Kind == DateTimeKind.Unspecified ? DateTimeKind.Local : value.Kind).ToUniversalTime();

    private static DownpatchManifestPair? NextHalo3Version(DownpatchManifestPair version)
    {
        return Halo3DownpatchCatalog
            .Where(v => v.EffectiveUtc > version.EffectiveUtc)
            .OrderBy(v => v.EffectiveUtc)
            .FirstOrDefault();
    }

    private DownpatchWorkspaceEntry? FindTrackedDownpatch(string gameKey, DownpatchManifestPair version)
    {
        return _downpatchIndex.Entries
            .Where(e => e.GameKey.Equals(gameKey, StringComparison.OrdinalIgnoreCase) &&
                        e.BaseManifestId == version.BaseManifestId &&
                        e.MultiplayerManifestId == version.MultiplayerManifestId)
            .OrderByDescending(e => Directory.Exists(e.FolderPath))
            .ThenByDescending(e => e.TrackedAtUtc)
            .FirstOrDefault();
    }

    private static string BuildSuggestedDownpatchFolder(string workspace, DownpatchManifestPair version, DownpatchManifestPair? next)
    {
        if (string.IsNullOrWhiteSpace(workspace))
            return "";

        var end = next is null ? "current" : next.EffectiveUtc.AddSeconds(-1).ToString("yyyyMMdd", CultureInfo.InvariantCulture);
        var folderName = $"Halo3_{version.EffectiveUtc:yyyyMMdd}_to_{end}_base-{ShortManifest(version.BaseManifestId)}_mp-{ShortManifest(version.MultiplayerManifestId)}";
        return Path.Combine(workspace, folderName);
    }

    private static string ShortManifest(string manifestId)
        => manifestId.Length <= 8 ? manifestId : manifestId[..8];

    private static string BuildNoEacLaunchCommand(string mccRoot)
    {
        if (string.IsNullOrWhiteSpace(mccRoot))
            return "";

        var exe = Path.Combine(mccRoot, "MCC", "Binaries", "Win64", "MCC-Win64-Shipping.exe");
        return $"Set-Location \"{mccRoot}\"{Environment.NewLine}& \"{exe}\" -eac-nop-loaded";
    }

    private void SetDownpatchLaunchCommand(string command)
    {
        if (TxtDownpatchLaunchCommand is not null)
            TxtDownpatchLaunchCommand.Text = command;
    }

    private void SetDownpatchPlaybackButtons(bool canLaunch)
    {
        if (BtnCopyDownpatchLaunchCommand is not null)
            BtnCopyDownpatchLaunchCommand.IsEnabled = canLaunch;
        if (BtnLaunchDownpatch is not null)
            BtnLaunchDownpatch.IsEnabled = canLaunch;
    }

    private void SetDownpatchDepotButtons(bool enabled)
    {
        if (BtnOpenSteamConsole is not null)
            BtnOpenSteamConsole.IsEnabled = enabled;
        if (BtnCopyNextDownpatchCommand is not null)
            BtnCopyNextDownpatchCommand.IsEnabled = enabled &&
                _downpatchWorkflowStep is not DownpatchWorkflowStep.WatchingBase
                    and not DownpatchWorkflowStep.WatchingMultiplayer
                    and not DownpatchWorkflowStep.Staging;
    }

    private TheaterClip? GetSelectedDownpatchClip()
        => _externalDownpatchClip ?? CboDownpatchClip?.SelectedItem as TheaterClip;

    private TheaterClip? BuildExternalDownpatchClip(string path)
    {
        try
        {
            var fi = new FileInfo(path);
            if (!fi.Exists || !fi.Extension.Equals(".mov", StringComparison.OrdinalIgnoreCase))
                return null;

            var header = ReadSavedFilmHeader(path, fi.LastWriteTime);
            var gameKey = header.GameKey;
            if (string.IsNullOrWhiteSpace(gameKey))
                gameKey = header.HeaderText.Contains("halo 3 saved film", StringComparison.OrdinalIgnoreCase) ? "Halo3" : "";

            if (string.IsNullOrWhiteSpace(gameKey))
                return null;

            var displayName = GameInfo.TryGetValue(gameKey, out var info)
                ? info.DisplayName
                : gameKey;
            var accent = GameInfo.TryGetValue(gameKey, out var colorInfo)
                ? colorInfo.Accent
                : Colors.Gray;
            var mapName = !string.IsNullOrWhiteSpace(header.MapId)
                ? header.MapId
                : Path.GetFileNameWithoutExtension(fi.Name);
            var mapDisplay = !string.IsNullOrWhiteSpace(header.MapDisplayName)
                ? header.MapDisplayName
                : ResolveMapName(gameKey, fi.Name);

            return new TheaterClip
            {
                Game = displayName,
                GameKey = gameKey,
                FileName = fi.Name,
                MapName = mapName,
                MapDisplayName = mapDisplay,
                FileSizeBytes = fi.Length,
                RecordedAt = header.Date.RecordedAt,
                RecordedAtSource = header.Date.Source,
                SourcePath = fi.FullName,
                BackupPath = fi.FullName,
                IsBackedUp = true,
                SourcePresent = true,
                GameBrush = new SolidColorBrush(accent),
                CustomName = Path.GetFileNameWithoutExtension(fi.Name)
            };
        }
        catch
        {
            return null;
        }
    }

    private void BtnScan_Click(object sender, RoutedEventArgs e) => InitialScan();

    private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(BackupRoot);
        Process.Start("explorer.exe", BackupRoot);
    }

    private void BtnDownpatchRefresh_Click(object sender, RoutedEventArgs e)
    {
        LoadDownpatchIndex();
        RefreshDownpatchUi();
    }

    private void BtnBrowseDownpatchWorkspace_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog
        {
            Title = "Select a folder for isolated MCC downpatch versions",
            InitialDirectory = Directory.Exists(TxtDownpatchWorkspace.Text.Trim())
                ? TxtDownpatchWorkspace.Text.Trim()
                : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        };
        if (dlg.ShowDialog() != true)
            return;

        TxtDownpatchWorkspace.Text = dlg.FolderName;
        App.SaveDownpatchWorkspacePath(dlg.FolderName);
        RefreshDownpatchUi();
    }

    private void BtnOpenDownpatchWorkspace_Click(object sender, RoutedEventArgs e)
    {
        var workspace = TxtDownpatchWorkspace.Text.Trim();
        if (string.IsNullOrWhiteSpace(workspace))
        {
            RefreshDownpatchUi();
            return;
        }

        Directory.CreateDirectory(workspace);
        Process.Start("explorer.exe", workspace);
    }

    private void CboDownpatchClip_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_settingExternalDownpatchClip && CboDownpatchClip.SelectedItem is not null)
        {
            _externalDownpatchClip = null;
            _downpatchWorkflowStep = DownpatchWorkflowStep.Idle;
            _downpatchLogWatchTimer.Stop();
        }
        RefreshDownpatchUi();
    }

    private void BtnBrowseDownpatchClip_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select a Halo saved-film .mov",
            Filter = "Halo saved films (*.mov)|*.mov|All files (*.*)|*.*",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        };
        if (dlg.ShowDialog() == true)
            LoadExternalDownpatchClip(dlg.FileName);
    }

    private void DownpatchClipDropZone_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = HasMovDrop(e) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void DownpatchClipDropZone_Drop(object sender, DragEventArgs e)
    {
        if (!HasMovDrop(e))
            return;

        var path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
        LoadExternalDownpatchClip(path);
        e.Handled = true;
    }

    private static bool HasMovDrop(DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            return false;

        var files = e.Data.GetData(DataFormats.FileDrop) as string[];
        return files?.Any(file => Path.GetExtension(file).Equals(".mov", StringComparison.OrdinalIgnoreCase)) == true;
    }

    private void LoadExternalDownpatchClip(string path)
    {
        var clip = BuildExternalDownpatchClip(path);
        if (clip is null)
        {
            TxtDownpatchReuse.Text = "That file does not look like a supported Halo saved-film .mov.";
            return;
        }

        string? postLoadStatus = null;
        if (clip.GameKey.Equals("Halo3", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var installedPath = InstallSafeTheaterClip(clip);
                postLoadStatus = $"Installed safe theater copy: {installedPath}";

                var installedClip = BuildExternalDownpatchClip(installedPath);
                if (installedClip is not null)
                    clip = installedClip;

                InitialScan();
            }
            catch (Exception ex)
            {
                postLoadStatus = $"Safe clip install failed: {ex.Message}";
            }
        }

        _externalDownpatchClip = clip;
        _downpatchWorkflowStep = DownpatchWorkflowStep.Idle;
        _downpatchLogWatchTimer.Stop();
        _settingExternalDownpatchClip = true;
        try
        {
            CboDownpatchClip.SelectedItem = null;
        }
        finally
        {
            _settingExternalDownpatchClip = false;
        }

        RefreshDownpatchUi();
        if (!string.IsNullOrWhiteSpace(postLoadStatus))
            TxtDownpatchReuse.Text = postLoadStatus;
    }

    private void TxtDownpatchWorkspace_TextChanged(object sender, TextChangedEventArgs e)
    {
        var workspace = TxtDownpatchWorkspace.Text.Trim();
        if (!string.IsNullOrWhiteSpace(workspace))
            App.SaveDownpatchWorkspacePath(workspace);
        RefreshDownpatchUi();
    }

    private async void DownpatchLogWatchTimer_Tick(object? sender, EventArgs e)
    {
        var clip = GetSelectedDownpatchClip();
        if (clip is null || ResolveHalo3Version(clip.RecordedAt) is not { } version)
            return;

        var log = ReadSteamConsoleLog();
        if (string.IsNullOrWhiteSpace(log))
            return;

        if (_downpatchWorkflowStep == DownpatchWorkflowStep.WatchingBase)
        {
            UpdateDownloadStartedText(log, MccBaseDepotId, version.BaseManifestId, "Base");
            if (HasDepotCompletion(log, MccBaseDepotId, version.BaseManifestId, _downpatchWatchStartedAt))
            {
                _downpatchWorkflowStep = DownpatchWorkflowStep.BaseComplete;
                _downpatchLogWatchTimer.Stop();
                TxtDownpatchReuse.Text = "Base depot complete. Copy the Multiplayer command next.";
                RefreshDownpatchUi();
            }
        }
        else if (_downpatchWorkflowStep == DownpatchWorkflowStep.WatchingMultiplayer)
        {
            UpdateDownloadStartedText(log, Halo3MultiplayerDepotId, version.MultiplayerManifestId, "Multiplayer");
            if (HasDepotCompletion(log, Halo3MultiplayerDepotId, version.MultiplayerManifestId, _downpatchWatchStartedAt))
            {
                _downpatchLogWatchTimer.Stop();
                _downpatchWorkflowStep = DownpatchWorkflowStep.Staging;
                TxtDownpatchReuse.Text = "Both depots complete. Preparing to copy files into the isolated downpatch folder...";
                RefreshDownpatchUi();
                bool staged = await StageDownpatchDepotsAsync(clip, version);
                _downpatchWorkflowStep = staged ? DownpatchWorkflowStep.Complete : DownpatchWorkflowStep.BaseComplete;
                RefreshDownpatchUi();
            }
        }
    }

    private static bool HasDepotCompletion(string log, string depotId, string manifestId, DateTime after)
    {
        var pattern = $@"\[(?<time>[^\]]+)\]\s+Depot download complete\s*:\s*"".*?depot_{Regex.Escape(depotId)}""\s*\(manifest {Regex.Escape(manifestId)}\)";
        foreach (Match match in Regex.Matches(log, pattern, RegexOptions.IgnoreCase))
        {
            if (DateTime.TryParseExact(
                    match.Groups["time"].Value,
                    "yyyy-MM-dd HH:mm:ss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal,
                    out var timestamp) &&
                timestamp >= after)
            {
                return true;
            }
        }

        return false;
    }

    private void UpdateDownloadStartedText(string log, string depotId, string manifestId, string label)
    {
        var pattern = $@"Downloading depot {Regex.Escape(depotId)} \((?<files>\d+) files, (?<mb>\d+) MB\)";
        var matches = Regex.Matches(log, pattern, RegexOptions.IgnoreCase);
        if (matches.Count == 0)
            return;

        var match = matches[matches.Count - 1];
        TxtDownpatchReuse.Text =
            $"{label} download detected: {match.Groups["files"].Value} files, {match.Groups["mb"].Value} MB. Waiting for manifest {manifestId} completion...";
    }

    private static string ReadSteamConsoleLog()
    {
        try
        {
            var steamapps = ResolveSteamAppsFolder(App.LoadMccInstallationPath());
            var logPath = EnumerateSteamappsCandidates(steamapps)
                .Select(path => Directory.GetParent(path)?.FullName)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => Path.Combine(path!, "logs", "console_log.txt"))
                .FirstOrDefault(File.Exists);
            if (string.IsNullOrWhiteSpace(logPath))
                return "";

            using var stream = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch
        {
            return "";
        }
    }

    private void BtnCopyNextDownpatchCommand_Click(object sender, RoutedEventArgs e)
    {
        var clip = GetSelectedDownpatchClip();
        if (clip is null || ResolveHalo3Version(clip.RecordedAt) is null)
            return;
        var selectedVersion = ResolveHalo3Version(clip.RecordedAt);
        if (selectedVersion is null ||
            string.IsNullOrWhiteSpace(selectedVersion.BaseManifestId) ||
            string.IsNullOrWhiteSpace(selectedVersion.MultiplayerManifestId))
        {
            TxtDownpatchReuse.Text = "This patch timestamp is known, but the complete base + multiplayer manifest pair is not in the toolbox catalog yet.";
            return;
        }

        if (_downpatchWorkflowStep is DownpatchWorkflowStep.Idle or DownpatchWorkflowStep.Complete)
        {
            Clipboard.SetText(TxtDownpatchBaseCommand.Text);
            _downpatchWatchStartedAt = DateTime.Now.AddSeconds(-5);
            _downpatchWorkflowStep = DownpatchWorkflowStep.WatchingBase;
            _downpatchLogWatchTimer.Start();
        }
        else if (_downpatchWorkflowStep == DownpatchWorkflowStep.BaseComplete)
        {
            Clipboard.SetText(TxtDownpatchMultiplayerCommand.Text);
            _downpatchWatchStartedAt = DateTime.Now.AddSeconds(-5);
            _downpatchWorkflowStep = DownpatchWorkflowStep.WatchingMultiplayer;
            _downpatchLogWatchTimer.Start();
        }

        RefreshDownpatchUi();
    }

    private void BtnOpenSteamConsole_Click(object sender, RoutedEventArgs e)
    {
        Process.Start(new ProcessStartInfo("steam://open/console") { UseShellExecute = true });
    }

    private void BtnCopyDownpatchLaunchCommand_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(TxtDownpatchLaunchCommand.Text))
            return;

        Clipboard.SetText(TxtDownpatchLaunchCommand.Text);
        TxtDownpatchReuse.Text = "Copied no-EAC launch command.";
    }

    private void BtnLaunchDownpatch_Click(object sender, RoutedEventArgs e)
    {
        var root = ResolvePlaybackRoot();
        if (string.IsNullOrWhiteSpace(root))
        {
            TxtDownpatchReuse.Text = "No launchable MCC folder is available yet.";
            return;
        }

        var exe = Path.Combine(root, "MCC", "Binaries", "Win64", "MCC-Win64-Shipping.exe");
        if (!File.Exists(exe))
        {
            TxtDownpatchReuse.Text = $"MCC shipping exe not found: {exe}";
            return;
        }

        string markerWarning = "";
        try
        {
            WriteSteamAppIdMarkers(root);
        }
        catch (Exception ex)
        {
            markerWarning = $" Could not write steam_appid.txt markers: {ex.Message}";
        }

        try
        {
            Process.Start(new ProcessStartInfo(exe)
            {
                WorkingDirectory = root,
                Arguments = "-eac-nop-loaded",
                UseShellExecute = false
            });
            TxtDownpatchReuse.Text = $"Launched no-EAC MCC from: {root}{markerWarning}";
        }
        catch (Exception ex)
        {
            TxtDownpatchReuse.Text = $"Launch failed: {ex.Message}";
        }
    }

    private string ResolvePlaybackRoot()
    {
        var clip = GetSelectedDownpatchClip();
        if (clip is null)
            return "";

        if (ToUtc(clip.RecordedAt) >= ToUtc(CurrentBuildNoDownpatchCutoff))
            return App.LoadMccInstallationPath();

        var version = ResolveHalo3Version(clip.RecordedAt);
        if (version is null)
            return "";

        var tracked = FindTrackedDownpatch(clip.GameKey, version);
        if (tracked is not null && Directory.Exists(tracked.FolderPath))
            return tracked.FolderPath;

        var workspace = TxtDownpatchWorkspace.Text.Trim();
        var suggested = BuildSuggestedDownpatchFolder(workspace, version, NextHalo3Version(version));
        return Directory.Exists(suggested) ? suggested : "";
    }

    private static string InstallSafeTheaterClip(TheaterClip clip)
    {
        if (!clip.GameKey.Equals("Halo3", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Safe theater install is currently supported for Halo 3 films only.");

        var source = File.Exists(clip.SourcePath) ? clip.SourcePath : clip.BackupPath;
        if (!File.Exists(source))
            throw new FileNotFoundException("Saved film source file was not found.", source);

        var header = ReadSavedFilmHeader(source, clip.RecordedAt);
        var movieFolder = Path.Combine(TheaterRoot, "Halo3", "Movie");
        Directory.CreateDirectory(movieFolder);

        if (IsInDirectory(source, movieFolder) && IsSafeHalo3MovieName(Path.GetFileName(source)))
            return source;

        var safeName = BuildSafeHalo3MovieName(source, header);
        var destination = MakeUniquePath(Path.Combine(movieFolder, safeName), source);
        File.Copy(source, destination, overwrite: false);
        File.SetCreationTimeUtc(destination, File.GetCreationTimeUtc(source));
        File.SetLastWriteTimeUtc(destination, File.GetLastWriteTimeUtc(source));
        File.SetLastAccessTimeUtc(destination, File.GetLastAccessTimeUtc(source));
        return destination;
    }

    private static bool IsSafeHalo3MovieName(string fileName)
        => Regex.IsMatch(fileName, @"^asq_[a-z0-9]{1,7}_[A-F0-9]{8}_[A-F0-9]{8}\.mov$", RegexOptions.IgnoreCase);

    private static bool IsInDirectory(string path, string directory)
    {
        var fullPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var fullDirectory = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return fullPath.StartsWith(fullDirectory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildSafeHalo3MovieName(string source, SavedFilmHeader header)
    {
        var mapId = string.IsNullOrWhiteSpace(header.MapId)
            ? Path.GetFileNameWithoutExtension(source)
            : header.MapId;
        var mapPrefix = NormalizeAsqMapPrefix(mapId);

        var bytes = File.ReadAllBytes(source);
        var hash = SHA256.HashData(bytes);
        var part1 = BitConverter.ToUInt32(hash, 0).ToString("X8", CultureInfo.InvariantCulture);
        var part2 = BitConverter.ToUInt32(hash, 4).ToString("X8", CultureInfo.InvariantCulture);
        return $"asq_{mapPrefix}_{part1}_{part2}.mov";
    }

    private static string NormalizeAsqMapPrefix(string value)
    {
        var normalized = new string(value
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());

        if (string.IsNullOrWhiteSpace(normalized))
            normalized = "unknown";

        return normalized.Length <= 7 ? normalized : normalized[..7];
    }

    private static string MakeUniquePath(string destination, string source)
    {
        if (!File.Exists(destination))
            return destination;

        try
        {
            var existing = new FileInfo(destination);
            var incoming = new FileInfo(source);
            if (existing.Length == incoming.Length)
                return Path.Combine(
                    Path.GetDirectoryName(destination)!,
                    $"{Path.GetFileNameWithoutExtension(destination)}_{DateTime.Now:HHmmss}{Path.GetExtension(destination)}");
        }
        catch { }

        var dir = Path.GetDirectoryName(destination)!;
        var name = Path.GetFileNameWithoutExtension(destination);
        var ext = Path.GetExtension(destination);
        for (int i = 2; i < 1000; i++)
        {
            var candidate = Path.Combine(dir, $"{name}_{i}{ext}");
            if (!File.Exists(candidate))
                return candidate;
        }

        return Path.Combine(dir, $"{name}_{Guid.NewGuid():N}{ext}");
    }

    private async Task<bool> StageDownpatchDepotsAsync(TheaterClip clip, DownpatchManifestPair version)
    {
        if (_downpatchStagingActive)
        {
            TxtDownpatchReuse.Text = "Downpatch staging is already copying files. Wait for that copy to finish before starting another one.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(version.BaseManifestId) || string.IsNullOrWhiteSpace(version.MultiplayerManifestId))
        {
            TxtDownpatchReuse.Text = "This patch timestamp is known, but the complete base + multiplayer manifest pair is not in the toolbox catalog yet.";
            return false;
        }

        var workspace = TxtDownpatchWorkspace.Text.Trim();
        if (string.IsNullOrWhiteSpace(workspace))
        {
            TxtDownpatchReuse.Text = "Choose a workspace folder before staging depots.";
            return false;
        }

        var steamapps = ResolveSteamAppsFolder(App.LoadMccInstallationPath());
        var baseDepot = FindDepotFolder(steamapps, MccBaseDepotId);
        var multiplayerDepot = FindDepotFolder(steamapps, Halo3MultiplayerDepotId);
        if (!Directory.Exists(baseDepot) || !Directory.Exists(multiplayerDepot))
        {
            TxtDownpatchReuse.Text = "Steam depot cache folders are missing. Run both download_depot commands first.";
            return false;
        }

        var targetFolder = BuildSuggestedDownpatchFolder(workspace, version, NextHalo3Version(version));
        BtnCopyNextDownpatchCommand.IsEnabled = false;
        _downpatchStagingActive = true;
        TxtDownpatchReuse.Text = $"Creating downpatch folder: {targetFolder}";
        try
        {
            await Task.Run(() => Directory.CreateDirectory(targetFolder));

            var progress = new Progress<CopyProgress>(p =>
            {
                TxtDownpatchReuse.Text =
                    $"Copying {p.Label}: {p.FilesCopied:N0}/{p.TotalFiles:N0} files, {FormatBytes(p.BytesCopied)} / {FormatBytes(p.TotalBytes)}";
            });

            TxtDownpatchReuse.Text = $"Copying MCC base depot files into: {targetFolder}";
            await Task.Run(() => CopyDirectoryContents(baseDepot, targetFolder, "MCC base depot", progress));

            TxtDownpatchReuse.Text = $"Copying Halo 3 multiplayer depot files into: {targetFolder}";
            await Task.Run(() => CopyDirectoryContents(multiplayerDepot, targetFolder, "Halo 3 multiplayer depot", progress));

            TxtDownpatchReuse.Text = "Writing Steam app markers and tracking the downpatch folder...";
            await Task.Run(() => WriteSteamAppIdMarkers(targetFolder));

            TrackDownpatchFolder(clip, version, targetFolder);
            TxtDownpatchReuse.Text = $"Staged and tracked reusable folder: {targetFolder}";
            return true;
        }
        catch (Exception ex)
        {
            TxtDownpatchReuse.Text = $"Stage failed: {ex.Message}";
            return false;
        }
        finally
        {
            _downpatchStagingActive = false;
            BtnCopyNextDownpatchCommand.IsEnabled = true;
            RefreshDownpatchUi();
        }
    }

    private void TrackDownpatchFolder(TheaterClip clip, DownpatchManifestPair version, string folder)
    {
        var existing = FindTrackedDownpatch(clip.GameKey, version);
        if (existing is not null)
        {
            existing.FolderPath = folder;
            existing.TrackedAtUtc = DateTime.UtcNow;
        }
        else
        {
            _downpatchIndex.Entries.Add(new DownpatchWorkspaceEntry
            {
                GameKey = clip.GameKey,
                VersionLabel = version.Label,
                EffectiveUtc = version.EffectiveUtc,
                BaseManifestId = version.BaseManifestId,
                MultiplayerManifestId = version.MultiplayerManifestId,
                FolderPath = folder,
                TrackedAtUtc = DateTime.UtcNow
            });
        }

        var marker = new DownpatchFolderMarker
        {
            GameKey = clip.GameKey,
            VersionLabel = version.Label,
            EffectiveUtc = version.EffectiveUtc,
            BaseManifestId = version.BaseManifestId,
            MultiplayerManifestId = version.MultiplayerManifestId,
            CreatedAtUtc = DateTime.UtcNow,
            SourceClip = clip.FileName
        };
        File.WriteAllText(
            Path.Combine(folder, "halo-toolbox-downpatch.json"),
            JsonSerializer.Serialize(marker, new JsonSerializerOptions { WriteIndented = true }));

        SaveDownpatchIndex();
    }

    private static string? ResolveSteamAppsFolder(string mccPath)
    {
        if (string.IsNullOrWhiteSpace(mccPath))
            return null;

        var current = new DirectoryInfo(mccPath);
        while (current is not null)
        {
            if (current.Name.Equals("steamapps", StringComparison.OrdinalIgnoreCase))
                return current.FullName;
            current = current.Parent;
        }

        var parent = Directory.GetParent(mccPath);
        if (parent?.Name.Equals("common", StringComparison.OrdinalIgnoreCase) == true)
            return parent.Parent?.FullName;

        return null;
    }

    private static string FindDepotFolder(string? preferredSteamapps, string depotId)
    {
        foreach (var steamapps in EnumerateSteamappsCandidates(preferredSteamapps))
        {
            var depot = Path.Combine(steamapps, "content", $"app_{HaloMccAppId}", $"depot_{depotId}");
            if (Directory.Exists(depot))
                return depot;
        }

        return preferredSteamapps is null
            ? ""
            : Path.Combine(preferredSteamapps, "content", $"app_{HaloMccAppId}", $"depot_{depotId}");
    }

    private static IEnumerable<string> EnumerateSteamappsCandidates(string? preferredSteamapps)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void Add(string? path)
        {
            if (!string.IsNullOrWhiteSpace(path) && seen.Add(path))
                _ = path;
        }

        Add(preferredSteamapps);
        Add(@"C:\Program Files (x86)\Steam\steamapps");
        Add(@"D:\SteamLibrary\steamapps");
        Add(@"E:\SteamLibrary\steamapps");

        foreach (var path in seen)
            yield return path;
    }

    private static void WriteSteamAppIdMarkers(string mccRoot)
    {
        Directory.CreateDirectory(mccRoot);
        File.WriteAllText(Path.Combine(mccRoot, "steam_appid.txt"), HaloMccAppId);

        var win64 = Path.Combine(mccRoot, "MCC", "Binaries", "Win64");
        Directory.CreateDirectory(win64);
        File.WriteAllText(Path.Combine(win64, "steam_appid.txt"), HaloMccAppId);
    }

    private static void CopyDirectoryContents(string source, string destination, string label, IProgress<CopyProgress>? progress)
    {
        foreach (var directory in Directory.EnumerateDirectories(source, "*", SearchOption.AllDirectories))
        {
            var relative = Path.GetRelativePath(source, directory);
            Directory.CreateDirectory(Path.Combine(destination, relative));
        }

        var files = Directory
            .EnumerateFiles(source, "*", SearchOption.AllDirectories)
            .Select(path => new FileInfo(path))
            .ToList();
        long totalBytes = files.Sum(file => file.Length);
        long copiedBytes = 0;
        int copiedFiles = 0;
        var lastReport = Stopwatch.StartNew();

        progress?.Report(new CopyProgress(label, copiedFiles, files.Count, copiedBytes, totalBytes));

        foreach (var file in files)
        {
            var relative = Path.GetRelativePath(source, file.FullName);
            var target = Path.Combine(destination, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            if (File.Exists(target) && new FileInfo(target).Length == file.Length)
            {
                copiedBytes += file.Length;
                copiedFiles++;
                progress?.Report(new CopyProgress(label, copiedFiles, files.Count, copiedBytes, totalBytes));
                continue;
            }

            if (File.Exists(target))
                File.Delete(target);

            if (TryCreateHardLink(target, file.FullName))
            {
                copiedBytes += file.Length;
                copiedFiles++;
                progress?.Report(new CopyProgress(label, copiedFiles, files.Count, copiedBytes, totalBytes));
                continue;
            }

            CopyFileWithProgress(file.FullName, target, bytes =>
            {
                copiedBytes += bytes;
                if (lastReport.ElapsedMilliseconds >= 500)
                {
                    progress?.Report(new CopyProgress(label, copiedFiles, files.Count, copiedBytes, totalBytes));
                    lastReport.Restart();
                }
            });
            File.SetLastWriteTimeUtc(target, file.LastWriteTimeUtc);
            copiedFiles++;
            progress?.Report(new CopyProgress(label, copiedFiles, files.Count, copiedBytes, totalBytes));
        }
    }

    private static bool TryCreateHardLink(string target, string source)
    {
        try
        {
            return CreateHardLink(target, source, IntPtr.Zero);
        }
        catch
        {
            return false;
        }
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);

    private static void CopyFileWithProgress(string source, string destination, Action<int> onBytesCopied)
    {
        const int BufferSize = 1024 * 1024;
        var buffer = new byte[BufferSize];
        using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize, FileOptions.SequentialScan);
        using var output = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, BufferSize, FileOptions.SequentialScan);

        int read;
        while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            output.Write(buffer, 0, read);
            onBytesCopied(read);
        }
    }

    private static string FormatBytes(long bytes)
    {
        string[] units = ["B", "KB", "MB", "GB", "TB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024 && unit < units.Length - 1)
        {
            value /= 1024;
            unit++;
        }

        return $"{value:N1} {units[unit]}";
    }

    private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
    {
        var visibleClips = GetVisibleClips();
        if (visibleClips.Count == 0)
        {
            UpdateSelectAllButton();
            return;
        }

        bool shouldSelectAll = visibleClips.Any(c => !c.IsSelected);
        foreach (var clip in visibleClips)
            clip.IsSelected = shouldSelectAll;

        UpdateSelectAllButton();
    }

    private void BtnDeleteSelected_Click(object sender, RoutedEventArgs e)
    {
        var selected = _clips.Where(c => c.IsSelected).ToList();
        if (selected.Count == 0)
        {
            UpdateStatus("No clips selected - click rows to select them.");
            UpdateSelectAllButton();
            return;
        }

        ConfirmAndDelete(selected, $"{selected.Count} selected clip(s)");
    }

    private void CboGame_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _filterGame = CboGame.SelectedIndex <= 0
            ? ""
            : GameKeys[CboGame.SelectedIndex - 1];
        _view?.Refresh();
        UpdateSelectAllButton();
    }

    private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        _filterSearch = TxtSearch.Text.Trim();
        _view?.Refresh();
        UpdateSelectAllButton();
    }

    private void CboSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        ApplySort();
    }

    private List<TheaterClip> GetVisibleClips()
        => _view.Cast<TheaterClip>().ToList();

    private void UpdateSelectAllButton()
    {
        if (BtnSelectAll is null) return;

        var visibleClips = GetVisibleClips();
        if (BtnDeleteSelected is not null)
            BtnDeleteSelected.IsEnabled = _clips.Any(c => c.IsSelected);

        if (visibleClips.Count == 0)
        {
            BtnSelectAll.Content = "SELECT ALL";
            BtnSelectAll.IsEnabled = false;
            return;
        }

        BtnSelectAll.IsEnabled = true;
        BtnSelectAll.Content = visibleClips.All(c => c.IsSelected)
            ? "DESELECT ALL"
            : "SELECT ALL";
    }
}

internal readonly record struct SavedFilmDate(DateTime RecordedAt, string Source);

internal readonly record struct SavedFilmHeader(
    SavedFilmDate Date,
    string GameKey,
    string MapId,
    string MapDisplayName,
    string HeaderText);

internal enum DownpatchWorkflowStep
{
    Idle,
    WatchingBase,
    BaseComplete,
    WatchingMultiplayer,
    Staging,
    Complete
}

internal sealed record DownpatchManifestPair(
    string Label,
    DateTime EffectiveUtc,
    string BaseManifestId,
    string MultiplayerManifestId);

internal sealed class DownpatchWorkspaceIndex
{
    public List<DownpatchWorkspaceEntry> Entries { get; set; } = new();
}

internal sealed class DownpatchWorkspaceEntry
{
    public string GameKey { get; set; } = "";
    public string VersionLabel { get; set; } = "";
    public DateTime EffectiveUtc { get; set; }
    public string BaseManifestId { get; set; } = "";
    public string MultiplayerManifestId { get; set; } = "";
    public string FolderPath { get; set; } = "";
    public DateTime TrackedAtUtc { get; set; }
}

internal sealed class DownpatchFolderMarker
{
    public string GameKey { get; set; } = "";
    public string VersionLabel { get; set; } = "";
    public DateTime EffectiveUtc { get; set; }
    public string BaseManifestId { get; set; } = "";
    public string MultiplayerManifestId { get; set; } = "";
    public DateTime CreatedAtUtc { get; set; }
    public string SourceClip { get; set; } = "";
}

internal readonly record struct CopyProgress(
    string Label,
    int FilesCopied,
    int TotalFiles,
    long BytesCopied,
    long TotalBytes);
