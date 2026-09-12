using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Navigation;
using System.Xml.Linq;
using Microsoft.Win32;

namespace HaloToolbox
{
    public partial class MainWindow : Window
    {
        private ObservableCollection<MapEntry> _maps = new();

        // -- Official Halo 3 multiplayer map filenames => display names --
        // Source: https://www.halopedia.org/Map_file  (Halo 3 section)
        // Only multiplayer maps are listed here. Campaign maps start with digits
        // and are filtered out in LoadMaps().
        private static readonly Dictionary<string, string> OfficialMaps = new(StringComparer.OrdinalIgnoreCase)
        {
            // Base game
            ["construct"]   = "Construct",
            ["salvation"]   = "Epitaph",
            ["guardian"]    = "Guardian",
            ["deadlock"]    = "High Ground",
            ["isolation"]   = "Isolation",
            ["zanzibar"]    = "Last Resort",
            ["chill"]       = "Narrows",
            ["shrine"]      = "Sandtrap",
            ["snowbound"]   = "Snowbound",
            ["cyberdyne"]   = "The Pit",
            ["riverworld"]  = "Valhalla",
            // Heroic Map Pack
            ["warehouse"]   = "Foundry",
            ["armory"]      = "Rat's Nest",
            ["bunkerworld"] = "Standoff",
            // Legendary Map Pack
            ["sidewinder"]  = "Avalanche",
            ["lockout"]     = "Blackout",
            ["ghosttown"]   = "Ghost Town",
            // Cold Storage DLC
            ["chillout"]    = "Cold Storage",
            // Mythic Map Pack
            ["descent"]     = "Assembly",
            ["spacecamp"]   = "Orbital",
            ["sandbox"]     = "Sandbox",
            // Mythic II Map Pack
            ["fortress"]    = "Citadel",
            ["docks"]       = "Longshore",
            ["midship"]     = "Heretic",
            // MCC-exclusive (Halo Online / Saber3D origin)
            ["s3d_waterfall"] = "Waterfall",
            ["s3d_edge"]      = "Edge",
            ["s3d_turf"]      = "Icebox",
        };

        // 343 / Saber3D maps for quick-disable button
        private static readonly HashSet<string> Map343Names = new(StringComparer.OrdinalIgnoreCase)
        {
            "s3d_edge", "s3d_waterfall", "s3d_turf"
        };

        // Shared / system files to always skip
        private static readonly HashSet<string> SystemMapNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "shared", "single_player_shared", "mainmenu", "nightmare", "intro"
        };

        private const string RemovedPrefix = "REMOVED_";

        // ── Stats Tab — shared HTTP client ───────────────────────────────────
        internal static readonly HttpClient StatsHttp = new() { Timeout = TimeSpan.FromSeconds(15) };

        // ── Stats Tab — file paths ───────────────────────────────────────────
        private static readonly string StatsWatchPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "Low",
            @"MCC\Temporary");
        private static readonly string StatsSettingsFile = App.StatsGamertagPath;
        private static readonly string StatsCacheFile    = App.StatsCachePath;
        private static readonly string StatsTokenFile    = App.StatsTokenPath;

        // ── Stats Tab — mutable state (always access under _statsLock) ───────
        private readonly object _statsLock = new();
        private string _statsGamertag = "";
        private StatsSessionStats _statsSession = new();
        private readonly ObservableCollection<StatsSessionGameRow> _statsSessionGames = new();
        private readonly ObservableCollection<StatsSessionPlayerHistoryRow> _statsSessionPlayers = new();
        private readonly Dictionary<string, StatsSessionPlayerAggregate> _statsSessionPlayerHistory = new(StringComparer.OrdinalIgnoreCase);
        private List<XElement> _statsLastPlayers = new();
        private string _statsLastFileSig = "";
        private string _statsSpartanToken = "";
        private bool _statsHwTokenExpired = false;
        private bool _statsWaypointUnavailable;
        private DateTimeOffset _statsTokenLastValidatedUtc = DateTimeOffset.MinValue;
        private readonly SemaphoreSlim _statsWaypointMaintenanceLock = new(1, 1);
        private string _statsLastGameTokenCandidate = "";
        private string _statsLastGameTokenGamertag = "";
        private DateTimeOffset _statsLastGameTokenProbeUtc;
        private bool _statsGameTokenCandidateResolved;
        private readonly System.Windows.Threading.DispatcherTimer _statsWaypointTokenTimer;
        private bool _statsCurrentLobbyScanRunning = false;
        private string _statsCurrentLobbyServerText = "";
        private string _statsLastGameServerText = "";
        private string _statsLastGameModeText = "";
        private string _statsLastGamePlayedText = "";
        private List<StatsPlayerRow> _statsCurrentLobbySnapshotRows = new();
        private List<StatsPlayerRow> _statsLastCompletedLobbyRows = new();

        // MCC's unified multiplayer-medal IDs from the carnage report. These
        // are shared metadata IDs, not Halo 3's old sequential 8-16 values.
        internal static readonly StatsMedalDefinition[] StatsMultikillMedals =
        {
            new("Double Kill",     62, "Resources/Medals/double-kill.png"),
            new("Triple Kill",    224, "Resources/Medals/triple-kill.png"),
            new("Overkill",       162, "Resources/Medals/overkill.png"),
            new("Killtacular",    140, "Resources/Medals/killtacular.png"),
            new("Killtrocity",    142, "Resources/Medals/killtrocity.png"),
            new("Killimanjaro",   134, "Resources/Medals/killimanjaro.png"),
            new("Killtastrophe",  141, "Resources/Medals/killtastrophe.png"),
            new("Killpocalypse",  139, "Resources/Medals/killpocalypse.png"),
            new("Killionaire",    137, "Resources/Medals/killionaire.png"),
        };

        // ── Stats Tab — lookup caches ────────────────────────────────────────
        private readonly Dictionary<string, string> _statsKd =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statsTotals =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statsGames =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statsRecentKd =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statsGamertagsByXuid =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, MatchmakingPlayerPing> _statsMatchmakingPings =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, StatsCachedPlayer> _statsPersistentCache =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Queue<string> _statsCacheOrder = new();

        // ── Stats Tab — UI collection ────────────────────────────────────────
        private readonly ObservableCollection<StatsPlayerRow> _statsCurrentLobbyRows = new();
        private readonly ObservableCollection<StatsPlayerRow> _statsLobbyRows = new();
        private readonly ObservableCollection<MatchmakingPopulationRow> _statsPopulationRows = new();
        private readonly List<MatchmakingPopulationSample> _statsPopulationHistory = new();
        private string _selectedPopulationHopperName = "";
        private int _populationHistoryMinutes = 60;
        private readonly System.Windows.Threading.DispatcherTimer _populationHistoryTimer = new()
        {
            Interval = TimeSpan.FromMinutes(1)
        };
        private readonly ObservableCollection<FirewallRuleRow> _firewallRuleRows = new();
        private bool _firewallRuleTableLoaded;
        private bool _firewallRuleTableRefreshing;
        private readonly SemaphoreSlim _statsPopulationRefreshLock = new(1, 1);
        private string _statsPopulationSortProperty = nameof(MatchmakingPopulationRow.Population);
        private ListSortDirection _statsPopulationSortDirection = ListSortDirection.Descending;
        private readonly List<string> _sessionLogLines = new();
        private readonly object _pendingLogLock = new();
        private readonly Queue<string> _pendingLogLines = new();
        private bool _logFlushScheduled;
        private const int MaxSessionLogLines = 2000;
        private readonly ProxyService _rejoinProxy = new();
        private readonly NetworkStatsMonitor _networkStatsMonitor = new();
        private readonly GameServerConnectionMonitor _gameServerConnectionMonitor = new();
        private readonly ObsOverlayServer _obsOverlayServer = new();
        private GameNetworkStatsOverlayWindow? _gameNetworkStatsOverlay;
        private GameNetworkStatsOverlayWindow? _matchmakingWaitOverlay;
        private GameNetworkStatsOverlayWindow? _sessionStatsOverlay;
        private GameNetworkStatsOverlayWindow? _combinedNetworkSessionOverlay;
        private OverlayRepositionHelpWindow? _overlayRepositionHelpWindow;
        private bool _networkStatsOverlayEnabled = true;
        private bool _matchmakingWaitOverlayEnabled = true;
        private bool _combinedNetworkSessionOverlayEnabled;
        private SmartMatchWaitEstimate? _smartMatchWaitEstimate;
        private int? _smartMatchHopperPopulation;
        private string _smartMatchHopperDisplayName = "";
        private readonly System.Windows.Threading.DispatcherTimer _matchmakingPopulationTimer;
        private DateTimeOffset _lastFullPopulationRefreshUtc = DateTimeOffset.MinValue;
        private bool _networkStatsOverlayMoveEnabled;
        private bool _obsBrowserOverlayEnabled;
        private bool _obsBrowserOverlaySessionStatsEnabled = true;
        private bool _networkStatsObsOnly;
        private bool _matchmakingWaitObsOnly;
        private bool _sessionStatsObsOnly;
        private bool _combinedNetworkSessionObsOnly;
        private NetworkStatsSnapshot? _lastNetworkStatsSnapshot;
        private NetworkTrafficSnapshot? _lastNetworkTrafficSnapshot;
        private ObsPostGameRecap? _postGameRecap;
        private Task? _supportSessionCheckTask;
        private Microsoft.Web.WebView2.Wpf.WebView2? _hiddenCookieChecker;
        private Mods? _modsTab;
        private Theater? _theaterTab;
        private Playlists? _playlistsTab;
        private Vpn? _vpnTab;
        private BanChecker? _banCheckerTab;
        private MatchHistory? _matchHistoryTab;
        private string _playlistsMccPath = App.DefaultMccPath;
        private bool _firstRenderInitializationQueued;
        private bool _statsInitialized;
        private Rect _lastOverlayRelativePlacement = new(0, 0, 1280.0 / 1920.0, 170.0 / 1080.0);
        private readonly Dictionary<string, Rect> _componentOverlayRelativePlacements = new(StringComparer.OrdinalIgnoreCase);
        private GameServerInfo? _lastNetworkStatsRelayServer;
        private GameServerInfo? _trustedDedicatedServer;
        private bool _mainWindowInitialized;
        private bool _rejoinWinHttpManualNeeded;
        private readonly object _rejoinCrashWatchLock = new();
        private readonly Dictionary<int, Process> _rejoinWatchedMccProcesses = new();
        private readonly System.Windows.Threading.DispatcherTimer _rejoinCrashWatchTimer;
        private SteamFirewallState _steamFirewallUiState = SteamFirewallState.Missing;
        private readonly SemaphoreSlim _steamFirewallAutoLock = new(1, 1);
        private readonly System.Windows.Threading.DispatcherTimer _steamFirewallAutoTimer;
        private bool _steamFirewallAutoEnabled;
        private bool _steamFirewallAutoPaused;
        private bool _steamFirewallAutoHeldForActiveMatch;
        private bool _steamFirewallAutoSuspendedForCrashRestore;
        private bool _rejoinFirewallCheckChanging;
        private bool _steamFirewallRulesPrepared;
        private bool _rejoinCampaignFirewallApplying;
        private bool _rejoinCampaignFirewallEnabled;
        private bool _firewallFixActionPending;
        private bool _closeFirewallCleanupStarted;
        private TabItem? _lastMainTab;
        private bool _restoringMainTabSelection;
        private ToolsPage _toolsPage = ToolsPage.Home;
        private string _homeMccStatusPath = "";
        private bool _homeMccStatusFound;
        private static readonly SemaphoreSlim SteamFirewallCommandLock = new(1, 1);
        private DateTime _steamFirewallAutoResumeAfterUtc = DateTime.MinValue;
        private const int SteamFirewallAutoSearchHoldSeconds = 180;
        private const int SteamFirewallAutoMatchFoundHoldSeconds = 5;
        private static readonly bool SteamFirewallFeatureEnabled = false;
        private const string RejoinFirewallCampaignLabel = "CAMPAIGN";
        private const string RejoinFirewallMatchmakingLabel = "MATCHMAKING AUTO";
        private const string RejoinFirewallDisabledSuffix = " (Disabled until Rejoin Fix is Enabled)";

        private const long MaxDiagnosticExportBytes = 25L * 1024 * 1024;
        private const string ToolboxRegistryPath = @"Software\HaloMCCToolbox";
        private const string RejoinFixProxyAddress = ProxyService.DefaultProxyAddress;
        private const string RejoinFixProxyCertificatePassword = "halointel-proxy";
        private static readonly int[] SteamFirewallPorts = { 3478, 4379 };
        private static readonly int[] RejoinCampaignFirewallPorts = { 3478 };
        private const string SteamFirewallRulePrefix = "Halo Toolbox - Block MCC P2P Port";
        private const string GlobalSteamFirewallRulePrefix = "Halo Toolbox - Block Steam P2P Port";
        private const string LegacyPort4379FirewallRulePrefix = "Halo Toolbox - Block Port 4379";
        private static readonly string ToolboxLocalAppDataRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HaloMCCToolbox");
        private static readonly string ToolboxRoamingAppDataRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HaloMCCToolbox");
        private static readonly string[] SteamFirewallRuleNames = SteamFirewallPorts
            .SelectMany(port => new[]
            {
                $"{SteamFirewallRulePrefix} {port} TCP Inbound",
                $"{SteamFirewallRulePrefix} {port} UDP Inbound",
                $"{SteamFirewallRulePrefix} {port} TCP Outbound",
                $"{SteamFirewallRulePrefix} {port} UDP Outbound"
            })
            .ToArray();
        private static readonly string[] LegacySteamFirewallRuleNames =
        {
            $"{LegacyPort4379FirewallRulePrefix} TCP Inbound",
            $"{LegacyPort4379FirewallRulePrefix} UDP Inbound",
            $"{LegacyPort4379FirewallRulePrefix} TCP Outbound",
            $"{LegacyPort4379FirewallRulePrefix} UDP Outbound"
        };
        private static readonly string[] GlobalSteamFirewallRuleNames = SteamFirewallPorts
            .SelectMany(port => new[]
            {
                $"{GlobalSteamFirewallRulePrefix} {port} TCP Inbound",
                $"{GlobalSteamFirewallRulePrefix} {port} UDP Inbound",
                $"{GlobalSteamFirewallRulePrefix} {port} TCP Outbound",
                $"{GlobalSteamFirewallRulePrefix} {port} UDP Outbound"
            })
            .ToArray();
        private static readonly string SteamFirewallStateFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HaloMCCToolbox",
            "steam-firewall-state.txt");

        private enum SteamFirewallState
        {
            Unknown,
            Missing,
            Disabled,
            Enabled,
            Partial
        }

        private enum MapToggleResult
        {
            Unchanged,
            Changed,
            Failed
        }

        private static readonly string[] DiagnosticExtensions =
        {
            ".log", ".txt", ".xml", ".json", ".dmp", ".runtime-xml", ".ue4stats"
        };

        // ── Per-game multiplayer map lists (Report tab) ──────────────────────
        private static readonly Dictionary<string, List<string>> GameMaps =
            new(StringComparer.OrdinalIgnoreCase)
        {
            ["Halo CE"] = new List<string>
            {
                "Battle Creek", "Blood Gulch", "Boarding Action", "Chill Out",
                "Chiron TL-34", "Danger Canyon", "Damnation", "Death Island",
                "Derelict", "Gephyrophobia", "Hang 'Em High", "Ice Fields",
                "Infinity", "Longest", "Prisoner", "Rat Race", "Sidewinder",
                "Timberland", "Wizard"
            },
            ["Halo 2"] = new List<string>
            {
                "Ascension", "Backwash", "Beaver Creek", "Burial Mounds",
                "Coagulation", "Colossus", "Containment", "Desolation", "Elongation",
                "Foundation", "Gemini", "Headlong", "Ivory Tower", "Lockout",
                "Midship", "Relic", "Sanctuary", "Terminal", "Tombstone", "Turf",
                "Uplift", "Warlock", "Waterworks", "Zanzibar"
            },
            ["Halo 2 Anniversary"] = new List<string>
            {
                "Ascension", "Backwash", "Beaver Creek", "Burial Mounds",
                "Coagulation", "Colossus", "Containment", "Desolation", "District",
                "Elongation", "Foundation", "Gemini", "Headlong", "Ivory Tower",
                "Lockout", "Midship", "Relic", "Sanctuary", "Terminal", "Tombstone",
                "Turf", "Uplift", "Warlock", "Waterworks", "Zanzibar"
            },
            ["Halo 3"] = new List<string>
            {
                "Assembly", "Avalanche", "Blackout", "Citadel", "Cold Storage",
                "Construct", "Edge", "Epitaph", "Foundry", "Ghost Town", "Guardian",
                "Heretic", "High Ground", "Icebox", "Isolation", "Last Resort",
                "Longshore", "Narrows", "Orbital", "Rat's Nest", "Sandbox",
                "Sandtrap", "Snowbound", "Standoff", "The Pit", "Valhalla", "Waterfall"
            },
            ["Halo Reach"] = new List<string>
            {
                "Anchor 9", "Battle Canyon", "Boardwalk", "Boneyard", "Breakneck",
                "Breakpoint", "Condemned", "Countdown", "Forge World", "Hemorrhage",
                "High Noon", "Highlands", "Powerhouse", "Reflection", "Ridgeline",
                "Solitary", "Spire", "Sword Base", "Tempest", "Unearthed", "Zealot"
            },
            ["Halo 4"] = new List<string>
            {
                "Abandon", "Adrift", "Complex", "Daybreak", "Erosion", "Exile",
                "Haven", "Harvest", "Impact", "Landfall", "Longbow", "Meltdown",
                "Monolith", "Perdition", "Pitfall", "Ragnarok", "Ravine", "Relay",
                "Shatter", "Shutdown", "Skyline", "Solace", "Vertigo", "Vortex",
                "Wreckage"
            },
        };

        // Games that support Film/Theater recording in MCC
        private static readonly HashSet<string> GamesWithTheater =
            new(StringComparer.OrdinalIgnoreCase) { "Halo 3", "Halo Reach", "Halo 4" };

        public MainWindow()
        {
            InitializeComponent();
            _matchmakingPopulationTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(10)
            };
            _matchmakingPopulationTimer.Tick += MatchmakingPopulationTimer_Tick;
            _populationHistoryTimer.Tick += async (_, _) => await StatsRefreshMatchmakingPopulationAsync();
            _statsWaypointTokenTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(1)
            };
            _statsWaypointTokenTimer.Tick += StatsWaypointTokenTimer_Tick;
            RestoreMainWindowPlacement();
            LoadSectionVisibility();
            _lastMainTab = MainTabs.SelectedItem as TabItem;
            MainTabs.SelectionChanged += MainTabs_SelectionChanged;
            if (App.LoadSetupPreference("OpenLastSection", true))
            {
                string lastSection = App.LoadLastMainSection();
                var savedTab = MainTabs.Items
                    .OfType<TabItem>()
                    .FirstOrDefault(tab => tab.Name == lastSection && tab.Visibility == Visibility.Visible);
                if (savedTab is not null)
                {
                    MainTabs.SelectedItem = savedTab;
                    _lastMainTab = savedTab;
                }
            }
            _networkStatsOverlayEnabled = App.LoadGameNetworkStatsOverlayEnabled();
            ChkNetworkStatsOverlay.IsChecked = _networkStatsOverlayEnabled;
            _matchmakingWaitOverlayEnabled = App.LoadMatchmakingWaitOverlayEnabled();
            ChkMatchmakingWaitOverlay.IsChecked = _matchmakingWaitOverlayEnabled;
            _combinedNetworkSessionOverlayEnabled = App.LoadCombinedNetworkSessionOverlayEnabled();
            CombinedNetworkSessionOverlayToggle.IsChecked = _combinedNetworkSessionOverlayEnabled;
            string savedRejoinFirewallMode = App.LoadRejoinFirewallMode();
            bool hadRejoinFirewallEnabled =
                string.Equals(savedRejoinFirewallMode, "Campaign", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(savedRejoinFirewallMode, "Matchmaking", StringComparison.OrdinalIgnoreCase);
            SetRejoinFirewallCheckbox(
                ChkRejoinFixFirewall,
                string.Equals(savedRejoinFirewallMode, "Campaign", StringComparison.OrdinalIgnoreCase));
            SetRejoinFirewallCheckbox(
                ChkRejoinFixFirewallMatchmaking,
                string.Equals(savedRejoinFirewallMode, "Matchmaking", StringComparison.OrdinalIgnoreCase));
            TxtMccPath.Text = App.LoadMccInstallationPath();
            _playlistsMccPath = TxtMccPath.Text;
            TxtMccPath.TextChanged += TxtMccPath_TextChanged;
            UpdateMccEditionUi();
            MapList.ItemsSource = _maps;
            FirewallRuleTable.ItemsSource = _firewallRuleRows;
            ShowToolsPage(ToolsPage.Home, selectToolsSection: false);
            SyncSidebarSelection(MainTabs.SelectedItem as TabItem ?? ToolsSection);
            AppendLog("[INFO]", "Halo MCC Toolbox started. Made by The FFA Panda.", "#00C8FF");
            ContentRendered += MainWindow_ContentRendered;
            _obsBrowserOverlayEnabled = App.LoadObsBrowserOverlayEnabled();
            _obsBrowserOverlaySessionStatsEnabled = App.LoadObsBrowserOverlaySessionStatsEnabled();
            _networkStatsObsOnly = App.LoadNetworkStatsObsOnlyEnabled();
            _matchmakingWaitObsOnly = App.LoadMatchmakingWaitObsOnlyEnabled();
            _sessionStatsObsOnly = App.LoadSessionStatsObsOnlyEnabled();
            _combinedNetworkSessionObsOnly = App.LoadCombinedNetworkSessionObsOnlyEnabled();
            StatsObsOverlayToggle.IsChecked = _obsBrowserOverlayEnabled;
            StatsObsSessionStatsToggle.IsChecked = _obsBrowserOverlaySessionStatsEnabled;
            NetworkOverlayStyleCombo.SelectedIndex =
                App.LoadGameOverlayVisualStyle("network") == GameOverlayVisualStyle.Modern ? 1 : 0;
            SessionOverlayStyleCombo.SelectedIndex =
                App.LoadGameOverlayVisualStyle("session") == GameOverlayVisualStyle.Modern ? 1 : 0;
            MatchmakingWaitOverlayStyleCombo.SelectedIndex =
                App.LoadGameOverlayVisualStyle("wait") == GameOverlayVisualStyle.Modern ? 1 : 0;
            NetworkStatsObsOnlyToggle.IsChecked = _networkStatsObsOnly;
            MatchmakingWaitObsOnlyToggle.IsChecked = _matchmakingWaitObsOnly;
            SessionStatsObsOnlyToggle.IsChecked = _sessionStatsObsOnly;
            CombinedNetworkSessionObsOnlyToggle.IsChecked = _combinedNetworkSessionObsOnly;
            StatsRefreshObsOverlayUi();
            _rejoinProxy.WinHttpManualSetRequired += (_, command) =>
                Dispatcher.InvokeAsync(() =>
                {
                    _rejoinWinHttpManualNeeded = true;
                    AppendLog("[REJOIN]", $"Proxy active, but MCC capture may need admin approval. Manual fallback: {command}", "#FF6A00");
                    UpdateRejoinFixUi();
                });
            _rejoinProxy.OnMatchSessionSaved += (_, _) =>
                Dispatcher.InvokeAsync(() =>
                {
                    _smartMatchWaitEstimate = null;
                    _smartMatchHopperPopulation = null;
                    _matchmakingPopulationTimer.Stop();
                    StatsMatchmakingWaitLabel.Text = "";
                    PublishObsOverlaySnapshot();
                    AppendLog("[REJOIN]", "Captured matchmaking session and saved it to Toolbox appdata.", "#00C8FF");
                    UpdateRejoinFixUi();
                    HoldSteamFirewallPausedForActiveMatch("match session captured");
                });
            _rejoinProxy.OnPlayerIdentityChanged += (_, _) =>
                Dispatcher.InvokeAsync(UpdateRejoinFixUi);
            _rejoinProxy.OnRejoinContextChanged += (_, _) =>
                Dispatcher.InvokeAsync(() =>
                {
                    UpdateRejoinFixUi();
                    if (_rejoinProxy.CurrentSquadMemberCount > 0)
                        HoldSteamFirewallPausedForActiveMatch("squad session active");
                });
            _rejoinProxy.OnCrashRestorePendingChanged += (_, pending) =>
                Dispatcher.InvokeAsync(() => HandleCrashRestoreFirewallStateChangedAsync(pending));
            _rejoinProxy.OnGameServerChanged += (_, serverInfo) =>
                Dispatcher.InvokeAsync(() => HandleTrustedGameServerChanged(serverInfo));
            _lobbyIdentityTimer.Tick += (_, _) => _ = StatsResolveLobbyIdentitiesAsync();
            _lobbyIdentityTimer.Start();
            Closed += (_, _) => _lobbyIdentityTimer.Stop();
            _rejoinProxy.OnMatchmakingPlayerPingsObserved += (_, pings) =>
            {
                Dispatcher.InvokeAsync(() =>
                {
                    lock (_statsLock)
                    {
                        _statsMatchmakingPings.Clear();
                        foreach (var ping in pings)
                        {
                            string normalizedXuid = StatsNormalizeXuid(ping.Xuid);
                            if (!string.IsNullOrWhiteSpace(normalizedXuid))
                            {
                                _statsMatchmakingPings[normalizedXuid] = ping;
                                StatsRememberGamertagForXuid(normalizedXuid, ping.Gamertag);
                            }
                        }
                    }

                    StatsRebuildCurrentLobbyRows();
                    StatsRebuildLobbyRows();
                    _ = StatsResolveLobbyIdentitiesAsync();
                    if (_rejoinProxy.IsRunning && pings.Count > 0)
                        _ = StatsFetchCurrentLobbyStats();
                });
            };
            _rejoinProxy.OnSmartMatchWaitEstimateChanged += (_, estimate) =>
                Dispatcher.InvokeAsync(() =>
                {
                    _smartMatchWaitEstimate = estimate;
                    _smartMatchHopperPopulation = null;
                    _smartMatchHopperDisplayName = EnsurePlaylistsTab().GetMatchmakingHoppers()
                        .FirstOrDefault(x => x.HopperName.Equals(
                            estimate.HopperName,
                            StringComparison.OrdinalIgnoreCase))?.DisplayName
                        ?? estimate.HopperName;
                    _matchmakingPopulationTimer.Start();
                    StatsMatchmakingWaitLabel.Text = estimate.WaitSeconds < 60
                        ? $"EST. WAIT ~{estimate.WaitSeconds} SEC"
                        : $"EST. WAIT ~{Math.Ceiling(estimate.WaitSeconds / 60.0):0} MIN · MAY TAKE A WHILE";
                    if (_matchmakingWaitOverlayEnabled && _rejoinProxy.IsRunning)
                    {
                        EnsureGameNetworkStatsOverlay();
                        _gameNetworkStatsOverlay?.SetPreferredProcessId(TryGetMccProcessId());
                        _gameNetworkStatsOverlay?.SetMoveMode(_networkStatsOverlayMoveEnabled);
                    }
                    PublishObsOverlaySnapshot();
                    AppendLog("[MATCH]", $"SmartMatch estimated wait: ~{estimate.WaitSeconds} seconds.", "#00C8FF");
                    _ = StatsRefreshMatchmakingPopulationAsync();
                });
            _rejoinProxy.OnSmartMatchWaitCancelled += (_, _) =>
                Dispatcher.InvokeAsync(() =>
                {
                    _smartMatchWaitEstimate = null;
                    _smartMatchHopperPopulation = null;
                    _smartMatchHopperDisplayName = "";
                    _matchmakingPopulationTimer.Stop();
                    StatsMatchmakingWaitLabel.Text = "";
                    PublishObsOverlaySnapshot();
                    AppendLog("[MATCH]", "Matchmaking ticket cancelled; wait estimate cleared.", "#C8D8E8");
                });
            _rejoinProxy.OnRequestCaptured += (_, entry) =>
                Dispatcher.InvokeAsync(() =>
                {
                    HandleSteamFirewallAutoSignal(entry);
                });
            _rejoinProxy.BanSpartanTokenChanged += (_, _) =>
                Dispatcher.InvokeAsync(() =>
                {
                    _banCheckerTab?.RefreshAuthorizationStatus();
                    _ = StatsMaintainWaypointTokenAsync();
                });
            _networkStatsMonitor.StatsUpdated += (_, snapshot) =>
                Dispatcher.InvokeAsync(() => UpdateNetworkStatsOverlay(snapshot));
            _gameServerConnectionMonitor.ActiveServerChanged += (_, serverInfo) =>
                Dispatcher.InvokeAsync(() => HandleNetworkStatsObservedServer(serverInfo));
            _gameServerConnectionMonitor.TrafficStatsUpdated += (_, snapshot) =>
                Dispatcher.InvokeAsync(() => UpdateNetworkTrafficOverlay(snapshot));
            _gameServerConnectionMonitor.StatusChanged += (_, status) =>
                Dispatcher.InvokeAsync(() => AppendLog("[NET]", status, "#4A5A6A"));
            VpnConnectionPresence.Changed += VpnConnectionPresence_Changed;
            Closed += (_, _) =>
            {
                _populationHistoryTimer.Stop();
                CloseOverlayRepositionHelp();
                VpnConnectionPresence.Changed -= VpnConnectionPresence_Changed;
                DisposeHiddenCookieChecker();
                _modsTab?.Dispose();
                _vpnTab?.Dispose();
                StopRejoinCrashWatcher();
                _gameServerConnectionMonitor.Dispose();
                _networkStatsMonitor.Dispose();
                _obsOverlayServer.Dispose();
                _rejoinProxy.Dispose();
                _matchmakingPopulationTimer.Stop();
                _statsWaypointTokenTimer.Stop();
            };
            Closing += MainWindow_Closing;
            StateChanged += (_, _) => UpdateMaximizeButton();
            UpdateRejoinFixUi();
            _rejoinCrashWatchTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _rejoinCrashWatchTimer.Tick += (_, _) =>
            {
                PollMccProcessesForRejoinCrashRestore();
            };
            _steamFirewallAutoTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(5)
            };
            _steamFirewallAutoTimer.Tick += async (_, _) => await SteamFirewallAutoTimer_TickAsync();
            InitializeSteamFirewallFeatureState();
            _mainWindowInitialized = true;
            if (!IsMicrosoftStoreInstallation && hadRejoinFirewallEnabled)
                Dispatcher.InvokeAsync(SynchronizeStartupFirewallStateAsync);
            Dispatcher.InvokeAsync(StartPendingRejoinFixAfterElevationAsync);

        }

        private void MainWindow_ContentRendered(object? sender, EventArgs e)
        {
            if (_firstRenderInitializationQueued)
                return;

            _firstRenderInitializationQueued = true;
            ContentRendered -= MainWindow_ContentRendered;

            // Input and rendering have higher dispatcher priority than these callbacks,
            // so the first window can be dragged immediately while startup work is
            // divided into small, independently scheduled pieces.
            Dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new Action(InitializeDeferredUiState));
            Dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new Action(StartDeferredBackgroundWork));
        }

        private void InitializeDeferredUiState()
        {
            _populationHistoryTimer.Start();
            ThemeToggleBtn.Content = App.IsDarkTheme ? "☾" : "☀";
            EnsureStatsInitialized();
        }

        private void StartDeferredBackgroundWork()
        {
            string mccPath = TxtMccPath.Text.Trim();
            var defaultMapsPath = Path.Combine(mccPath, "halo3", "maps");
            if (Directory.Exists(defaultMapsPath))
            {
                AppendLog("[INFO]", "Loading maps in background...", "#4A5A6A");
                _ = Task.Run(() => LoadMaps(mccPath));
            }

            _ = Task.Run(StatsMonitorLoop);
            _ = RefreshPopulationAfterFirstInputAsync();
        }

        private async Task RefreshPopulationAfterFirstInputAsync()
        {
            // Keep the initial dispatcher free long enough for the first drag/click.
            await Task.Delay(1000);
            if (Dispatcher.HasShutdownStarted || Dispatcher.HasShutdownFinished)
                return;

            await StatsRefreshMatchmakingPopulationAsync();
        }

        private void EnsureStatsInitialized()
        {
            if (_statsInitialized)
                return;

            _statsInitialized = true;
            StatsInitialize();
            StatsPopulationList.ItemsSource = _statsPopulationRows;
        }

        private Mods EnsureModsTab()
        {
            if (_modsTab is not null)
                return _modsTab;

            _modsTab = new Mods();
            ModsHost.Content = _modsTab;
            return _modsTab;
        }

        private Theater EnsureTheaterTab()
        {
            if (_theaterTab is not null)
                return _theaterTab;

            _theaterTab = new Theater();
            _theaterTab.SetMccInstallationPath(TxtMccPath.Text.Trim());
            TheaterHost.Content = _theaterTab;
            return _theaterTab;
        }

        private Playlists EnsurePlaylistsTab()
        {
            if (_playlistsTab is not null)
                return _playlistsTab;

            _playlistsTab = new Playlists();
            _playlistsTab.SetMccInstallationPath(_playlistsMccPath);
            PlaylistsHost.Content = _playlistsTab;
            return _playlistsTab;
        }

        private enum ToolsPage
        {
            Home,
            Maps,
            Fixes,
            Overlays,
            Firewall
        }

        private void SidebarNavigation_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not RadioButton { Tag: string target })
                return;

            if (TryParseToolsPage(target, out var toolsPage))
            {
                ShowToolsPage(toolsPage);
                return;
            }

            TabItem? section = target switch
            {
                "H3Mods" => H3ModsSection,
                "Playlists" => PlaylistsSection,
                "Vpn" => VpnSection,
                "Stats" => StatsSection,
                "Population" => PopulationSection,
                "MatchHistory" => MatchHistorySection,
                "BanChecker" => BanCheckerSection,
                "Theater" => TheaterSection,
                "Downpatch" => DownpatchSection,
                "Report" => ReportSection,
                "Log" => LogSection,
                "About" => AboutSection,
                _ => null
            };

            if (section is null || section.Visibility != Visibility.Visible)
            {
                SyncSidebarSelection(MainTabs.SelectedItem as TabItem ?? ToolsSection);
                return;
            }

            MainTabs.SelectedItem = section;
            SyncSidebarSelection(MainTabs.SelectedItem as TabItem ?? ToolsSection);
        }

        private void ToolsHomeTask_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button { Tag: string target })
                return;

            if (TryParseToolsPage(target, out var toolsPage))
            {
                ShowToolsPage(toolsPage);
                return;
            }

            TabItem? section = target switch
            {
                "Vpn" => VpnSection,
                "BanChecker" => BanCheckerSection,
                _ => null
            };
            if (section is not null && section.Visibility == Visibility.Visible)
            {
                MainTabs.SelectedItem = section;
                SyncSidebarSelection(section);
            }
        }

        private static bool TryParseToolsPage(string value, out ToolsPage page)
        {
            page = value switch
            {
                "Home" => ToolsPage.Home,
                "Maps" => ToolsPage.Maps,
                "Fixes" => ToolsPage.Fixes,
                "Overlays" => ToolsPage.Overlays,
                "Firewall" => ToolsPage.Firewall,
                _ => ToolsPage.Home
            };
            return value is "Home" or "Maps" or "Fixes" or "Overlays" or "Firewall";
        }

        private void ShowToolsPage(ToolsPage page, bool selectToolsSection = true)
        {
            _toolsPage = page;

            ToolsHomePanel.Visibility = page == ToolsPage.Home ? Visibility.Visible : Visibility.Collapsed;

            bool showFixes = page == ToolsPage.Fixes;
            FixRepairSectionLabel.Visibility = showFixes ? Visibility.Visible : Visibility.Collapsed;
            CleanCredentialsCard.Visibility = showFixes ? Visibility.Visible : Visibility.Collapsed;
            EacRepairCard.Visibility = showFixes ? Visibility.Visible : Visibility.Collapsed;
            AudioRepairCard.Visibility = showFixes ? Visibility.Visible : Visibility.Collapsed;

            bool showOverlays = page == ToolsPage.Overlays;
            bool showFirewall = page == ToolsPage.Firewall;
            AdvancedFeaturesCard.Visibility = showOverlays || showFirewall
                ? Visibility.Visible
                : Visibility.Collapsed;
            AdvancedFeaturesHeaderText.Text = showFirewall ? "FIREWALL FIX" : "OVERLAYS";
            OverlaysControlsPanel.Visibility = showOverlays ? Visibility.Visible : Visibility.Collapsed;
            OverlayPreviewsCard.Visibility = showOverlays ? Visibility.Visible : Visibility.Collapsed;
            if (showOverlays)
                UpdateOverlayPreviews();
            FirewallControlsPanel.Visibility = showFirewall ? Visibility.Visible : Visibility.Collapsed;
            SteamFirewallCard.Visibility = showFirewall && SteamFirewallFeatureEnabled
                ? Visibility.Visible
                : Visibility.Collapsed;
            if (showFirewall)
            {
                UpdateFirewallPageStatus();
                if (!_firewallRuleTableLoaded && !_firewallRuleTableRefreshing)
                    _ = RefreshFirewallRuleTableAsync(logResults: false);
            }

            bool showMaps = page == ToolsPage.Maps;
            MccInstallationPathCard.Visibility = showMaps ? Visibility.Visible : Visibility.Collapsed;
            MapSelectorPanel.Visibility = showMaps ? Visibility.Visible : Visibility.Collapsed;

            if (selectToolsSection && !ReferenceEquals(MainTabs.SelectedItem, ToolsSection))
                MainTabs.SelectedItem = ToolsSection;

            SyncSidebarSelection(ToolsSection);
            if (page == ToolsPage.Home)
                UpdateToolboxStatus();
        }

        private void SyncSidebarSelection(TabItem selectedTab)
        {
            RadioButton? selectedButton = ReferenceEquals(selectedTab, ToolsSection)
                ? _toolsPage switch
                {
                    ToolsPage.Home => SidebarHomeButton,
                    ToolsPage.Maps => SidebarMapsButton,
                    ToolsPage.Fixes => SidebarFixesButton,
                    ToolsPage.Overlays => SidebarOverlaysButton,
                    ToolsPage.Firewall => SidebarFirewallButton,
                    _ => SidebarHomeButton
                }
                : ReferenceEquals(selectedTab, H3ModsSection) ? SidebarModsButton
                : ReferenceEquals(selectedTab, PlaylistsSection) ? SidebarPlaylistsButton
                : ReferenceEquals(selectedTab, VpnSection) ? SidebarVpnButton
                : ReferenceEquals(selectedTab, StatsSection) ? SidebarStatsButton
                : ReferenceEquals(selectedTab, PopulationSection) ? SidebarPopulationButton
                : ReferenceEquals(selectedTab, MatchHistorySection) ? SidebarMatchHistoryButton
                : ReferenceEquals(selectedTab, BanCheckerSection) ? SidebarBanCheckerButton
                : ReferenceEquals(selectedTab, DownpatchSection) ? SidebarDownpatchButton
                : ReferenceEquals(selectedTab, TheaterSection) ? SidebarTheaterButton
                : ReferenceEquals(selectedTab, ReportSection) ? SidebarReportButton
                : ReferenceEquals(selectedTab, LogSection) ? SidebarLogButton
                : ReferenceEquals(selectedTab, AboutSection) ? SidebarAboutButton
                : null;

            if (selectedButton is not null)
                selectedButton.IsChecked = true;
        }

        private void UpdateToolboxStatus()
        {
            if (HomeMccStatus is null)
                return;

            string mccPath = TxtMccPath.Text.Trim();
            if (!string.Equals(_homeMccStatusPath, mccPath, StringComparison.OrdinalIgnoreCase))
            {
                _homeMccStatusPath = mccPath;
                _homeMccStatusFound = Directory.Exists(Path.Combine(mccPath, "halo3", "maps"));
            }
            bool mccFound = _homeMccStatusFound;
            SetToolboxStatusText(
                HomeMccStatus,
                mccFound ? App.GetMccInstallationLabel(mccPath) : "NOT FOUND",
                mccFound ? "#39FF14" : "#FF6A00");

            bool vpnConnected = !string.IsNullOrWhiteSpace(VpnConnectionPresence.ConnectedRegion);
            SetToolboxStatusText(
                HomeVpnStatus,
                vpnConnected
                    ? $"CONNECTED · {VpnConnectionPresence.ConnectedRegion.ToUpperInvariant()}"
                    : "NOT CONNECTED",
                vpnConnected ? "#39FF14" : "#71869A");

            SetToolboxStatusText(
                HomeFeaturesStatus,
                _rejoinProxy.IsRunning ? "● ON" : "● OFF",
                _rejoinProxy.IsRunning ? "#39FF14" : "#71869A");
        }

        private bool IsMicrosoftStoreInstallation =>
            App.GetMccInstallationKind(TxtMccPath?.Text) == MccInstallationKind.MicrosoftStore;

        private void UpdateMccEditionUi()
        {
            if (TxtMccEditionBadge is null || TxtMccPath is null)
                return;

            var kind = App.GetMccInstallationKind(TxtMccPath.Text);
            string label = kind switch
            {
                MccInstallationKind.Steam => "STEAM",
                MccInstallationKind.MicrosoftStore => "MICROSOFT STORE",
                _ => "NOT DETECTED"
            };

            TxtMccEditionBadge.Text = $"MCC: {label}";
            TxtMccEditionBadge.Foreground = Brush(kind == MccInstallationKind.Unknown ? "#FF6A00" : "#39FF14");
            _theaterTab?.SetMccInstallationPath(TxtMccPath.Text.Trim());

            bool steamOnlyFeaturesAvailable = kind == MccInstallationKind.Steam;
            ChkRejoinFixFirewall.Visibility = steamOnlyFeaturesAvailable ? Visibility.Visible : Visibility.Collapsed;
            ChkRejoinFixFirewallMatchmaking.Visibility = steamOnlyFeaturesAvailable ? Visibility.Visible : Visibility.Collapsed;

            if (!steamOnlyFeaturesAvailable)
            {
                SetRejoinFirewallCheckbox(ChkRejoinFixFirewall, false);
                SetRejoinFirewallCheckbox(ChkRejoinFixFirewallMatchmaking, false);
                _steamFirewallAutoEnabled = false;
                _steamFirewallAutoPaused = false;
                _rejoinCampaignFirewallApplying = false;
                _rejoinCampaignFirewallEnabled = false;
            }

            if (_mainWindowInitialized)
            {
                UpdateRejoinFirewallOptionAvailability(_rejoinProxy.IsRunning);
                UpdateRejoinFirewallStatus();
                UpdateToolboxStatus();
            }
        }

        private static void SetToolboxStatusText(TextBlock target, string text, string color)
        {
            if (string.Equals(target.Text, text, StringComparison.Ordinal))
                return;

            target.Text = text;
            target.Foreground = Brush(color);
        }

        private void VpnConnectionPresence_Changed(object? sender, EventArgs e) =>
            Dispatcher.InvokeAsync(() =>
            {
                PublishObsOverlaySnapshot();
                UpdateToolboxStatus();
            });

        private Vpn EnsureVpnTab()
        {
            if (_vpnTab is not null)
                return _vpnTab;

            _vpnTab = new Vpn(() => EnsureCompanionServicesRunningAsync("MCC VPN"));
            VpnHost.Content = _vpnTab;
            return _vpnTab;
        }

        private BanChecker EnsureBanCheckerTab()
        {
            if (_banCheckerTab is not null)
                return _banCheckerTab;

            _banCheckerTab = new BanChecker(
                StatsCheckBanTargetsAsync,
                GetBanCheckerAuthorizationState,
                () => ShowToolsPage(ToolsPage.Home));
            BanCheckerHost.Content = _banCheckerTab;
            return _banCheckerTab;
        }

        private MatchHistory EnsureMatchHistoryTab()
        {
            EnsureStatsInitialized();
            if (_matchHistoryTab is not null)
                return _matchHistoryTab;

            string initialGamertag;
            lock (_statsLock)
                initialGamertag = _statsGamertag;

            _matchHistoryTab = new MatchHistory(
                initialGamertag,
                () =>
                {
                    lock (_statsLock)
                        return _statsSpartanToken;
                });
            MatchHistoryHost.Content = _matchHistoryTab;
            return _matchHistoryTab;
        }

        private void EnsureSelectedSectionContent(TabItem selectedTab)
        {
            if (ReferenceEquals(selectedTab, H3ModsSection))
                EnsureModsTab();
            else if (ReferenceEquals(selectedTab, TheaterSection) || ReferenceEquals(selectedTab, DownpatchSection))
            {
                var theater = EnsureTheaterTab();
                bool downpatch = ReferenceEquals(selectedTab, DownpatchSection);
                var host = downpatch ? DownpatchHost : TheaterHost;
                if (!ReferenceEquals(host.Content, theater))
                {
                    TheaterHost.Content = null;
                    DownpatchHost.Content = null;
                    host.Content = theater;
                }
                theater.ShowDownpatchPage(downpatch);
            }
            else if (ReferenceEquals(selectedTab, PlaylistsSection))
                EnsurePlaylistsTab();
            else if (ReferenceEquals(selectedTab, VpnSection))
                EnsureVpnTab();
            else if (ReferenceEquals(selectedTab, StatsSection))
                EnsureStatsInitialized();
            else if (ReferenceEquals(selectedTab, PopulationSection))
            {
                EnsureStatsInitialized();
            }
            else if (ReferenceEquals(selectedTab, MatchHistorySection))
                EnsureMatchHistoryTab();
            else if (ReferenceEquals(selectedTab, BanCheckerSection))
                EnsureBanCheckerTab().RefreshAuthorizationStatus();
        }

        public async Task OpenVpnSectionAndConnectAsync()
        {
            VpnSection.Visibility = Visibility.Visible;
            ShowVpnSection.IsChecked = true;
            App.SaveMainSectionVisible("Vpn", true);
            MainTabs.SelectedItem = VpnSection;
            _lastMainTab = VpnSection;
            SyncSidebarSelection(VpnSection);
            await EnsureVpnTab().ConnectFromRelaunchAsync();
        }

        // ------------------------------------------
        // Window chrome
        // ------------------------------------------
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left)
                return;

            if (e.ClickCount == 2)
            {
                ToggleMaximizeRestore();
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
                return;

            if (WindowState == WindowState.Maximized)
                RestoreForTitleBarDrag(e);

            try
            {
                DragMove();
            }
            catch (InvalidOperationException)
            {
                // Windows can cancel mouse capture while restoring from maximized.
            }
        }

        private void MinBtn_Click(object sender, RoutedEventArgs e) =>
            WindowState = WindowState.Minimized;

        private void MaxBtn_Click(object sender, RoutedEventArgs e) =>
            ToggleMaximizeRestore();

        private void CloseBtn_Click(object sender, RoutedEventArgs e) =>
            Close();

        private async void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            SaveMainWindowPlacement();

            if (_closeFirewallCleanupStarted)
                return;

            bool shouldCleanFirewall = _steamFirewallAutoEnabled
                || _rejoinCampaignFirewallEnabled
                || _steamFirewallUiState is SteamFirewallState.Enabled or SteamFirewallState.Partial;

            if (!shouldCleanFirewall)
                return;

            e.Cancel = true;
            _closeFirewallCleanupStarted = true;
            SetStatus("Closing: disabling Toolbox firewall rules...", "#FF6A00");
            AppendLog("[FIREWALL]", "Closing Toolbox; disabling MCC P2P firewall rules first.", "#FF6A00");

            try
            {
                DisableSteamFirewallAutoMode(logStatus: false);
                await DisableRejoinFirewallRulesAsync(logStatus: false);
                AppendLog("[FIREWALL]", "MCC P2P firewall rules disabled for shutdown.", "#39FF14");
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Shutdown firewall cleanup failed: {ex.Message}", "#FF2D55");
            }
            finally
            {
                Close();
            }
        }

        private void ThemeToggleBtn_Click(object sender, RoutedEventArgs e)
        {
            App.ToggleTheme();
            ThemeToggleBtn.Content = App.IsDarkTheme ? "☾" : "☀";
        }

        private void MainTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // SelectionChanged events from controls inside a tab bubble through the TabControl.
            if (e.OriginalSource != MainTabs || _restoringMainTabSelection)
                return;

            // Use the tab reported by this event rather than MainTabs.SelectedItem,
            // which can still point at the previous tab during a selection transition.
            if (e.AddedItems.Count != 1 || e.AddedItems[0] is not TabItem selectedTab)
                return;

            // WebView2 profile startup can briefly occupy the WPF UI thread. Keep it
            // out of the application launch path and initialize it only when needed.
            if (_mainWindowInitialized && ReferenceEquals(selectedTab, ReportSection))
                _ = EnsureSupportSessionCheckedAsync();

            if (!ReferenceEquals(selectedTab, H3ModsSection) || IsRunningAsAdministrator())
            {
                EnsureSelectedSectionContent(selectedTab);
                _lastMainTab = selectedTab;
                SyncSidebarSelection(selectedTab);
                if (App.LoadSetupPreference("OpenLastSection", true))
                    App.SaveLastMainSection(selectedTab.Name);
                return;
            }

            // Do not leave the admin-only control active while the UAC decision is pending.
            _restoringMainTabSelection = true;
            MainTabs.SelectedItem = _lastMainTab ?? ToolsSection;
            _restoringMainTabSelection = false;
            SyncSidebarSelection(MainTabs.SelectedItem as TabItem ?? ToolsSection);

            var result = ToolboxDialog.Show(
                "Mods requires the Toolbox to run as Administrator.\n\nRelaunch as Administrator now?",
                "Mods -- Halo MCC Toolbox",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                RelaunchAsAdministrator();
                Close();
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                AppendLog("[INFO]", "Mods relaunch cancelled at administrator prompt.", "#4A5A6A");
                SetStatus("Mods requires Administrator.", "#4A5A6A");
            }
            catch (Exception ex)
            {
                ToolboxDialog.Show(
                    $"Could not relaunch the Toolbox as Administrator.\n\n{ex.Message}",
                    "Mods -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                AppendLog("[ERROR]", $"Could not relaunch as Administrator: {ex.Message}", "#FF2D55");
            }
        }

        private void SectionSettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            SectionSettingsPopup.IsOpen = !SectionSettingsPopup.IsOpen;
        }

        private (string Key, FrameworkElement Target, CheckBox CheckBox)[] OptionalSections =>
        [
            ("Maps", SidebarMapsButton, ShowMapsSection),
            ("H3Mods", H3ModsSection, ShowH3ModsSection),
            ("Playlists", PlaylistsSection, ShowPlaylistsSection),
            ("Vpn", VpnSection, ShowVpnSection),
            ("Overlays", SidebarOverlaysButton, ShowOverlaysSection),
            ("Firewall", SidebarFirewallButton, ShowFirewallSection),
            ("Stats", StatsSection, ShowStatsSection),
            ("Population", PopulationSection, ShowPopulationSection),
            ("MatchHistory", MatchHistorySection, ShowMatchHistorySection),
            ("BanChecker", BanCheckerSection, ShowBanCheckerSection),
            ("Theater", TheaterSection, ShowTheaterSection),
            ("Downpatch", DownpatchSection, ShowDownpatchSection),
            ("Report", ReportSection, ShowReportSection),
            ("Fixes", SidebarFixesButton, ShowFixesSection),
            ("Log", LogSection, ShowLogSection),
        ];

        private void LoadSectionVisibility()
        {
            foreach (var section in OptionalSections)
                SetSectionVisibility(section.Target, section.CheckBox, section.Key, App.LoadMainSectionVisible(section.Key));
            AboutSection.Visibility = Visibility.Visible;
            UpdateToggleAllSectionsButton();
        }

        private void SectionVisibility_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not CheckBox checkBox || checkBox.Tag is not string key)
                return;

            var section = OptionalSections.First(s => s.Key == key);
            SetSectionVisibility(section.Target, checkBox, key, checkBox.IsChecked == true, save: true);
            if (checkBox.IsChecked != true)
                ReturnToDashboardIfHidden();
            UpdateToggleAllSectionsButton();
        }

        private void ToggleAllSections_Click(object sender, RoutedEventArgs e)
        {
            bool makeVisible = !AreAllOptionalSectionsVisible();
            foreach (var section in OptionalSections)
                SetSectionVisibility(section.Target, section.CheckBox, section.Key, makeVisible, save: true);
            if (!makeVisible)
                ShowToolsPage(ToolsPage.Home);
            UpdateToggleAllSectionsButton();
        }

        private void ReturnToDashboardIfHidden()
        {
            if (MainTabs.SelectedItem is TabItem { Visibility: not Visibility.Visible } ||
                OptionalSections.Any(s => s.Target is RadioButton { IsChecked: true, Visibility: not Visibility.Visible }))
                ShowToolsPage(ToolsPage.Home);
        }

        private bool AreAllOptionalSectionsVisible() =>
            OptionalSections.All(s => s.CheckBox.IsChecked == true);

        private void UpdateToggleAllSectionsButton()
        {
            ToggleAllSectionsBtn.Content = AreAllOptionalSectionsVisible() ? "HIDE ALL" : "SHOW ALL";
        }

        private static void SetSectionVisibility(FrameworkElement section, CheckBox checkBox, string sectionName, bool visible, bool save = false)
        {
            section.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            checkBox.IsChecked = visible;
            if (save)
                App.SaveMainSectionVisible(sectionName, visible);
        }

        private void ToggleMaximizeRestore()
        {
            WindowState = WindowState == WindowState.Maximized
                ? WindowState.Normal
                : WindowState.Maximized;
        }

        private void UpdateMaximizeButton()
        {
            if (MaxBtn is null)
                return;

            bool isMaximized = WindowState == WindowState.Maximized;
            MaxBtn.Content = isMaximized ? "❐" : "□";
            MaxBtn.ToolTip = isMaximized ? "Restore" : "Maximize";
        }

        private void RestoreMainWindowPlacement()
        {
            var placement = App.LoadMainWindowPlacement();
            if (placement is null)
                return;

            double width = Math.Clamp(placement.Value.Width, MinWidth, SystemParameters.VirtualScreenWidth);
            double height = Math.Clamp(placement.Value.Height, MinHeight, SystemParameters.VirtualScreenHeight);
            double left = ClampWindowCoordinate(
                placement.Value.Left,
                SystemParameters.VirtualScreenLeft,
                SystemParameters.VirtualScreenLeft + SystemParameters.VirtualScreenWidth,
                width);
            double top = ClampWindowCoordinate(
                placement.Value.Top,
                SystemParameters.VirtualScreenTop,
                SystemParameters.VirtualScreenTop + SystemParameters.VirtualScreenHeight,
                height);

            Width = width;
            Height = height;
            Left = left;
            Top = top;

            if (placement.Value.IsMaximized)
                WindowState = WindowState.Maximized;
        }

        private void SaveMainWindowPlacement()
        {
            Rect bounds = WindowState == WindowState.Maximized || WindowState == WindowState.Minimized
                ? RestoreBounds
                : new Rect(Left, Top, ActualWidth > 0 ? ActualWidth : Width, ActualHeight > 0 ? ActualHeight : Height);

            if (double.IsNaN(bounds.Left) || double.IsNaN(bounds.Top) ||
                double.IsNaN(bounds.Width) || double.IsNaN(bounds.Height) ||
                bounds.Width < MinWidth || bounds.Height < MinHeight)
            {
                return;
            }

            App.SaveMainWindowPlacement(new App.WindowPlacement(
                bounds.Left,
                bounds.Top,
                bounds.Width,
                bounds.Height,
                WindowState == WindowState.Maximized));
        }

        private static double ClampWindowCoordinate(double value, double min, double max, double size)
        {
            const double visibleEdge = 80;
            double lower = min - size + visibleEdge;
            double upper = max - visibleEdge;
            if (lower > upper)
                return min;

            return Math.Clamp(value, lower, upper);
        }

        private void RestoreForTitleBarDrag(MouseButtonEventArgs e)
        {
            Point mouseOnWindow = e.GetPosition(this);
            Point mouseOnScreen = PointToScreen(mouseOnWindow);
            var source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget is not null)
                mouseOnScreen = source.CompositionTarget.TransformFromDevice.Transform(mouseOnScreen);

            Rect restoreBounds = RestoreBounds;
            double horizontalRatio = ActualWidth > 0
                ? Math.Clamp(mouseOnWindow.X / ActualWidth, 0.0, 1.0)
                : 0.5;

            WindowState = WindowState.Normal;
            Left = mouseOnScreen.X - (restoreBounds.Width * horizontalRatio);
            Top = Math.Max(0, mouseOnScreen.Y - mouseOnWindow.Y);
        }

        // ------------------------------------------
        // LOG helpers
        // ------------------------------------------
        private void AppendLog(string tag, string message, string colorHex = "#C8D8E8")
        {
            // Producers never wait for the UI, and bursts share one background dispatch.
            lock (_pendingLogLock)
            {
                _pendingLogLines.Enqueue($"[{DateTime.Now:HH:mm:ss}] {tag} {message}");
                while (_pendingLogLines.Count > MaxSessionLogLines)
                    _pendingLogLines.Dequeue();
                if (_logFlushScheduled) return;
                _logFlushScheduled = true;
            }
            Dispatcher.InvokeAsync(FlushPendingLog, System.Windows.Threading.DispatcherPriority.Background);
        }

        private void FlushPendingLog()
        {
            string[] lines;
            lock (_pendingLogLock)
            {
                lines = _pendingLogLines.ToArray();
                _pendingLogLines.Clear();
                _logFlushScheduled = false;
            }
            if (lines.Length == 0) return;
            _sessionLogLines.AddRange(lines);
            if (_sessionLogLines.Count > MaxSessionLogLines)
            {
                _sessionLogLines.RemoveRange(0, _sessionLogLines.Count - MaxSessionLogLines);
                TxtLog.Text = string.Join(Environment.NewLine, _sessionLogLines) + Environment.NewLine;
            }
            else
                TxtLog.AppendText(string.Join(Environment.NewLine, lines) + Environment.NewLine);
            TxtLog.ScrollToEnd();
        }

        private void SetStatus(string msg, string colorHex = "#4A5A6A")
        {
            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = msg;
                TxtStatus.Foreground = Brush(colorHex);
            });
        }

        private void StartNetworkStatsOverlay(string targetIp, GameServerInfo? serverInfo = null)
        {
            if (!_rejoinProxy.IsRunning)
            {
                _gameServerConnectionMonitor.Stop();
                _networkStatsMonitor.Stop();
                CloseGameNetworkStatsOverlay();
                ClearNetworkOverlaySnapshots();
                _lastNetworkStatsRelayServer = null;
                PublishObsOverlaySnapshot();
                return;
            }

            _gameServerConnectionMonitor.Start();

            if (!_networkStatsOverlayEnabled && !_matchmakingWaitOverlayEnabled &&
                !_combinedNetworkSessionOverlayEnabled &&
                !_obsBrowserOverlaySessionStatsEnabled && !_obsBrowserOverlayEnabled)
            {
                _networkStatsMonitor.Stop();
                CloseGameNetworkStatsOverlay();
                ClearNetworkOverlaySnapshots();
                PublishObsOverlaySnapshot();
                return;
            }

            if (_networkStatsOverlayEnabled || _matchmakingWaitOverlayEnabled ||
                _combinedNetworkSessionOverlayEnabled || _obsBrowserOverlaySessionStatsEnabled)
            {
                EnsureGameNetworkStatsOverlay();
                foreach (var overlay in AllGameOverlays())
                {
                    overlay.SetPreferredProcessId(TryGetMccProcessId());
                    overlay.SetMoveMode(_networkStatsOverlayMoveEnabled);
                }
            }
            else
            {
                CloseGameNetworkStatsOverlay();
            }

            if (_obsBrowserOverlayEnabled)
                EnsureOverlaySourceServer(logStatus: false);

            if (string.IsNullOrWhiteSpace(targetIp))
            {
                ClearNetworkStatsOverlayDisplay();
                return;
            }

            if (_networkStatsOverlayEnabled || _combinedNetworkSessionOverlayEnabled)
                _gameNetworkStatsOverlay?.UpdateServer(serverInfo ?? _rejoinProxy.CurrentGameServerInfo);

            _networkStatsMonitor.Start(targetIp);
            PublishObsOverlaySnapshot();
            var port = serverInfo?.Ports.FirstOrDefault()?.Num;
            var endpoint = port is > 0 ? $"{targetIp}:{port}" : targetIp;
            AppendLog("[NET]", $"Monitoring server latency for {endpoint}.", "#00C8FF");
        }

        private void HandleTrustedGameServerChanged(GameServerInfo? serverInfo)
        {
            if (serverInfo is not null && !string.IsNullOrWhiteSpace(serverInfo.IPv4Address))
            {
                _trustedDedicatedServer = serverInfo;
                var port = serverInfo.Ports.FirstOrDefault()?.Num;
                var endpoint = port is > 0 ? $"{serverInfo.IPv4Address}:{port}" : serverInfo.IPv4Address;
                HoldSteamFirewallPausedForActiveMatch("dedicated server active");
                AppendLog("[GUARD]", $"Trusted dedicated server set to {endpoint}.", "#39FF14");
            }
            else
            {
                _trustedDedicatedServer = null;
            }

            UpdateStatsServerLabels(serverInfo);
            StartNetworkStatsOverlay(serverInfo?.IPv4Address ?? "", serverInfo);
        }

        private void HandleNetworkStatsObservedServer(GameServerInfo? serverInfo)
        {
            if (serverInfo is not null && !string.IsNullOrWhiteSpace(serverInfo.IPv4Address))
            {
                _lastNetworkStatsRelayServer = serverInfo;
                UpdateStatsServerLabels(serverInfo);
                StartNetworkStatsOverlay(serverInfo.IPv4Address, serverInfo);
                return;
            }

            _lastNetworkStatsRelayServer = null;
            UpdateStatsServerLabels(_trustedDedicatedServer ?? _rejoinProxy.CurrentGameServerInfo);
            ClearNetworkStatsOverlayDisplay();
        }

        private void UpdateStatsServerLabels(GameServerInfo? serverInfo)
        {
            if (serverInfo is null)
                return;

            string region = GameServerRegionResolver.GetRegionLabel(serverInfo);
            if (string.IsNullOrWhiteSpace(region))
                region = serverInfo.Region;

            if (string.IsNullOrWhiteSpace(region) ||
                region.Equals("ACTIVE UDP", StringComparison.OrdinalIgnoreCase))
                return;

            string label = $"Server - {region}";
            _statsCurrentLobbyServerText = label;
            _statsLastGameServerText = label;
            StatsCurrentLobbyServerLabel.Text = label;
            StatsLastGameServerLabel.Text = StatsFormatLastGameHeader();
            PublishObsOverlaySnapshot();
        }

        private string StatsFormatLastGameHeader() => string.Join(
            "  ·  ",
            new[] { _statsLastGameModeText, _statsLastGameServerText, _statsLastGamePlayedText }
                .Where(value => !string.IsNullOrWhiteSpace(value)));

        private void ClearNetworkStatsOverlayDisplay()
        {
            if ((!_networkStatsOverlayEnabled && !_matchmakingWaitOverlayEnabled &&
                 !_combinedNetworkSessionOverlayEnabled && !_obsBrowserOverlayEnabled) || !_rejoinProxy.IsRunning)
                return;

            if (_networkStatsOverlayEnabled || _matchmakingWaitOverlayEnabled || _combinedNetworkSessionOverlayEnabled)
            {
                EnsureGameNetworkStatsOverlay();
                _gameNetworkStatsOverlay?.SetPreferredProcessId(TryGetMccProcessId());
                _gameNetworkStatsOverlay?.SetMoveMode(_networkStatsOverlayMoveEnabled);
            }
            _networkStatsMonitor.Stop();
            ClearNetworkOverlaySnapshots();
            if (_networkStatsOverlayEnabled || _combinedNetworkSessionOverlayEnabled)
                _gameNetworkStatsOverlay?.ClearStats();
            PublishObsOverlaySnapshot();
        }

        private void ClearNetworkOverlaySnapshots()
        {
            _lastNetworkStatsSnapshot = null;
            _lastNetworkTrafficSnapshot = null;
        }

        private void UpdateNetworkStatsOverlay(NetworkStatsSnapshot snapshot)
        {
            _lastNetworkStatsSnapshot = snapshot;
            if (_networkStatsOverlayEnabled || _combinedNetworkSessionOverlayEnabled)
            {
                EnsureGameNetworkStatsOverlay();
                _gameNetworkStatsOverlay?.UpdateStats(snapshot);
            }
            PublishObsOverlaySnapshot();
        }

        private void UpdateNetworkTrafficOverlay(NetworkTrafficSnapshot snapshot)
        {
            _lastNetworkTrafficSnapshot = snapshot;
            if ((!_networkStatsOverlayEnabled && !_combinedNetworkSessionOverlayEnabled && !_obsBrowserOverlayEnabled) || !_rejoinProxy.IsRunning)
                return;

            if (_networkStatsOverlayEnabled || _combinedNetworkSessionOverlayEnabled)
            {
                EnsureGameNetworkStatsOverlay();
                _gameNetworkStatsOverlay?.UpdateTrafficStats(snapshot);
            }
            PublishObsOverlaySnapshot();
        }

        private string GetNetworkStatsTargetIp()
        {
            if (_lastNetworkStatsRelayServer is not null &&
                !string.IsNullOrWhiteSpace(_lastNetworkStatsRelayServer.IPv4Address))
            {
                return _lastNetworkStatsRelayServer.IPv4Address;
            }

            return _rejoinProxy.CurrentGameServerIp;
        }

        private GameServerInfo? GetNetworkStatsTargetServerInfo()
        {
            return _lastNetworkStatsRelayServer ?? _rejoinProxy.CurrentGameServerInfo;
        }

        private void EnsureGameNetworkStatsOverlay()
        {
            if (_networkStatsOverlayEnabled && !_networkStatsObsOnly && _gameNetworkStatsOverlay is null)
                _gameNetworkStatsOverlay = CreateComponentOverlay("network");
            else if ((!_networkStatsOverlayEnabled || _networkStatsObsOnly) && _gameNetworkStatsOverlay is not null)
                CloseComponentOverlay(ref _gameNetworkStatsOverlay);

            if (_matchmakingWaitOverlayEnabled && !_matchmakingWaitObsOnly && _matchmakingWaitOverlay is null)
                _matchmakingWaitOverlay = CreateComponentOverlay("wait");
            else if ((!_matchmakingWaitOverlayEnabled || _matchmakingWaitObsOnly) && _matchmakingWaitOverlay is not null)
                CloseComponentOverlay(ref _matchmakingWaitOverlay);

            if (_obsBrowserOverlaySessionStatsEnabled && !_sessionStatsObsOnly && _sessionStatsOverlay is null)
                _sessionStatsOverlay = CreateComponentOverlay("session");
            else if ((!_obsBrowserOverlaySessionStatsEnabled || _sessionStatsObsOnly) && _sessionStatsOverlay is not null)
                CloseComponentOverlay(ref _sessionStatsOverlay);

            if (_combinedNetworkSessionOverlayEnabled && !_combinedNetworkSessionObsOnly && _combinedNetworkSessionOverlay is null)
                _combinedNetworkSessionOverlay = CreateComponentOverlay("combined");
            else if ((!_combinedNetworkSessionOverlayEnabled || _combinedNetworkSessionObsOnly) && _combinedNetworkSessionOverlay is not null)
                CloseComponentOverlay(ref _combinedNetworkSessionOverlay);

            foreach (var overlay in AllGameOverlays())
            {
                overlay.SetPreferredProcessId(TryGetMccProcessId());
                overlay.SetMoveMode(_networkStatsOverlayMoveEnabled);
                overlay.UpdateSessionStats(BuildObsOverlaySnapshot());
            }
            PublishObsOverlaySnapshot();
        }

        private GameNetworkStatsOverlayWindow CreateComponentOverlay(string component)
        {
            var overlay = new GameNetworkStatsOverlayWindow(
                component,
                App.LoadGameOverlayVisualStyle(component))
            {
                Owner = this
            };
            overlay.RelativePlacementChanged += (_, placement) =>
                ComponentOverlay_RelativePlacementChanged(component, placement);
            overlay.Closed += (_, _) =>
            {
                if (ReferenceEquals(_gameNetworkStatsOverlay, overlay)) _gameNetworkStatsOverlay = null;
                if (ReferenceEquals(_matchmakingWaitOverlay, overlay)) _matchmakingWaitOverlay = null;
                if (ReferenceEquals(_sessionStatsOverlay, overlay)) _sessionStatsOverlay = null;
                if (ReferenceEquals(_combinedNetworkSessionOverlay, overlay)) _combinedNetworkSessionOverlay = null;
            };
            overlay.Show();
            return overlay;
        }

        private IEnumerable<GameNetworkStatsOverlayWindow> AllGameOverlays()
        {
            if (_gameNetworkStatsOverlay is not null) yield return _gameNetworkStatsOverlay;
            if (_matchmakingWaitOverlay is not null) yield return _matchmakingWaitOverlay;
            if (_sessionStatsOverlay is not null) yield return _sessionStatsOverlay;
            if (_combinedNetworkSessionOverlay is not null) yield return _combinedNetworkSessionOverlay;
        }

        private void GameNetworkStatsOverlay_RelativePlacementChanged(object? sender, Rect placement)
        {
            _lastOverlayRelativePlacement = placement;
            PublishObsOverlaySnapshot();
        }

        private void ComponentOverlay_RelativePlacementChanged(string component, Rect placement)
        {
            _componentOverlayRelativePlacements[component] = placement;
            PublishObsOverlaySnapshot();
        }

        private static int? TryGetMccProcessId()
        {
            try
            {
                using var process = MccProcessLocator.GetLatestRuntimeProcess(App.LoadMccInstallationPath());
                return process?.Id;
            }
            catch
            {
                return null;
            }
        }

        private void CloseGameNetworkStatsOverlay()
        {
            var overlays = AllGameOverlays().ToList();
            _gameNetworkStatsOverlay = null;
            _matchmakingWaitOverlay = null;
            _sessionStatsOverlay = null;
            _combinedNetworkSessionOverlay = null;
            foreach (var overlay in overlays) overlay.Close();
        }

        private static void CloseComponentOverlay(ref GameNetworkStatsOverlayWindow? overlay)
        {
            var closing = overlay;
            overlay = null;
            closing?.Close();
        }

        private void ChkNetworkStatsOverlay_Checked(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayEnabled = true;
            if (!_mainWindowInitialized)
                return;

            App.SaveGameNetworkStatsOverlayEnabled(true);
            if (!_rejoinProxy.IsRunning)
            {
                UpdateRejoinFixUi();
                return;
            }
            StartNetworkStatsOverlay(GetNetworkStatsTargetIp(), GetNetworkStatsTargetServerInfo());
            AppendLog("[NET]", "Game network stats overlay enabled.", "#00C8FF");
            UpdateRejoinFixUi();
        }

        private void ChkNetworkStatsOverlay_Unchecked(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayEnabled = false;
            if (!_mainWindowInitialized)
                return;

            App.SaveGameNetworkStatsOverlayEnabled(false);
            _networkStatsOverlayMoveEnabled = false;
            CloseOverlayRepositionHelp();
            CloseComponentOverlay(ref _gameNetworkStatsOverlay);
            if (!_combinedNetworkSessionOverlayEnabled && !_obsBrowserOverlayEnabled)
            {
                _gameServerConnectionMonitor.Stop();
                _networkStatsMonitor.Stop();
            }
            PublishObsOverlaySnapshot();
            if (!_matchmakingWaitOverlayEnabled && !_combinedNetworkSessionOverlayEnabled &&
                !_obsBrowserOverlaySessionStatsEnabled && !_obsBrowserOverlayEnabled)
                _obsOverlayServer.Stop();
            AppendLog("[NET]", "Game network stats overlay disabled.", "#C8D8E8");
            UpdateRejoinFixUi();
        }

        private void ChkMatchmakingWaitOverlay_Checked(object sender, RoutedEventArgs e)
        {
            _matchmakingWaitOverlayEnabled = true;
            if (!_mainWindowInitialized) return;
            App.SaveMatchmakingWaitOverlayEnabled(true);
            if (!_rejoinProxy.IsRunning)
            {
                UpdateRejoinFixUi();
                return;
            }
            if (_rejoinProxy.IsRunning)
            {
                EnsureGameNetworkStatsOverlay();
                _gameNetworkStatsOverlay?.SetPreferredProcessId(TryGetMccProcessId());
            }
            PublishObsOverlaySnapshot();
            AppendLog("[MATCH]", "Matchmaking wait estimate overlay enabled.", "#00C8FF");
            UpdateRejoinFixUi();
        }

        private void ChkMatchmakingWaitOverlay_Unchecked(object sender, RoutedEventArgs e)
        {
            _matchmakingWaitOverlayEnabled = false;
            if (!_mainWindowInitialized) return;
            App.SaveMatchmakingWaitOverlayEnabled(false);
            CloseComponentOverlay(ref _matchmakingWaitOverlay);
            PublishObsOverlaySnapshot();
            if (!_networkStatsOverlayEnabled && !_combinedNetworkSessionOverlayEnabled &&
                !_obsBrowserOverlaySessionStatsEnabled && !_obsBrowserOverlayEnabled)
                _obsOverlayServer.Stop();
            AppendLog("[MATCH]", "Matchmaking wait estimate overlay disabled.", "#C8D8E8");
            UpdateRejoinFixUi();
        }

        private void CombinedNetworkSessionOverlayToggle_Checked(object sender, RoutedEventArgs e)
        {
            _combinedNetworkSessionOverlayEnabled = true;
            if (!_mainWindowInitialized)
                return;

            App.SaveCombinedNetworkSessionOverlayEnabled(true);
            if (!_rejoinProxy.IsRunning)
            {
                UpdateRejoinFixUi();
                return;
            }

            StartNetworkStatsOverlay(GetNetworkStatsTargetIp(), GetNetworkStatsTargetServerInfo());
            AppendLog("[NET]", "Combined Network / Session overlay enabled.", "#00C8FF");
            UpdateRejoinFixUi();
        }

        private void CombinedNetworkSessionOverlayToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _combinedNetworkSessionOverlayEnabled = false;
            if (!_mainWindowInitialized)
                return;

            App.SaveCombinedNetworkSessionOverlayEnabled(false);
            CloseComponentOverlay(ref _combinedNetworkSessionOverlay);
            if (!_networkStatsOverlayEnabled && !_obsBrowserOverlayEnabled)
            {
                _gameServerConnectionMonitor.Stop();
                _networkStatsMonitor.Stop();
            }
            PublishObsOverlaySnapshot();
            if (!_networkStatsOverlayEnabled && !_matchmakingWaitOverlayEnabled &&
                !_obsBrowserOverlaySessionStatsEnabled && !_obsBrowserOverlayEnabled)
                _obsOverlayServer.Stop();
            AppendLog("[NET]", "Combined Network / Session overlay disabled.", "#C8D8E8");
            UpdateRejoinFixUi();
        }

        private void OverlayStyleCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_mainWindowInitialized || sender is not ComboBox combo || combo.SelectedIndex < 0)
                return;

            string component = combo.Tag as string ?? "";
            if (component != "network" && component != "session" && component != "wait")
                return;

            var visualStyle = combo.SelectedIndex == 1
                ? GameOverlayVisualStyle.Modern
                : GameOverlayVisualStyle.Classic;
            App.SaveGameOverlayVisualStyle(component, visualStyle);

            var overlay = component == "network"
                ? _gameNetworkStatsOverlay
                : component == "wait" ? _matchmakingWaitOverlay : _sessionStatsOverlay;
            overlay?.SetVisualStyle(visualStyle);
            PublishObsOverlaySnapshot();
            AppendLog(
                "[NET]",
                $"{(component == "network" ? "Network" : component == "wait" ? "Matchmaking wait" : "Session")} overlay style changed to {visualStyle}.",
                "#00C8FF");
            UpdateOverlayPreviews();
        }

        private void UpdateOverlayPreviews()
        {
            if (NetworkPreviewCard is null)
                return;

            bool networkModern = NetworkOverlayStyleCombo.SelectedIndex == 1;
            bool sessionModern = SessionOverlayStyleCombo.SelectedIndex == 1;
            bool waitModern = MatchmakingWaitOverlayStyleCombo.SelectedIndex == 1;
            MatchmakingWaitPreviewClassicPanel.Visibility = waitModern ? Visibility.Collapsed : Visibility.Visible;
            MatchmakingWaitPreviewModernPanel.Visibility = waitModern ? Visibility.Visible : Visibility.Collapsed;
            NetworkPreviewClassicPanel.Visibility = networkModern ? Visibility.Collapsed : Visibility.Visible;
            NetworkPreviewModernPanel.Visibility = networkModern ? Visibility.Visible : Visibility.Collapsed;
            SessionPreviewClassicPanel.Visibility = sessionModern ? Visibility.Collapsed : Visibility.Visible;
            SessionPreviewModernPanel.Visibility = sessionModern ? Visibility.Visible : Visibility.Collapsed;

            SetPreviewState(
                NetworkPreviewCard,
                NetworkPreviewStateText,
                _networkStatsOverlayEnabled,
                _networkStatsObsOnly,
                networkModern ? "MODERN" : "CLASSIC");
            SetPreviewState(
                MatchmakingWaitPreviewCard,
                MatchmakingWaitPreviewStateText,
                _matchmakingWaitOverlayEnabled,
                _matchmakingWaitObsOnly,
                waitModern ? "MODERN" : "CLASSIC");
            SetPreviewState(
                SessionPreviewCard,
                SessionPreviewStateText,
                _obsBrowserOverlaySessionStatsEnabled,
                _sessionStatsObsOnly,
                sessionModern ? "MODERN" : "CLASSIC");
            SetPreviewState(
                CombinedPreviewCard,
                CombinedPreviewStateText,
                _combinedNetworkSessionOverlayEnabled,
                _combinedNetworkSessionObsOnly,
                null);
        }

        private static void SetPreviewState(
            Border card,
            TextBlock label,
            bool enabled,
            bool obsOnly,
            string? visualStyle)
        {
            string state = enabled ? (obsOnly ? "OBS ONLY" : "ON") : "OFF";
            label.Text = string.IsNullOrWhiteSpace(visualStyle) ? state : $"{state} · {visualStyle}";
            label.Foreground = Brush(enabled ? "#39FF14" : "#71869A");
            card.Opacity = enabled ? 1.0 : 0.58;
        }

        private void ChkNetworkStatsOverlayMove_Checked(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayMoveEnabled = true;
            if (!_mainWindowInitialized)
                return;

            if ((_networkStatsOverlayEnabled || _combinedNetworkSessionOverlayEnabled) && _rejoinProxy.IsRunning)
                StartNetworkStatsOverlay(GetNetworkStatsTargetIp(), GetNetworkStatsTargetServerInfo());

            foreach (var overlay in AllGameOverlays()) overlay.SetMoveMode(true);
            ShowOverlayRepositionHelp();
            AppendLog("[NET]", "Overlay drag mode enabled. Drag a docked group together, or Shift+drag one overlay to detach it.", "#00C8FF");
        }

        private void ChkNetworkStatsOverlayMove_Unchecked(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayMoveEnabled = false;
            if (!_mainWindowInitialized)
                return;

            foreach (var overlay in AllGameOverlays()) overlay.SetMoveMode(false);
            CloseOverlayRepositionHelp();
            AppendLog("[NET]", "Overlay drag mode disabled; overlay is click-through.", "#C8D8E8");
        }

        private void BtnNetworkStatsOverlayMove_Click(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayMoveEnabled = !_networkStatsOverlayMoveEnabled;
            if (_networkStatsOverlayMoveEnabled)
                EnsureGameNetworkStatsOverlay();

            foreach (var overlay in AllGameOverlays()) overlay.SetMoveMode(_networkStatsOverlayMoveEnabled);
            if (_networkStatsOverlayMoveEnabled)
                ShowOverlayRepositionHelp();
            else
                CloseOverlayRepositionHelp();
            BtnNetworkStatsOverlayMove.Content = _networkStatsOverlayMoveEnabled
                ? "FINISH"
                : "REPOSITION";
            AppendLog(
                "[NET]",
                _networkStatsOverlayMoveEnabled
                    ? "Overlay drag mode enabled. Drag a docked group together, or Shift+drag one overlay to detach it."
                    : "Overlay drag mode disabled; overlay is click-through.",
                _networkStatsOverlayMoveEnabled ? "#00C8FF" : "#C8D8E8");

            if (!_networkStatsOverlayMoveEnabled && !_rejoinProxy.IsRunning)
                CloseGameNetworkStatsOverlay();
        }

        private void ShowOverlayRepositionHelp()
        {
            CloseOverlayRepositionHelp();
            var helpWindow = new OverlayRepositionHelpWindow
            {
                Owner = this
            };
            helpWindow.Closed += (_, _) =>
            {
                if (ReferenceEquals(_overlayRepositionHelpWindow, helpWindow))
                    _overlayRepositionHelpWindow = null;
            };
            _overlayRepositionHelpWindow = helpWindow;
            helpWindow.Show();
        }

        private void CloseOverlayRepositionHelp()
        {
            var helpWindow = _overlayRepositionHelpWindow;
            _overlayRepositionHelpWindow = null;
            helpWindow?.Close();
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            lock (_pendingLogLock) _pendingLogLines.Clear();
            TxtLog.Clear();
            _sessionLogLines.Clear();
            AppendLog("[INFO]", "Log cleared.", "#4A5A6A");
        }

        private void BtnCopyLog_Click(object sender, RoutedEventArgs e)
        {
            FlushPendingLog();
            if (_sessionLogLines.Count == 0)
            {
                SetStatus("No log lines to copy.", "#FF6A00");
                return;
            }

            Clipboard.SetText(string.Join(Environment.NewLine, _sessionLogLines));
            SetStatus($"Copied {_sessionLogLines.Count} log line(s) to clipboard.", "#39FF14");
        }

        private async void BtnFirewallCheck_Click(object sender, RoutedEventArgs e)
        {
            await RefreshFirewallRuleTableAsync(logResults: true);
        }

        private async void FirewallRulesRefreshButton_Click(object sender, RoutedEventArgs e)
        {
            await RefreshFirewallRuleTableAsync(logResults: false);
        }

        private async Task RefreshFirewallRuleTableAsync(bool logResults)
        {
            if (_firewallRuleTableRefreshing)
                return;

            _firewallRuleTableRefreshing = true;
            BtnFirewallCheck.IsEnabled = false;
            FirewallRulesRefreshButton.IsEnabled = false;
            FirewallRuleTableSummary.Text = "Checking Windows Firewall rules...";
            FirewallRuleTableSummary.Foreground = Brush("#00C8FF");
            if (logResults)
                AppendLog("[FIREWALL]", "Running Firewall Check for all Toolbox firewall rules...", "#00C8FF");
            SetStatus("Checking Toolbox firewall rules...", "#00C8FF");

            try
            {
                string expectedProgram = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
                var rows = new List<FirewallRuleRow>();

                foreach (var group in GetFirewallCheckRuleGroups())
                {
                    if (logResults)
                        AppendLog("[FIREWALL]", group.Header, "#C8D8E8");

                    foreach (string ruleName in group.RuleNames)
                    {
                        string port = ExtractExpectedPortFromRuleName(ruleName) ?? "—";
                        string protocol = ruleName.Contains(" UDP ", StringComparison.OrdinalIgnoreCase) ? "UDP" : "TCP";
                        string direction = ruleName.EndsWith("Inbound", StringComparison.OrdinalIgnoreCase) ? "INBOUND" : "OUTBOUND";
                        string mode = port == "3478" ? "CAMPAIGN" : "MATCHMAKING";
                        var result = await RunNetshAsync("advfirewall", "firewall", "show", "rule", $"name={ruleName}", "verbose");

                        if (result.ExitCode != 0 || !NetshRuleExists(result.Output))
                        {
                            rows.Add(new FirewallRuleRow
                            {
                                Mode = mode,
                                Port = port,
                                Protocol = protocol,
                                Direction = direction,
                                State = "MISSING",
                                Action = "—",
                                Definition = "Rule not installed",
                                StateBrush = Brush("#71869A")
                            });
                            if (logResults)
                                LogFirewallRuleMissing(ruleName);
                            continue;
                        }

                        string enabled = ExtractNetshField(result.Output, "Enabled") ?? "Unknown";
                        string action = ExtractNetshField(result.Output, "Action") ?? "Unknown";
                        string program = ExtractNetshField(result.Output, "Program") ?? "Any";
                        var problems = GetFirewallRuleDefinitionProblems(ruleName, result.Output, expectedProgram);
                        bool isBlocked = enabled.Equals("Yes", StringComparison.OrdinalIgnoreCase)
                            && action.Equals("Block", StringComparison.OrdinalIgnoreCase);
                        string state = problems.Count > 0
                            ? "REPAIR"
                            : isBlocked ? "BLOCKED" : "OPEN";
                        string stateColor = problems.Count > 0
                            ? "#FF6A00"
                            : isBlocked ? "#39FF14" : "#71869A";

                        rows.Add(new FirewallRuleRow
                        {
                            Mode = mode,
                            Port = port,
                            Protocol = protocol,
                            Direction = direction,
                            State = state,
                            Action = action.ToUpperInvariant(),
                            Definition = problems.Count == 0
                                ? $"OK · {Path.GetFileName(program)}"
                                : string.Join(" · ", problems),
                            StateBrush = Brush(stateColor)
                        });

                        if (logResults)
                        {
                            LogFirewallRuleStatus(ruleName, result.Output, logProgram: false);
                            LogFirewallRuleDefinitionProblems(ruleName, result.Output, expectedProgram);
                        }
                    }
                }

                _firewallRuleRows.Clear();
                foreach (FirewallRuleRow row in rows)
                    _firewallRuleRows.Add(row);

                int blocked = rows.Count(row => row.State == "BLOCKED");
                int open = rows.Count(row => row.State == "OPEN");
                int missing = rows.Count(row => row.State == "MISSING");
                int repair = rows.Count(row => row.State == "REPAIR");
                FirewallRuleTableSummary.Text = $"{rows.Count} rules · {blocked} blocked · {open} open · {missing} missing · {repair} repair";
                FirewallRuleTableSummary.Foreground = Brush(repair > 0 ? "#FF6A00" : blocked > 0 ? "#39FF14" : "#C8D8E8");
                _firewallRuleTableLoaded = true;

                if (logResults)
                {
                    AppendLog("[FIREWALL]", $"Expected Program: {expectedProgram}", "#4A5A6A");
                    AppendLog("[FIREWALL]", "Firewall Check complete.", "#39FF14");
                }
                SetStatus("Firewall Check complete.", "#39FF14");
            }
            catch (Exception ex)
            {
                FirewallRuleTableSummary.Text = $"Check failed · {ex.Message}";
                FirewallRuleTableSummary.Foreground = Brush("#FF2D55");
                SetStatus("Firewall Check failed.", "#FF2D55");
                AppendLog("[ERROR]", $"Firewall Check failed: {ex.Message}", "#FF2D55");
            }
            finally
            {
                _firewallRuleTableRefreshing = false;
                BtnFirewallCheck.IsEnabled = true;
                FirewallRulesRefreshButton.IsEnabled = true;
            }
        }

        private static IEnumerable<(string Header, string[] RuleNames)> GetFirewallCheckRuleGroups()
        {
            yield return ("Campaign rules (port 3478)", GetPortRuleNames(SteamFirewallRulePrefix, 3478));
            yield return ("Matchmaking rules (port 4379)", GetPortRuleNames(SteamFirewallRulePrefix, 4379));
        }

        private static string[] GetPortRuleNames(string prefix, int port) => new[]
        {
            $"{prefix} {port} TCP Inbound",
            $"{prefix} {port} UDP Inbound",
            $"{prefix} {port} TCP Outbound",
            $"{prefix} {port} UDP Outbound"
        };

        private static bool NetshRuleExists(string output)
        {
            return output.Contains("Rule Name:", StringComparison.OrdinalIgnoreCase);
        }

        private void LogFirewallRuleMissing(string ruleName)
        {
            AppendLog("[FIREWALL]", $"{GetFirewallRuleShortName(ruleName).PadRight(17)}  MISSING", "#4A5A6A");
        }

        private void LogFirewallRuleStatus(string ruleName, string netshOutput, bool logProgram = true)
        {
            string enabled = ExtractNetshField(netshOutput, "Enabled") ?? "Unknown";
            string action = ExtractNetshField(netshOutput, "Action") ?? "Unknown";
            string direction = ExtractNetshField(netshOutput, "Direction") ?? "Unknown";
            string protocol = ExtractNetshField(netshOutput, "Protocol") ?? "Unknown";
            string localPort = ExtractNetshField(netshOutput, "LocalPort") ?? "Any";
            string remotePort = ExtractNetshField(netshOutput, "RemotePort") ?? "Any";
            string program = ExtractNetshField(netshOutput, "Program") ?? "Any";

            string color = enabled.Equals("Yes", StringComparison.OrdinalIgnoreCase) &&
                           action.Equals("Block", StringComparison.OrdinalIgnoreCase)
                ? "#39FF14"
                : "#FF6A00";

            string shortName = GetFirewallRuleShortName(ruleName);
            string port = localPort.Equals("Any", StringComparison.OrdinalIgnoreCase) ? remotePort : localPort;
            string enabledText = enabled.Equals("Yes", StringComparison.OrdinalIgnoreCase) ? "ENABLED " : "DISABLED";
            string actionText = action.ToUpperInvariant().PadRight(5);

            AppendLog("[FIREWALL]", $"{shortName.PadRight(17)}  Port {port.PadRight(5)}  {enabledText}  {actionText}", color);
            if (logProgram)
                AppendLog("[FIREWALL]", $"Program: {program}", "#4A5A6A");
        }

        private void LogFirewallRuleDefinitionProblems(string ruleName, string netshOutput, string expectedProgram)
        {
            var problems = GetFirewallRuleDefinitionProblems(ruleName, netshOutput, expectedProgram);
            foreach (string problem in problems)
                AppendLog("[FIREWALL]", $"{GetFirewallRuleShortName(ruleName).PadRight(17)}  REPAIR NEEDED - {problem}", "#FF6A00");
        }

        private static List<string> GetFirewallRuleDefinitionProblems(string ruleName, string netshOutput, string expectedProgram)
        {
            var problems = new List<string>();
            string expectedDirection = ruleName.EndsWith("Inbound", StringComparison.OrdinalIgnoreCase) ? "In" : "Out";
            string expectedProtocol = ruleName.Contains(" UDP ", StringComparison.OrdinalIgnoreCase) ? "UDP" : "TCP";
            string expectedPortLabel = expectedDirection.Equals("In", StringComparison.OrdinalIgnoreCase) ? "LocalPort" : "RemotePort";
            string? expectedPort = ExtractExpectedPortFromRuleName(ruleName);

            AddFieldProblem(problems, netshOutput, "Action", "Block");
            AddFieldProblem(problems, netshOutput, "Direction", expectedDirection);
            AddFieldProblem(problems, netshOutput, "Protocol", expectedProtocol);
            AddFieldProblem(problems, netshOutput, "Program", expectedProgram);
            if (!string.IsNullOrWhiteSpace(expectedPort))
                AddFieldProblem(problems, netshOutput, expectedPortLabel, expectedPort);

            return problems;
        }

        private static string? ExtractExpectedPortFromRuleName(string ruleName)
        {
            foreach (int port in SteamFirewallPorts)
            {
                if (ruleName.Contains($" {port} ", StringComparison.OrdinalIgnoreCase))
                    return port.ToString();
            }

            return null;
        }

        private static void AddFieldProblem(List<string> problems, string netshOutput, string field, string expected)
        {
            string? actual = ExtractNetshField(netshOutput, field);
            if (string.IsNullOrWhiteSpace(actual))
            {
                problems.Add($"{field} is unreadable");
                return;
            }

            if (!actual.Equals(expected, StringComparison.OrdinalIgnoreCase))
                problems.Add($"{field} expected {expected}, found {actual}");
        }

        private static string GetFirewallRuleShortName(string ruleName)
        {
            string shortName = ruleName
                .Replace(SteamFirewallRulePrefix, "MCC", StringComparison.OrdinalIgnoreCase)
                .Replace(GlobalSteamFirewallRulePrefix, "GLOBAL", StringComparison.OrdinalIgnoreCase)
                .Replace(LegacyPort4379FirewallRulePrefix, "LEGACY 4379", StringComparison.OrdinalIgnoreCase)
                .Replace("Inbound", "IN", StringComparison.OrdinalIgnoreCase)
                .Replace("Outbound", "OUT", StringComparison.OrdinalIgnoreCase);

            return string.Join(' ', shortName
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries))
                .ToUpperInvariant();
        }

        private static string? ExtractNetshField(string output, string label)
        {
            foreach (string line in output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                int colon = line.IndexOf(':');
                if (colon < 0)
                    continue;

                string key = line[..colon].Trim();
                if (!key.Equals(label, StringComparison.OrdinalIgnoreCase))
                    continue;

                return line[(colon + 1)..].Trim();
            }

            return null;
        }

        private static async Task<(int ExitCode, string Output)> RunNetshAsync(params string[] arguments)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "netsh.exe",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            foreach (string argument in arguments)
                startInfo.ArgumentList.Add(argument);

            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Failed to start netsh.");

            string output = await process.StandardOutput.ReadToEndAsync();
            string error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            string text = string.IsNullOrWhiteSpace(error) ? output : $"{output}\n{error}";
            return (process.ExitCode, text.Trim());
        }

        private void BtnExportLogs_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Title            = "Save Diagnostics ZIP",
                Filter           = "ZIP Archive (*.zip)|*.zip",
                FileName         = $"MCC_Logs_{DateTime.Now:yyyyMMdd_HHmmss}.zip",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (dlg.ShowDialog() != true) return;

            var zipPath = dlg.FileName;
            var mccPath = TxtMccPath.Text.Trim();
            BtnExportLogs.IsEnabled = false;
            AppendLog("[RUN]", "Building diagnostics log bundle...", "#FF6A00");
            SetStatus("Exporting diagnostics logs...", "#FF6A00");

            Task.Run(() =>
            {
                try
                {
                    if (File.Exists(zipPath)) File.Delete(zipPath);

                    var manifest = new StringBuilder();
                    var exportTime = DateTime.Now;
                    var sessionLog = GetSessionLogSnapshot();

                    WriteManifestHeader(manifest, exportTime, mccPath);

                    using var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create);

                    AddTextEntry(zip, "toolbox/session_log.txt", sessionLog);
                    AppendManifestInclude(manifest, "toolbox/session_log.txt", sessionLog.Length, exportTime, "current session log");
                    AppendLog("[ZIP]", "toolbox/session_log.txt", "#C8D8E8");

                    foreach (var file in RejoinFixPaths.GetExportFiles())
                    {
                        try
                        {
                            var fileInfo = new FileInfo(file);
                            if (fileInfo.Length > MaxDiagnosticExportBytes)
                            {
                                AppendManifestSkipped(manifest, file,
                                    $"Skipped oversized file ({FormatSize(fileInfo.Length)} > {FormatSize(MaxDiagnosticExportBytes)}).");
                                continue;
                            }

                            var entryPath = CombineZipPath("toolbox/rejoin_fix", Path.GetFileName(file));
                            zip.CreateEntryFromFile(file, entryPath, CompressionLevel.Fastest);
                            AppendManifestInclude(manifest, entryPath, fileInfo.Length, fileInfo.LastWriteTime, file);
                            AppendLog("[ZIP]", entryPath, "#C8D8E8");
                        }
                        catch (Exception ex)
                        {
                            AppendManifestError(manifest, file, ex.Message);
                            AppendLog("[WARN]", $"Skipped {Path.GetFileName(file)}: {ex.Message}", "#FF6A00");
                        }
                    }

                    var probeRoots = BuildDiagnosticProbeRoots();
                    foreach (var probe in probeRoots)
                    {
                        AppendManifestProbe(manifest, probe.Label, probe.RootPath);

                        if (!Directory.Exists(probe.RootPath))
                        {
                            AppendManifestMissing(manifest, probe.RootPath);
                            continue;
                        }

                        var files = SafeEnumerateFiles(probe.RootPath)
                            .Where(path => probe.Include(path))
                            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        if (probe.LimitToLatest && files.Count > 1)
                        {
                            var latest = files
                                .Select(path => new FileInfo(path))
                                .OrderByDescending(info => info.LastWriteTimeUtc)
                                .First();

                            foreach (var skipped in files.Where(path => !path.Equals(latest.FullName, StringComparison.OrdinalIgnoreCase)))
                                AppendManifestSkipped(manifest, skipped, "Skipped because only the latest file from this source is exported.");

                            files = new List<string> { latest.FullName };
                        }

                        if (files.Count == 0)
                        {
                            AppendManifestSkipped(manifest, probe.RootPath, "No matching diagnostic files found.");
                            continue;
                        }

                        foreach (var file in files)
                        {
                            try
                            {
                                var fileInfo = new FileInfo(file);
                                if (fileInfo.Length > MaxDiagnosticExportBytes)
                                {
                                    AppendManifestSkipped(manifest, file,
                                        $"Skipped oversized file ({FormatSize(fileInfo.Length)} > {FormatSize(MaxDiagnosticExportBytes)}).");
                                    continue;
                                }

                                var relative = Path.GetRelativePath(probe.RootPath, file);
                                var entryPath = CombineZipPath(probe.ZipRoot, relative);
                                zip.CreateEntryFromFile(file, entryPath, CompressionLevel.Fastest);
                                AppendManifestInclude(manifest, entryPath, fileInfo.Length, fileInfo.LastWriteTime, file);
                                AppendLog("[ZIP]", entryPath, "#C8D8E8");
                            }
                            catch (Exception ex)
                            {
                                AppendManifestError(manifest, file, ex.Message);
                                AppendLog("[WARN]", $"Skipped {Path.GetFileName(file)}: {ex.Message}", "#FF6A00");
                            }
                        }
                    }

                    AppendManifestPrivacyNotes(manifest);
                    AddTextEntry(zip, "manifest.txt", manifest.ToString());
                    AppendLog("[ZIP]", "manifest.txt", "#C8D8E8");

                    var info = new FileInfo(zipPath);
                    var sizeTxt = FormatSize(info.Length);
                    AppendLog("[DONE]", $"Diagnostics ZIP created: {sizeTxt}  =>  {zipPath}", "#39FF14");

                    Dispatcher.Invoke(() =>
                    {
                        SetStatus("Diagnostics ZIP created.", "#39FF14");
                        var open = ToolboxDialog.Show(
                            $"Diagnostics ZIP created.\n\nSaved to:\n{zipPath}\n\nOpen containing folder?",
                            "Logs Exported -- Halo MCC Toolbox",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Information);

                        if (open == MessageBoxResult.Yes)
                            Process.Start("explorer.exe", $"/select,\"{zipPath}\"");
                    });
                }
                catch (Exception ex)
                {
                    AppendLog("[ERROR]", $"Log export failed: {ex.Message}", "#FF2D55");
                    Dispatcher.Invoke(() =>
                    {
                        SetStatus("Failed to export diagnostics logs.", "#FF2D55");
                        ToolboxDialog.Show($"Failed to export logs:\n\n{ex.Message}",
                            "Error -- Halo MCC Toolbox", MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
                finally
                {
                    Dispatcher.Invoke(() => BtnExportLogs.IsEnabled = true);
                }
            });
        }

        private async void BtnRemoveToolboxTraces_Click(object sender, RoutedEventArgs e)
        {
            var confirm = ToolboxDialog.Show(
                "This will remove Halo MCC Toolbox data from this PC:\n\n" +
                "- Toolbox Local/Roaming AppData, including WebView2 logins and Rejoin Fix files\n" +
                "- Toolbox registry settings\n" +
                "- Toolbox firewall rules and proxy certificate\n" +
                "- Legacy stats cache/token files saved beside the app\n\n" +
                "It will not delete MCC clips, maps, screenshots, or game files.\n\nContinue?",
                "Remove Toolbox Traces -- Halo MCC Toolbox",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (confirm != MessageBoxResult.Yes)
            {
                AppendLog("[INFO]", "Toolbox trace cleanup cancelled.", "#4A5A6A");
                return;
            }

            BtnRemoveToolboxTraces.IsEnabled = false;
            BtnExportLogs.IsEnabled = false;
            AppendLog("[RUN]", "Removing Halo MCC Toolbox traces from this PC...", "#FF6A00");
            SetStatus("Removing Toolbox traces...", "#FF6A00");

            try
            {
                await RemoveToolboxTracesAsync();
                AppendLog("[DONE]", "Toolbox trace cleanup complete. Restart the app to recreate fresh settings.", "#39FF14");
                SetStatus("Toolbox traces removed.", "#39FF14");

                ToolboxDialog.Show(
                    "Halo MCC Toolbox traces were removed.\n\nRestart the app if you want to keep using it with fresh settings.",
                    "Toolbox Traces Removed -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Toolbox trace cleanup failed: {ex.Message}", "#FF2D55");
                SetStatus("Toolbox trace cleanup failed.", "#FF2D55");
                ToolboxDialog.Show(
                    $"Toolbox trace cleanup failed:\n\n{ex.Message}",
                    "Cleanup Failed -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                BtnRemoveToolboxTraces.IsEnabled = true;
                BtnExportLogs.IsEnabled = true;
            }
        }

        private async Task RemoveToolboxTracesAsync()
        {
            CloseGameNetworkStatsOverlay();
            _gameServerConnectionMonitor.Stop();
            _networkStatsMonitor.Stop();
            _steamFirewallAutoTimer.Stop();

            if (_rejoinProxy.IsRunning)
            {
                StopRejoinCrashWatcher();
                _rejoinProxy.Stop();
                AppendLog("[CLEAN]", "Stopped Rejoin Fix proxy and restored proxy settings.", "#C8D8E8");
            }

            if (ClearToolboxWinInetProxyIfPresent())
                AppendLog("[CLEAN]", $"Cleared WinINet proxy setting: {RejoinFixProxyAddress}", "#C8D8E8");

            var elevatedRemovals = await RemoveElevatedToolboxTracesAsync();
            foreach (var removal in elevatedRemovals)
                AppendLog("[CLEAN]", removal, "#C8D8E8");

            if (elevatedRemovals.Count == 0)
                AppendLog("[CLEAN]", "No elevated firewall, WinHTTP proxy, or certificate traces were found.", "#4A5A6A");

            DeleteRegistrySubKeyTree(Registry.CurrentUser, ToolboxRegistryPath, "HKCU\\" + ToolboxRegistryPath);

            DisposeHiddenCookieChecker();

            DeleteFileIfExists(Path.GetFullPath(StatsSettingsFile));
            DeleteFileIfExists(Path.GetFullPath(StatsCacheFile));
            DeleteFileIfExists(Path.GetFullPath(StatsTokenFile));

            var baseDirectory = AppContext.BaseDirectory;
            DeleteFileIfExists(Path.Combine(baseDirectory, StatsSettingsFile));
            DeleteFileIfExists(Path.Combine(baseDirectory, StatsCacheFile));
            DeleteFileIfExists(Path.Combine(baseDirectory, StatsTokenFile));

            DeleteDirectoryIfSafe(ToolboxRoamingAppDataRoot);
            DeleteDirectoryIfSafe(ToolboxLocalAppDataRoot);
        }

        private static async Task<IReadOnlyList<string>> RemoveElevatedToolboxTracesAsync()
        {
            string allRuleNames = string.Join(", ", SteamFirewallRuleNames
                .Concat(LegacySteamFirewallRuleNames)
                .Concat(GlobalSteamFirewallRuleNames)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(QuotePowerShellString));
            string certPath = Path.Combine(ToolboxLocalAppDataRoot, "RejoinFix", "proxy-root.pfx");
            Directory.CreateDirectory(ToolboxLocalAppDataRoot);
            string resultPath = Path.Combine(ToolboxLocalAppDataRoot, $"trace-cleanup-{Guid.NewGuid():N}.txt");

            string script = $@"
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$removed = New-Object System.Collections.Generic.List[string]
$ruleNames = @({allRuleNames})
foreach ($name in $ruleNames) {{
    $rules = @(Get-NetFirewallRule -DisplayName $name -ErrorAction SilentlyContinue)
    foreach ($rule in $rules) {{
        $removed.Add(""Removed firewall rule: $($rule.DisplayName)"")
        $rule | Remove-NetFirewallRule -ErrorAction SilentlyContinue
    }}
}}

$pfxPath = {QuotePowerShellString(certPath)}
if (Test-Path -LiteralPath $pfxPath) {{
    try {{
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($pfxPath, {QuotePowerShellString(RejoinFixProxyCertificatePassword)}, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)
        $certs = @(Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
            Where-Object {{ $_.Thumbprint -eq $cert.Thumbprint }})
        foreach ($trustedCert in $certs) {{
            $removed.Add(""Removed trusted certificate: $($trustedCert.Subject) [$($trustedCert.Thumbprint)]"")
            $trustedCert | Remove-Item -ErrorAction SilentlyContinue
        }}
    }} catch {{ }}
}}

$winHttp = (& netsh winhttp show proxy) -join ""`n""
if ($winHttp -match '127\.0\.0\.1:(?:8888|19999)') {{
    & netsh winhttp reset proxy | Out-Null
    $removed.Add(""Reset WinHTTP proxy: {RejoinFixProxyAddress}"")
}}

$removed | Set-Content -LiteralPath {QuotePowerShellString(resultPath)} -Encoding UTF8
exit 0";

            try
            {
                await RunPowerShellAsync(script, elevated: !IsRunningAsAdministrator(), timeoutMs: 30000);
                if (!File.Exists(resultPath))
                    return Array.Empty<string>();

                return File.ReadAllLines(resultPath)
                    .Select(line => line.Trim())
                    .Where(line => !string.IsNullOrWhiteSpace(line))
                    .ToArray();
            }
            finally
            {
                try
                {
                    if (File.Exists(resultPath))
                        File.Delete(resultPath);
                }
                catch
                {
                    // The appdata cleanup immediately after this removes the parent directory if needed.
                }
            }
        }

        private void DisposeHiddenCookieChecker()
        {
            var checker = _hiddenCookieChecker;
            if (checker is null)
                return;

            _hiddenCookieChecker = null;
            try
            {
                MainRootGrid.Children.Remove(checker);
                checker.Dispose();
            }
            catch (Exception ex)
            {
                AppendLog("[WARN]", $"Could not release WebView2 before cleanup: {ex.Message}", "#FF6A00");
            }
        }

        private Microsoft.Web.WebView2.Wpf.WebView2 EnsureHiddenCookieChecker()
        {
            if (_hiddenCookieChecker is not null)
                return _hiddenCookieChecker;

            // Keep WebView2 out of the startup visual tree. Constructing it from XAML
            // can initialize Edge while the main window is rendering and block input.
            var checker = new Microsoft.Web.WebView2.Wpf.WebView2
            {
                Width = 0,
                Height = 0,
                IsHitTestVisible = false,
                Focusable = false
            };
            Grid.SetRowSpan(checker, 3);
            MainRootGrid.Children.Add(checker);
            _hiddenCookieChecker = checker;
            return checker;
        }

        private bool ClearToolboxWinInetProxyIfPresent()
        {
            try
            {
                var recovery = ToolboxWinInetProxy.RecoverStaleProxy();
                return recovery is StaleProxyRecoveryResult.RestoredSavedSettings or
                    StaleProxyRecoveryResult.DisabledLegacyProxy;
            }
            catch (Exception ex)
            {
                AppendLog("[WARN]", $"Could not clear Toolbox WinINet proxy setting: {ex.Message}", "#FF6A00");
                return false;
            }
        }

        private void DeleteRegistrySubKeyTree(RegistryKey root, string subKey, string label)
        {
            try
            {
                using var existing = root.OpenSubKey(subKey);
                if (existing is null)
                    return;

                root.DeleteSubKeyTree(subKey, throwOnMissingSubKey: false);
                AppendLog("[CLEAN]", $"Deleted registry key: {label}", "#C8D8E8");
            }
            catch (Exception ex)
            {
                AppendLog("[WARN]", $"Could not delete registry key {label}: {ex.Message}", "#FF6A00");
            }
        }

        private void DeleteDirectoryIfSafe(string path)
        {
            try
            {
                var fullPath = Path.GetFullPath(path);
                if (!IsToolboxAppDataDirectory(fullPath))
                    throw new InvalidOperationException($"Refusing to delete unexpected path: {fullPath}");

                if (!Directory.Exists(fullPath))
                    return;

                Directory.Delete(fullPath, recursive: true);
                AppendLog("[CLEAN]", $"Deleted directory: {fullPath}", "#C8D8E8");
            }
            catch (Exception ex)
            {
                AppendLog("[WARN]", $"Could not delete directory {path}: {ex.Message}", "#FF6A00");
            }
        }

        private static bool IsToolboxAppDataDirectory(string fullPath)
        {
            return IsSamePath(fullPath, ToolboxLocalAppDataRoot) ||
                   IsSamePath(fullPath, ToolboxRoamingAppDataRoot);
        }

        private void DeleteFileIfExists(string path)
        {
            try
            {
                var fullPath = Path.GetFullPath(path);
                var fileName = Path.GetFileName(fullPath);
                if (!string.Equals(fileName, StatsSettingsFile, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(fileName, StatsCacheFile, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(fileName, StatsTokenFile, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Refusing to delete unexpected file: {fullPath}");

                if (!File.Exists(fullPath))
                    return;

                File.Delete(fullPath);
                AppendLog("[CLEAN]", $"Deleted file: {fullPath}", "#C8D8E8");
            }
            catch (Exception ex)
            {
                AppendLog("[WARN]", $"Could not delete file {path}: {ex.Message}", "#FF6A00");
            }
        }

        private static bool IsSamePath(string left, string right)
        {
            return string.Equals(
                Path.TrimEndingDirectorySeparator(Path.GetFullPath(left)),
                Path.TrimEndingDirectorySeparator(Path.GetFullPath(right)),
                StringComparison.OrdinalIgnoreCase);
        }

        private string GetSessionLogSnapshot()
        {
            return Dispatcher.Invoke(() =>
            {
                var lines = _sessionLogLines.Count == 0
                    ? new[] { $"[{DateTime.Now:HH:mm:ss}] [INFO] Log export started before any session entries existed." }
                    : _sessionLogLines.ToArray();

                return string.Join(Environment.NewLine, lines) + Environment.NewLine;
            });
        }

        private static void AddTextEntry(ZipArchive zip, string entryPath, string contents)
        {
            var entry = zip.CreateEntry(entryPath, CompressionLevel.Fastest);
            using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
            writer.Write(contents);
        }

        private static IEnumerable<string> SafeEnumerateFiles(string rootPath)
        {
            try
            {
                return Directory.EnumerateFiles(rootPath, "*", SearchOption.AllDirectories);
            }
            catch
            {
                return Enumerable.Empty<string>();
            }
        }

        private static List<DiagnosticProbeRoot> BuildDiagnosticProbeRoots()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            var mccCrashReportPath = Path.Combine(App.LoadMccInstallationPath(), "crash_report");

            var probes = new List<DiagnosticProbeRoot>
            {
                new("MCC Crash Report", mccCrashReportPath, "mcc/crash_report",
                    path => IsMccCrashReportFile(path), true),
                new("Easy Anti-Cheat", Path.Combine(appData, "EasyAntiCheat"), "eac",
                    path => IsDiagnosticFile(path)),
                new("Steam Logs", Path.Combine(programFilesX86, @"Steam\logs"), "steam/logs",
                    path => IsRelevantSteamLog(path)),
            };

            return probes;
        }

        private static bool IsDiagnosticFile(string path)
        {
            if (IsSensitiveDiagnosticFile(path))
                return false;

            var ext = Path.GetExtension(path);
            if (DiagnosticExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
                return true;

            var fileName = Path.GetFileName(path);
            return fileName.Contains(".log.", StringComparison.OrdinalIgnoreCase)
                || fileName.Contains("crash", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSensitiveDiagnosticFile(string path)
        {
            return Path.GetFileName(path).Equals(StatsTokenFile, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsMccCrashReportFile(string path)
        {
            var ext = Path.GetExtension(path);
            return ext.Equals(".dmp", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".bmp", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".png", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRelevantSteamLog(string path)
        {
            var fileName = Path.GetFileName(path);
            return fileName.Equals("gameprocess_log.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.Equals("gameprocess_log.previous.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.Equals("content_log.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.Equals("content_log.previous.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.Equals("appinfo_log.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.Equals("appinfo_log.previous.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.Equals("cloud_log.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.Equals("cloud_log.previous.txt", StringComparison.OrdinalIgnoreCase)
                || fileName.StartsWith("connection_log_976730", StringComparison.OrdinalIgnoreCase);
        }

        private static string CombineZipPath(string zipRoot, string relativePath)
        {
            var normalized = relativePath.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
            return $"{zipRoot.TrimEnd('/')}/{normalized}";
        }

        private static string FormatSize(long bytes)
        {
            double size = bytes;
            string[] units = { "B", "KB", "MB", "GB" };
            int unit = 0;
            while (size >= 1024 && unit < units.Length - 1)
            {
                size /= 1024;
                unit++;
            }

            return unit == 0 ? $"{size:0} {units[unit]}" : $"{size:0.0} {units[unit]}";
        }

        private static void WriteManifestHeader(StringBuilder manifest, DateTime exportTime, string mccPath)
        {
            manifest.AppendLine("HALO MCC TOOLBOX -- DIAGNOSTICS EXPORT");
            manifest.AppendLine($"Generated: {exportTime:yyyy-MM-dd HH:mm:ss}");
            manifest.AppendLine($"Configured MCC Path: {mccPath}");
            manifest.AppendLine($"Per-file size cap: {FormatSize(MaxDiagnosticExportBytes)}");
            manifest.AppendLine();
        }

        private static void AppendManifestProbe(StringBuilder manifest, string label, string rootPath)
        {
            manifest.AppendLine($"[PROBE] {label}");
            manifest.AppendLine($"Path: {rootPath}");
        }

        private static void AppendManifestInclude(StringBuilder manifest, string entryPath, long size, DateTime lastWrite, string source)
        {
            manifest.AppendLine($"[INCLUDED] {entryPath}");
            manifest.AppendLine($"  Source: {source}");
            manifest.AppendLine($"  Size: {FormatSize(size)}");
            manifest.AppendLine($"  Last Write: {lastWrite:yyyy-MM-dd HH:mm:ss}");
        }

        private static void AppendManifestMissing(StringBuilder manifest, string path)
        {
            manifest.AppendLine($"[MISSING] {path}");
            manifest.AppendLine();
        }

        private static void AppendManifestSkipped(StringBuilder manifest, string path, string reason)
        {
            manifest.AppendLine($"[SKIPPED] {path}");
            manifest.AppendLine($"  Reason: {reason}");
        }

        private static void AppendManifestError(StringBuilder manifest, string path, string error)
        {
            manifest.AppendLine($"[ERROR] {path}");
            manifest.AppendLine($"  Message: {error}");
        }

        private static void AppendManifestPrivacyNotes(StringBuilder manifest)
        {
            manifest.AppendLine();
            manifest.AppendLine("[EXCLUDED BY DEFAULT]");
            manifest.AppendLine("- MCC Saved\\webcache, mcc/logs, and temp_reports");
            manifest.AppendLine("- MCC carnagereports and gamecollections");
            manifest.AppendLine("- Steam userdata");
            manifest.AppendLine("- Steam logs not clearly tied to Halo MCC");
            manifest.AppendLine("- Generic caches unrelated to diagnostics");
        }

        private sealed record DiagnosticProbeRoot(
            string Label,
            string RootPath,
            string ZipRoot,
            Func<string, bool> Include,
            bool LimitToLatest = false);

        // ------------------------------------------
        // TOOL: Clean XBL credentials + webcache
        // ------------------------------------------
        private void BtnCleanCreds_Click(object sender, RoutedEventArgs e)
        {
            var result = ToolboxDialog.Show(
                "This will delete your stored Xbox Live credentials and MCC webcache files.\n\nMake sure MCC is closed before continuing.\n\nProceed?",
                "Confirm -- Halo MCC Toolbox",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                AppendLog("[INFO]", "Operation cancelled by user.", "#4A5A6A");
                return;
            }

            AppendLog("[RUN]", "Starting XBL credential + webcache cleanup...", "#FF6A00");
            SetStatus("Running cleanup...", "#FF6A00");
            BtnCleanCreds.IsEnabled = false;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    AppendLog("[STEP]", "Deleting Xbl credentials via cmdkey...", "#C8D8E8");

                    var psi = new ProcessStartInfo("cmd.exe")
                    {
                        Arguments = "/C cmdkey /list",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    string cmdkeyOutput;
                    using (var proc = Process.Start(psi)!)
                    {
                        cmdkeyOutput = proc.StandardOutput.ReadToEnd();
                        proc.WaitForExit();
                    }

                    foreach (var line in cmdkeyOutput.Split('\n'))
                    {
                        if (!line.Contains("Xbl", StringComparison.OrdinalIgnoreCase)) continue;
                        string? target = null;
                        foreach (var part in line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            if (part.StartsWith("LegacyGeneric:", StringComparison.OrdinalIgnoreCase) ||
                                part.ToLower().Contains("xbl"))
                            {
                                target = part.TrimEnd(':');
                                break;
                            }
                        }
                        if (target != null)
                        {
                            using var delProc = Process.Start(new ProcessStartInfo("cmdkey.exe")
                            {
                                Arguments = $"/delete:{target}",
                                UseShellExecute = false,
                                CreateNoWindow = true
                            });
                            delProc?.WaitForExit();
                            AppendLog("[CRED]", $"Deleted: {target}", "#39FF14");
                        }
                    }

                    // Run the original batch script too
                    var batchPath = Path.Combine(Path.GetTempPath(), "mcc_clean_temp.bat");
                    File.WriteAllText(batchPath, BuildCleanupBatch(), Encoding.ASCII);
                    using (var batchProc = Process.Start(new ProcessStartInfo("cmd.exe")
                    {
                        Arguments = $"/C \"{batchPath}\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    })!)
                    {
                        foreach (var ln in batchProc.StandardOutput.ReadToEnd().Split('\n'))
                            if (!string.IsNullOrWhiteSpace(ln))
                                AppendLog("[BAT]", ln.Trim(), "#C8D8E8");
                        batchProc.WaitForExit();
                    }
                    try { File.Delete(batchPath); } catch { }

                    // Delete webcache directly
                    var webcachePath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                        "AppData", "LocalLow", "MCC", "Saved", "webcache");

                    if (Directory.Exists(webcachePath))
                    {
                        var files = Directory.GetFiles(webcachePath);
                        int deleted = 0;
                        foreach (var f in files)
                        {
                            try { File.Delete(f); deleted++; }
                            catch (Exception ex)
                            { AppendLog("[WARN]", $"Could not delete {Path.GetFileName(f)}: {ex.Message}", "#FF6A00"); }
                        }
                        AppendLog("[STEP]", $"Webcache: deleted {deleted}/{files.Length} files.", "#39FF14");
                    }
                    else
                    {
                        AppendLog("[INFO]", $"Webcache folder not found: {webcachePath}", "#4A5A6A");
                    }

                    AppendLog("[DONE]", "Cleanup complete! Restart MCC and sign in again.", "#39FF14");
                    SetStatus("Cleanup complete.", "#39FF14");
                    Dispatcher.Invoke(() =>
                        ToolboxDialog.Show("Cleanup complete!\n\nXBL credentials and webcache have been cleared.\nRestart Halo MCC and sign in again.",
                            "Done -- Halo MCC Toolbox", MessageBoxButton.OK, MessageBoxImage.Information));
                }
                catch (Exception ex)
                {
                    AppendLog("[ERROR]", $"Cleanup failed: {ex.Message}", "#FF2D55");
                    SetStatus("Error during cleanup.", "#FF2D55");
                }
                finally
                {
                    Dispatcher.Invoke(() => BtnCleanCreds.IsEnabled = true);
                }
            });
        }

        private static string BuildCleanupBatch() =>
@"@echo off
echo Deleting Xbl credentials...
for /F ""tokens=1,2 delims= "" %%F in ('cmdkey /list ^| findstr Xbl') do cmdkey /delete %%G
echo Xbl credentials deleted.
echo Deleting webcache files...
del /q /f ""%userprofile%\AppData\LocalLow\MCC\Saved\webcache\*""
echo Webcache files deleted.
echo All tasks complete.
";

        // ------------------------------------------
        // TOOL: Repair EasyAntiCheat
        // ------------------------------------------
        private void BtnRepairEAC_Click(object sender, RoutedEventArgs e)
        {
            // Find EAC setup relative to the configured MCC path first,
            // then fall back to the Steam default.
            var mccBase      = TxtMccPath.Text.Trim();
            var eacInMcc     = Path.Combine(mccBase, "EasyAntiCheat", "EasyAntiCheat_EOS_Setup.exe");
            var eacDefault   = Path.Combine(
                App.DefaultMccPath,
                "EasyAntiCheat", "EasyAntiCheat_EOS_Setup.exe");

            var eacPath = File.Exists(eacInMcc)   ? eacInMcc
                        : File.Exists(eacDefault)  ? eacDefault
                        : null;

            if (eacPath == null)
            {
                var msg = "EasyAntiCheat EOS setup executable not found.\n\n" +
                          "Expected location:\n" +
                          $"{eacInMcc}\n\n" +
                          "Make sure your MCC installation path is set correctly.";
                AppendLog("[ERROR]", "EasyAntiCheat_EOS_Setup.exe not found.", "#FF2D55");
                ToolboxDialog.Show(msg, "EAC Not Found -- Halo MCC Toolbox",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = ToolboxDialog.Show(
                "This will launch the EasyAntiCheat EOS setup tool.\n\n" +
                "When it opens:\n" +
                "  1. Click  \"Repair Service\"\n" +
                "  2. Wait for it to complete\n" +
                "  3. Relaunch MCC\n\n" +
                "Make sure MCC is closed before continuing.\n\nProceed?",
                "Repair EAC -- Halo MCC Toolbox",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (confirm != MessageBoxResult.Yes)
            {
                AppendLog("[INFO]", "EAC repair cancelled by user.", "#4A5A6A");
                return;
            }

            AppendLog("[RUN]", $"Launching EAC EOS setup: {eacPath}", "#FF6A00");
            SetStatus("Launching EasyAntiCheat EOS repair...", "#FF6A00");

            try
            {
                // EAC setup requires elevation to repair the service
                var psi = new ProcessStartInfo(eacPath)
                {
                    UseShellExecute = true,   // needed for Verb = runas
                    Verb            = "runas" // request UAC elevation
                };
                Process.Start(psi);
                AppendLog("[INFO]", "EAC EOS setup launched. Follow the on-screen prompts to Repair Service.", "#39FF14");
                SetStatus("EAC EOS setup launched.", "#39FF14");
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Failed to launch EAC setup: {ex.Message}", "#FF2D55");
                SetStatus("Failed to launch EAC setup.", "#FF2D55");
                ToolboxDialog.Show($"Could not launch EasyAntiCheat EOS setup:\n\n{ex.Message}",
                    "Error -- Halo MCC Toolbox", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------------------------------
        // TOOL: Repair MCC audio device selection
        // ------------------------------------------
        private void BtnRepairAudioDevices_Click(object sender, RoutedEventArgs e)
        {
            var mccProcesses = MccProcessLocator.GetRuntimeProcesses(TxtMccPath.Text.Trim()).ToArray();
            var isMccRunning = mccProcesses.Length > 0;
            foreach (var process in mccProcesses)
                process.Dispose();

            if (isMccRunning)
            {
                AppendLog("[AUDIO]", "Repair blocked because MCC is running.", "#FF6A00");
                ToolboxDialog.Show(
                    "Close Halo: The Master Chief Collection before running this fix.",
                    "MCC Is Running -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var settingsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "AppData", "LocalLow", "MCC", "Saved", "Config", "WindowsNoEditor",
                "GameUserSettings.ini");

            if (!File.Exists(settingsPath))
            {
                AppendLog("[ERROR]", $"MCC settings file not found: {settingsPath}", "#FF2D55");
                ToolboxDialog.Show(
                    $"MCC's settings file was not found:\n\n{settingsPath}\n\nLaunch MCC once, close it, and try again.",
                    "Settings Not Found -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var confirm = ToolboxDialog.Show(
                "This will back up MCC's settings and clear its saved audio output device. " +
                "MCC will detect the current Windows output device the next time it launches.\n\n" +
                "No other MCC settings will be changed.\n\nProceed?",
                "Repair Audio Devices -- Halo MCC Toolbox",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (confirm != MessageBoxResult.Yes)
            {
                AppendLog("[AUDIO]", "Audio device repair cancelled by user.", "#4A5A6A");
                return;
            }

            BtnRepairAudioDevices.IsEnabled = false;
            SetStatus("Repairing MCC audio device settings...", "#FF6A00");

            var attributes = File.GetAttributes(settingsPath);
            var wasReadOnly = attributes.HasFlag(FileAttributes.ReadOnly);

            try
            {
                string settings;
                Encoding encoding;
                using (var reader = new StreamReader(settingsPath, Encoding.UTF8, true))
                {
                    settings = reader.ReadToEnd();
                    encoding = reader.CurrentEncoding;
                }

                var match = Regex.Match(settings, @"(?m)^AudioOutputDevice=.*$");
                if (!match.Success)
                    throw new InvalidDataException("AudioOutputDevice setting was not found.");

                if (match.Value == "AudioOutputDevice=")
                {
                    AppendLog("[AUDIO]", "MCC audio output device is already reset.", "#39FF14");
                    SetStatus("Audio device setting is already reset.", "#39FF14");
                    ToolboxDialog.Show(
                        "MCC's saved audio output device is already cleared. No changes were needed.",
                        "Repair Audio Devices -- Halo MCC Toolbox",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                var backupPath = settingsPath + ".audio-repair-" +
                                 DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".bak";
                File.Copy(settingsPath, backupPath, false);

                if (wasReadOnly)
                    File.SetAttributes(settingsPath, attributes & ~FileAttributes.ReadOnly);

                var repairedSettings = settings.Remove(match.Index, match.Length)
                                               .Insert(match.Index, "AudioOutputDevice=");
                File.WriteAllText(settingsPath, repairedSettings, encoding);

                AppendLog("[AUDIO]", "Cleared MCC's saved audio output device.", "#39FF14");
                AppendLog("[BACKUP]", backupPath, "#4A5A6A");
                SetStatus("MCC audio devices repaired.", "#39FF14");
                ToolboxDialog.Show(
                    "Audio device repair complete.\n\nLaunch MCC to test the fix.",
                    "Done -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Audio device repair failed: {ex.Message}", "#FF2D55");
                SetStatus("Audio device repair failed.", "#FF2D55");
                ToolboxDialog.Show(
                    $"Could not repair MCC's audio device setting:\n\n{ex.Message}",
                    "Error -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                if (wasReadOnly && File.Exists(settingsPath))
                {
                    try { File.SetAttributes(settingsPath, File.GetAttributes(settingsPath) | FileAttributes.ReadOnly); }
                    catch (Exception ex)
                    { AppendLog("[WARN]", $"Could not restore read-only state: {ex.Message}", "#FF6A00"); }
                }

                BtnRepairAudioDevices.IsEnabled = true;
            }
        }

        private void StartRejoinCrashWatcher()
        {
            PollMccProcessesForRejoinCrashRestore();
            _rejoinCrashWatchTimer.Start();
            RejoinFixDiagnostics.Info("restore", "MCC process watcher started; all exits during a saved match will arm crash restore.");
        }

        private void StopRejoinCrashWatcher()
        {
            _rejoinCrashWatchTimer.Stop();

            lock (_rejoinCrashWatchLock)
            {
                foreach (var process in _rejoinWatchedMccProcesses.Values)
                {
                    try { process.Exited -= MccProcess_Exited; } catch { }
                    try { process.Dispose(); } catch { }
                }

                _rejoinWatchedMccProcesses.Clear();
            }
        }

        private void PollMccProcessesForRejoinCrashRestore()
        {
            if (!_rejoinProxy.IsRunning)
            {
                StopRejoinCrashWatcher();
                UpdateRejoinFixUi();
                return;
            }

            Process[] processes;
            try
            {
                processes = MccProcessLocator.GetRuntimeProcesses(TxtMccPath.Text.Trim()).ToArray();
            }
            catch (Exception ex)
            {
                RejoinFixDiagnostics.Warn("restore", $"Failed to poll MCC process state: {ex.Message}");
                return;
            }

            var liveProcessIds = new HashSet<int>(processes.Select(process => process.Id));
            List<int> vanishedProcessIds = new();

            lock (_rejoinCrashWatchLock)
            {
                foreach (int watchedPid in _rejoinWatchedMccProcesses.Keys.ToList())
                {
                    if (liveProcessIds.Contains(watchedPid))
                        continue;

                    vanishedProcessIds.Add(watchedPid);
                    if (_rejoinWatchedMccProcesses.Remove(watchedPid, out var vanishedProcess))
                    {
                        try { vanishedProcess.Exited -= MccProcess_Exited; } catch { }
                        try { vanishedProcess.Dispose(); } catch { }
                    }
                }
            }

            foreach (int vanishedPid in vanishedProcessIds)
            {
                RejoinFixDiagnostics.Warn(
                    "restore",
                    $"MCC process vanished before the Exited event fired pid={vanishedPid}; treating as unexpected exit.");
                Dispatcher.InvokeAsync(() => ArmRejoinCrashRestoreFromMccExit(vanishedPid, null));
            }

            foreach (var process in processes)
            {
                bool keepProcess = false;
                try
                {
                    lock (_rejoinCrashWatchLock)
                    {
                        if (_rejoinWatchedMccProcesses.ContainsKey(process.Id))
                            continue;

                        process.EnableRaisingEvents = true;
                        process.Exited += MccProcess_Exited;
                        _rejoinWatchedMccProcesses[process.Id] = process;
                        keepProcess = true;
                    }

                    if (process.HasExited)
                        MccProcess_Exited(process, EventArgs.Empty);
                }
                catch (Exception ex)
                {
                    RejoinFixDiagnostics.Warn("restore", $"Failed to watch MCC process: {ex.Message}");
                }
                finally
                {
                    if (!keepProcess)
                    {
                        try { process.Dispose(); } catch { }
                    }
                }
            }
        }

        private void MccProcess_Exited(object? sender, EventArgs e)
        {
            if (sender is not Process process)
                return;

            int pid = 0;
            int exitCode = 0;
            bool hasExitCode = false;

            try { pid = process.Id; } catch { }
            try
            {
                exitCode = process.ExitCode;
                hasExitCode = true;
            }
            catch
            {
                // If Windows will not give us an exit code, treat disappearance as abnormal.
            }

            lock (_rejoinCrashWatchLock)
            {
                if (pid != 0)
                    _rejoinWatchedMccProcesses.Remove(pid);
            }

            try { process.Exited -= MccProcess_Exited; } catch { }
            try { process.Dispose(); } catch { }

            RejoinFixDiagnostics.Warn(
                "restore",
                $"Observed MCC process exit{FormatExitCode(pid, hasExitCode ? exitCode : null)}; evaluating saved match for restore.");

            Dispatcher.InvokeAsync(() => ArmRejoinCrashRestoreFromMccExit(pid, hasExitCode ? exitCode : null));
        }

        private void ArmRejoinCrashRestoreFromMccExit(int pid, int? exitCode)
        {
            if (!_rejoinProxy.IsRunning)
                return;

            var matchSession = TryLoadSavedRejoinMatchSession();
            if (matchSession is null)
            {
            RejoinFixDiagnostics.Warn(
                "restore",
                $"MCC exited{FormatExitCode(pid, exitCode)}, but no saved matchmaking session was available.");
                return;
            }

            _rejoinProxy.SetPendingCrashRestore(matchSession);
            RejoinFixDiagnostics.Warn(
                "restore",
                $"MCC exited{FormatExitCode(pid, exitCode)}; armed crash restore for {matchSession.TemplateName}/{matchSession.SessionShort}.");
            AppendLog("[REJOIN]", $"MCC exited; armed crash restore for {matchSession.SessionShort}.", "#FF6A00");
            SetStatus("Rejoin crash restore armed.", "#FF6A00");
            UpdateRejoinFixUi();
        }

        private static string FormatExitCode(int pid, int? exitCode)
        {
            string pidPart = pid == 0 ? "" : $" pid={pid}";
            string codePart = exitCode.HasValue ? $" exit={exitCode.Value}" : " exit=unknown";
            return $"{pidPart}{codePart}";
        }

        private static SavedHandleInfo? TryLoadSavedRejoinMatchSession()
        {
            try
            {
                if (!File.Exists(RejoinFixPaths.LastMatchSessionFile))
                    return null;

                var json = File.ReadAllText(RejoinFixPaths.LastMatchSessionFile);
                return JsonSerializer.Deserialize<SavedHandleInfo>(json);
            }
            catch (Exception ex)
            {
                RejoinFixDiagnostics.Warn("restore", $"Failed to load saved matchmaking session: {ex.Message}");
                return null;
            }
        }

        private void UpdateRejoinFixUi()
        {
            bool isRunning = _rejoinProxy.IsRunning;
            bool hasSavedState = File.Exists(RejoinFixPaths.LastHandleFile)
                || File.Exists(RejoinFixPaths.LastMatchSessionFile)
                || File.Exists(RejoinFixPaths.LastSquadStateFile)
                || File.Exists(RejoinFixPaths.LastGameServerFile);
            string gamertagSuffix = string.IsNullOrWhiteSpace(_rejoinProxy.CurrentPlayerGamertag)
                ? ""
                : $" ({_rejoinProxy.CurrentPlayerGamertag})";
            string modeLabel = _rejoinProxy.CurrentRejoinModeLabel;
            int squadMemberCount = _rejoinProxy.CurrentSquadMemberCount;
            string modeSuffix = squadMemberCount > 0
                ? $" ({squadMemberCount} member{(squadMemberCount == 1 ? "" : "s")})"
                : "";

            BtnRejoinFix.Content = isRunning ? "STOP PROXY" : "START PROXY";
            PopulationStartProxyButton.Content = isRunning ? "PROXY RUNNING" : "START PROXY";
            PopulationStartProxyButton.IsEnabled = !isRunning;
            HomeProxyToggleButton.Content = isRunning ? "STOP PROXY" : "START PROXY";
            SidebarFeaturesToggleButton.Content = isRunning ? "STOP" : "START";
            SidebarFeaturesToggleButton.ToolTip = isRunning
                ? "Stop the MCC data proxy without leaving this page"
                : "Start the MCC data proxy without leaving this page";
            SetToolboxStatusText(
                HomeFeaturesStatus,
                isRunning ? "● ON" : "● OFF",
                isRunning ? "#39FF14" : "#71869A");
            TxtRejoinRecoveryStatus.Text = isRunning
                ? "● REJOIN RECOVERY · CORE · ACTIVE"
                : "○ REJOIN RECOVERY · CORE · STARTS WITH PROXY";
            TxtRejoinRecoveryStatus.Foreground = Brush(isRunning ? "#39FF14" : "#4A5A6A");
            bool hasPlayerVisibleOverlay =
                (_networkStatsOverlayEnabled && !_networkStatsObsOnly) ||
                (_matchmakingWaitOverlayEnabled && !_matchmakingWaitObsOnly) ||
                (_obsBrowserOverlaySessionStatsEnabled && !_sessionStatsObsOnly) ||
                (_combinedNetworkSessionOverlayEnabled && !_combinedNetworkSessionObsOnly);
            BtnNetworkStatsOverlayMove.Visibility =
                hasPlayerVisibleOverlay
                ? Visibility.Visible
                : Visibility.Collapsed;
            BtnNetworkStatsOverlayMove.IsEnabled = hasPlayerVisibleOverlay;

            if (isRunning && _rejoinWinHttpManualNeeded)
            {
                TxtRejoinFixStatus.Text = $"● ACTIVE{gamertagSuffix} - MCC capture may still need admin proxy approval";
                TxtRejoinFixStatus.Foreground = Brush("#FF6A00");
            }
            else if (isRunning)
            {
                TxtRejoinFixStatus.Text = $"● ACTIVE{gamertagSuffix} - MCC companion features are available";
                TxtRejoinFixStatus.Foreground = Brush("#39FF14");
            }
            else if (hasSavedState)
            {
                TxtRejoinFixStatus.Text = $"○ STOPPED{gamertagSuffix} - saved capture files are available for diagnostics";
                TxtRejoinFixStatus.Foreground = Brush("#C8D8E8");
            }
            else
            {
                TxtRejoinFixStatus.Text = $"○ STOPPED{gamertagSuffix}";
                TxtRejoinFixStatus.Foreground = Brush("#4A5A6A");
            }

            HomeProxyStatusText.Text = TxtRejoinFixStatus.Text;
            HomeProxyStatusText.Foreground = TxtRejoinFixStatus.Foreground;

            if (squadMemberCount > 0)
            {
                TxtRejoinFixMode.Visibility = Visibility.Visible;
                TxtRejoinFixMode.Text = $"PATH: {modeLabel}{modeSuffix}";
                TxtRejoinFixMode.Foreground = modeLabel switch
                {
                    "PARTY" => Brush("#00C8FF"),
                    "SOLO" => Brush("#39FF14"),
                    _ => Brush("#C8D8E8")
                };
            }
            else
            {
                TxtRejoinFixMode.Text = "";
                TxtRejoinFixMode.Visibility = Visibility.Collapsed;
            }

            HomeProxyModeText.Text = TxtRejoinFixMode.Text;
            HomeProxyModeText.Foreground = TxtRejoinFixMode.Foreground;
            HomeProxyModeText.Visibility = TxtRejoinFixMode.Visibility;

            UpdateRejoinFirewallOptionAvailability(isRunning);
            UpdateRejoinFirewallStatus();
            StatsUpdateCurrentLobbyVisibility(isRunning);
            UpdateToolboxStatus();
            UpdateOverlayPreviews();
        }

        private void StatsUpdateCurrentLobbyVisibility(bool isRejoinFixRunning)
        {
            if (isRejoinFixRunning)
            {
                StatsCurrentLobbyHeader.Visibility = Visibility.Visible;
                StatsCurrentLobbyList.Visibility = Visibility.Visible;
                StatsCurrentLobbySplitter.Visibility = Visibility.Visible;
                StatsCurrentLobbyHeaderRow.Height = GridLength.Auto;
                StatsCurrentLobbyListRow.Height = new GridLength(260);
                StatsCurrentLobbyListRow.MinHeight = 170;
                StatsCurrentLobbySplitterRow.Height = new GridLength(5);
                return;
            }

            StatsCurrentLobbyHeader.Visibility = Visibility.Collapsed;
            StatsCurrentLobbyList.Visibility = Visibility.Collapsed;
            StatsCurrentLobbySplitter.Visibility = Visibility.Collapsed;
            StatsCurrentLobbyHeaderRow.Height = new GridLength(0);
            StatsCurrentLobbyListRow.Height = new GridLength(0);
            StatsCurrentLobbyListRow.MinHeight = 0;
            StatsCurrentLobbySplitterRow.Height = new GridLength(0);
        }

        private void UpdateRejoinFirewallOptionAvailability(bool isRejoinFixRunning)
        {
            if (IsMicrosoftStoreInstallation)
            {
                ChkRejoinFixFirewall.IsEnabled = false;
                ChkRejoinFixFirewallMatchmaking.IsEnabled = false;
                ChkRejoinFixFirewall.Visibility = Visibility.Collapsed;
                ChkRejoinFixFirewallMatchmaking.Visibility = Visibility.Collapsed;
                ChkRejoinFixFirewall.ToolTip = "Steam-only feature";
                ChkRejoinFixFirewallMatchmaking.ToolTip = "Steam-only feature";
                return;
            }

            ChkRejoinFixFirewall.Visibility = Visibility.Visible;
            ChkRejoinFixFirewallMatchmaking.Visibility = Visibility.Visible;
            ChkRejoinFixFirewall.IsEnabled = true;
            ChkRejoinFixFirewallMatchmaking.IsEnabled = true;
            ChkRejoinFixFirewall.Content = RejoinFirewallCampaignLabel;
            ChkRejoinFixFirewallMatchmaking.Content = RejoinFirewallMatchmakingLabel;
            string? pendingReason = isRejoinFixRunning
                ? null
                : "Selection saved. This activates when the MCC data proxy starts.";
            ChkRejoinFixFirewall.ToolTip = pendingReason;
            ChkRejoinFixFirewallMatchmaking.ToolTip = pendingReason;
        }

        private void UpdateRejoinFirewallStatus(string? overrideText = null, string? overrideColor = null)
        {
            if (IsMicrosoftStoreInstallation)
            {
                TxtRejoinFirewallStatus.Text = "FIREWALL: STEAM ONLY - not used for Microsoft Store MCC";
                TxtRejoinFirewallStatus.Foreground = Brush("#4A5A6A");
                PublishNetworkFirewallStatus();
                return;
            }

            if (!string.IsNullOrWhiteSpace(overrideText))
            {
                TxtRejoinFirewallStatus.Text = overrideText;
                TxtRejoinFirewallStatus.Foreground = Brush(overrideColor ?? "#C8D8E8");
                PublishNetworkFirewallStatus();
                return;
            }

            if (_steamFirewallAutoSuspendedForCrashRestore)
            {
                TxtRejoinFirewallStatus.Text = "FIREWALL: REJOIN RESTORE - ports are open for crash rejoin";
                TxtRejoinFirewallStatus.Foreground = Brush("#00C8FF");
                PublishNetworkFirewallStatus();
                return;
            }

            if (ChkRejoinFixFirewall.IsChecked == true)
            {
                if (_rejoinCampaignFirewallApplying)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: CAMPAIGN PENDING - applying port block";
                    TxtRejoinFirewallStatus.Foreground = Brush("#FF6A00");
                }
                else if (_rejoinCampaignFirewallEnabled)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: CAMPAIGN ON - port 3478 is blocked; invites will not function";
                    TxtRejoinFirewallStatus.Foreground = Brush("#39FF14");
                }
                else
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: CAMPAIGN PENDING - waiting to apply port block";
                    TxtRejoinFirewallStatus.Foreground = Brush("#FF6A00");
                }
            }
            else if (ChkRejoinFixFirewallMatchmaking.IsChecked == true)
            {
                if (_steamFirewallAutoPaused)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: MATCHMAKING PAUSED - ports are open while MCC searches/connects";
                    TxtRejoinFirewallStatus.Foreground = Brush("#00C8FF");
                }
                else if (_steamFirewallAutoEnabled && _steamFirewallUiState == SteamFirewallState.Enabled)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: MATCHMAKING ON - ports 3478 and 4379 are blocked until matchmaking traffic is detected";
                    TxtRejoinFirewallStatus.Foreground = Brush("#39FF14");
                }
                else if (_steamFirewallAutoEnabled)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: MATCHMAKING PENDING - waiting to apply port block";
                    TxtRejoinFirewallStatus.Foreground = Brush("#FF6A00");
                }
                else
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: MATCHMAKING OFF";
                    TxtRejoinFirewallStatus.Foreground = Brush("#4A5A6A");
                }
            }
            else
            {
                if (_steamFirewallUiState == SteamFirewallState.Enabled)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: ON - rules are enabled but no Rejoin firewall mode is selected";
                    TxtRejoinFirewallStatus.Foreground = Brush("#FF6A00");
                    PublishNetworkFirewallStatus();
                    return;
                }

                if (_steamFirewallUiState == SteamFirewallState.Partial)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: PARTIAL - rules are mixed; restart Rejoin Fix to repair";
                    TxtRejoinFirewallStatus.Foreground = Brush("#FF6A00");
                    PublishNetworkFirewallStatus();
                    return;
                }

                TxtRejoinFirewallStatus.Text = _steamFirewallRulesPrepared
                    ? "FIREWALL: READY - rules installed, currently off"
                    : "FIREWALL: OFF - rules will be prepared when Rejoin Fix starts as admin";
                TxtRejoinFirewallStatus.Foreground = Brush(_steamFirewallRulesPrepared ? "#C8D8E8" : "#4A5A6A");
            }

            PublishNetworkFirewallStatus();
        }

        private void PublishNetworkFirewallStatus()
        {
            UpdateFirewallPageStatus();
            if (_mainWindowInitialized)
                PublishObsOverlaySnapshot();
        }

        private void UpdateFirewallPageStatus()
        {
            if (FirewallHeroStatusText is null)
                return;

            bool isSteam = !IsMicrosoftStoreInstallation;
            bool proxyRunning = _rejoinProxy.IsRunning;
            bool campaignSelected = ChkRejoinFixFirewall.IsChecked == true;
            bool matchmakingSelected = ChkRejoinFixFirewallMatchmaking.IsChecked == true;
            bool modeSelected = campaignSelected || matchmakingSelected;
            bool runtimeActive = campaignSelected
                ? _rejoinCampaignFirewallEnabled
                : matchmakingSelected && _steamFirewallAutoEnabled;

            string heroText;
            string heroColor;
            if (!isSteam)
            {
                heroText = "FIREWALL FIX IS UNAVAILABLE";
                heroColor = "#71869A";
            }
            else if (runtimeActive)
            {
                heroText = "FIREWALL FIX IS ACTIVE";
                heroColor = "#39FF14";
            }
            else if (modeSelected)
            {
                heroText = proxyRunning ? "FIREWALL FIX IS STARTING" : "FIREWALL FIX IS ARMED";
                heroColor = proxyRunning ? "#FF6A00" : "#00C8FF";
            }
            else
            {
                heroText = "FIREWALL FIX IS DISABLED";
                heroColor = "#C8D8E8";
            }

            SetToolboxStatusText(FirewallHeroStatusText, heroText, heroColor);
            string ports = campaignSelected ? "PORT 3478" : "PORTS 3478 + 4379";
            string rules = _steamFirewallRulesPrepared ? "RULES READY" : "RULES NOT INSTALLED";
            FirewallHeroMetaText.Text = isSteam
                ? $"MCC ONLY · STEAM · {ports} · {rules}"
                : "STEAM INSTALLATION REQUIRED";

            string rulesText;
            string rulesColor;
            switch (_steamFirewallUiState)
            {
                case SteamFirewallState.Enabled:
                    rulesText = campaignSelected ? "ENABLED · PORT 3478 BLOCKED" : "ENABLED · PORTS BLOCKED";
                    rulesColor = "#39FF14";
                    break;
                case SteamFirewallState.Partial:
                    rulesText = "PARTIAL · REPAIR REQUIRED";
                    rulesColor = "#FF6A00";
                    break;
                case SteamFirewallState.Unknown:
                    rulesText = "UNKNOWN · CHECK LOG";
                    rulesColor = "#FF6A00";
                    break;
                default:
                    rulesText = "DISABLED · PORTS OPEN";
                    rulesColor = "#71869A";
                    break;
            }
            SetToolboxStatusText(FirewallRulesStatusText, rulesText, rulesColor);

            string automationText;
            string automationColor;
            if (!modeSelected)
            {
                automationText = "DISABLED";
                automationColor = "#71869A";
            }
            else if (!proxyRunning)
            {
                automationText = "ARMED · WAITING FOR PROXY";
                automationColor = "#00C8FF";
            }
            else if (campaignSelected)
            {
                automationText = _rejoinCampaignFirewallEnabled ? "CAMPAIGN · ACTIVE" : "CAMPAIGN · STARTING";
                automationColor = _rejoinCampaignFirewallEnabled ? "#39FF14" : "#FF6A00";
            }
            else if (_steamFirewallAutoSuspendedForCrashRestore)
            {
                automationText = "PAUSED · REJOIN RESTORE";
                automationColor = "#00C8FF";
            }
            else if (_steamFirewallAutoPaused)
            {
                automationText = "PAUSED · MATCHMAKING";
                automationColor = "#00C8FF";
            }
            else if (_steamFirewallAutoEnabled)
            {
                automationText = "MATCHMAKING · MONITORING";
                automationColor = "#39FF14";
            }
            else
            {
                automationText = "MATCHMAKING · STARTING";
                automationColor = "#FF6A00";
            }
            SetToolboxStatusText(FirewallAutomationStatusText, automationText, automationColor);
            SetToolboxStatusText(FirewallProxyStatusText, proxyRunning ? "RUNNING" : "STOPPED", proxyRunning ? "#39FF14" : "#71869A");
            SetToolboxStatusText(FirewallRejoinStatusText, proxyRunning ? "ACTIVE" : "STANDBY", proxyRunning ? "#39FF14" : "#71869A");
            SetToolboxStatusText(
                FirewallAdminStatusText,
                IsRunningAsAdministrator() ? "READY" : "PROMPT ON ENABLE",
                IsRunningAsAdministrator() ? "#39FF14" : "#00C8FF");

            FirewallFixToggleButton.Content = runtimeActive ? "DISABLE FIX" : "ENABLE FIX";
            FirewallFixToggleButton.IsEnabled = isSteam && !_firewallFixActionPending;
        }

        private async void FirewallFixToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsMicrosoftStoreInstallation || _firewallFixActionPending)
                return;

            _firewallFixActionPending = true;
            UpdateFirewallPageStatus();
            try
            {
                bool campaignSelected = ChkRejoinFixFirewall.IsChecked == true;
                bool matchmakingSelected = ChkRejoinFixFirewallMatchmaking.IsChecked == true;
                bool runtimeActive = campaignSelected
                    ? _rejoinCampaignFirewallEnabled
                    : matchmakingSelected && _steamFirewallAutoEnabled;

                if (runtimeActive)
                {
                    SetRejoinFirewallCheckbox(ChkRejoinFixFirewall, false);
                    SetRejoinFirewallCheckbox(ChkRejoinFixFirewallMatchmaking, false);
                    App.SaveRejoinFirewallMode("Disabled");
                    DisableSteamFirewallAutoMode(logStatus: false);
                    await DisableRejoinFirewallRulesAsync(logStatus: true);
                    return;
                }

                if (!campaignSelected && !matchmakingSelected)
                {
                    SetRejoinFirewallCheckbox(ChkRejoinFixFirewallMatchmaking, true);
                    App.SaveRejoinFirewallMode("Matchmaking");
                    matchmakingSelected = true;
                }

                if (!_rejoinProxy.IsRunning)
                {
                    await EnsureCompanionServicesRunningAsync("Firewall Fix");
                    return;
                }

                if (campaignSelected)
                {
                    DisableSteamFirewallAutoMode(logStatus: false);
                    await ApplyRejoinFirewallOptionAsync();
                }
                else if (matchmakingSelected)
                {
                    await EnableSteamFirewallAutoModeAsync(ensureObserverRunning: true);
                }
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Firewall Fix action failed: {ex.Message}", "#FF2D55");
                SetStatus("Firewall Fix action failed.", "#FF2D55");
                UpdateRejoinFirewallStatus("FIREWALL: ACTION FAILED - check the log for details", "#FF2D55");
            }
            finally
            {
                _firewallFixActionPending = false;
                UpdateRejoinFixUi();
                UpdateFirewallPageStatus();
            }
        }

        private async Task RefreshSteamFirewallUiAsync()
        {
            if (!SteamFirewallFeatureEnabled)
            {
                InitializeSteamFirewallFeatureState();
                return;
            }

            UpdateSteamFirewallUi(SteamFirewallState.Unknown);

            try
            {
                var state = await GetSteamFirewallStateAsync();
                UpdateSteamFirewallUi(state);
            }
            catch (Exception ex)
            {
                TxtSteamFirewallStatus.Text = $"UNKNOWN - could not read firewall status: {ex.Message}";
                TxtSteamFirewallStatus.Foreground = Brush("#FF6A00");
                BtnSteamFirewallFix.Content = "RETRY";
                BtnSteamFirewallFix.IsEnabled = true;
            }
        }

        private void InitializeSteamFirewallFeatureState()
        {
            if (SteamFirewallFeatureEnabled)
            {
                SteamFirewallCard.Visibility = _toolsPage == ToolsPage.Firewall
                    ? Visibility.Visible
                    : Visibility.Collapsed;
                ChkSteamFirewallAuto.IsEnabled = true;
                UpdateSteamFirewallUi(SteamFirewallState.Unknown);
                return;
            }

            _steamFirewallAutoEnabled = false;
            _steamFirewallAutoPaused = false;
            _steamFirewallAutoTimer.Stop();
            _steamFirewallUiState = SteamFirewallState.Disabled;

            SteamFirewallCard.Visibility = Visibility.Collapsed;
            ChkSteamFirewallAuto.IsChecked = false;
            ChkSteamFirewallAuto.IsEnabled = false;
            BtnSteamFirewallFix.Content = "DISABLED";
            BtnSteamFirewallFix.IsEnabled = false;
            TxtSteamFirewallStatus.Text = "DISABLED - MCC P2P Firewall Fix is unavailable in this build";
            TxtSteamFirewallStatus.Foreground = Brush("#4A5A6A");
        }

        private async Task SynchronizeStartupFirewallStateAsync()
        {
            try
            {
                var actualState = await GetSteamFirewallStateAsync();
                _steamFirewallRulesPrepared = actualState is SteamFirewallState.Disabled or SteamFirewallState.Enabled or SteamFirewallState.Partial;
                SetSteamFirewallRuntimeState(actualState);

                if (actualState is not (SteamFirewallState.Enabled or SteamFirewallState.Partial))
                    return;

                AppendLog("[FIREWALL]", "Found leftover or mixed MCC P2P firewall rules from a previous run; disabling them for a clean start.", "#FF6A00");
                await DisableRejoinFirewallRulesAsync(logStatus: false);
                _steamFirewallRulesPrepared = true;
                AppendLog("[FIREWALL]", "Startup firewall cleanup complete. Rules are installed and disabled.", "#39FF14");
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                AppendLog("[FIREWALL]", "Startup firewall cleanup was cancelled at the administrator prompt; run Firewall Check if MCC connectivity looks wrong.", "#FF6A00");
                SetSteamFirewallRuntimeState(SteamFirewallState.Unknown);
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Startup firewall state check failed: {ex.Message}", "#FF2D55");
                SetSteamFirewallRuntimeState(SteamFirewallState.Unknown);
            }
        }

        private void UpdateSteamFirewallUi(SteamFirewallState state)
        {
            if (!SteamFirewallFeatureEnabled)
            {
                InitializeSteamFirewallFeatureState();
                return;
            }

            _steamFirewallUiState = state;

            switch (state)
            {
                case SteamFirewallState.Enabled:
                    TxtSteamFirewallStatus.Text = "ON - ports 3478 and 4379 are blocked for Halo MCC only";
                    TxtSteamFirewallStatus.Foreground = Brush("#39FF14");
                    BtnSteamFirewallFix.Content = "DISABLE";
                    BtnSteamFirewallFix.IsEnabled = true;
                    break;
                case SteamFirewallState.Disabled:
                    TxtSteamFirewallStatus.Text = "OFF - MCC-only firewall rules exist but are disabled";
                    TxtSteamFirewallStatus.Foreground = Brush("#C8D8E8");
                    BtnSteamFirewallFix.Content = "ENABLE";
                    BtnSteamFirewallFix.IsEnabled = true;
                    break;
                case SteamFirewallState.Missing:
                    TxtSteamFirewallStatus.Text = "OFF - MCC-only firewall rules have not been created yet";
                    TxtSteamFirewallStatus.Foreground = Brush("#4A5A6A");
                    BtnSteamFirewallFix.Content = "ENABLE";
                    BtnSteamFirewallFix.IsEnabled = true;
                    break;
                case SteamFirewallState.Partial:
                    TxtSteamFirewallStatus.Text = "PARTIAL - MCC-only firewall rules are incomplete or mixed; click enable to repair";
                    TxtSteamFirewallStatus.Foreground = Brush("#FF6A00");
                    BtnSteamFirewallFix.Content = "ENABLE";
                    BtnSteamFirewallFix.IsEnabled = true;
                    break;
                default:
                    TxtSteamFirewallStatus.Text = "CHECKING - reading firewall status";
                    TxtSteamFirewallStatus.Foreground = Brush("#4A5A6A");
                    BtnSteamFirewallFix.Content = "CHECKING";
                    BtnSteamFirewallFix.IsEnabled = false;
                    break;
            }
        }

        private static SteamFirewallState LoadSteamFirewallUiState()
        {
            try
            {
                if (!File.Exists(SteamFirewallStateFile))
                    return SteamFirewallState.Missing;

                string value = File.ReadAllText(SteamFirewallStateFile).Trim();
                return string.Equals(value, "Enabled", StringComparison.OrdinalIgnoreCase)
                    ? SteamFirewallState.Enabled
                    : SteamFirewallState.Disabled;
            }
            catch
            {
                return SteamFirewallState.Missing;
            }
        }

        private static void SaveSteamFirewallUiState(bool enabled)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SteamFirewallStateFile)!);
                File.WriteAllText(SteamFirewallStateFile, enabled ? "Enabled" : "Disabled");
            }
            catch
            {
                // UI hint only; firewall state is still changed by the elevated command.
            }
        }

        private async void ChkSteamFirewallAuto_Checked(object sender, RoutedEventArgs e)
        {
            if (!SteamFirewallFeatureEnabled)
            {
                InitializeSteamFirewallFeatureState();
                return;
            }

            if (!IsRunningAsAdministrator())
            {
                ChkSteamFirewallAuto.IsChecked = false;
                ToolboxDialog.Show(
                    "Auto mode needs the Toolbox to run as Administrator so it can toggle MCC firewall rules without interrupting matchmaking.\n\nThe Toolbox will relaunch as Administrator now.",
                    "MCC P2P Firewall Auto -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? "",
                        UseShellExecute = true,
                        Verb = "runas"
                    });
                    Close();
                }
                catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
                {
                    AppendLog("[INFO]", "Firewall Auto mode cancelled at administrator prompt.", "#4A5A6A");
                    SetStatus("Firewall Auto mode cancelled.", "#4A5A6A");
                }
                catch (Exception ex)
                {
                    AppendLog("[ERROR]", $"Could not relaunch as Administrator: {ex.Message}", "#FF2D55");
                    SetStatus("Could not enable Firewall Auto mode.", "#FF2D55");
                }

                return;
            }

            _steamFirewallAutoEnabled = true;
            _steamFirewallAutoTimer.Start();
            AppendLog("[FIREWALL]", "Auto protection enabled. Firewall, proxy observer, and crash restore are armed.", "#00C8FF");
            SetStatus("Firewall Auto protection enabled.", "#00C8FF");

            try
            {
                string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
                await SetSteamFirewallEnabledAsync(true, mccExePath);
                SaveSteamFirewallUiState(true);
                UpdateSteamFirewallUi(SteamFirewallState.Enabled);
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Could not enable firewall protection automatically: {ex.Message}", "#FF2D55");
            }

            await EnsureRejoinObserverRunningAsync();
        }

        private async void ChkRejoinFixFirewall_Checked(object sender, RoutedEventArgs e)
        {
            if (IsMicrosoftStoreInstallation)
            {
                SetRejoinFirewallCheckbox(ChkRejoinFixFirewall, false);
                return;
            }

            if (!_mainWindowInitialized || _rejoinFirewallCheckChanging)
                return;

            SetRejoinFirewallCheckbox(ChkRejoinFixFirewallMatchmaking, false);
            App.SaveRejoinFirewallMode("Campaign");
            if (!_rejoinProxy.IsRunning)
            {
                UpdateRejoinFixUi();
                return;
            }
            DisableSteamFirewallAutoMode(logStatus: false);
            _rejoinCampaignFirewallEnabled = false;
            UpdateRejoinFirewallStatus();

            try
            {
                await ApplyRejoinFirewallOptionAsync();
            }
            catch (Exception ex)
            {
                _rejoinCampaignFirewallApplying = false;
                _rejoinCampaignFirewallEnabled = false;
                AppendLog("[ERROR]", $"Could not enable Firewall Fix (Campaign): {ex.Message}", "#FF2D55");
                SetStatus("Firewall Fix (Campaign) failed.", "#FF2D55");
                SetRejoinFirewallCheckbox(ChkRejoinFixFirewall, false);
                UpdateRejoinFirewallStatus();
            }
        }

        private void ChkRejoinFixFirewall_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!_mainWindowInitialized || _rejoinFirewallCheckChanging)
                return;

            if (ChkRejoinFixFirewallMatchmaking.IsChecked != true)
                App.SaveRejoinFirewallMode("Disabled");
            _rejoinCampaignFirewallEnabled = false;
            if (!_rejoinProxy.IsRunning)
            {
                UpdateRejoinFixUi();
                return;
            }

            _ = DisableRejoinFirewallRulesAsync(logStatus: true);
        }

        private async void ChkRejoinFixFirewallMatchmaking_Checked(object sender, RoutedEventArgs e)
        {
            if (IsMicrosoftStoreInstallation)
            {
                SetRejoinFirewallCheckbox(ChkRejoinFixFirewallMatchmaking, false);
                return;
            }

            if (!_mainWindowInitialized || _rejoinFirewallCheckChanging)
                return;

            SetRejoinFirewallCheckbox(ChkRejoinFixFirewall, false);
            App.SaveRejoinFirewallMode("Matchmaking");
            if (!_rejoinProxy.IsRunning)
            {
                UpdateRejoinFixUi();
                return;
            }
            _rejoinCampaignFirewallApplying = false;
            _rejoinCampaignFirewallEnabled = false;
            UpdateRejoinFirewallStatus();

            await EnableSteamFirewallAutoModeAsync(ensureObserverRunning: true);
        }

        private void ChkSteamFirewallAuto_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!SteamFirewallFeatureEnabled)
            {
                InitializeSteamFirewallFeatureState();
                return;
            }

            _steamFirewallAutoEnabled = false;
            _steamFirewallAutoPaused = false;
            _steamFirewallAutoTimer.Stop();
            AppendLog("[FIREWALL]", "Auto protection disabled. Manual firewall state is left as-is.", "#C8D8E8");
            SetStatus("Firewall Auto protection disabled.", "#C8D8E8");
        }

        private void ChkRejoinFixFirewallMatchmaking_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!_mainWindowInitialized || _rejoinFirewallCheckChanging)
                return;

            if (ChkRejoinFixFirewall.IsChecked != true)
                App.SaveRejoinFirewallMode("Disabled");
            if (!_rejoinProxy.IsRunning)
            {
                UpdateRejoinFixUi();
                return;
            }

            DisableSteamFirewallAutoMode(logStatus: true);
        }

        private async Task EnableSteamFirewallAutoModeAsync(bool ensureObserverRunning)
        {
            _steamFirewallAutoEnabled = true;
            _steamFirewallAutoPaused = false;
            _steamFirewallAutoHeldForActiveMatch = false;
            _steamFirewallAutoTimer.Start();
            AppendLog("[FIREWALL]", "Firewall Fix (Matchmaking) enabled. Firewall, proxy observer, and crash restore are armed.", "#00C8FF");
            SetStatus("Firewall Fix (Matchmaking) enabled.", "#00C8FF");
            UpdateRejoinFirewallStatus("FIREWALL: MATCHMAKING PENDING - enabling port block", "#FF6A00");

            try
            {
                string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
                await SetSteamFirewallEnabledAsync(true, mccExePath);
                SaveSteamFirewallUiState(true);
                SetSteamFirewallRuntimeState(SteamFirewallState.Enabled);
                UpdateRejoinFirewallStatus();
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Could not enable Firewall Fix (Matchmaking): {ex.Message}", "#FF2D55");
                _steamFirewallAutoEnabled = false;
                _steamFirewallAutoPaused = false;
                _steamFirewallAutoHeldForActiveMatch = false;
                _steamFirewallAutoSuspendedForCrashRestore = false;
                _steamFirewallAutoTimer.Stop();
                SetRejoinFirewallCheckbox(ChkRejoinFixFirewallMatchmaking, false);
                UpdateRejoinFirewallStatus("FIREWALL: MATCHMAKING FAILED - port block was not applied", "#FF2D55");
                SetStatus("Firewall Fix (Matchmaking) failed.", "#FF2D55");
            }

            if (ensureObserverRunning && _steamFirewallAutoEnabled)
                await EnsureRejoinObserverRunningAsync();
        }

        private void DisableSteamFirewallAutoMode(bool logStatus)
        {
            _steamFirewallAutoEnabled = false;
            _steamFirewallAutoPaused = false;
            _steamFirewallAutoHeldForActiveMatch = false;
            _steamFirewallAutoSuspendedForCrashRestore = false;
            _steamFirewallAutoTimer.Stop();
            UpdateRejoinFirewallStatus();
            if (!logStatus)
                return;

            AppendLog("[FIREWALL]", "Firewall Fix (Matchmaking) disabled. Manual firewall state is left as-is.", "#C8D8E8");
            SetStatus("Firewall Fix (Matchmaking) disabled.", "#C8D8E8");
        }

        private async Task DisableRejoinFirewallRulesAsync(bool logStatus)
        {
            try
            {
                string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
                await SetSteamFirewallEnabledAsync(false, mccExePath);
                SaveSteamFirewallUiState(false);
                SetSteamFirewallRuntimeState(SteamFirewallState.Disabled);
                _rejoinCampaignFirewallApplying = false;
                _rejoinCampaignFirewallEnabled = false;
                UpdateRejoinFirewallStatus();

                if (logStatus)
                {
                    AppendLog("[FIREWALL]", "Rejoin firewall rules disabled.", "#C8D8E8");
                    SetStatus("Firewall fixes disabled.", "#C8D8E8");
                }
            }
            catch (Exception ex)
            {
                _rejoinCampaignFirewallApplying = false;
                AppendLog("[ERROR]", $"Could not disable Rejoin firewall rules: {ex.Message}", "#FF2D55");
                SetStatus("Firewall fixes disable failed.", "#FF2D55");
            }
        }

        private async Task EnsureSteamFirewallRulesPreparedAsync()
        {
            string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
            UpdateRejoinFirewallStatus(
                _steamFirewallRulesPrepared
                    ? "FIREWALL: SETUP - confirming rules are disabled"
                    : "FIREWALL: SETUP - preparing disabled rules for later toggles",
                "#FF6A00");
            AppendLog("[FIREWALL]", _steamFirewallRulesPrepared
                ? "Confirming MCC P2P firewall rules are disabled before Rejoin Fix continues..."
                : "Preparing MCC P2P firewall rules for Campaign and Matchmaking toggles...", "#FF6A00");

            await SetSteamFirewallEnabledAsync(false, mccExePath);
            SaveSteamFirewallUiState(false);
            SetSteamFirewallRuntimeState(SteamFirewallState.Disabled);
            _steamFirewallRulesPrepared = true;

            AppendLog("[FIREWALL]", $"MCC P2P firewall rules are ready and disabled for {Path.GetFileName(mccExePath)}.", "#39FF14");
            UpdateRejoinFirewallStatus();
        }

        private void SetSteamFirewallRuntimeState(SteamFirewallState state)
        {
            _steamFirewallUiState = state;
            if (SteamFirewallFeatureEnabled)
                UpdateSteamFirewallUi(state);
            UpdateRejoinFirewallStatus();
        }

        private void SetRejoinFirewallCheckbox(CheckBox checkBox, bool isChecked)
        {
            _rejoinFirewallCheckChanging = true;
            try
            {
                checkBox.IsChecked = isChecked;
            }
            finally
            {
                _rejoinFirewallCheckChanging = false;
            }
        }

        private void HandleSteamFirewallAutoSignal(ProxyCaptureEntry entry)
        {
            if (!_steamFirewallAutoEnabled)
                return;

            if (_steamFirewallAutoSuspendedForCrashRestore)
                return;

            if (_steamFirewallAutoPaused && IsSteamFirewallAutoResumeSignal(entry))
            {
                ScheduleSteamFirewallAutoResume(SteamFirewallAutoMatchFoundHoldSeconds, "lobby connection confirmed");
                return;
            }

            if (_steamFirewallAutoPaused && IsSteamFirewallAutoDisableSignal(entry))
            {
                if (!_steamFirewallAutoHeldForActiveMatch)
                    ScheduleSteamFirewallAutoResume(SteamFirewallAutoSearchHoldSeconds, "matchmaking traffic still active");
                return;
            }

            if (_steamFirewallAutoPaused)
                return;

            if (!IsSteamFirewallAutoDisableSignal(entry))
                return;

            _ = PauseSteamFirewallForMatchmakingAsync(entry);
        }

        private async Task HandleCrashRestoreFirewallStateChangedAsync(bool pending)
        {
            if (pending)
            {
                await SuspendSteamFirewallForCrashRestoreAsync();
                return;
            }

            await ResumeSteamFirewallAfterCrashRestoreAsync();
        }

        private async Task SuspendSteamFirewallForCrashRestoreAsync()
        {
            if (_steamFirewallAutoSuspendedForCrashRestore)
                return;

            bool firewallMayBlockRestore = _steamFirewallAutoEnabled ||
                _steamFirewallUiState is SteamFirewallState.Enabled or SteamFirewallState.Partial ||
                ChkRejoinFixFirewall.IsChecked == true ||
                ChkRejoinFixFirewallMatchmaking.IsChecked == true;

            if (!firewallMayBlockRestore)
                return;

            // Crash restore is a hard firewall suspension. Wait for any in-flight
            // auto transition so this request cannot be dropped while the normal
            // matchmaking timer is changing the same rules.
            await _steamFirewallAutoLock.WaitAsync();

            try
            {
                _steamFirewallAutoSuspendedForCrashRestore = true;
                _steamFirewallAutoPaused = _steamFirewallAutoEnabled;
                _steamFirewallAutoHeldForActiveMatch = true;
                _steamFirewallAutoResumeAfterUtc = DateTime.MaxValue;

                string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
                UpdateRejoinFirewallStatus("FIREWALL: REJOIN RESTORE - opening ports for crash rejoin", "#00C8FF");
                AppendLog("[FIREWALL]", "Crash restore armed; opening MCC P2P firewall rules until rejoin finishes or times out.", "#00C8FF");

                await SetSteamFirewallEnabledAsync(false, mccExePath);
                SaveSteamFirewallUiState(false);
                SetSteamFirewallRuntimeState(SteamFirewallState.Disabled);
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Could not open firewall for crash restore: {ex.Message}", "#FF2D55");
                SetStatus("Crash restore firewall open failed.", "#FF6A00");
            }
            finally
            {
                _steamFirewallAutoLock.Release();
            }
        }

        private async Task ResumeSteamFirewallAfterCrashRestoreAsync()
        {
            if (!_steamFirewallAutoSuspendedForCrashRestore)
                return;

            _steamFirewallAutoSuspendedForCrashRestore = false;
            _steamFirewallAutoHeldForActiveMatch = false;

            if (!_steamFirewallAutoEnabled || ChkRejoinFixFirewallMatchmaking.IsChecked != true)
            {
                _steamFirewallAutoPaused = false;
                UpdateRejoinFirewallStatus();
                return;
            }

            _steamFirewallAutoPaused = true;
            _steamFirewallAutoResumeAfterUtc = DateTime.UtcNow;
            AppendLog("[FIREWALL]", "Crash restore ended; resuming MCC P2P firewall auto protection.", "#00C8FF");
            await ResumeSteamFirewallAfterMatchmakingAsync();
        }

        private async Task EnsureRejoinObserverRunningAsync()
        {
            if (_rejoinProxy.IsRunning)
                return;

            try
            {
                RejoinFixPaths.EnsureRootDirectory();
                _rejoinWinHttpManualNeeded = false;
                RejoinFixDiagnostics.Info("proxy", "Auto protection started the proxy observer.");
                await _rejoinProxy.StartAsync();
                StartRejoinCrashWatcher();
                StartNetworkStatsOverlay(_rejoinProxy.CurrentGameServerIp);
                AppendLog("[REJOIN]", $"Proxy observer active for protection on 127.0.0.1:{_rejoinProxy.Port}.", "#39FF14");
                UpdateRejoinFixUi();
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Auto protection could not start proxy observer: {ex.Message}", "#FF2D55");
                SetStatus("Protection observer failed to start.", "#FF2D55");
            }
        }

        private static bool IsSteamFirewallAutoDisableSignal(ProxyCaptureEntry entry)
        {
            string path = entry.Path;

            // A successful SmartMatch hopper POST is the first definitive signal
            // that MCC entered a matchmaking queue. The hopper path itself does
            // not contain the word "Matchmaking", so relying on the legacy path
            // checks below leaves the firewall enabled throughout the search.
            bool isSmartMatchQueueStart =
                entry.StatusCode is >= 200 and < 300 &&
                entry.Method.Equals("POST", StringComparison.OrdinalIgnoreCase) &&
                entry.Host.Equals("smartmatch.xboxlive.com", StringComparison.OrdinalIgnoreCase) &&
                path.Contains("/hoppers/", StringComparison.OrdinalIgnoreCase);

            return isSmartMatchQueueStart
                || path.Contains("Party/RequestParty", StringComparison.OrdinalIgnoreCase)
                || path.Contains("Matchmaking", StringComparison.OrdinalIgnoreCase)
                || path.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSteamFirewallAutoResumeSignal(ProxyCaptureEntry entry)
        {
            if (entry.StatusCode < 200 || entry.StatusCode >= 300)
                return false;

            return entry.Path.Contains("Party/RequestParty", StringComparison.OrdinalIgnoreCase)
                || entry.Path.Contains("/CascadeMatchmaking/sessions/", StringComparison.OrdinalIgnoreCase)
                || entry.Path.Contains("/CascadeSquadSession/sessions/", StringComparison.OrdinalIgnoreCase)
                || entry.Path.Contains("/handles", StringComparison.OrdinalIgnoreCase);
        }

        private async Task PauseSteamFirewallForMatchmakingAsync(ProxyCaptureEntry entry)
        {
            if (!_steamFirewallAutoEnabled || _steamFirewallUiState != SteamFirewallState.Enabled)
                return;

            if (!await _steamFirewallAutoLock.WaitAsync(0))
                return;

            try
            {
                if (!_steamFirewallAutoEnabled || _steamFirewallUiState != SteamFirewallState.Enabled)
                    return;

                string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
                TxtSteamFirewallStatus.Text = "AUTO - pausing MCC port block for matchmaking";
                TxtSteamFirewallStatus.Foreground = Brush("#FF6A00");
                UpdateRejoinFirewallStatus("FIREWALL: MATCHMAKING PAUSING - opening ports for matchmaking", "#FF6A00");
                BtnSteamFirewallFix.Content = "PAUSING";
                BtnSteamFirewallFix.IsEnabled = false;

                await SetSteamFirewallEnabledAsync(false, mccExePath);
                SaveSteamFirewallUiState(false);
                _steamFirewallAutoPaused = true;
                _steamFirewallAutoHeldForActiveMatch = false;
                _steamFirewallAutoResumeAfterUtc = DateTime.UtcNow.AddSeconds(SteamFirewallAutoSearchHoldSeconds);
                SetSteamFirewallRuntimeState(SteamFirewallState.Disabled);
                TxtSteamFirewallStatus.Text = "AUTO - disabled while MCC searches/connects";
                TxtSteamFirewallStatus.Foreground = Brush("#00C8FF");
                UpdateRejoinFirewallStatus("FIREWALL: MATCHMAKING PAUSED - ports are open while MCC searches/connects", "#00C8FF");
                AppendLog("[FIREWALL]", $"Auto-disabled MCC port block after matchmaking signal: {entry.Host}{entry.Path}", "#00C8FF");
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Firewall Auto pause failed: {ex.Message}", "#FF2D55");
                SetStatus("Firewall Auto pause failed.", "#FF2D55");
            }
            finally
            {
                _steamFirewallAutoLock.Release();
                BtnSteamFirewallFix.IsEnabled = true;
            }
        }

        private void ScheduleSteamFirewallAutoResume(int holdSeconds, string reason)
        {
            if (_steamFirewallAutoSuspendedForCrashRestore ||
                !_steamFirewallAutoEnabled ||
                !_steamFirewallAutoPaused)
                return;

            _steamFirewallAutoResumeAfterUtc = DateTime.UtcNow.AddSeconds(holdSeconds);
            TxtSteamFirewallStatus.Text = $"AUTO - re-enabling soon ({reason})";
            TxtSteamFirewallStatus.Foreground = Brush("#00C8FF");
            UpdateRejoinFirewallStatus($"FIREWALL: MATCHMAKING PENDING - re-enabling soon ({reason})", "#00C8FF");
        }

        private void HoldSteamFirewallPausedForActiveMatch(string reason)
        {
            // A server assignment or squad update is only rejoin progress, not
            // proof that MCC completed its UDP/DTLS connection. Do not let those
            // normal match signals shorten the crash-restore firewall suspension.
            if (_steamFirewallAutoSuspendedForCrashRestore ||
                !_steamFirewallAutoEnabled ||
                !_steamFirewallAutoPaused)
                return;

            _steamFirewallAutoHeldForActiveMatch = true;
            _steamFirewallAutoResumeAfterUtc = DateTime.UtcNow.AddSeconds(SteamFirewallAutoMatchFoundHoldSeconds);
            TxtSteamFirewallStatus.Text = $"AUTO - re-enabling soon ({reason})";
            TxtSteamFirewallStatus.Foreground = Brush("#00C8FF");
            UpdateRejoinFirewallStatus($"FIREWALL: MATCHMAKING PENDING - re-enabling soon ({reason})", "#00C8FF");
            AppendLog("[FIREWALL]", $"Detected active match signal; re-enabling MCC port block soon: {reason}.", "#00C8FF");
        }

        private async Task SteamFirewallAutoTimer_TickAsync()
        {
            if (_steamFirewallAutoSuspendedForCrashRestore ||
                !_steamFirewallAutoEnabled ||
                !_steamFirewallAutoPaused ||
                DateTime.UtcNow < _steamFirewallAutoResumeAfterUtc)
                return;

            await ResumeSteamFirewallAfterMatchmakingAsync();
        }

        private async Task ResumeSteamFirewallAfterMatchmakingAsync()
        {
            if (!await _steamFirewallAutoLock.WaitAsync(0))
                return;

            try
            {
                // Recheck after acquiring the lock. Crash restore may have been
                // armed while this resume was waiting behind another transition.
                if (_steamFirewallAutoSuspendedForCrashRestore ||
                    !_steamFirewallAutoEnabled ||
                    !_steamFirewallAutoPaused)
                    return;

                string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
                TxtSteamFirewallStatus.Text = "AUTO - re-enabling MCC port block";
                TxtSteamFirewallStatus.Foreground = Brush("#FF6A00");
                UpdateRejoinFirewallStatus("FIREWALL: MATCHMAKING PENDING - re-enabling port block", "#FF6A00");
                BtnSteamFirewallFix.Content = "ENABLING";
                BtnSteamFirewallFix.IsEnabled = false;

                await SetSteamFirewallEnabledAsync(true, mccExePath);
                SaveSteamFirewallUiState(true);
                _steamFirewallAutoPaused = false;
                _steamFirewallAutoHeldForActiveMatch = false;
                SetSteamFirewallRuntimeState(SteamFirewallState.Enabled);
                AppendLog("[FIREWALL]", "Auto re-enabled MCC port block after matchmaking quiet period.", "#39FF14");
            }
            catch (Exception ex)
            {
                _steamFirewallAutoResumeAfterUtc = DateTime.UtcNow.AddSeconds(30);
                AppendLog("[ERROR]", $"Firewall Auto resume failed: {ex.Message}", "#FF2D55");
                SetStatus("Firewall Auto resume failed; retrying.", "#FF6A00");
            }
            finally
            {
                _steamFirewallAutoLock.Release();
                BtnSteamFirewallFix.IsEnabled = true;
            }
        }

        private async void BtnSteamFirewallFix_Click(object sender, RoutedEventArgs e)
        {
            if (!SteamFirewallFeatureEnabled)
            {
                InitializeSteamFirewallFeatureState();
                return;
            }

            BtnSteamFirewallFix.IsEnabled = false;

            try
            {
                bool enable = _steamFirewallUiState != SteamFirewallState.Enabled;
                string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());

                TxtSteamFirewallStatus.Text = enable
                    ? "APPLYING - waiting for administrator approval"
                    : "DISABLING - waiting for administrator approval";
                TxtSteamFirewallStatus.Foreground = Brush("#FF6A00");
                BtnSteamFirewallFix.Content = enable ? "ENABLING" : "DISABLING";

                await SetSteamFirewallEnabledAsync(enable, mccExePath);
                SaveSteamFirewallUiState(enable);
                AppendLog("[FIREWALL]", enable
                    ? $"MCC P2P firewall fix enabled for ports 3478 and 4379 ({Path.GetFileName(mccExePath)})."
                    : "MCC P2P firewall fix disabled for ports 3478 and 4379.", enable ? "#39FF14" : "#C8D8E8");
                SetStatus(enable ? "MCC P2P firewall fix enabled." : "MCC P2P firewall fix disabled.",
                    enable ? "#39FF14" : "#C8D8E8");
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                AppendLog("[INFO]", "MCC P2P firewall fix cancelled at administrator prompt.", "#4A5A6A");
                SetStatus("MCC P2P firewall fix cancelled.", "#4A5A6A");
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"MCC P2P firewall fix failed: {ex.Message}", "#FF2D55");
                SetStatus("MCC P2P firewall fix failed.", "#FF2D55");
                ToolboxDialog.Show(
                    $"MCC P2P Firewall Fix could not be changed:\n\n{ex.Message}",
                    "MCC P2P Firewall Fix -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                await RefreshSteamFirewallUiAsync();
            }
        }

        private static async Task<SteamFirewallState> GetSteamFirewallStateAsync()
        {
            string ruleNames = string.Join(", ", SteamFirewallRuleNames.Select(QuotePowerShellString));
            string legacyRuleNames = string.Join(", ", LegacySteamFirewallRuleNames.Select(QuotePowerShellString));
            string globalRuleNames = string.Join(", ", GlobalSteamFirewallRuleNames.Select(QuotePowerShellString));
            string script = $@"
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$names = @({ruleNames})
$legacyNames = @({legacyRuleNames})
$globalNames = @({globalRuleNames})
$existing = @()
foreach ($name in $names) {{
    $rule = Get-NetFirewallRule -DisplayName $name -ErrorAction SilentlyContinue
    if ($rule) {{ $existing += $rule }}
}}
$legacy = @()
foreach ($name in $legacyNames) {{
    $rule = Get-NetFirewallRule -DisplayName $name -ErrorAction SilentlyContinue
    if ($rule) {{ $legacy += $rule }}
}}
$global = @()
foreach ($name in $globalNames) {{
    $rule = Get-NetFirewallRule -DisplayName $name -ErrorAction SilentlyContinue
    if ($rule) {{ $global += $rule }}
}}
$globalEnabled = @($global | Where-Object {{ $_.Enabled -eq 'True' }}).Count -gt 0
if ($existing.Count -eq 0) {{
    if ($globalEnabled -or @($legacy | Where-Object {{ $_.Enabled -eq 'True' }}).Count -gt 0) {{
        'Partial'
    }} else {{
        'Missing'
    }}
}} elseif ($existing.Count -lt $names.Count) {{
    $allRules = @($existing) + @($legacy) + @($global)
    if (@($allRules | Where-Object {{ $_.Enabled -eq 'True' }}).Count -gt 0) {{
        'Partial'
    }} else {{
        'Partial'
    }}
}} else {{
    $enabledCount = @($existing | Where-Object {{ $_.Enabled -eq 'True' }}).Count
    if ($enabledCount -eq $existing.Count) {{
        'Enabled'
    }} elseif ($enabledCount -eq 0 -and -not $globalEnabled -and @($legacy | Where-Object {{ $_.Enabled -eq 'True' }}).Count -eq 0) {{
        'Disabled'
    }} else {{
        'Partial'
    }}
}}";

            string output = (await RunPowerShellAsync(script, elevated: false)).Trim();
            return Enum.TryParse(output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault(),
                ignoreCase: true,
                out SteamFirewallState state)
                ? state
                : SteamFirewallState.Unknown;
        }

        private static string ResolveMccExecutablePath(string mccRoot)
        {
            if (string.IsNullOrWhiteSpace(mccRoot))
                throw new InvalidOperationException("Set the MCC installation path before enabling this fix.");

            var candidates = new[]
            {
                Path.Combine(mccRoot, "MCC", "Binaries", "Win64", "MCC-Win64-Shipping.exe"),
                Path.Combine(mccRoot, "MCC", "Binaries", "Win64", "MCC.exe"),
                Path.Combine(mccRoot, "MCC-Win64-Shipping.exe"),
                Path.Combine(mccRoot, "MCC.exe"),
            };

            string? exePath = candidates.FirstOrDefault(File.Exists);
            if (exePath is not null)
                return exePath;

            throw new FileNotFoundException(
                "Could not find the Halo MCC executable. Check the MCC installation path.",
                candidates[0]);
        }

        private static bool IsRunningAsAdministrator()
        {
            try
            {
                var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private static void RelaunchAsAdministrator()
        {
            string? executablePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (string.IsNullOrWhiteSpace(executablePath))
                throw new InvalidOperationException("Could not find the Toolbox executable path.");

            Process.Start(new ProcessStartInfo
            {
                FileName = executablePath,
                UseShellExecute = true,
                Verb = "runas"
            });
        }

        private async Task ApplyRejoinFirewallOptionAsync()
        {
            if (ChkRejoinFixFirewall.IsChecked != true)
                return;

            if (!IsRunningAsAdministrator())
                throw new InvalidOperationException("Firewall Fix (Campaign) requires the Toolbox to run as Administrator.");

            string mccExePath = ResolveMccExecutablePath(TxtMccPath.Text.Trim());
            AppendLog("[FIREWALL]", "Enabling MCC P2P firewall fix for Rejoin Fix...", "#FF6A00");
            AppendLog("[FIREWALL]", "Toolbox is running as Administrator.", "#C8D8E8");
            SetStatus("Enabling MCC P2P firewall fix...", "#FF6A00");
            _rejoinCampaignFirewallApplying = true;
            _rejoinCampaignFirewallEnabled = false;
            UpdateRejoinFirewallStatus();

            bool applied = false;
            try
            {
                await SetSteamFirewallEnabledAsync(true, mccExePath, RejoinCampaignFirewallPorts, SteamFirewallPorts);
                SaveSteamFirewallUiState(true);
                SetSteamFirewallRuntimeState(SteamFirewallState.Enabled);
                _rejoinCampaignFirewallEnabled = true;
                applied = true;
            }
            finally
            {
                _rejoinCampaignFirewallApplying = false;
                if (!applied)
                    _rejoinCampaignFirewallEnabled = false;
                UpdateRejoinFirewallStatus();
            }

            AppendLog("[FIREWALL]", $"MCC P2P firewall fix enabled for Rejoin Fix ({Path.GetFileName(mccExePath)}).", "#39FF14");
            SetStatus("Firewall Fix (Campaign) enabled.", "#39FF14");
        }

        private static async Task SetSteamFirewallEnabledAsync(
            bool enabled,
            string mccExePath,
            IReadOnlyCollection<int>? activePorts = null,
            IReadOnlyCollection<int>? cleanupPorts = null)
        {
            activePorts ??= SteamFirewallPorts;
            cleanupPorts ??= activePorts;
            string ports = string.Join(", ", activePorts);
            string cleanupPortsText = string.Join(", ", cleanupPorts);
            string rulePrefix = SteamFirewallRulePrefix;
            string legacyRulePrefix = LegacyPort4379FirewallRulePrefix;
            string globalRulePrefix = GlobalSteamFirewallRulePrefix;
            string targetEnabled = enabled ? "yes" : "no";
            string quotedMccExePath = QuotePowerShellString(mccExePath);

            string script = $@"
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$ports = @({ports})
$cleanupPorts = @({cleanupPortsText})
$rulePrefix = {QuotePowerShellString(rulePrefix)}
$legacyRulePrefix = {QuotePowerShellString(legacyRulePrefix)}
$globalRulePrefix = {QuotePowerShellString(globalRulePrefix)}
$mccExePath = {quotedMccExePath}

function Disable-RuleIfPresent([string]$name) {{
    & netsh advfirewall firewall set rule ""name=$name"" new enable=no | Out-Null
}}

function Invoke-Netsh([string[]]$arguments) {{
    $output = & netsh @arguments 2>&1
    $message = ($output | Out-String).Trim()
    return [pscustomobject]@{{
        ExitCode = $LASTEXITCODE
        Text = $message
    }}
}}

function Invoke-NetshChecked([string]$label, [string[]]$arguments) {{
    $result = Invoke-Netsh $arguments
    $message = $result.Text
    if ($result.ExitCode -ne 0) {{
        if ([string]::IsNullOrWhiteSpace($message)) {{
            $message = 'netsh returned no output'
        }}

        throw ""NETSH_FAILED: $label failed with exit code $($result.ExitCode). $message""
    }}

    return $message
}}

function Test-NetshRuleExists([string]$name) {{
    $result = Invoke-Netsh @('advfirewall', 'firewall', 'show', 'rule', ""name=$name"", 'verbose')
    if ($result.ExitCode -ne 0) {{
        return $false
    }}

    return $result.Text -match 'Rule Name:\s+'
}}

function Ensure-Rule(
    [string]$name,
    [string]$direction,
    [string]$protocol,
    [string]$portSide,
    [int]$port,
    [bool]$ruleEnabled) {{
    $netshEnabledValue = if ($ruleEnabled) {{ 'yes' }} else {{ 'no' }}
    $netshDirection = if ($direction -eq 'Inbound') {{ 'dir=in' }} else {{ 'dir=out' }}
    $netshPortArgument = if ($portSide -eq 'Local') {{ ""localport=$port"" }} else {{ ""remoteport=$port"" }}

    if (Test-NetshRuleExists $name) {{
        Invoke-NetshChecked ""set $name"" @(
            'advfirewall',
            'firewall',
            'set',
            'rule',
            ""name=$name"",
            'new',
            ""enable=$netshEnabledValue"") | Out-Null
    }} else {{
        Invoke-NetshChecked ""add $name"" @(
            'advfirewall',
            'firewall',
            'add',
            'rule',
            ""name=$name"",
            $netshDirection,
            'action=block',
            ""program=$mccExePath"",
            ""protocol=$protocol"",
            $netshPortArgument,
            'profile=any',
            ""enable=$netshEnabledValue"") | Out-Null
    }}
}}

$ruleEnabled = [string]::Equals('{targetEnabled}', 'yes', [System.StringComparison]::OrdinalIgnoreCase)
foreach ($port in $ports) {{
    Ensure-Rule ""$rulePrefix $port TCP Inbound"" 'Inbound' 'TCP' 'Local' $port $ruleEnabled
    Ensure-Rule ""$rulePrefix $port UDP Inbound"" 'Inbound' 'UDP' 'Local' $port $ruleEnabled
    Ensure-Rule ""$rulePrefix $port TCP Outbound"" 'Outbound' 'TCP' 'Remote' $port $ruleEnabled
    Ensure-Rule ""$rulePrefix $port UDP Outbound"" 'Outbound' 'UDP' 'Remote' $port $ruleEnabled
}}

Disable-RuleIfPresent ""$legacyRulePrefix TCP Inbound""
Disable-RuleIfPresent ""$legacyRulePrefix UDP Inbound""
Disable-RuleIfPresent ""$legacyRulePrefix TCP Outbound""
Disable-RuleIfPresent ""$legacyRulePrefix UDP Outbound""

foreach ($port in $cleanupPorts) {{
    Disable-RuleIfPresent ""$globalRulePrefix $port TCP Inbound""
    Disable-RuleIfPresent ""$globalRulePrefix $port UDP Inbound""
    Disable-RuleIfPresent ""$globalRulePrefix $port TCP Outbound""
    Disable-RuleIfPresent ""$globalRulePrefix $port UDP Outbound""

    if ($ports -notcontains $port) {{
        Disable-RuleIfPresent ""$rulePrefix $port TCP Inbound""
        Disable-RuleIfPresent ""$rulePrefix $port UDP Inbound""
        Disable-RuleIfPresent ""$rulePrefix $port TCP Outbound""
        Disable-RuleIfPresent ""$rulePrefix $port UDP Outbound""
    }}
}}

exit 0";

            if (!await SteamFirewallCommandLock.WaitAsync(TimeSpan.FromSeconds(5)))
                throw new TimeoutException("Another firewall command is still running. Close/reopen Toolbox or wait for the previous Windows firewall prompt to finish.");

            try
            {
                await RunPowerShellAsync(script, elevated: !IsRunningAsAdministrator(), timeoutMs: 30000);
            }
            finally
            {
                SteamFirewallCommandLock.Release();
            }
        }

        private static async Task VerifySteamFirewallRulesAsync(
            bool enabled,
            string mccExePath,
            IReadOnlyCollection<int> activePorts,
            IReadOnlyCollection<int> cleanupPorts)
        {
            string activeRuleNames = string.Join(", ", activePorts.SelectMany(port => new[]
            {
                $"{SteamFirewallRulePrefix} {port} TCP Inbound",
                $"{SteamFirewallRulePrefix} {port} UDP Inbound",
                $"{SteamFirewallRulePrefix} {port} TCP Outbound",
                $"{SteamFirewallRulePrefix} {port} UDP Outbound"
            }).Select(QuotePowerShellString));

            string cleanupRuleNames = string.Join(", ", cleanupPorts
                .Except(activePorts)
                .SelectMany(port => new[]
                {
                    $"{SteamFirewallRulePrefix} {port} TCP Inbound",
                    $"{SteamFirewallRulePrefix} {port} UDP Inbound",
                    $"{SteamFirewallRulePrefix} {port} TCP Outbound",
                    $"{SteamFirewallRulePrefix} {port} UDP Outbound"
                })
                .Select(QuotePowerShellString));

            string expectedEnabled = enabled ? "Yes" : "No";
            string quotedMccExePath = QuotePowerShellString(mccExePath);

            string script = $@"
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$activeNames = @({activeRuleNames})
$cleanupNames = @({cleanupRuleNames})
$expectedEnabled = '{expectedEnabled}'
$mccExePath = {quotedMccExePath}
$problems = New-Object System.Collections.Generic.List[string]

function Invoke-Netsh([string[]]$arguments) {{
    $output = & netsh @arguments 2>&1
    $message = ($output | Out-String).Trim()
    return [pscustomobject]@{{
        ExitCode = $LASTEXITCODE
        Text = $message
    }}
}}

function Read-NetshRule([string]$name) {{
    Invoke-Netsh @('advfirewall', 'firewall', 'show', 'rule', ""name=$name"", 'verbose')
}}

function Read-Field([string]$text, [string]$label, [string]$pattern) {{
    $match = [regex]::Match($text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {{
        $problems.Add(""could not read $label"")
        return $null
    }}

    return $match.Groups[1].Value.Trim()
}}

function Assert-Field([string]$ruleName, [string]$text, [string]$label, [string]$pattern, [string]$expected) {{
    $actual = Read-Field $text $label $pattern
    if ($null -eq $actual) {{
        $problems.Add(""could not read $label for $ruleName"")
        return
    }}

    if (-not [string]::Equals($actual, $expected, [System.StringComparison]::OrdinalIgnoreCase)) {{
        $problems.Add(""wrong $label for $ruleName (expected $expected, found $actual)"")
    }}
}}

foreach ($name in $activeNames) {{
    $result = Read-NetshRule $name
    if ($result.ExitCode -ne 0 -or $result.Text -notmatch 'Rule Name:\s+') {{
        $problems.Add(""missing active rule: $name. $($result.Text)"")
        continue
    }}

    $expectedDirection = if ($name -match 'Inbound$') {{ 'In' }} else {{ 'Out' }}
    $expectedProtocol = if ($name -match ' UDP ') {{ 'UDP' }} else {{ 'TCP' }}
    $expectedPortSide = if ($expectedDirection -eq 'In') {{ 'LocalPort' }} else {{ 'RemotePort' }}
    $portMatch = [regex]::Match($name, ' Port (\d+) ', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $portMatch.Success) {{
        $problems.Add(""could not infer expected port for $name"")
        continue
    }}

    Assert-Field $name $result.Text 'enabled state' 'Enabled:\s+(Yes|No)' $expectedEnabled
    Assert-Field $name $result.Text 'action' 'Action:\s+(\S+)' 'Block'
    Assert-Field $name $result.Text 'direction' 'Direction:\s+(\S+)' $expectedDirection
    Assert-Field $name $result.Text 'protocol' 'Protocol:\s+(\S+)' $expectedProtocol
    Assert-Field $name $result.Text 'program' 'Program:\s+(.+)' $mccExePath
    Assert-Field $name $result.Text $expectedPortSide ""$($expectedPortSide):\s+(\S+)"" $portMatch.Groups[1].Value
}}

foreach ($name in $cleanupNames) {{
    $result = Read-NetshRule $name
    if ($result.ExitCode -ne 0 -or $result.Text -notmatch 'Rule Name:\s+') {{
        continue
    }}

    $enabledMatch = [regex]::Match($result.Text, 'Enabled:\s+(Yes|No)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($enabledMatch.Success -and [string]::Equals($enabledMatch.Groups[1].Value, 'Yes', [System.StringComparison]::OrdinalIgnoreCase)) {{
        $problems.Add(""cleanup rule still enabled: $name"")
    }}
}}

if ($problems.Count -gt 0) {{
    Write-Output (""VERIFY_FAILED: "" + ($problems -join '; '))
    exit 1
}}

'Verified'";

            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    await RunPowerShellAsync(script, elevated: false, timeoutMs: 10000);
                    return;
                }
                catch when (attempt < 3)
                {
                    await Task.Delay(500);
                }
            }
        }

        private static async Task<string> RunPowerShellAsync(string script, bool elevated, int timeoutMs = 5000)
        {
            string? elevatedScriptPath = null;
            string? elevatedTranscriptPath = null;
            string encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encoded}",
                UseShellExecute = elevated,
                CreateNoWindow = !elevated
            };

            if (elevated)
            {
                Directory.CreateDirectory(ToolboxLocalAppDataRoot);
                elevatedScriptPath = Path.Combine(ToolboxLocalAppDataRoot, $"firewall-command-{Guid.NewGuid():N}.ps1");
                elevatedTranscriptPath = Path.Combine(ToolboxLocalAppDataRoot, "firewall-command-result.txt");
                string wrappedScript = $@"
$ErrorActionPreference = 'Continue'
Start-Transcript -Path {QuotePowerShellString(elevatedTranscriptPath)} -Force | Out-Null
try {{
{script}
}} catch {{
    Write-Error ($_.Exception | Format-List * -Force | Out-String)
    exit 1
}} finally {{
    try {{ Stop-Transcript | Out-Null }} catch {{ }}
}}";

                File.WriteAllText(elevatedScriptPath, wrappedScript, Encoding.UTF8);
                startInfo.Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"{elevatedScriptPath}\"";
                startInfo.Verb = "runas";
            }
            else
            {
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
            }

            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Failed to start PowerShell.");

            if (elevated)
            {
                var elevatedWaitTask = process.WaitForExitAsync();
                if (await Task.WhenAny(elevatedWaitTask, Task.Delay(timeoutMs)) != elevatedWaitTask)
                {
                    try
                    {
                        process.Kill(entireProcessTree: true);
                    }
                    catch
                    {
                        // Best effort cleanup; elevated child processes may outlive us if Windows denies termination.
                    }

                    throw new TimeoutException("Elevated PowerShell firewall command timed out.");
                }

                if (process.ExitCode != 0)
                {
                    string transcript = "";
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(elevatedTranscriptPath) && File.Exists(elevatedTranscriptPath))
                            transcript = File.ReadAllText(elevatedTranscriptPath);
                    }
                    catch
                    {
                        transcript = "";
                    }

                    transcript = CleanPowerShellError(transcript);
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(transcript)
                        ? $"Elevated firewall command failed with exit code {process.ExitCode}."
                        : transcript);
                }

                return "";
            }

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();
            var waitTask = process.WaitForExitAsync();
            if (await Task.WhenAny(waitTask, Task.Delay(timeoutMs)) != waitTask)
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // Best effort cleanup; caller will fall back to an actionable UI state.
                }

                throw new TimeoutException("PowerShell status check timed out.");
            }

            string output = await outputTask;
            string error = await errorTask;

            if (process.ExitCode != 0)
            {
                error = CleanPowerShellError(error);
                string failureText = string.IsNullOrWhiteSpace(error)
                    ? CleanPowerShellError(output)
                    : error;
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(failureText)
                    ? $"PowerShell exited with code {process.ExitCode}."
                    : failureText);
            }

            return output;
        }

        private static string CleanPowerShellError(string error)
        {
            if (string.IsNullOrWhiteSpace(error))
                return "";

            string cleaned = error.Trim();
            if (cleaned.StartsWith("#< CLIXML", StringComparison.OrdinalIgnoreCase))
            {
                var lines = cleaned
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(line =>
                        !line.Contains("System.Management.Automation.PSCustomObject", StringComparison.OrdinalIgnoreCase) &&
                        !line.Contains("Preparing modules for first use", StringComparison.OrdinalIgnoreCase) &&
                        !line.Contains("Completed", StringComparison.OrdinalIgnoreCase) &&
                        !line.Contains("progress", StringComparison.OrdinalIgnoreCase) &&
                        !line.Contains("Get-NetFirewallRule", StringComparison.OrdinalIgnoreCase) &&
                        !line.StartsWith("#< CLIXML", StringComparison.OrdinalIgnoreCase) &&
                        !line.StartsWith("<Objs ", StringComparison.OrdinalIgnoreCase) &&
                        !line.StartsWith("</Objs>", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                cleaned = string.Join(Environment.NewLine, lines).Trim();
            }

            return cleaned;
        }

        private static string QuotePowerShellString(string value) => $"'{value.Replace("'", "''")}'";

        private async Task StartRejoinFixAsync()
        {
            RejoinFixPaths.EnsureRootDirectory();
            _rejoinWinHttpManualNeeded = false;
            RejoinFixDiagnostics.Info("proxy", "Activation requested from Toolbox UI.");
            AppendLog("[REJOIN]", "Starting Rejoin Fix proxy...", "#FF6A00");
            SetStatus("Starting Rejoin Fix...", "#FF6A00");

            await _rejoinProxy.StartAsync();
            StartRejoinCrashWatcher();
            StartNetworkStatsOverlay(_rejoinProxy.CurrentGameServerIp);
            AppendLog("[REJOIN]", $"Rejoin Fix active on 127.0.0.1:{_rejoinProxy.Port}. Restart MCC now.", "#39FF14");
            SetStatus("Rejoin Fix active.", "#39FF14");

            if (IsMicrosoftStoreInstallation)
            {
                UpdateRejoinFirewallStatus();
                return;
            }

            bool firewallRequested = ChkRejoinFixFirewall.IsChecked == true ||
                ChkRejoinFixFirewallMatchmaking.IsChecked == true;
            if (!firewallRequested)
            {
                UpdateRejoinFirewallStatus();
                return;
            }

            try
            {
                await EnsureSteamFirewallRulesPreparedAsync();
                await ApplyRejoinFirewallOptionAsync();
                if (ChkRejoinFixFirewallMatchmaking.IsChecked == true)
                    await EnableSteamFirewallAutoModeAsync(ensureObserverRunning: false);
            }
            catch (Exception ex)
            {
                _rejoinCampaignFirewallApplying = false;
                _rejoinCampaignFirewallEnabled = false;
                AppendLog("[ERROR]", $"Firewall setup after Rejoin Fix start failed: {ex.Message}", "#FF2D55");
                SetStatus("Rejoin Fix active; firewall setup failed.", "#FF6A00");
                UpdateRejoinFirewallStatus("FIREWALL: SETUP FAILED - Rejoin Fix is still active", "#FF2D55");
            }
        }

        private async Task<bool> EnsureCompanionServicesRunningAsync(string requestedFeature)
        {
            if (_rejoinProxy.IsRunning)
                return true;

            try
            {
                if (!IsRunningAsAdministrator())
                {
                    ToolboxDialog.Show(
                        $"{requestedFeature} needs the MCC data proxy. The Toolbox will relaunch as Administrator and start it automatically.\n\nIf MCC is currently open, restart MCC afterward so traffic capture can take effect.",
                        "MCC Data Proxy -- Halo MCC Toolbox",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    App.SavePendingRejoinFixAutoStart(true);
                    RelaunchAsAdministrator();
                    Close();
                    return true;
                }

                await StartRejoinFixAsync();
                return _rejoinProxy.IsRunning;
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                HandleAdministratorRelaunchCancelled();
                return false;
            }
            catch (Exception ex)
            {
                RejoinFixDiagnostics.Error("proxy", $"Automatic service activation for {requestedFeature} failed: {ex.Message}");
                AppendLog("[ERROR]", $"MCC data proxy failed: {ex.Message}", "#FF2D55");
                SetStatus("MCC data proxy failed to start.", "#FF2D55");
                ToolboxDialog.Show(
                    $"The MCC data proxy could not start:\n\n{ex.Message}",
                    "MCC Data Proxy -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
            finally
            {
                // This path is also used by VPN and overlay auto-start, so it
                // must keep the shared Services controls in sync without
                // relying on the Start Services button's click handler.
                UpdateRejoinFixUi();
            }
        }

        private async Task StartPendingRejoinFixAfterElevationAsync()
        {
            if (!App.ConsumePendingRejoinFixAutoStart())
                return;

            if (!IsRunningAsAdministrator() || _rejoinProxy.IsRunning)
                return;

            BtnRejoinFix.IsEnabled = false;
            HomeProxyToggleButton.IsEnabled = false;
            SidebarFeaturesToggleButton.IsEnabled = false;
            try
            {
                await StartRejoinFixAsync();
            }
            catch (Exception ex)
            {
                RejoinFixDiagnostics.Error("proxy", $"Automatic activation after elevation failed: {ex.Message}");
                AppendLog("[ERROR]", $"Rejoin Fix failed: {ex.Message}", "#FF2D55");
                SetStatus("Rejoin Fix failed to start.", "#FF2D55");
                ToolboxDialog.Show(
                    $"Rejoin Fix could not start:\n\n{ex.Message}",
                    "Rejoin Fix -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                UpdateRejoinFixUi();
                BtnRejoinFix.IsEnabled = true;
                HomeProxyToggleButton.IsEnabled = true;
                SidebarFeaturesToggleButton.IsEnabled = true;
            }
        }

        private async void BtnRejoinFix_Click(object sender, RoutedEventArgs e)
        {
            BtnRejoinFix.IsEnabled = false;
            HomeProxyToggleButton.IsEnabled = false;
            SidebarFeaturesToggleButton.IsEnabled = false;

            try
            {
                if (_rejoinProxy.IsRunning)
                {
                    bool rejoinFirewallWasEnabled = ChkRejoinFixFirewall.IsChecked == true
                        || ChkRejoinFixFirewallMatchmaking.IsChecked == true
                        || _steamFirewallAutoEnabled
                        || _rejoinCampaignFirewallEnabled;
                    StopRejoinCrashWatcher();
                    _rejoinProxy.Stop();
                    StartNetworkStatsOverlay("");
                    _obsOverlayServer.Stop();
                    _rejoinWinHttpManualNeeded = false;
                    DisableSteamFirewallAutoMode(logStatus: false);
                    if (rejoinFirewallWasEnabled)
                        await DisableRejoinFirewallRulesAsync(logStatus: false);
                    _rejoinCampaignFirewallApplying = false;
                    _rejoinCampaignFirewallEnabled = false;
                    AppendLog("[REJOIN]", "MCC data proxy stopped.", "#C8D8E8");
                    SetStatus("MCC data proxy stopped.", "#C8D8E8");
                }
                else
                {
                    if (!IsRunningAsAdministrator())
                    {
                        ToolboxDialog.Show(
                            "The MCC data proxy needs the Toolbox to run as Administrator so it can use the system proxy and firewall settings.\n\nIf MCC is currently open, restart it afterward so traffic capture can take effect.\n\nThe Toolbox will relaunch as Administrator now.",
                            "MCC Data Proxy -- Halo MCC Toolbox",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);

                        App.SavePendingRejoinFixAutoStart(true);
                        RelaunchAsAdministrator();
                        Close();
                        return;
                    }

                    await StartRejoinFixAsync();
                }
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                HandleAdministratorRelaunchCancelled();
            }
            catch (Exception ex)
            {
                RejoinFixDiagnostics.Error("proxy", $"Activation failed: {ex.Message}");
                AppendLog("[ERROR]", $"Rejoin Fix failed: {ex.Message}", "#FF2D55");
                SetStatus("Rejoin Fix failed to start.", "#FF2D55");
                ToolboxDialog.Show(
                    $"Rejoin Fix could not start:\n\n{ex.Message}",
                    "Rejoin Fix -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                UpdateRejoinFixUi();
                BtnRejoinFix.IsEnabled = true;
                HomeProxyToggleButton.IsEnabled = true;
                SidebarFeaturesToggleButton.IsEnabled = true;
            }
        }

        private void HandleAdministratorRelaunchCancelled()
        {
            App.SavePendingRejoinFixAutoStart(false);
            AppendLog("[INFO]", "MCC data proxy cancelled at administrator prompt.", "#4A5A6A");
            SetStatus("MCC data proxy requires Administrator.", "#4A5A6A");
        }

        // ------------------------------------------
        // MAP SELECTOR
        // ------------------------------------------
        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFolderDialog
            {
                Title = "Select your Halo MCC installation folder",
                InitialDirectory = TxtMccPath.Text.Trim()
            };
            if (dlg.ShowDialog() == true)
            {
                TxtMccPath.Text = MccInstallationResolver.NormalizeRoot(dlg.FolderName);
                SaveMccInstallationPath();
            }
        }

        private void BtnLoadMaps_Click(object sender, RoutedEventArgs e)
        {
            SaveMccInstallationPath();
            LoadMaps(TxtMccPath.Text.Trim());
        }

        private void TxtMccPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            _playlistsMccPath = TxtMccPath.Text;
            _playlistsTab?.SetMccInstallationPath(_playlistsMccPath);
            App.SaveMccInstallationPath(TxtMccPath.Text);
            UpdateMccEditionUi();
            UpdateToolboxStatus();
        }

        private void SaveMccInstallationPath()
        {
            var mccPath = MccInstallationResolver.NormalizeRoot(TxtMccPath.Text);
            if (!string.Equals(TxtMccPath.Text, mccPath, StringComparison.OrdinalIgnoreCase))
                TxtMccPath.Text = mccPath;
            App.SaveMccInstallationPath(mccPath);
            _playlistsMccPath = mccPath;
            _playlistsTab?.SetMccInstallationPath(mccPath);
            _theaterTab?.SetMccInstallationPath(mccPath);
            UpdateMccEditionUi();
        }

        private void LoadMaps(string mccPath)
        {
            var mapsPath = Path.Combine(mccPath, "halo3", "maps");

            if (!Directory.Exists(mapsPath))
            {
                Dispatcher.InvokeAsync(() =>
                {
                    _homeMccStatusPath = mccPath;
                    _homeMccStatusFound = false;
                    TxtMapStatus.Text = $"Maps folder not found: {mapsPath}";
                    TxtMapStatus.Foreground = Brush("#FF2D55");
                    UpdateToolboxStatus();
                });
                AppendLog("[ERROR]", $"Halo 3 maps folder not found: {mapsPath}", "#FF2D55");
                return;
            }

            AppendLog("[INFO]", $"Scanning: {mapsPath}", "#00C8FF");

            var officialEntries = new List<MapEntry>();
            var moddedEntries   = new List<MapEntry>();

            foreach (var file in Directory.GetFiles(mapsPath, "*.map", SearchOption.TopDirectoryOnly))
            {
                var fileName  = Path.GetFileNameWithoutExtension(file);
                bool isRemoved = fileName.StartsWith(RemovedPrefix, StringComparison.OrdinalIgnoreCase);
                var baseName  = isRemoved ? fileName[RemovedPrefix.Length..] : fileName;

                // Skip system/shared maps
                if (SystemMapNames.Contains(baseName)) continue;

                // Skip campaign maps -- filenames starting with a digit (010_jungle, etc.)
                if (baseName.Length > 0 && char.IsDigit(baseName[0])) continue;

                var entry = new MapEntry
                {
                    FileName    = file,
                    BaseName    = baseName,
                    IsEnabled   = !isRemoved,
                    IsModded    = false,
                };

                if (OfficialMaps.TryGetValue(baseName, out var friendlyName))
                {
                    entry.DisplayName = friendlyName;
                    officialEntries.Add(entry);
                }
                else
                {
                    // Unknown file -- treat as modded map, show filename as display name
                    entry.DisplayName = baseName;
                    entry.IsModded    = true;
                    moddedEntries.Add(entry);
                }
            }

            int officialCount = officialEntries.Count;
            int moddedCount   = moddedEntries.Count;

            // Marshal all collection updates to the UI thread
            Dispatcher.InvokeAsync(() =>
            {
                _homeMccStatusPath = mccPath;
                _homeMccStatusFound = true;
                _maps.Clear();

                // Add official maps sorted alphabetically
                foreach (var e in officialEntries.OrderBy(e => e.DisplayName, StringComparer.OrdinalIgnoreCase))
                    _maps.Add(e);

                // Add modded maps separator + entries (sorted alphabetically)
                if (moddedEntries.Count > 0)
                {
                    _maps.Add(new MapEntry
                    {
                        DisplayName = "-- MODDED MAPS --",
                        IsHeader    = true,
                        IsEnabled   = true,
                        FileName    = "",
                        BaseName    = "",
                    });

                    foreach (var e in moddedEntries.OrderBy(e => e.DisplayName, StringComparer.OrdinalIgnoreCase))
                        _maps.Add(e);
                }

                // Update UI status
                if (officialCount == 0 && moddedCount == 0)
                {
                    TxtMapStatus.Text = $"No multiplayer map files found in: {mapsPath}";
                    TxtMapStatus.Foreground = Brush("#FF6A00");
                    AppendLog("[WARN]", "No maps found. Check your MCC path.", "#FF6A00");
                }
                else
                {
                    var msg = moddedCount > 0
                        ? $"Loaded {officialCount} official maps, {moddedCount} modded maps."
                        : $"Loaded {officialCount} maps.";
                    TxtMapStatus.Text = "";
                    AppendLog("[INFO]", msg, "#39FF14");
                    SetStatus(msg, "#39FF14");
                }

                UpdateToolboxStatus();
            });
        }

        // Clicking a row toggles its enabled state (headers are non-interactive)
        private void MapRow_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is MapEntry map && !map.IsHeader)
                SetMapEnabled(map, !map.IsEnabled);
        }

        private void BtnEnableAll_Click(object sender, RoutedEventArgs e)
        {
            var maps = _maps.Where(m => !m.IsHeader).ToList();
            if (maps.Count == 0)
            {
                ShowNoMapsLoadedMessage();
                return;
            }

            ApplyMapState(maps, true, "Enable All");
        }

        private void BtnDisableAll_Click(object sender, RoutedEventArgs e)
        {
            var maps = _maps.Where(m => !m.IsHeader).ToList();
            if (maps.Count == 0)
            {
                ShowNoMapsLoadedMessage();
                return;
            }

            ApplyMapState(maps, false, "Disable All");
        }

        private void BtnDisable343_Click(object sender, RoutedEventArgs e)
        {
            var maps = _maps.Where(m => !m.IsHeader && Map343Names.Contains(m.BaseName)).ToList();
            if (maps.Count == 0)
            {
                AppendLog("[WARN]", "No 343 maps found. Load maps first.", "#FF6A00");
                return;
            }

            ApplyMapState(maps, false, "Disable 343 Maps");
        }

        private void ShowNoMapsLoadedMessage()
        {
            ToolboxDialog.Show("No maps loaded. Load your maps first.", "Halo MCC Toolbox",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void ApplyMapState(IReadOnlyCollection<MapEntry> maps, bool enabled, string actionLabel)
        {
            int changed = 0, unchanged = 0, fail = 0;
            foreach (var map in maps)
            {
                var result = SetMapEnabled(map, enabled);
                if (result == MapToggleResult.Changed)
                    changed++;
                else if (result == MapToggleResult.Unchanged)
                    unchanged++;
                else
                    fail++;
            }

            var col = fail > 0 ? "#FF6A00" : "#39FF14";
            SetStatus($"{actionLabel}: {changed} changed, {unchanged} already set, {fail} failed.", col);
            AppendLog("[DONE]", $"{actionLabel}: {changed} changed, {unchanged} already set, {fail} failed.", col);
        }

        private MapToggleResult SetMapEnabled(MapEntry map, bool enabled)
        {
            if (map.IsHeader)
                return MapToggleResult.Unchanged;

            try
            {
                var dir = Path.GetDirectoryName(map.FileName);
                if (string.IsNullOrWhiteSpace(dir))
                    throw new InvalidOperationException("Map path is missing.");

                var ext = Path.GetExtension(map.FileName);
                var target = enabled
                    ? Path.Combine(dir, map.BaseName + ext)
                    : Path.Combine(dir, RemovedPrefix + map.BaseName + ext);

                if (map.FileName.Equals(target, StringComparison.OrdinalIgnoreCase))
                {
                    map.IsEnabled = enabled;
                    return MapToggleResult.Unchanged;
                }

                File.Move(map.FileName, target);
                AppendLog(enabled ? "[ENABLE]" : "[REMOVE]",
                    $"{Path.GetFileName(map.FileName)}  =>  {Path.GetFileName(target)}",
                    enabled ? "#39FF14" : "#FF2D55");
                map.FileName = target;
                map.IsEnabled = enabled;
                SetStatus($"{map.DisplayName} {(enabled ? "enabled" : "disabled")}.", enabled ? "#39FF14" : "#FF2D55");
                return MapToggleResult.Changed;
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Failed to rename {map.BaseName}: {ex.Message}", "#FF2D55");
                SetStatus($"Failed to toggle {map.DisplayName}.", "#FF2D55");
                return MapToggleResult.Failed;
            }
        }

        private static SolidColorBrush Brush(string hex) =>
            (SolidColorBrush)new BrushConverter().ConvertFrom(hex)!;

        // ======================================================
        // REPORT TAB -- state
        // ======================================================

        private ObservableCollection<PlayerEntry> _players = new();
        private string? _carnageFilePath;   // full path to loaded XML
        private string? _lastReportZipPath; // full path to most recently built ZIP
        private string  _selectedGame = "Halo 3";

        // Returns the per-game theater Movie folder path (empty string if not supported)
        private static string GetTheaterRoot(string game)
        {
            var up = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var folder = game switch
            {
                "Halo 3"     => "Halo3",
                "Halo Reach" => "HaloReach",
                "Halo 4"     => "Halo4",
                _            => ""
            };
            return string.IsNullOrEmpty(folder) ? "" :
                Path.Combine(up, "AppData", "LocalLow", "MCC", "Temporary", "UserContent", folder, "Movie");
        }

        // Theater .mov filenames follow the pattern:  asq_<first7chars_of_internal_name>_<hash>.mov
        // e.g.  guardian    => asq_guardia_xxxx.mov
        //       salvation   => asq_salvati_xxxx.mov
        //       chillout    => asq_chillou_xxxx.mov
        //       chill       => asq_chill_xxxx.mov   (only 5 chars, keeps underscore)
        //       s3d_waterfall => asq_s3d_wat_xxxx.mov  (truncated after 7 chars of base)
        // We match by checking if the filename STARTS WITH the prefix (case-insensitive).
        private static readonly Dictionary<string, string> MapToTheaterPrefix =
            new(StringComparer.OrdinalIgnoreCase)
        {
            ["Avalanche"]    = "asq_sidewin",
            ["Assembly"]     = "asq_descent",
            ["Blackout"]     = "asq_lockout",
            ["Citadel"]      = "asq_fortres",
            ["Cold Storage"] = "asq_chillou",
            ["Construct"]    = "asq_constru",
            ["Edge"]         = "asq_s3d_edg",
            ["Epitaph"]      = "asq_salvati",
            ["Foundry"]      = "asq_warehou",
            ["Ghost Town"]   = "asq_ghostto",
            ["Guardian"]     = "asq_guardia",
            ["Heretic"]      = "asq_midship",
            ["High Ground"]  = "asq_deadloc",
            ["Icebox"]       = "asq_s3d_tur",
            ["Isolation"]    = "asq_isolati",
            ["Last Resort"]  = "asq_zanziba",
            ["Longshore"]    = "asq_docks_",
            ["Narrows"]      = "asq_chill_",   // "chill" is only 5 chars -- trailing _ prevents matching "chillou"
            ["Orbital"]      = "asq_spaceca",
            ["Rat's Nest"]   = "asq_armory_",
            ["Sandbox"]      = "asq_sandbox",
            ["Sandtrap"]     = "asq_shrine_",
            ["Snowbound"]    = "asq_snowbou",
            ["Standoff"]     = "asq_bunkerw",
            ["The Pit"]      = "asq_cyberde",
            ["Valhalla"]     = "asq_riverwo",
            ["Waterfall"]    = "asq_s3d_wat",
        };

        // ------------------------------------------
        // Load carnage report XML
        // ------------------------------------------
        private void ReportScoreboard_Loaded(object sender, RoutedEventArgs e)
        {
            // Keep the selected match and draft intact when returning to this tab.
            if (_players.Count > 0) return;
            try
            {
                var tempDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    "AppData", "LocalLow", "MCC", "Temporary");
                var latest = Directory.Exists(tempDir)
                    ? Directory.GetFiles(tempDir, "mpcarnagereport*.xml")
                        .OrderByDescending(File.GetLastWriteTimeUtc).FirstOrDefault()
                    : null;
                if (latest == null)
                {
                    TxtReportStatus.Text = "No recent game found. Play a match, then click LOAD LAST GAME.";
                    return;
                }
                _carnageFilePath = latest;
                ParseCarnageReport(latest);
            }
            catch (Exception ex)
            {
                TxtReportStatus.Text = "Could not load the last game: " + ex.Message;
            }
        }
        private void BtnLoadCarnage_Click(object sender, RoutedEventArgs e)
        {
            var tempDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "AppData", "LocalLow", "MCC", "Temporary");

            if (!Directory.Exists(tempDir))
            {
                ToolboxDialog.Show($"MCC Temporary folder not found:\n{tempDir}",
                    "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Find the most recently modified mpcarnagereport*.xml
            var carnageFiles = Directory.GetFiles(tempDir, "mpcarnagereport*.xml", SearchOption.TopDirectoryOnly)
                .OrderByDescending(File.GetLastWriteTime)
                .ToArray();

            if (carnageFiles.Length == 0)
            {
                // Let user browse manually
                var dlg = new OpenFileDialog
                {
                    Title            = "Select a carnage report XML",
                    Filter           = "Carnage Report XML|mpcarnagereport*.xml|All XML|*.xml",
                    InitialDirectory = tempDir
                };
                if (dlg.ShowDialog() != true) return;
                _carnageFilePath = dlg.FileName;
            }
            else
            {
                _carnageFilePath = carnageFiles[0];
            }

            ParseCarnageReport(_carnageFilePath);
        }

        private void ParseCarnageReport(string xmlPath)
        {
            try
            {
                var xml  = XDocument.Load(xmlPath);
                var root = xml.Root;
                if (root == null) throw new Exception("Empty XML file.");

                // -- Game metadata --------------------------------------------------
                // GameTypeName uses the same string as both element name and attribute name
                var gameTypeName = root.Element("GameTypeName")?.Attribute("GameTypeName")?.Value
                                ?? root.Element("GameTypeName")?.Value
                                ?? "Unknown";

                // No map name is stored in the XML -- we rely on manual map selection in the form.
                // Derive a label from the filename as a hint (e.g. mpcarnagereport1_3528_0_0)
                var fileHint = Path.GetFileNameWithoutExtension(xmlPath);

                var isMatchmaking = root.Element("IsMatchmaking")?.Attribute("IsMatchmaking")?.Value ?? "false";
                var isTeams       = root.Element("IsTeamsEnabled")?.Attribute("IsTeamsEnabled")?.Value ?? "false";

                // File write time is the closest we have to a game timestamp
                var gameDate = File.GetLastWriteTime(xmlPath).ToString("yyyy-MM-dd  HH:mm");

                // -- Update info bar ------------------------------------------------
                TxtGameMap.Text    = "-- select below --";
                TxtGameMode.Text   = gameTypeName;
                TxtGameDate.Text   = gameDate;
                TxtCarnageFile.Text = Path.GetFileName(xmlPath);
                GameInfoBar.Visibility = Visibility.Visible;

                // -- Parse players --------------------------------------------------
                _players.Clear();
                ScoreboardList.ItemsSource = _players;

                var playerElements = root.Element("Players")?.Elements("Player").ToList()
                                  ?? new List<XElement>();

                if (playerElements.Count == 0)
                    throw new Exception("No <Player> elements found inside <Players>.\n\nThe file may be from a different game or is malformed.");

                var entries = new List<PlayerEntry>();
                foreach (var el in playerElements)
                {
                    // All stats are XML attributes directly on <Player>
                    string Attr(string name) => el.Attribute(name)?.Value ?? "";
                    int    Int(string name)  => int.TryParse(Attr(name), out var v) ? v : 0;

                    entries.Add(new PlayerEntry
                    {
                        Gamertag   = Attr("mGamertagText"),
                        XboxUserId = Attr("mXboxUserId"),   // important for reporting -- real ID
                        Score      = Int("Score"),
                        Kills      = Int("mKills"),
                        Deaths     = Int("mDeaths"),
                        Assists    = Int("mAssists"),
                        Betrayals  = Int("mBetrayals"),
                        Suicides   = Int("mSuicides"),
                        Team       = Int("mTeamId") switch { 0 => "Red", 1 => "Blue", 2 => "Green", 3 => "Yellow", _ => Attr("mTeamId") },
                        Completed  = Attr("mCompletedGame") == "1",
                    });
                }

                // Sort: score desc, then kills desc
                foreach (var p in entries.OrderByDescending(p => p.Score).ThenByDescending(p => p.Kills))
                    _players.Add(p);

                // -- Populate map combo ---------------------------------------------
                var currentGame = (CboGameTitle.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Halo 3";
                PopulateReportMapCombo(null, currentGame); // no map in XML -- user must pick

                TxtSelectedPlayer.Text       = "Click a player on the scoreboard above";
                TxtSelectedPlayer.Foreground = Brush("#4A5A6A");
                TxtReportStatus.Text         = $"Loaded {_players.Count} players  .  {gameTypeName}  .  {gameDate}";
                TxtReportStatus.Foreground   = Brush("#39FF14");

                AppendLog("[REPORT]", $"Loaded {_players.Count} players. Game type: {gameTypeName}. File: {Path.GetFileName(xmlPath)}", "#00C8FF");
            }
            catch (Exception ex)
            {
                ToolboxDialog.Show($"Failed to parse carnage report:\n\n{ex.Message}\n\nPath: {xmlPath}",
                    "Parse Error", MessageBoxButton.OK, MessageBoxImage.Error);
                AppendLog("[ERROR]", $"Carnage parse failed: {ex.Message}", "#FF2D55");
            }
        }

        private void PopulateReportMapCombo(string? preselect, string? game = null)
        {
            // Guard: CboReportMap may not yet exist if SelectionChanged fires during InitializeComponent
            if (CboReportMap == null) return;

            game ??= (CboGameTitle.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Halo 3";

            CboReportMap.Items.Clear();
            var maps = GameMaps.TryGetValue(game, out var list)
                ? list.OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                : Enumerable.Empty<string>();

            foreach (var name in maps)
            {
                var item = new ComboBoxItem { Content = name };
                CboReportMap.Items.Add(item);
                if (!string.IsNullOrEmpty(preselect) &&
                    name.Equals(preselect, StringComparison.OrdinalIgnoreCase))
                    CboReportMap.SelectedItem = item;
            }
            var other = new ComboBoxItem { Content = "Other / Unknown" };
            CboReportMap.Items.Add(other);
            if (CboReportMap.SelectedIndex < 0)
                CboReportMap.SelectedIndex = 0;

            // Wire change event (remove first to avoid double-subscribe)
            CboReportMap.SelectionChanged -= CboReportMap_SelectionChanged;
            CboReportMap.SelectionChanged += CboReportMap_SelectionChanged;
            UpdateTheaterCount();
        }

        private void CboGameTitle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var game = (CboGameTitle.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Halo 3";
            _selectedGame = game;
            PopulateReportMapCombo(null, game);
        }

        private void CboReportMap_SelectionChanged(object sender, SelectionChangedEventArgs e)
            => UpdateTheaterCount();

        private void UpdateTheaterCount()
        {
            // Guard: controls may not yet exist during InitializeComponent ordering
            if (CboGameTitle == null || CboReportMap == null || TxtTheaterCount == null) return;

            var game = (CboGameTitle.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";

            // Show/hide the entire theater row based on whether the game supports Film mode
            if (TheaterPanel != null)
                TheaterPanel.Visibility = GamesWithTheater.Contains(game)
                    ? Visibility.Visible
                    : Visibility.Collapsed;

            if (!GamesWithTheater.Contains(game)) return;

            var mapName = (CboReportMap.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (string.IsNullOrEmpty(mapName) || mapName == "Other / Unknown")
            {
                TxtTheaterCount.Text       = "— select a map first —";
                TxtTheaterCount.Foreground = Brush("#4A5A6A");
                return;
            }

            var files = GetTheaterFilesForMap(mapName);
            if (files.Length == 0)
            {
                TxtTheaterCount.Text       = "0 files found";
                TxtTheaterCount.Foreground = Brush("#FF6A00");
            }
            else
            {
                TxtTheaterCount.Text       = $"{files.Length} .mov file(s) found  ✓";
                TxtTheaterCount.Foreground = Brush("#39FF14");
            }
        }

        // ------------------------------------------
        // Halo Support session status check
        // ------------------------------------------
        private Task EnsureSupportSessionCheckedAsync()
            => _supportSessionCheckTask ??= CheckSupportSessionAsync();

        /// <summary>
        /// Checks whether a Halo Support / Microsoft Account session is stored in the
        /// persistent WebView2 profile and updates TxtSupportSessionStatus accordingly.
        ///
        /// Strategy (two independent signals — either one = green):
        ///   1. login.live.com  — look for RPSSecAuth / MSPAuth (MS Account "stay signed in")
        ///   2. support.halowaypoint.com — look for Zendesk session / auth cookies
        ///
        /// The hidden WebView2 is created lazily and shares the same CoreWebView2Environment
        /// as HaloReportWindow so it reads from the same on-disk cookie store. It must not
        /// be declared in MainWindow.xaml because that can block the first render on Edge.
        /// </summary>
        private async Task CheckSupportSessionAsync()
        {
            // Show "checking..." while the async work runs
            Dispatcher.Invoke(() =>
            {
                TxtSupportSessionStatus.Text       = "● checking session…";
                TxtSupportSessionStatus.Foreground = Brush("#4A5A6A");
            });

            try
            {
                // Initialize the hidden WebView2 with the shared persistent environment.
                // EnsureCoreWebView2Async is idempotent — safe to call multiple times.
                var checker = EnsureHiddenCookieChecker();
                var env = await WebViewEnvironmentManager.GetOrCreateAsync();
                await checker.EnsureCoreWebView2Async(env);

                var mgr = checker.CoreWebView2.CookieManager;

                // ── Signal 1: Microsoft Account "Stay signed in" cookies ──────────────
                // RPSSecAuth and MSPAuth are the persistent auth cookies set by
                // login.live.com when the user chooses "Stay signed in".
                var liveCookies = await mgr.GetCookiesAsync("https://login.live.com");
                bool hasMsAuth = liveCookies.Any(c =>
                    c.Name.Equals("RPSSecAuth", StringComparison.OrdinalIgnoreCase) ||
                    c.Name.Equals("MSPAuth",    StringComparison.OrdinalIgnoreCase) ||
                    c.Name.Equals("MSCC",       StringComparison.OrdinalIgnoreCase));

                // ── Signal 2: Zendesk / Halo Support session cookies ──────────────────
                var haloCookies = await mgr.GetCookiesAsync("https://support.halowaypoint.com");
                bool hasHaloSession = haloCookies.Any(c =>
                    c.Name.IndexOf("session",  StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.Name.IndexOf("auth",     StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.Name.IndexOf("zendesk",  StringComparison.OrdinalIgnoreCase) >= 0);

                bool isLoggedIn = hasMsAuth || hasHaloSession;

                Dispatcher.Invoke(() =>
                {
                    if (isLoggedIn)
                    {
                        TxtSupportSessionStatus.Text       = "● session active";
                        TxtSupportSessionStatus.Foreground = Brush("#39FF14");
                    }
                    else
                    {
                        TxtSupportSessionStatus.Text       = "● login required";
                        TxtSupportSessionStatus.Foreground = Brush("#FF2D55");
                    }
                });
            }
            catch
            {
                // Swallow — this is a best-effort status check, not critical path
                Dispatcher.Invoke(() =>
                {
                    TxtSupportSessionStatus.Text       = "● status unknown";
                    TxtSupportSessionStatus.Foreground = Brush("#FF6A00");
                });
            }
        }

        private string[] GetTheaterFilesForMap(string friendlyName)
        {
            var game = (CboGameTitle.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Halo 3";
            var root = GetTheaterRoot(game);
            if (!Directory.Exists(root)) return Array.Empty<string>();

            // Halo 3: filter by per-map filename prefix (asq_ pattern)
            if (string.Equals(game, "Halo 3", StringComparison.OrdinalIgnoreCase) &&
                MapToTheaterPrefix.TryGetValue(friendlyName, out var prefix))
            {
                return Directory.GetFiles(root, "*.mov", SearchOption.AllDirectories)
                    .Where(f => Path.GetFileName(f).StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            }

            // Reach / H4: no prefix map available — return all .mov files in the game's folder
            return Directory.GetFiles(root, "*.mov", SearchOption.AllDirectories);
        }

        // ------------------------------------------
        // Scoreboard row click -- select/deselect player
        // ------------------------------------------
        private void ScoreboardRow_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is PlayerEntry clicked)
            {
                // Toggle -- clicking an already-selected player deselects
                foreach (var p in _players) p.IsSelected = false;
                if (clicked != null)
                {
                    clicked.IsSelected = true;
                    TxtSelectedPlayer.Text       = clicked.Gamertag;
                    TxtSelectedPlayer.Foreground = Brush("#FF2D55");
                    TxtReportStatus.Text         = $"Reporting: {clicked.Gamertag}  --  Fill in the form below and click BUILD REPORT ZIP.";
                    TxtReportStatus.Foreground   = Brush("#FF6A00");
                }
            }
        }

        // ------------------------------------------
        // Build Report ZIP
        // ------------------------------------------
        private void BtnBuildReport_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            var suspect = _players.FirstOrDefault(p => p.IsSelected);
            if (suspect == null)
            {
                ToolboxDialog.Show("Please select the player to report on the scoreboard first.",
                    "No Player Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var mapName = (CboReportMap.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Unknown";
            var reportReason = (CboReportReason.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Not specified";
            var notes = TxtReportNotes.Text.Trim();

            if (CboReportReason.SelectedIndex < 0)
            {
                ToolboxDialog.Show("Please select a report reason.", "Missing Info",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Ask where to save
            var safeTag = string.Concat(suspect.Gamertag.Select(c =>
                Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
            var dlg = new SaveFileDialog
            {
                Title            = "Save Player Report ZIP",
                Filter           = "ZIP Archive (*.zip)|*.zip",
                FileName         = $"PlayerReport_{safeTag}_{DateTime.Now:yyyyMMdd_HHmmss}.zip",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (dlg.ShowDialog() != true) return;

            var zipPath = dlg.FileName;
            _lastReportZipPath = zipPath; // remember for Explorer highlight when submitting
            BtnBuildReport.IsEnabled = false;

            // Gather theater files
            var theaterFiles = mapName == "Other / Unknown"
                ? Array.Empty<string>()
                : GetTheaterFilesForMap(mapName);

            // Snapshot all players for the report
            var allPlayers   = _players.ToList();
            var carnagePath  = _carnageFilePath;
            var selectedGame = _selectedGame;
            var gameMode     = TxtGameMode.Text;
            var gameDate     = TxtGameDate.Text;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    if (File.Exists(zipPath)) File.Delete(zipPath);

                    using var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create);

                    // -- 1. Human-readable report TXT ----------------------
                    var sb = new StringBuilder();
                    sb.AppendLine("=======================================================");
                    sb.AppendLine("  HALO MCC -- PLAYER REPORT");
                    sb.AppendLine("  Generated by Halo MCC Toolbox  /  The FFA Panda");
                    sb.AppendLine($"  {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine("=======================================================");
                    sb.AppendLine();
                    sb.AppendLine("[ REPORTED PLAYER ]");
                    sb.AppendLine($"  Gamertag    : {suspect.Gamertag}");
                    sb.AppendLine($"  Xbox User ID: {suspect.XboxUserId}");
                    sb.AppendLine($"  Behavior    : {reportReason}");
                    sb.AppendLine();
                    sb.AppendLine("[ GAME DETAILS ]");
                    sb.AppendLine($"  Game      : {selectedGame}");
                    sb.AppendLine($"  Map       : {mapName}");
                    sb.AppendLine($"  Mode      : {gameMode}");
                    sb.AppendLine($"  Date/Time : {gameDate}");
                    sb.AppendLine();
                    if (!string.IsNullOrEmpty(notes))
                    {
                        sb.AppendLine("[ DESCRIPTION ]");
                        foreach (var line in notes.Split('\n'))
                            sb.AppendLine($"  {line.TrimEnd()}");
                        sb.AppendLine();
                    }
                    sb.AppendLine("[ FULL SCOREBOARD ]");
                    sb.AppendLine($"  {"GAMERTAG",-24} {"SCORE",6} {"KILLS",6} {"DEATHS",7} {"ASSISTS",8} {"BETR",5} {"TEAM",7}");
                    sb.AppendLine($"  {new string('-', 68)}");
                    foreach (var p in allPlayers)
                    {
                        var marker = p.IsSelected ? " << REPORTED" : "";
                        sb.AppendLine($"  {p.Gamertag,-24} {p.Score,6} {p.Kills,6} {p.Deaths,7} {p.Assists,8} {p.Betrayals,5} {p.Team,7}{marker}");
                    }
                    sb.AppendLine();
                    sb.AppendLine("[ XBOX USER IDs  (for reporting to 343 / Microsoft) ]");
                    foreach (var p in allPlayers)
                    {
                        var marker = p.IsSelected ? " << REPORTED" : "";
                        sb.AppendLine($"  {p.Gamertag,-24}  {p.XboxUserId}{marker}");
                    }
                    sb.AppendLine();
                    if (theaterFiles.Length > 0)
                    {
                        sb.AppendLine("[ THEATER FILES INCLUDED ]");
                        foreach (var f in theaterFiles)
                            sb.AppendLine($"  {Path.GetFileName(f)}");
                        sb.AppendLine();
                    }
                    sb.AppendLine("[ HOW TO REPORT ]");
                    sb.AppendLine("  1. Go to https://www.halowaypoint.com/en-us/support");
                    sb.AppendLine("  2. Submit a player report with this information.");
                    sb.AppendLine("  3. Attach the carnage report XML and theater files from this ZIP.");
                    sb.AppendLine("  4. You can also report via the in-game Recent Players list.");

                    var reportEntry = zip.CreateEntry("report.txt");
                    using (var writer = new StreamWriter(reportEntry.Open(), Encoding.UTF8))
                        writer.Write(sb.ToString());

                    AppendLog("[ZIP]", "report.txt", "#C8D8E8");

                    // -- 2. Carnage report XML -----------------------------
                    if (!string.IsNullOrEmpty(carnagePath) && File.Exists(carnagePath))
                    {
                        zip.CreateEntryFromFile(carnagePath,
                            $"carnage_report/{Path.GetFileName(carnagePath)}",
                            CompressionLevel.Fastest);
                        AppendLog("[ZIP]", Path.GetFileName(carnagePath), "#C8D8E8");
                    }

                    // -- 3. Theater .mov files -----------------------------
                    foreach (var mov in theaterFiles)
                    {
                        zip.CreateEntryFromFile(mov,
                            $"theater_files/{Path.GetFileName(mov)}",
                            CompressionLevel.Fastest);
                        AppendLog("[ZIP]", $"theater_files/{Path.GetFileName(mov)}", "#C8D8E8");
                    }

                    var info    = new FileInfo(zipPath);
                    var sizeKb  = info.Length / 1024.0;
                    var sizeTxt = sizeKb >= 1024 ? $"{sizeKb/1024:F1} MB" : $"{sizeKb:F0} KB";

                    AppendLog("[DONE]",
                        $"Report ZIP created: {theaterFiles.Length} theater file(s), {sizeTxt}  =>  {zipPath}",
                        "#39FF14");

                    Dispatcher.Invoke(() =>
                    {
                        TxtReportStatus.Text       = $"Report built -- {theaterFiles.Length} theater file(s), {sizeTxt}.";
                        TxtReportStatus.Foreground = Brush("#39FF14");

                        var open = ToolboxDialog.Show(
                            $"Report ZIP created!\n\n" +
                            $"  Suspect   : {suspect.Gamertag}\n" +
                            $"  Game      : {selectedGame}\n" +
                            $"  Map       : {mapName}\n" +
                            $"  Behavior  : {reportReason}\n" +
                            $"  Theater   : {theaterFiles.Length} file(s) included\n" +
                            $"  Size      : {sizeTxt}\n\n" +
                            $"Saved to:\n{zipPath}\n\nOpen containing folder?",
                            "Report Built -- Halo MCC Toolbox",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Information);

                        if (open == MessageBoxResult.Yes)
                            Process.Start("explorer.exe", $"/select,\"{zipPath}\"");
                    });
                }
                catch (Exception ex)
                {
                    AppendLog("[ERROR]", $"Report build failed: {ex.Message}", "#FF2D55");
                    Dispatcher.Invoke(() =>
                    {
                        TxtReportStatus.Text       = $"Failed: {ex.Message}";
                        TxtReportStatus.Foreground = Brush("#FF2D55");
                        ToolboxDialog.Show($"Failed to build report:\n\n{ex.Message}",
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
                finally
                {
                    Dispatcher.Invoke(() => BtnBuildReport.IsEnabled = true);
                }
            });
        }



        // ------------------------------------------
        // Open Halo Support ticket form (WebView2 popup)
        // ------------------------------------------
        private void BtnSubmitHalo_Click(object sender, RoutedEventArgs e)
        {
            var suspect = _players.FirstOrDefault(p => p.IsSelected);
            if (suspect == null)
            {
                ToolboxDialog.Show("Please select the player to report on the scoreboard first.",
                    "No Player Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (CboReportReason.SelectedIndex < 0)
            {
                ToolboxDialog.Show("Please select a report reason first.",
                    "Missing Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var mapName   = (CboReportMap.SelectedItem  as ComboBoxItem)?.Content?.ToString() ?? "Unknown";
            var reportReason = (CboReportReason.SelectedItem  as ComboBoxItem)?.Content?.ToString() ?? "";
            var gameType  = TxtGameMode.Text;
            var gameDate  = TxtGameDate.Text.Trim();

            // Build scoreboard text
            var sbText = new StringBuilder();
            sbText.AppendLine($"{"GAMERTAG",-24} {"SCORE",6} {"KILLS",6} {"DEATHS",7} {"ASST",6} {"TEAM",6}");
            sbText.AppendLine(new string('-', 60));
            foreach (var p in _players)
            {
                var marker = p.IsSelected ? " << REPORTED" : "";
                sbText.AppendLine($"{p.Gamertag,-24} {p.Score,6} {p.Kills,6} {p.Deaths,7} {p.Assists,6} {p.Team,6}{marker}");
            }

            var win = new HaloReportWindow
            {
                Owner           = this,
                SuspectGamertag = suspect.Gamertag,
                SuspectXboxId   = suspect.XboxUserId,
                ReportReason    = reportReason,
                BehaviorLabel   = (CboReportReason.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "",
                GameTitle       = _selectedGame,
                MapName         = mapName,
                GameType        = gameType,
                GameDate        = gameDate,
                Notes           = TxtReportNotes.Text.Trim(),
                Scoreboard      = sbText.ToString(),
                ZipPath         = _lastReportZipPath ?? "",
            };
            // Re-check session status when the support window closes so we reflect
            // any login that just happened (or a session that was revoked).
            win.Closed += (_, _) =>
            {
                _supportSessionCheckTask = null;
                _ = EnsureSupportSessionCheckedAsync();
            };

            win.Show();
            AppendLog("[REPORT]", "Opened Halo Support form for: " + suspect.Gamertag, "#00C8FF");

            // Open Explorer with the ZIP highlighted so the user can drag it into the form
            if (!string.IsNullOrEmpty(_lastReportZipPath) && File.Exists(_lastReportZipPath))
            {
                System.Diagnostics.Process.Start("explorer.exe",
                    "/select,\"" + _lastReportZipPath + "\"");
                AppendLog("[REPORT]", "Opened Explorer -- drag the ZIP into the browser to attach it.", "#FFD700");
            }
        }

        private static int ParseInt(string? s)
            => int.TryParse(s, out var i) ? i : 0;

        // ══════════════════════════════════════════════════════════════════════
        // Stats Tab — Initialization
        // ══════════════════════════════════════════════════════════════════════

        private void StatsInitialize()
        {
            StatsHttp.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "Mozilla/5.0");
            StatsCurrentLobbyList.ItemsSource = _statsCurrentLobbyRows;
            StatsLobbyList.ItemsSource = _statsLobbyRows;
            StatsSessionTimelineList.ItemsSource = _statsSessionGames;
            StatsSessionPlayersList.ItemsSource = _statsSessionPlayers;
            StatsShowLobbyView();

            StatsLoadGamertag();
            StatsLoadPersistentCache();
            StatsRefreshLifetimeUI();
            StatsLoadSpartanToken();

            StatsGamertagBox.Text = _statsGamertag;
            StatsInitializeSignature();
            StatsUpdateHwStatus();

            if (!string.IsNullOrWhiteSpace(_statsGamertag))
            {
                _statsWaypointTokenTimer.Start();
                _ = StatsInitializeWaypointSessionAsync(_statsGamertag);
            }

            StatsLoadLastGameOnStartup();
        }

        private void StatsLoadLastGameOnStartup()
        {
            if (!Directory.Exists(StatsWatchPath))
                return;

            var file = StatsLatestCarnageFile();
            if (file is null)
                return;

            StatsSetStatus($"Loading last game: {file.Name}");
            Task.Run(() => StatsTryProcessFile(file.FullName, countTowardSession: false));
        }

        // ══════════════════════════════════════════════════════════════════════
        // Stats Tab — Event handlers
        // ══════════════════════════════════════════════════════════════════════

        private void StatsApplyBtn_Click(object sender, RoutedEventArgs e)
            => StatsApplyGamertag();

        private void StatsGamertagBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) StatsApplyGamertag();
        }

        private async void StatsSyncBtn_Click(object sender, RoutedEventArgs e)
        {
            string gt; lock (_statsLock) { gt = _statsGamertag; }
            if (string.IsNullOrWhiteSpace(gt))
                return;

            // A manual sync is also an explicit request to retry the persisted
            // Waypoint session. This lets a temporary silent-refresh failure heal
            // immediately instead of waiting for the next maintenance interval.
            await StatsMaintainWaypointTokenAsync();
            await StatsFetchStats(gt);
        }

        private void StatsResetBtn_Click(object sender, RoutedEventArgs e)
        {
            lock (_statsLock)
            {
                _statsSession.Reset();
                _statsSessionPlayerHistory.Clear();
                _postGameRecap = null;
            }
            Dispatcher.Invoke(() =>
            {
                _statsSessionGames.Clear();
                _statsSessionPlayers.Clear();
                StatsSessionGameCountLabel.Text = "0";
                StatsSessionPlayerCountLabel.Text = "0";
            });
            StatsRefreshSessionUI();
            StatsSetStatus("Session reset.");
        }

        private void StatsLobbyView_Click(object sender, RoutedEventArgs e) => StatsShowLobbyView();

        private void StatsSessionView_Click(object sender, RoutedEventArgs e) => StatsShowSessionView();

        private void StatsSessionGamesTab_Click(object sender, RoutedEventArgs e)
        {
            StatsSessionGamesPanel.Visibility = Visibility.Visible;
            StatsSessionPlayersPanel.Visibility = Visibility.Collapsed;
            StatsSessionGamesTab.IsChecked = true;
            StatsSessionPlayersTab.IsChecked = false;
        }

        private void StatsSessionPlayersTab_Click(object sender, RoutedEventArgs e)
        {
            StatsSessionGamesPanel.Visibility = Visibility.Collapsed;
            StatsSessionPlayersPanel.Visibility = Visibility.Visible;
            StatsSessionGamesTab.IsChecked = false;
            StatsSessionPlayersTab.IsChecked = true;
        }

        private void StatsShowLobbyView()
        {
            if (StatsLobbyContent is null || StatsSessionContent is null) return;
            StatsLobbyContent.Visibility = Visibility.Visible;
            StatsSessionContent.Visibility = Visibility.Collapsed;
            StatsLobbyViewBtn.IsChecked = true;
            StatsSessionViewBtn.IsChecked = false;
        }

        private void StatsShowSessionView()
        {
            if (StatsLobbyContent is null || StatsSessionContent is null) return;
            StatsLobbyContent.Visibility = Visibility.Collapsed;
            StatsSessionContent.Visibility = Visibility.Visible;
            StatsLobbyViewBtn.IsChecked = false;
            StatsSessionViewBtn.IsChecked = true;
        }

        private void StatsScanBtn_Click(object sender, RoutedEventArgs e)
            => _ = StatsFetchLobbyStats();

        private async void StatsLastGameBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(StatsWatchPath)) { StatsSetStatus("MCC folder not found."); return; }
            var f = StatsLatestCarnageFile();
            if (f == null) { StatsSetStatus("No carnage report found."); return; }

            StatsLastGameBtn.IsEnabled = false;
            StatsLastGameBtn.Content = "↻ LOADING...";
            StatsSetStatus($"Loading {f.Name}…");
            try
            {
                await Task.Run(() => StatsTryProcessFile(f.FullName));
            }
            catch (Exception ex)
            {
                StatsSetStatus($"Could not load {f.Name}: {ex.Message}");
            }
            finally
            {
                StatsLastGameBtn.Content = "↺ LAST GAME";
                StatsLastGameBtn.IsEnabled = true;
            }
        }

        private void StatsObsOverlayToggle_Checked(object sender, RoutedEventArgs e)
        {
            _obsBrowserOverlayEnabled = true;
            if (!_mainWindowInitialized)
                return;

            App.SaveObsBrowserOverlayEnabled(true);
            if (!_rejoinProxy.IsRunning)
            {
                StatsRefreshObsOverlayUi();
                UpdateRejoinFixUi();
                return;
            }
            EnsureOverlaySourceServer(logStatus: true);
            StartNetworkStatsOverlay(GetNetworkStatsTargetIp(), GetNetworkStatsTargetServerInfo());
            StatsRefreshObsOverlayUi();
            PublishObsOverlaySnapshot();
            TryCopyObsOverlayUrlToClipboard();
        }

        private void StatsObsOverlayToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _obsBrowserOverlayEnabled = false;
            if (!_mainWindowInitialized)
                return;

            App.SaveObsBrowserOverlayEnabled(false);
            if (!_networkStatsOverlayEnabled && !_combinedNetworkSessionOverlayEnabled)
                _obsOverlayServer.Stop();
            if (!_networkStatsOverlayEnabled && !_combinedNetworkSessionOverlayEnabled)
                StartNetworkStatsOverlay(GetNetworkStatsTargetIp(), GetNetworkStatsTargetServerInfo());
            StatsRefreshObsOverlayUi();
            StatsSetStatus("OBS overlay stopped.");
        }

        private void NetworkStatsObsOnlyToggle_Checked(object sender, RoutedEventArgs e)
        {
            _networkStatsObsOnly = true;
            if (!_mainWindowInitialized)
                return;

            App.SaveNetworkStatsObsOnlyEnabled(true);
            CloseComponentOverlay(ref _gameNetworkStatsOverlay);
            EnsureOverlaySourceServer(logStatus: false);
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Network Stats are now OBS-only.");
        }

        private void NetworkStatsObsOnlyToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _networkStatsObsOnly = false;
            if (!_mainWindowInitialized)
                return;

            App.SaveNetworkStatsObsOnlyEnabled(false);
            if (_networkStatsOverlayEnabled && _rejoinProxy.IsRunning)
                EnsureGameNetworkStatsOverlay();
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Network Stats are visible in-game and in OBS.");
        }

        private void MatchmakingWaitObsOnlyToggle_Checked(object sender, RoutedEventArgs e)
        {
            _matchmakingWaitObsOnly = true;
            if (!_mainWindowInitialized)
                return;

            App.SaveMatchmakingWaitObsOnlyEnabled(true);
            CloseComponentOverlay(ref _matchmakingWaitOverlay);
            EnsureOverlaySourceServer(logStatus: false);
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Matchmaking Wait is now OBS-only.");
        }

        private void MatchmakingWaitObsOnlyToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _matchmakingWaitObsOnly = false;
            if (!_mainWindowInitialized)
                return;

            App.SaveMatchmakingWaitObsOnlyEnabled(false);
            if (_matchmakingWaitOverlayEnabled && _rejoinProxy.IsRunning)
                EnsureGameNetworkStatsOverlay();
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Matchmaking Wait is visible in-game and in OBS.");
        }

        private void SessionStatsObsOnlyToggle_Checked(object sender, RoutedEventArgs e)
        {
            _sessionStatsObsOnly = true;
            if (!_mainWindowInitialized)
                return;

            App.SaveSessionStatsObsOnlyEnabled(true);
            CloseComponentOverlay(ref _sessionStatsOverlay);
            EnsureOverlaySourceServer(logStatus: false);
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Session Stats and medals are now OBS-only.");
        }

        private void SessionStatsObsOnlyToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _sessionStatsObsOnly = false;
            if (!_mainWindowInitialized)
                return;

            App.SaveSessionStatsObsOnlyEnabled(false);
            if (_obsBrowserOverlaySessionStatsEnabled && _rejoinProxy.IsRunning)
                EnsureGameNetworkStatsOverlay();
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Session Stats and medals are visible in-game and in OBS.");
        }

        private void CombinedNetworkSessionObsOnlyToggle_Checked(object sender, RoutedEventArgs e)
        {
            _combinedNetworkSessionObsOnly = true;
            if (!_mainWindowInitialized)
                return;

            App.SaveCombinedNetworkSessionObsOnlyEnabled(true);
            CloseComponentOverlay(ref _combinedNetworkSessionOverlay);
            EnsureOverlaySourceServer(logStatus: false);
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Combined Network / Session is now OBS-only.");
        }

        private void CombinedNetworkSessionObsOnlyToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _combinedNetworkSessionObsOnly = false;
            if (!_mainWindowInitialized)
                return;

            App.SaveCombinedNetworkSessionObsOnlyEnabled(false);
            if (_combinedNetworkSessionOverlayEnabled && _rejoinProxy.IsRunning)
                EnsureGameNetworkStatsOverlay();
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
            StatsSetStatus("Combined Network / Session is visible in-game and in OBS.");
        }

        private void StatsObsSessionStatsToggle_Checked(object sender, RoutedEventArgs e)
        {
            _obsBrowserOverlaySessionStatsEnabled = true;
            App.SaveObsBrowserOverlaySessionStatsEnabled(true);
            if (_mainWindowInitialized && _rejoinProxy.IsRunning)
                EnsureGameNetworkStatsOverlay();
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
        }

        private void StatsObsSessionStatsToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _obsBrowserOverlaySessionStatsEnabled = false;
            App.SaveObsBrowserOverlaySessionStatsEnabled(false);
            CloseComponentOverlay(ref _sessionStatsOverlay);
            PublishObsOverlaySnapshot();
            UpdateRejoinFixUi();
        }

        private void StatsHwAuthBtn_Click(object sender, RoutedEventArgs e)
        {
            string gt; lock (_statsLock) { gt = _statsGamertag; }
            var win = new StatsAuthWindow(gt, silent: false) { Owner = this };
            if (win.ShowDialog() == true && !string.IsNullOrEmpty(win.CapturedToken))
                StatsApplyCapturedToken(win.CapturedToken!);
        }

        private async void StatsLobbyList_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            StatsPlayerRow? row = sender switch
            {
                ListView { SelectedItem: StatsPlayerRow selected } => selected,
                _ => null
            };
            if (row is null) return;
            var history = EnsureMatchHistoryTab();
            MainTabs.SelectedItem = MatchHistorySection;
            SyncSidebarSelection(MatchHistorySection);
            await history.LoadAsync(row.Gamertag);
        }

        private BanCheckerAuthorizationState GetBanCheckerAuthorizationState()
        {
            if (_rejoinProxy.TryGetLatestBanSpartanToken(out _, out var capturedAtUtc, out var sourceHost))
            {
                string captured = capturedAtUtc.ToLocalTime().ToString("MMM d, h:mm tt");
                string source = string.IsNullOrWhiteSpace(sourceHost) ? "MCC" : sourceHost;
                return new BanCheckerAuthorizationState(
                    true,
                    "READY",
                    $"Captured from {source} on {captured}. If MCC rejects it, the token is cleared and the refresh steps become required.");
            }

            string nextStep = _rejoinProxy.IsRunning
                ? "Rejoin Recovery is running. Launch MCC, sign in, and open Multiplayer to capture a fresh token."
                : "No usable token is stored. Start the MCC Data Proxy on Dashboard, then launch MCC and open Multiplayer.";
            return new BanCheckerAuthorizationState(false, "AUTHORIZATION NEEDED", nextStep);
        }

        private async Task<IReadOnlyList<BanCheckDisplayResult>> StatsCheckBanTargetsAsync(IReadOnlyList<string> targets)
        {
            if (!_rejoinProxy.TryGetLatestBanSpartanToken(out var token, out var capturedAtUtc, out _))
            {
                StatsSetStatus("Start the MCC Data Proxy on Dashboard, then let MCC make a ban summary request before using Ban Checker.");
                throw new InvalidOperationException(
                    "MCC authorization is missing or expired. Start the MCC Data Proxy on Dashboard, " +
                    "launch MCC, sign in, and open Multiplayer. Then return here and refresh the authorization status.");
            }

            StatsSetStatus($"Resolving {targets.Count} Ban Checker target(s)...");
            var resolved = await Task.WhenAll(targets.Select(async target =>
            {
                string xuid = await StatsResolveEnteredTargetToXuidAsync(target, token);
                return (Target: target, Xuid: xuid);
            }));

            var output = resolved.Where(item => string.IsNullOrWhiteSpace(item.Xuid))
                .Select(item => new BanCheckDisplayResult
                {
                    Target = item.Target,
                    Result = "NOT FOUND",
                    Details = "Could not resolve XUID; check the spelling."
                }).ToList();

            var found = resolved.Where(item => !string.IsNullOrWhiteSpace(item.Xuid)).ToList();
            if (found.Count > 0)
            {
                var checks = await StatsFetchBanSummariesAsync(found.Select(item => item.Xuid).ToList(), token);
                for (int i = 0; i < found.Count; i++)
                {
                    var check = checks[i];
                    output.Add(new BanCheckDisplayResult
                    {
                        Target = found[i].Target,
                        Result = check.HasActiveBans ? "BANNED" : "CLEAR",
                        Details = $"XUID {found[i].Xuid} — {check.Message.Replace(Environment.NewLine, " ")}"
                    });
                }
            }

            StatsSetStatus($"Ban Checker checked {targets.Count} player(s) using token captured {capturedAtUtc.LocalDateTime:g}.");
            return targets.Select(target => output.First(result => result.Target.Equals(target, StringComparison.OrdinalIgnoreCase))).ToList();
        }

        private async Task<string> StatsResolveEnteredTargetToXuidAsync(string target, string token)
        {
            string normalized = StatsNormalizeXuid(target);
            if (StatsLooksLikeXuid(normalized))
                return normalized;

            string cached = StatsResolveEnteredTargetToCachedXuid(target);
            if (!string.IsNullOrWhiteSpace(cached))
                return cached;

            StatsSetStatus($"Resolving XUID for {target}...");
            string resolved = await StatsFetchXuidForGamertagAsync(target, token);
            if (!string.IsNullOrWhiteSpace(resolved))
            {
                StatsRememberGamertagForXuid(resolved, target);
                return resolved;
            }

            return "";
        }

        private string StatsResolveEnteredTargetToCachedXuid(string target)
        {
            lock (_statsLock)
            {
                foreach (var (xuid, gamertag) in _statsGamertagsByXuid)
                {
                    if (gamertag.Equals(target, StringComparison.OrdinalIgnoreCase))
                        return StatsNormalizeXuid(xuid);
                }

                foreach (var row in _statsCurrentLobbySnapshotRows.Concat(_statsLastCompletedLobbyRows))
                {
                    if (row.Gamertag.Equals(target, StringComparison.OrdinalIgnoreCase))
                        return StatsNormalizeXuid(row.Xuid);
                }

                foreach (var player in _statsLastPlayers)
                {
                    string gamertag = player.Attribute("mGamertagText")?.Value ?? "";
                    if (gamertag.Equals(target, StringComparison.OrdinalIgnoreCase))
                        return StatsNormalizeXuid(player.Attribute("mXboxUserId")?.Value ?? "");
                }
            }

            return "";
        }

        private static async Task<string> StatsFetchXuidForGamertagAsync(string gamertag, string token)
        {
            string escaped = Uri.EscapeDataString(gamertag);
            string mccXuid = await StatsFetchXuidFromMccServiceRecordAsync(escaped, token);
            if (StatsLooksLikeXuid(mccXuid))
                return mccXuid;

            string[] urls =
            {
                $"https://api.geysermc.org/v2/xbox/xuid/{escaped}",
                $"https://playerdb.co/api/player/xbox/{escaped}",
            };

            foreach (string url in urls)
            {
                string xuid = await StatsTryFetchXuidFromUrlAsync(url);
                if (StatsLooksLikeXuid(xuid))
                    return xuid;
            }

            return "";
        }

        private static async Task<string> StatsFetchXuidFromMccServiceRecordAsync(string escapedGamertag, string token)
        {
            try
            {
                string url = $"https://mccapi.svc.halowaypoint.com/hmcc/users/gt({escapedGamertag})/service-record";
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.TryAddWithoutValidation("x-343-authorization-spartan", token);
                req.Headers.TryAddWithoutValidation("Accept", "application/json");
                req.Headers.TryAddWithoutValidation("User-Agent", "Mozilla/5.0");

                using var res = await StatsHttp.SendAsync(req);
                string body = await res.Content.ReadAsStringAsync();
                if (!res.IsSuccessStatusCode || string.IsNullOrWhiteSpace(body))
                    return "";

                return StatsExtractXuidFromLookupResponse(body);
            }
            catch
            {
                return "";
            }
        }

        private static async Task<string> StatsTryFetchXuidFromUrlAsync(string url)
        {
            try
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.TryAddWithoutValidation("Accept", "application/json,text/plain,text/html,*/*");
                req.Headers.TryAddWithoutValidation("User-Agent", "Mozilla/5.0");
                req.Headers.TryAddWithoutValidation("Referer", "https://cxkes.me/xbox/xuid");

                using var res = await StatsHttp.SendAsync(req);
                string body = await res.Content.ReadAsStringAsync();
                if (!res.IsSuccessStatusCode || string.IsNullOrWhiteSpace(body))
                    return "";

                return StatsExtractXuidFromLookupResponse(body);
            }
            catch
            {
                return "";
            }
        }

        private static string StatsExtractXuidFromLookupResponse(string body)
        {
            try
            {
                using var doc = JsonDocument.Parse(body);
                string fromJson = StatsFindXuidInJson(doc.RootElement);
                if (!string.IsNullOrWhiteSpace(fromJson))
                    return fromJson;
            }
            catch { }

            var match = System.Text.RegularExpressions.Regex.Match(body, @"\b(253327\d{10})\b");
            return match.Success ? match.Groups[1].Value : "";
        }

        private static string StatsFindXuidInJson(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var property in element.EnumerateObject())
                    {
                        if (property.Name.Contains("xuid", StringComparison.OrdinalIgnoreCase))
                        {
                            string candidate = property.Value.ValueKind == JsonValueKind.String
                                ? property.Value.GetString() ?? ""
                                : property.Value.GetRawText();
                            candidate = StatsNormalizeXuid(candidate);
                            if (StatsLooksLikeXuid(candidate))
                                return candidate;
                        }

                        string nested = StatsFindXuidInJson(property.Value);
                        if (!string.IsNullOrWhiteSpace(nested))
                            return nested;
                    }
                    break;

                case JsonValueKind.Array:
                    foreach (var item in element.EnumerateArray())
                    {
                        string nested = StatsFindXuidInJson(item);
                        if (!string.IsNullOrWhiteSpace(nested))
                            return nested;
                    }
                    break;

                case JsonValueKind.String:
                    string value = StatsNormalizeXuid(element.GetString() ?? "");
                    if (StatsLooksLikeXuid(value))
                        return value;
                    break;
            }

            return "";
        }

        private async Task<(bool HasActiveBans, string Message, string Status)> StatsFetchBanSummaryAsync(string xuid, string token)
        {
            string url =
                $"https://banprocessor.svc.halowaypoint.com/hmcc/bansummary" +
                $"?targets=xuid({xuid}),Authenticated(Device)";

            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.TryAddWithoutValidation("X-343-Authorization-Spartan", token);
            req.Headers.TryAddWithoutValidation("Accept", "application/json");
            req.Headers.TryAddWithoutValidation("User-Agent", "cpprestsdk/2.9.0");

            using var res = await StatsHttp.SendAsync(req);
            string body = await res.Content.ReadAsStringAsync();

            if (res.StatusCode == HttpStatusCode.BadRequest)
                return (false, "HTTP 400: request shape is wrong.", "Ban Checker: bad request shape.");
            if (res.StatusCode == HttpStatusCode.Unauthorized)
            {
                _rejoinProxy.ClearBanSpartanToken(token);
                return (false, "HTTP 401: token expired, wrong, or not an MCC in-game token.", "Ban Checker: token rejected.");
            }
            if (res.StatusCode == HttpStatusCode.NotFound)
                return (false, "HTTP 404: wrong endpoint path.", "Ban Checker: endpoint not found.");
            if (!res.IsSuccessStatusCode)
                return (false, $"HTTP {(int)res.StatusCode} {res.ReasonPhrase}\n\n{body}", $"Ban Checker: HTTP {(int)res.StatusCode}.");

            return StatsParseBanSummaryResponse(xuid, body);
        }

        private async Task<IReadOnlyList<(bool HasActiveBans, string Message, string Status)>> StatsFetchBanSummariesAsync(
            IReadOnlyList<string> xuids,
            string token)
        {
            string targets = string.Join(",", xuids.Select(xuid => $"xuid({xuid})"));
            string url = $"https://banprocessor.svc.halowaypoint.com/hmcc/bansummary?targets={targets},Authenticated(Device)";

            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.TryAddWithoutValidation("X-343-Authorization-Spartan", token);
            req.Headers.TryAddWithoutValidation("Accept", "application/json");
            req.Headers.TryAddWithoutValidation("User-Agent", "cpprestsdk/2.9.0");

            using var res = await StatsHttp.SendAsync(req);
            string body = await res.Content.ReadAsStringAsync();
            if (res.StatusCode == HttpStatusCode.Unauthorized)
            {
                _rejoinProxy.ClearBanSpartanToken(token);
                throw new InvalidOperationException(
                    "MCC rejected the stored authorization. It has been cleared. Start the MCC Data Proxy on Dashboard, " +
                    "launch MCC, sign in, and open Multiplayer to capture a new token.");
            }
            if (!res.IsSuccessStatusCode)
                throw new HttpRequestException($"Ban Checker returned HTTP {(int)res.StatusCode} {res.ReasonPhrase}.\n\n{body}");

            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("Results", out var results) || results.ValueKind != JsonValueKind.Array)
                throw new InvalidDataException("The Ban Checker response did not contain a Results array.");

            var responseItems = results.EnumerateArray().ToList();
            return xuids.Select((xuid, index) => index < responseItems.Count
                    ? StatsParseBanSummaryResponse(xuid, $"{{\"Results\":[{responseItems[index].GetRawText()}]}}")
                    : (false, "The service returned no result for this player.", "Ban Checker: no result."))
                .ToList();
        }

        private static (bool HasActiveBans, string Message, string Status) StatsParseBanSummaryResponse(string xuid, string body)
        {
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("Results", out var results) ||
                results.ValueKind != JsonValueKind.Array ||
                results.GetArrayLength() == 0)
            {
                return (false, "HTTP 200, but no Results were returned.", "Ban Checker: no results.");
            }

            var target = results.EnumerateArray().First();
            int resultCode = target.TryGetProperty("ResultCode", out var codeEl) && codeEl.TryGetInt32(out var code)
                ? code
                : -1;

            if (!target.TryGetProperty("Result", out var result) ||
                !result.TryGetProperty("BansInEffect", out var bans) ||
                bans.ValueKind != JsonValueKind.Array ||
                bans.GetArrayLength() == 0)
            {
                return (false, $"Not banned.\n\nNo active bans for xuid({xuid}).\nResultCode: {resultCode}", "Ban Checker: not banned.");
            }

            var banLines = bans.EnumerateArray().Select((ban, index) =>
            {
                int typeValue = ban.TryGetProperty("Type", out var typeEl) && typeEl.TryGetInt32(out var parsedType) ? parsedType : -1;
                int scopeValue = ban.TryGetProperty("Scope", out var scopeEl) && scopeEl.TryGetInt32(out var parsedScope) ? parsedScope : -1;
                string until = "unknown";
                if (ban.TryGetProperty("EnforceUntilUtc", out var untilObj) &&
                    untilObj.TryGetProperty("ISO8601Date", out var dateEl))
                    until = dateEl.GetString() ?? until;

                return $"{index + 1}. Type {typeValue}, Scope {scopeValue}, Until {until}";
            });

            return (true, $"BANNED.\n\nActive bans:\n{string.Join(Environment.NewLine, banLines)}\n\nResultCode: {resultCode}", "Ban Checker: active ban found.");
        }

        private void ExternalHyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        private void TryStartObsOverlayServer(bool logStatus)
        {
            EnsureOverlaySourceServer(logStatus);
        }

        private void EnsureOverlaySourceServer(bool logStatus)
        {
            try
            {
                _obsOverlayServer.Start();
                if (logStatus)
                    StatsSetStatus($"OBS overlay ready: {_obsOverlayServer.Url}");
                if (logStatus)
                    AppendLog("[OBS]", $"Browser source ready at {_obsOverlayServer.Url}", "#00C8FF");
            }
            catch (Exception ex)
            {
                if (_obsBrowserOverlayEnabled)
                {
                    _obsBrowserOverlayEnabled = false;
                    App.SaveObsBrowserOverlayEnabled(false);
                    Dispatcher.InvokeAsync(() => StatsObsOverlayToggle.IsChecked = false);
                }
                StatsSetStatus($"OBS overlay could not start: {ex.Message}");
                AppendLog("[OBS]", $"Overlay server could not start: {ex.Message}", "#FF2D55");
            }
        }

        private void TryCopyObsOverlayUrlToClipboard()
        {
            if (!_obsOverlayServer.IsRunning)
                return;

            try
            {
                string urls = string.Join(Environment.NewLine,
                    $"Network: {_obsOverlayServer.ComponentUrl("network", "obs")}",
                    $"Wait: {_obsOverlayServer.ComponentUrl("wait", "obs")}",
                    $"Session: {_obsOverlayServer.ComponentUrl("session", "obs")}",
                    $"Combined Network / Session: {_obsOverlayServer.ComponentUrl("combined", "obs")}");
                Clipboard.SetText(urls);
                StatsSetStatus("Four independent OBS overlay URLs copied.");
            }
            catch (Exception ex)
            {
                StatsSetStatus($"OBS overlay ready, but clipboard copy failed: {ex.Message}");
            }
        }

        private void StatsObsOverlayCopyButton_Click(object sender, RoutedEventArgs e)
        {
            TryCopyObsOverlayUrlToClipboard();
        }

        private void StatsRefreshObsOverlayUi()
        {
            StatsObsOverlayToggle.Content = "OBS BROWSER OVERLAY";
            StatsObsOverlayUrlLabel.Text = _obsBrowserOverlayEnabled
                ? $"NETWORK  {_obsOverlayServer.ComponentUrl("network", "obs")}\n" +
                  $"WAIT     {_obsOverlayServer.ComponentUrl("wait", "obs")}\n" +
                  $"SESSION  {_obsOverlayServer.ComponentUrl("session", "obs")}\n" +
                  $"COMBINED {_obsOverlayServer.ComponentUrl("combined", "obs")}"
                : "";
            StatsObsOverlayUrlsPanel.Visibility = _obsBrowserOverlayEnabled
                ? Visibility.Visible
                : Visibility.Collapsed;
            // Session-stat visibility applies to every overlay and is intentionally
            // independent of whether the OBS browser-source link is enabled.
            StatsObsSessionStatsToggle.IsEnabled = true;
        }

        private void PublishObsOverlaySnapshot()
        {
            if (!Dispatcher.CheckAccess())
            {
                _ = Dispatcher.InvokeAsync(
                    PublishObsOverlaySnapshot,
                    System.Windows.Threading.DispatcherPriority.Background);
                return;
            }

            var snapshot = BuildObsOverlaySnapshot();

            foreach (var overlay in AllGameOverlays()) overlay.UpdateSessionStats(snapshot);

            if (_obsOverlayServer.IsRunning)
                _obsOverlayServer.Update(snapshot);
        }

        private ObsOverlaySnapshot BuildObsOverlaySnapshot()
        {
            int wins, losses, games, bestSpree;
            long kills, deaths;
            Dictionary<string, int> medals;
            ObsPostGameRecap? recap;
            lock (_statsLock)
            {
                wins = _statsSession.Wins;
                losses = _statsSession.Losses;
                games = _statsSession.GamesPlayed;
                kills = _statsSession.Kills;
                deaths = _statsSession.Deaths;
                bestSpree = _statsSession.BestSpree;
                medals = new Dictionary<string, int>(_statsSession.MultikillCounts, StringComparer.OrdinalIgnoreCase);
                recap = _postGameRecap;
            }

            double kd = deaths > 0 ? (double)kills / deaths : kills;
            var serverInfo = GetNetworkStatsTargetServerInfo();
            string serverLabel = GameServerRegionResolver.GetRegionLabel(serverInfo);
            if (string.IsNullOrWhiteSpace(serverLabel))
                serverLabel = _statsCurrentLobbyServerText.Replace("Server - ", "", StringComparison.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(serverLabel))
                serverLabel = "SERVER: --";

            var rttHistory = _lastNetworkStatsSnapshot?.RttHistory
                .Select(x => x.HasValue ? (int?)Math.Clamp(x.Value, int.MinValue, int.MaxValue) : null)
                .ToArray() ?? Array.Empty<int?>();
            bool showFirewallStatus = ChkRejoinFixFirewall.IsChecked == true ||
                ChkRejoinFixFirewallMatchmaking.IsChecked == true;
            string firewallStatus = showFirewallStatus ? TxtRejoinFirewallStatus.Text : "";
            string firewallStatusColor = showFirewallStatus &&
                TxtRejoinFirewallStatus.Foreground is SolidColorBrush firewallStatusBrush
                    ? firewallStatusBrush.Color.ToString()
                    : "";

            return new ObsOverlaySnapshot(
                ShowSessionStats: _obsBrowserOverlaySessionStatsEnabled,
                ShowNetworkStats: _networkStatsOverlayEnabled,
                ShowCombinedOverlay: _combinedNetworkSessionOverlayEnabled,
                ShowMatchmakingWait: _matchmakingWaitOverlayEnabled && _smartMatchWaitEstimate is not null,
                MatchmakingWaitSeconds: _smartMatchWaitEstimate?.WaitSeconds,
                MatchmakingWaitModern: App.LoadGameOverlayVisualStyle("wait") == GameOverlayVisualStyle.Modern,
                MatchmakingPopulation: _smartMatchHopperPopulation,
                MatchmakingPlaylistName: _smartMatchHopperDisplayName,
                MatchmakingSearchScope: _smartMatchWaitEstimate?.HopperName.Contains(
                    "Ranked",
                    StringComparison.OrdinalIgnoreCase) == true
                        ? "all ranks"
                        : "all gametypes",
                MatchmakingStartedAtUtc: _smartMatchWaitEstimate?.CapturedAtUtc,
                MatchmakingExpiresAtUtc: _smartMatchWaitEstimate is null
                    ? null
                    : _smartMatchWaitEstimate.CapturedAtUtc.AddSeconds(_smartMatchWaitEstimate.GiveUpSeconds + 10),
                ServerLabel: serverLabel,
                VpnRegion: VpnConnectionPresence.ConnectedRegion,
                FirewallStatus: firewallStatus,
                FirewallStatusColor: firewallStatusColor,
                RttMs: _lastNetworkStatsSnapshot?.RttMs is long rtt ? (int?)Math.Clamp(rtt, int.MinValue, int.MaxValue) : null,
                JitterMs: _lastNetworkStatsSnapshot?.JitterMs,
                PacketLossPercent: _lastNetworkStatsSnapshot?.PacketLossPercent ?? 0,
                RttHistoryMs: rttHistory,
                UploadKilobytesPerSecond: _lastNetworkTrafficSnapshot?.UploadKilobytesPerSecond,
                DownloadKilobytesPerSecond: _lastNetworkTrafficSnapshot?.DownloadKilobytesPerSecond,
                UploadPacketsPerSecond: _lastNetworkTrafficSnapshot?.UploadPacketsPerSecond,
                DownloadPacketsPerSecond: _lastNetworkTrafficSnapshot?.DownloadPacketsPerSecond,
                Wins: wins,
                Losses: losses,
                GamesPlayed: games,
                Kills: kills,
                Deaths: deaths,
                SessionKd: kd.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
                PostGameRecap: recap,
                BestSpree: bestSpree,
                DoubleKills: medals.GetValueOrDefault("Double Kill"),
                TripleKills: medals.GetValueOrDefault("Triple Kill"),
                Overkills: medals.GetValueOrDefault("Overkill"),
                Killtaculars: medals.GetValueOrDefault("Killtacular"),
                Killtrocities: medals.GetValueOrDefault("Killtrocity"),
                Killimanjaros: medals.GetValueOrDefault("Killimanjaro"),
                Killtastrophes: medals.GetValueOrDefault("Killtastrophe"),
                Killpocalypses: medals.GetValueOrDefault("Killpocalypse"),
                Killionaires: medals.GetValueOrDefault("Killionaire"),
                OverlayLeftRatio: _lastOverlayRelativePlacement.X,
                OverlayTopRatio: _lastOverlayRelativePlacement.Y,
                OverlayWidthRatio: _lastOverlayRelativePlacement.Width,
                OverlayHeightRatio: _lastOverlayRelativePlacement.Height,
                OverlayPlacements: _componentOverlayRelativePlacements.ToDictionary(
                    pair => pair.Key,
                    pair => new ObsOverlayPlacement(pair.Value.X, pair.Value.Y, pair.Value.Width, pair.Value.Height),
                    StringComparer.OrdinalIgnoreCase));
        }

        // ══════════════════════════════════════════════════════════════════════
        // Stats Tab — Core logic
        // ══════════════════════════════════════════════════════════════════════

        private void StatsPopulationRefresh_Click(object sender, RoutedEventArgs e) =>
            _ = StatsRefreshMatchmakingPopulationAsync();

        private void StatsPopulationSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StatsPopulationList.SelectedItem is not MatchmakingPopulationRow selected) return;
            _selectedPopulationHopperName = selected.HopperName;
            UpdatePopulationChart();
        }

        private void PopulationRangeButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Controls.Primitives.ToggleButton { Tag: string value } button ||
                !int.TryParse(value, out int minutes)) return;
            _populationHistoryMinutes = minutes;
            PopulationRange15Button.IsChecked = ReferenceEquals(button, PopulationRange15Button);
            PopulationRange30Button.IsChecked = ReferenceEquals(button, PopulationRange30Button);
            PopulationRange60Button.IsChecked = ReferenceEquals(button, PopulationRange60Button);
            PopulationRange180Button.IsChecked = ReferenceEquals(button, PopulationRange180Button);
            UpdatePopulationChart();
        }

        private void UpdatePopulationChart()
        {
            var selected = _statsPopulationRows.FirstOrDefault(row => row.HopperName.Equals(
                _selectedPopulationHopperName, StringComparison.OrdinalIgnoreCase))
                ?? _statsPopulationRows.FirstOrDefault(row => row.DisplayName.Equals("Social 4s", StringComparison.OrdinalIgnoreCase))
                ?? _statsPopulationRows.FirstOrDefault();
            if (selected is not null)
            {
                _selectedPopulationHopperName = selected.HopperName;
                if (!ReferenceEquals(StatsPopulationList.SelectedItem, selected))
                    StatsPopulationList.SelectedItem = selected;
            }
            var cutoff = DateTimeOffset.Now.AddMinutes(-_populationHistoryMinutes);
            var samples = _statsPopulationHistory.Where(sample => sample.CapturedAt >= cutoff &&
                sample.HopperName.Equals(_selectedPopulationHopperName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(sample => sample.CapturedAt).ToList();
            StatsPopulationChart.Samples = samples;
            StatsPopulationChartTitle.Text = $"POPULATION HISTORY · {(selected?.DisplayName ?? "Social 4s").ToUpperInvariant()}";
            StatsPopulationChartSummary.Text = samples.Count == 0
                ? $"LAST {_populationHistoryMinutes} MINUTES · WAITING FOR SAMPLES"
                : $"LAST {_populationHistoryMinutes} MINUTES · {samples.Count} SAMPLES · LATEST {samples[^1].Population:N0} SEARCHING";
        }

        private void StatsPopulationGraph_Click(object sender, RoutedEventArgs e)
        {
            new PopulationHistoryWindow(_statsPopulationHistory) { Owner = this }.ShowDialog();
        }

        private void PopulationStartProxy_Click(object sender, RoutedEventArgs e)
        {
            if (_rejoinProxy.IsRunning || !BtnRejoinFix.IsEnabled) return;
            PopulationStartProxyButton.IsEnabled = false;
            BtnRejoinFix_Click(sender, e);
        }

        private async void MatchmakingPopulationTimer_Tick(object? sender, EventArgs e)
        {
            var estimate = _smartMatchWaitEstimate;
            if (estimate is null || !_rejoinProxy.IsRunning ||
                DateTimeOffset.UtcNow >= estimate.CapturedAtUtc.AddSeconds(estimate.GiveUpSeconds + 10))
            {
                _matchmakingPopulationTimer.Stop();
                _smartMatchHopperPopulation = null;
                PublishObsOverlaySnapshot();
                return;
            }

            if (DateTimeOffset.UtcNow - _lastFullPopulationRefreshUtc >= TimeSpan.FromSeconds(60))
            {
                await StatsRefreshMatchmakingPopulationAsync();
                return;
            }

            if (string.IsNullOrWhiteSpace(estimate.HopperName))
                return;

            var result = await _rejoinProxy.GetHopperStatisticsAsync(estimate.HopperName);
            if (ReferenceEquals(estimate, _smartMatchWaitEstimate) && string.IsNullOrWhiteSpace(result.Error))
            {
                _smartMatchHopperPopulation = result.Population;
                PublishObsOverlaySnapshot();
            }
        }

        private void StatsPopulationHeader_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not GridViewColumnHeader header || header.Tag is not string property)
                return;

            _statsPopulationSortDirection = property.Equals(_statsPopulationSortProperty, StringComparison.Ordinal)
                ? (_statsPopulationSortDirection == ListSortDirection.Ascending
                    ? ListSortDirection.Descending
                    : ListSortDirection.Ascending)
                : ListSortDirection.Ascending;
            _statsPopulationSortProperty = property;

            var view = CollectionViewSource.GetDefaultView(_statsPopulationRows);
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(new SortDescription(property, _statsPopulationSortDirection));
        }

        private async Task StatsRefreshMatchmakingPopulationAsync()
        {
            if (!await _statsPopulationRefreshLock.WaitAsync(0))
                return;

            try
            {
                _lastFullPopulationRefreshUtc = DateTimeOffset.UtcNow;
                if (!_rejoinProxy.HasPopulationAuthorization)
                {
                    StatsPopulationStatusLabel.Text = "Waiting for authorization — start the MCC Data Proxy, launch MCC, start matchmaking, then cancel and return to the lobby. Collection resumes automatically.";
                    _statsPopulationRows.Clear();
                    _statsPopulationHistory.RemoveAll(sample => sample.CapturedAt < DateTimeOffset.Now.AddHours(-3));
                    UpdatePopulationChart();
                    return;
                }
                var hoppers = EnsurePlaylistsTab().GetMatchmakingHoppers();
                if (hoppers.Count == 0)
                {
                    StatsPopulationStatusLabel.Text = "No hopper names were found in MCC's playlist XML.";
                    return;
                }

                StatsPopulationStatusLabel.Text = $"Refreshing {hoppers.Count} hoppers…";
                using var queryGate = new SemaphoreSlim(4, 4);
                var tasks = hoppers.Select(async hopper =>
                {
                    await queryGate.WaitAsync();
                    try
                    {
                        return new
                        {
                            Hopper = hopper,
                            Result = await _rejoinProxy.GetHopperStatisticsAsync(hopper.HopperName)
                        };
                    }
                    finally
                    {
                        queryGate.Release();
                    }
                });
                var results = await Task.WhenAll(tasks);

                // Live chart history is memory-only and limited to the last three hours.
                var capturedAt = DateTimeOffset.Now;
                _statsPopulationHistory.AddRange(results.Where(item =>
                    string.IsNullOrWhiteSpace(item.Result.Error) && item.Result.Population.HasValue)
                    .Select(item => new MatchmakingPopulationSample(capturedAt, item.Hopper.HopperName,
                        item.Hopper.DisplayName, item.Result.Population!.Value)));
                _statsPopulationHistory.RemoveAll(sample => sample.CapturedAt < capturedAt.AddHours(-3));

                _statsPopulationRows.Clear();
                foreach (var item in results
                    .Select(x => new MatchmakingPopulationRow(x.Hopper, x.Result))
                    .OrderByDescending(x => x.Population ?? -1)
                    .ThenBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase))
                {
                    _statsPopulationRows.Add(item);
                }

                UpdatePopulationChart();
                _smartMatchHopperPopulation = _smartMatchWaitEstimate is null
                    ? null
                    : results.FirstOrDefault(x => x.Hopper.HopperName.Equals(
                        _smartMatchWaitEstimate.HopperName,
                        StringComparison.OrdinalIgnoreCase))?.Result.Population;
                PublishObsOverlaySnapshot();

                int successful = results.Count(x => string.IsNullOrWhiteSpace(x.Result.Error));
                StatsPopulationStatusLabel.Text = successful > 0
                    ? $"Live hopper statistics · {successful}/{results.Length} available · updated {DateTime.Now:h:mm:ss tt}"
                    : results.Select(x => x.Result.Error).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))
                        ?? "Hopper statistics are unavailable.";
            }
            finally
            {
                _statsPopulationRefreshLock.Release();
            }
        }

        private void StatsApplyGamertag()
        {
            string gt = StatsGamertagBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(gt)) return;
            lock (_statsLock) { _statsGamertag = gt; _statsSession.Reset(); }
            App.SavePlayerGamertag(gt);
            _statsWaypointTokenTimer.Start();
            _ = StatsMaintainWaypointTokenAsync();
            try
            {
                Directory.CreateDirectory(App.ToolboxDataRoot);
                File.WriteAllText(StatsSettingsFile, gt);
            }
            catch { }
            StatsRefreshSessionUI();
            StatsSetStatus("Fetching stats…");
            _ = StatsFetchStats(gt);
            string tok; lock (_statsLock) { tok = _statsSpartanToken; }
            if (!string.IsNullOrEmpty(tok))
                _ = StatsFetchRecentStatsAsync(gt, tok);
        }

        private void StatsApplyCapturedToken(string token)
        {
            lock (_statsLock)
            {
                _statsSpartanToken = token;
                _statsHwTokenExpired = false;
                _statsWaypointUnavailable = false;
                _statsTokenLastValidatedUtc = DateTimeOffset.UtcNow;
            }
            StatsSaveToken(token);
            StatsUpdateHwStatus();
            StatsSetStatus("HW token captured.");
            string gt; lock (_statsLock) { gt = _statsGamertag; }
            if (!string.IsNullOrWhiteSpace(gt))
            {
                _ = StatsFetchStats(gt);
                _ = StatsFetchRecentStatsAsync(gt, token);
            }
        }

        private async Task StatsInitializeWaypointSessionAsync(string gamertag)
        {
            // Validate before fetching stats: a saved token is not proof of a live session.
            await StatsMaintainWaypointTokenAsync();
            await StatsFetchStats(gamertag);

            string token;
            bool expired;
            lock (_statsLock)
            {
                token = _statsSpartanToken;
                expired = _statsHwTokenExpired;
            }
            if (!string.IsNullOrWhiteSpace(token) && !expired)
                _ = StatsFetchRecentStatsAsync(gamertag, token);

        }

        private async void StatsWaypointTokenTimer_Tick(object? sender, EventArgs e) =>
            await StatsMaintainWaypointTokenAsync();

        private async Task StatsMaintainWaypointTokenAsync()
        {
            if (!await _statsWaypointMaintenanceLock.WaitAsync(0))
                return;

            try
            {
                string token;
                string gamertag;
                bool expired;
                DateTimeOffset validatedUtc;
                lock (_statsLock)
                {
                    token = _statsSpartanToken;
                    gamertag = _statsGamertag;
                    expired = _statsHwTokenExpired;
                    validatedUtc = _statsTokenLastValidatedUtc;
                }

                if (string.IsNullOrWhiteSpace(gamertag))
                    return;

                // MCC's ban-check token also authorizes Waypoint stats. Validate a
                // fresh candidate before replacing a working website connection.
                if (await StatsTryAdoptGameTokenAsync(gamertag, token))
                    return;

                if (string.IsNullOrWhiteSpace(token) || expired) return;

                bool unavailable;
                lock (_statsLock) { unavailable = _statsWaypointUnavailable; }
                if (unavailable || validatedUtc == DateTimeOffset.MinValue ||
                    DateTimeOffset.UtcNow - validatedUtc >= TimeSpan.FromMinutes(15))
                {
                    WaypointTokenProbeResult probe = await StatsProbeWaypointTokenAsync(gamertag, token);
                    if (probe == WaypointTokenProbeResult.Valid)
                    {
                        lock (_statsLock)
                        {
                            if (_statsSpartanToken == token)
                            {
                                _statsTokenLastValidatedUtc = DateTimeOffset.UtcNow;
                                _statsWaypointUnavailable = false;
                            }
                        }
                        StatsUpdateHwStatus();
                    }
                    else if (probe == WaypointTokenProbeResult.Unauthorized)
                    {
                        lock (_statsLock)
                        {
                            if (_statsSpartanToken == token) _statsHwTokenExpired = true;
                        }
                        StatsUpdateHwStatus();
                        return;
                    }
                    else
                    {
                        // A timeout, DNS failure, or 5xx response is not an auth
                        // failure. Keep the current session and try next tick.
                        return;
                    }
                }
            }
            finally
            {
                _statsWaypointMaintenanceLock.Release();
            }
        }

        // Called only while holding _statsWaypointMaintenanceLock. The one-minute
        // timer retries transient failures and picks up captures received mid-probe.
        private async Task<bool> StatsTryAdoptGameTokenAsync(string gamertag, string currentToken)
        {
            if (!_rejoinProxy.TryGetLatestBanSpartanToken(out var candidate, out _, out _) ||
                string.Equals(candidate, currentToken, StringComparison.Ordinal))
                return false;

            bool sameCandidate = candidate == _statsLastGameTokenCandidate &&
                gamertag == _statsLastGameTokenGamertag;
            if (sameCandidate && (_statsGameTokenCandidateResolved ||
                DateTimeOffset.UtcNow - _statsLastGameTokenProbeUtc < TimeSpan.FromMinutes(1)))
                return false;

            _statsLastGameTokenCandidate = candidate;
            _statsLastGameTokenGamertag = gamertag;
            _statsLastGameTokenProbeUtc = DateTimeOffset.UtcNow;
            _statsGameTokenCandidateResolved = false;

            var result = await StatsProbeWaypointTokenAsync(gamertag, candidate);
            if (result == WaypointTokenProbeResult.Unauthorized)
            {
                // A repeated request carrying the same expired token isn't a renewal.
                _statsGameTokenCandidateResolved = true;
                return false;
            }
            if (result != WaypointTokenProbeResult.Valid)
                return false;

            // Do not overwrite a manual connection, changed gamertag, or a newer
            // game capture that arrived while the validation request was in flight.
            lock (_statsLock)
            {
                if (_statsSpartanToken != currentToken || _statsGamertag != gamertag)
                    return false;
            }
            if (!_rejoinProxy.TryGetLatestBanSpartanToken(out var latest, out _, out _) || latest != candidate)
                return false;

            _statsGameTokenCandidateResolved = true;
            StatsApplyCapturedToken(candidate);
            StatsSetStatus("Waypoint connected automatically from MCC.");
            AppendLog("[HW]", "Validated and saved a game-captured Waypoint token.", "#00C8FF");
            return true;
        }

        private async Task<WaypointTokenProbeResult> StatsProbeWaypointTokenAsync(
            string gamertag,
            string token)
        {
            try
            {
                string url =
                    $"https://mccapi.svc.halowaypoint.com/hmcc/users/gt({Uri.EscapeDataString(gamertag)})/service-record";
                using var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.TryAddWithoutValidation("x-343-authorization-spartan", token);
                request.Headers.TryAddWithoutValidation("Accept", "application/json");
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                using HttpResponseMessage response = await StatsHttp.SendAsync(
                    request, HttpCompletionOption.ResponseHeadersRead, timeout.Token);
                if (response.StatusCode == HttpStatusCode.Unauthorized)
                    return WaypointTokenProbeResult.Unauthorized;
                lock (_statsLock)
                {
                    if (_statsSpartanToken == token)
                        _statsWaypointUnavailable = !response.IsSuccessStatusCode;
                }
                StatsUpdateHwStatus();
                return response.IsSuccessStatusCode
                    ? WaypointTokenProbeResult.Valid
                    : WaypointTokenProbeResult.TransientFailure;
            }
            catch
            {
                lock (_statsLock)
                {
                    if (_statsSpartanToken == token)
                        _statsWaypointUnavailable = true;
                }
                StatsUpdateHwStatus();
                return WaypointTokenProbeResult.TransientFailure;
            }
        }

        private enum WaypointTokenProbeResult
        {
            Valid,
            Unauthorized,
            TransientFailure
        }

        private void StatsSetStatus(string msg)
        {
            if (Dispatcher.CheckAccess())
                StatsStatusLabel.Text = msg;
            else
                _ = Dispatcher.InvokeAsync(() => StatsStatusLabel.Text = msg);
        }

        private void StatsUpdateHwStatus()
        {
            // Stats requests also complete on worker threads. Create the brushes
            // on the UI thread, not just the controls that receive those brushes.
            if (!Dispatcher.CheckAccess())
            {
                _ = Dispatcher.InvokeAsync(StatsUpdateHwStatus);
                return;
            }

            string text, buttonText; Brush color;
            lock (_statsLock)
            {
                if (string.IsNullOrEmpty(_statsSpartanToken))
                    (text, buttonText, color) = ("NOT CONNECTED", "CONNECT WAYPOINT", new SolidColorBrush(Color.FromRgb(0x71, 0x86, 0x9A)));
                else if (_statsHwTokenExpired)
                    (text, buttonText, color) = ("EXPIRED", "RECONNECT WAYPOINT", new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55)));
                else if (_statsWaypointUnavailable)
                    (text, buttonText, color) = ("UNAVAILABLE", "RETRY WAYPOINT", new SolidColorBrush(Color.FromRgb(0x71, 0x86, 0x9A)));
                else if (_statsTokenLastValidatedUtc == DateTimeOffset.MinValue)
                    (text, buttonText, color) = ("CHECKING...", "CHECKING...", new SolidColorBrush(Color.FromRgb(0x71, 0x86, 0x9A)));
                else
                    (text, buttonText, color) = ("CONNECTED", "WAYPOINT CONNECTED", new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF)));
            }
            StatsHwStatusLabel.Text = text;
            StatsHwStatusLabel.Foreground = color;
            StatsHwAuthBtn.Content = buttonText;
        }

        private void StatsRefreshSessionUI()
        {
            if (!Dispatcher.CheckAccess())
            {
                _ = Dispatcher.InvokeAsync(StatsRefreshSessionUI);
                return;
            }

            int wins, losses, games, bestSpree, longestWinStreak;
            long kills, deaths;
            double bestGameKd;
            string bestGameScore;
            Dictionary<string, int> medals;
            lock (_statsLock)
            {
                wins = _statsSession.Wins; losses = _statsSession.Losses;
                games = _statsSession.GamesPlayed;
                kills = _statsSession.Kills; deaths = _statsSession.Deaths;
                bestSpree = _statsSession.BestSpree;
                longestWinStreak = _statsSession.LongestWinStreak;
                bestGameKd = _statsSession.BestGameKd;
                bestGameScore = _statsSession.BestGameScore;
                medals = new Dictionary<string, int>(_statsSession.MultikillCounts, StringComparer.OrdinalIgnoreCase);
            }
            double kdr = deaths > 0 ? (double)kills / deaths : kills;
            double winRate = games > 0 ? (double)wins / games * 100 : 0;

            StatsWinsLabel.Text       = $"{wins}W";
            StatsLossesLabel.Text     = $"{losses}L";
            StatsGamesLabel.Text      = $"{games} game{(games == 1 ? "" : "s")}";
            StatsSessionKdLabel.Text  = kdr.ToString("F2");
            StatsSessionKillsLabel.Text = $"{kills:N0}K";
            StatsSessionDeathsLabel.Text = $"{deaths:N0}D";

            StatsDashboardRecord.Text = $"{wins}W–{losses}L";
            StatsDashboardWinRate.Text = games > 0 ? $"{winRate:F0}%" : "—";
            StatsDashboardKd.Text = games > 0 ? kdr.ToString("F2") : "—";
            StatsDashboardKills.Text = kills.ToString("N0");
            StatsDashboardDeaths.Text = deaths.ToString("N0");
            StatsDashboardBestSpree.Text = bestSpree.ToString();
            StatsDashboardBestGame.Text = string.IsNullOrEmpty(bestGameScore) ? "—" : bestGameScore;
            StatsDashboardBestKd.Text = games > 0 ? bestGameKd.ToString("F2") : "—";
            StatsDashboardWinStreak.Text = longestWinStreak.ToString();

            StatsMedalDouble.Text = medals.GetValueOrDefault("Double Kill").ToString();
            StatsMedalTriple.Text = medals.GetValueOrDefault("Triple Kill").ToString();
            StatsMedalOverkill.Text = medals.GetValueOrDefault("Overkill").ToString();
            StatsMedalKilltacular.Text = medals.GetValueOrDefault("Killtacular").ToString();
            StatsMedalKilltrocity.Text = medals.GetValueOrDefault("Killtrocity").ToString();
            StatsMedalKillimanjaro.Text = medals.GetValueOrDefault("Killimanjaro").ToString();
            StatsMedalKilltastrophe.Text = medals.GetValueOrDefault("Killtastrophe").ToString();
            StatsMedalKillpocalypse.Text = medals.GetValueOrDefault("Killpocalypse").ToString();
            StatsMedalKillionaire.Text = medals.GetValueOrDefault("Killionaire").ToString();

            PublishObsOverlaySnapshot();
        }

        private void StatsRefreshLifetimeUI()
        {
            string gt, kd, totals;
            lock (_statsLock)
            {
                gt = _statsGamertag;
                kd = _statsKd.GetValueOrDefault(gt, "—");
                totals = _statsTotals.GetValueOrDefault(gt, "");
            }
            Dispatcher.InvokeAsync(() =>
            {
                StatsLifetimeKdLabel.Text    = kd;
                StatsLifetimeTotalsLabel.Text = totals;
            });
        }

        private void StatsRebuildCurrentLobbyRows()
        {
            Dictionary<string, MatchmakingPlayerPing> pingSnap;
            Dictionary<string, string> kdSnap, totSnap, gamesSnap;
            Dictionary<string, string> gamertagsByXuid;
            string myGt;

            lock (_statsLock)
            {
                pingSnap = new Dictionary<string, MatchmakingPlayerPing>(_statsMatchmakingPings, StringComparer.OrdinalIgnoreCase);
                kdSnap = new Dictionary<string, string>(_statsKd, StringComparer.OrdinalIgnoreCase);
                totSnap = new Dictionary<string, string>(_statsTotals, StringComparer.OrdinalIgnoreCase);
                gamesSnap = new Dictionary<string, string>(_statsGames, StringComparer.OrdinalIgnoreCase);
                gamertagsByXuid = new Dictionary<string, string>(_statsGamertagsByXuid, StringComparer.OrdinalIgnoreCase);
                myGt = _statsGamertag;
            }

            var cutoff = DateTime.UtcNow - TimeSpan.FromMinutes(30);
            var freshPings = pingSnap.Values
                .Where(p => p.ObservedAt >= cutoff)
                .ToList();
            var squadLabels = StatsBuildSquadLabels(freshPings.Select(StatsGetSquadKey));

            var rows = freshPings
                .OrderBy(p => StatsSquadSortKey(StatsGetSquadKey(p), squadLabels))
                .ThenBy(p => StatsResolveGamertag(p, gamertagsByXuid))
                .ThenBy(p => StatsNormalizeXuid(p.Xuid))
                .Select(p =>
                {
                    string gt = StatsResolveGamertag(p, gamertagsByXuid);
                    string squadKey = StatsGetSquadKey(p);
                    string xuid = StatsNormalizeXuid(p.Xuid);
                    return new StatsPlayerRow
                    {
                        Gamertag = gt,
                        Xuid = xuid,
                        Team = "",
                        KD = kdSnap.GetValueOrDefault(gt, "—"),
                        Totals = totSnap.GetValueOrDefault(gt, ""),
                        GamesPlayed = gamesSnap.GetValueOrDefault(gt, ""),
                        BestServer = StatsFormatServerRegion(p.Region),
                        Ping = p.DisplayPing,
                        SquadId = squadKey,
                        SquadLabel = StatsFormatSquadLabel(squadKey, squadLabels),
                        SkillPercentile = StatsFormatSkillPercentile(p.AverageGroupSkillPercentile),
                        IsMe = gt.Equals(myGt, StringComparison.OrdinalIgnoreCase),
                    };
                })
                .ToList();

            StatsFillMissingSkillPercentilesFromSquads(rows);

            if (rows.Count > 0)
            {
                lock (_statsLock)
                {
                    _statsCurrentLobbySnapshotRows = rows.Select(StatsClonePlayerRow).ToList();
                }
            }

            Dispatcher.InvokeAsync(() =>
            {
                _statsCurrentLobbyRows.Clear();
                foreach (var row in rows)
                    _statsCurrentLobbyRows.Add(row);
            });
            PublishObsOverlaySnapshot();
        }

        private void StatsRebuildLobbyRows()
        {
            List<XElement> players; string myGt;
            Dictionary<string, string> kdSnap, totSnap, gamesSnap, recentKdSnap;
            Dictionary<string, MatchmakingPlayerPing> pingSnap;
            List<StatsPlayerRow> completedLobbyRows;

            lock (_statsLock)
            {
                players      = _statsLastPlayers.ToList();
                myGt         = _statsGamertag;
                kdSnap       = new Dictionary<string, string>(_statsKd,       StringComparer.OrdinalIgnoreCase);
                totSnap      = new Dictionary<string, string>(_statsTotals,   StringComparer.OrdinalIgnoreCase);
                gamesSnap    = new Dictionary<string, string>(_statsGames,    StringComparer.OrdinalIgnoreCase);
                recentKdSnap  = new Dictionary<string, string>(_statsRecentKd, StringComparer.OrdinalIgnoreCase);
                pingSnap      = new Dictionary<string, MatchmakingPlayerPing>(_statsMatchmakingPings, StringComparer.OrdinalIgnoreCase);
                completedLobbyRows = _statsLastCompletedLobbyRows.Select(StatsClonePlayerRow).ToList();
            }

            var completedByXuid = completedLobbyRows
                .Where(r => !string.IsNullOrWhiteSpace(r.Xuid))
                .GroupBy(r => StatsNormalizeXuid(r.Xuid), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            var completedByGamertag = completedLobbyRows
                .Where(r => !string.IsNullOrWhiteSpace(r.Gamertag))
                .GroupBy(r => r.Gamertag, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            bool isFfa = players.Select(p =>
                p.Attribute("mTeamIndex")?.Value ?? p.Attribute("mTeamId")?.Value ?? "0")
                .Distinct().Count() <= 1;

            var rows = players
                .OrderBy(p => p.Attribute("mTeamIndex")?.Value ?? p.Attribute("mTeamId")?.Value ?? "0")
                .ThenBy(p => int.TryParse(p.Attribute("mStanding")?.Value, out int s) ? s : 99)
                .ThenByDescending(p => ParseInt(p.Attribute("Score")?.Value))
                .ThenByDescending(p => ParseInt(p.Attribute("mKills")?.Value))
                .Select(p =>
                {
                    string gt    = p.Attribute("mGamertagText")?.Value ?? "Unknown";
                    string xuid  = p.Attribute("mXboxUserId")?.Value ?? "";
                    string normalizedXuid = StatsNormalizeXuid(xuid);
                    string team  = isFfa ? "FFA"
                        : p.Attribute("mTeamIndex")?.Value ?? p.Attribute("mTeamId")?.Value ?? "0";
                    string kd       = kdSnap.GetValueOrDefault(gt, "—");
                    string recentKd = recentKdSnap.GetValueOrDefault(gt, "");
                    int standing = int.TryParse(p.Attribute("mStanding")?.Value, out int parsedStanding)
                        ? parsedStanding
                        : 99;
                    completedByXuid.TryGetValue(normalizedXuid, out var completedLobbyRow);
                    if (completedLobbyRow is null)
                        completedByGamertag.TryGetValue(gt, out completedLobbyRow);
                    pingSnap.TryGetValue(normalizedXuid, out var playerPing);
                    if (playerPing is not null && DateTime.UtcNow - playerPing.ObservedAt > TimeSpan.FromHours(2))
                        playerPing = null;

                    kd = StatsPreferCopiedStat(completedLobbyRow?.KD, kd);
                    string totals = StatsPreferCopiedStat(completedLobbyRow?.Totals, totSnap.GetValueOrDefault(gt, ""));
                    string gamesPlayed = StatsPreferCopiedStat(completedLobbyRow?.GamesPlayed, gamesSnap.GetValueOrDefault(gt, ""));
                    string bestServer = StatsPreferCopiedDisplay(completedLobbyRow?.BestServer, StatsFormatServerRegion(playerPing?.Region ?? ""));
                    string ping = StatsPreferCopiedDisplay(completedLobbyRow?.Ping, playerPing is null ? "—" : playerPing.DisplayPing);

                    string trend = "";
                    if (!string.IsNullOrEmpty(recentKd) &&
                        double.TryParse(recentKd, out double rkd) &&
                        double.TryParse(kd, out double lkd))
                    {
                        trend = rkd > lkd + 0.05 ? "▲"
                              : rkd < lkd - 0.05 ? "▼"
                              : "≈";
                    }

                    return new StatsPlayerRow
                    {
                        Gamertag      = gt,
                        Xuid          = normalizedXuid,
                        Team          = team,
                        KD            = kd,
                        Totals        = totals,
                        GamesPlayed   = gamesPlayed,
                        BestServer    = bestServer,
                        Ping          = ping,
                        SquadId       = completedLobbyRow?.SquadId ?? "",
                        SquadLabel    = completedLobbyRow?.SquadLabel ?? "",
                        SkillPercentile = StatsPreferCopiedDisplay(completedLobbyRow?.SkillPercentile, "—"),
                        IsMe          = gt.Equals(myGt, StringComparison.OrdinalIgnoreCase),
                        IsScanning    = kd == "…",
                        Standing      = standing,
                        MatchResult   = StatsFormatMatchResult(standing, isFfa),
                        MatchScore    = ParseInt(p.Attribute("Score")?.Value),
                        MatchKills    = ParseInt(p.Attribute("mKills")?.Value),
                        MatchDeaths   = ParseInt(p.Attribute("mDeaths")?.Value),
                        MatchAssists  = ParseInt(p.Attribute("mAssists")?.Value),
                        MatchObjectiveStats = StatsFormatMatchCustomStats(p),
                        RecentKD      = recentKd,
                        RecentTrend   = trend,
                    };
                })
                .ToList();

            // Weighted team averages
            var teamStats = rows
                .Where(r => r.Team != "FFA" && double.TryParse(r.KD, out _))
                .GroupBy(r => r.Team)
                .ToDictionary(
                    g => g.Key,
                    g =>
                    {
                        double weightedKdSum = 0, totalWeight = 0, gamesSum = 0; int count = 0;
                        foreach (var r in g)
                        {
                            double kd = double.Parse(r.KD);
                            long games = long.TryParse(r.GamesPlayed.Replace(",", ""), out long gp) ? gp : 0;
                            double weight = games > 0 ? games : 1;
                            weightedKdSum += kd * weight; totalWeight += weight;
                            gamesSum += games; count++;
                        }
                        double avgKd    = totalWeight > 0 ? weightedKdSum / totalWeight : 0;
                        double avgGames = count > 0 ? gamesSum / count : 0;
                        return (avgKd, avgGames);
                    });

            Dispatcher.InvokeAsync(() =>
            {
                _statsLobbyRows.Clear();
                foreach (var r in rows) _statsLobbyRows.Add(r);

                if (!isFfa && teamStats.Count >= 2 &&
                    teamStats.TryGetValue("0", out var s0) &&
                    teamStats.TryGetValue("1", out var s1))
                {
                    StatsTeam0AvgLabel.Text   = s0.avgKd.ToString("F2");
                    StatsTeam1AvgLabel.Text   = s1.avgKd.ToString("F2");
                    StatsTeam0GamesLabel.Text = s0.avgGames > 0 ? $"~{s0.avgGames:N0} avg games" : "";
                    StatsTeam1GamesLabel.Text = s1.avgGames > 0 ? $"~{s1.avgGames:N0} avg games" : "";
                    bool t0Favored = s0.avgKd > s1.avgKd;
                    StatsTeam0FavoredLabel.Text = t0Favored  ? "▲ FAVORED" : "";
                    StatsTeam1FavoredLabel.Text = !t0Favored ? "▲ FAVORED" : "";
                    StatsTeamSummaryBar.Visibility = Visibility.Visible;
                }
                else
                {
                    StatsTeamSummaryBar.Visibility = Visibility.Collapsed;
                }
            });
        }

        private static string StatsFormatMatchResult(int standing, bool isFfa)
        {
            if (standing < 0 || standing >= 99)
                return "—";

            if (isFfa)
                return $"#{standing + 1}";

            return standing switch
            {
                0 => "WIN",
                1 => "LOSS",
                _ => $"#{standing + 1}"
            };
        }

        private static string StatsFormatMatchCustomStats(XElement player)
        {
            var stats = player.Element("CustomStats")?
                .Elements("CustomStat")
                .Select(stat => new
                {
                    Name = StatsFormatCustomStatName(stat.Attribute("mStatName")?.Value ?? ""),
                    Value = stat.Attribute("mValueForDisplay")?.Value?.Trim() ?? ""
                })
                .Where(stat => !string.IsNullOrWhiteSpace(stat.Name) &&
                               !string.IsNullOrWhiteSpace(stat.Value))
                .Select(stat => $"{stat.Name}: {stat.Value}")
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            return stats.Count == 0 ? "—" : string.Join("  ·  ", stats);
        }

        private static string StatsFormatCustomStatName(string value)
        {
            string name = value.Trim().TrimStart('$');
            if (string.IsNullOrWhiteSpace(name))
                return "";

            name = name.Replace('_', ' ');
            name = Regex.Replace(name, "(?<=[a-z0-9])(?=[A-Z])", " ");
            name = Regex.Replace(name, "\\s+", " ").Trim();
            return name.ToUpperInvariant();
        }

        private static string StatsPreferCopiedStat(string? copied, string fallback)
        {
            if (string.IsNullOrWhiteSpace(copied) || copied == "—" || copied == "…")
                return fallback;

            return copied;
        }

        private static string StatsPreferCopiedDisplay(string? copied, string fallback)
        {
            if (string.IsNullOrWhiteSpace(copied) || copied == "—")
                return fallback;

            return copied;
        }

        private static StatsPlayerRow StatsClonePlayerRow(StatsPlayerRow row) => new()
        {
            Gamertag = row.Gamertag,
            Xuid = row.Xuid,
            Team = row.Team,
            KD = row.KD,
            Totals = row.Totals,
            GamesPlayed = row.GamesPlayed,
            BestServer = row.BestServer,
            Ping = row.Ping,
            SquadId = row.SquadId,
            SquadLabel = row.SquadLabel,
            SkillPercentile = row.SkillPercentile,
            IsMe = row.IsMe,
            IsScanning = row.IsScanning,
            Standing = row.Standing,
            MatchResult = row.MatchResult,
            MatchScore = row.MatchScore,
            MatchKills = row.MatchKills,
            MatchDeaths = row.MatchDeaths,
            MatchAssists = row.MatchAssists,
            MatchObjectiveStats = row.MatchObjectiveStats,
            RecentKD = row.RecentKD,
            RecentTrend = row.RecentTrend,
        };

        private static string StatsNormalizeXuid(string xuid)
        {
            if (string.IsNullOrWhiteSpace(xuid))
                return "";

            string trimmed = xuid.Trim();
            if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                return ulong.TryParse(
                        trimmed[2..],
                        System.Globalization.NumberStyles.HexNumber,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out ulong value)
                    ? value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : trimmed;
            }

            return trimmed;
        }

        private static string StatsNormalizeTeamForDisplay(string team)
        {
            if (string.IsNullOrWhiteSpace(team))
                return "—";

            string trimmed = team.Trim();
            return trimmed switch
            {
                "0" => "0",
                "1" => "1",
                _ when trimmed.Equals("red", StringComparison.OrdinalIgnoreCase) => "0",
                _ when trimmed.Equals("blue", StringComparison.OrdinalIgnoreCase) => "1",
                _ => trimmed
            };
        }

        private static string StatsNormalizeTeamForSort(string team)
        {
            string normalized = StatsNormalizeTeamForDisplay(team);
            return normalized == "—" ? "9" : normalized;
        }

        private void StatsRememberGamertagForXuid(string xuid, string gamertag)
        {
            string normalizedXuid = StatsNormalizeXuid(xuid);
            if (string.IsNullOrWhiteSpace(normalizedXuid) || string.IsNullOrWhiteSpace(gamertag))
                return;

            string trimmedGamertag = gamertag.Trim();
            if (StatsLooksLikeXuid(trimmedGamertag))
                return;

            lock (_statsLock)
                _statsGamertagsByXuid[normalizedXuid] = trimmedGamertag;
        }

        private static string StatsResolveGamertag(
            MatchmakingPlayerPing ping,
            IReadOnlyDictionary<string, string> gamertagsByXuid)
        {
            string normalizedXuid = StatsNormalizeXuid(ping.Xuid);
            if (!string.IsNullOrWhiteSpace(ping.Gamertag) && !StatsLooksLikeXuid(ping.Gamertag))
                return ping.Gamertag.Trim();

            if (!string.IsNullOrWhiteSpace(normalizedXuid) &&
                gamertagsByXuid.TryGetValue(normalizedXuid, out var cachedGamertag) &&
                !string.IsNullOrWhiteSpace(cachedGamertag))
            {
                return cachedGamertag;
            }

            return string.IsNullOrWhiteSpace(normalizedXuid) ? "Name unavailable" : $"XUID {StatsShortXuid(normalizedXuid)}";
        }

        private static bool StatsLooksLikeXuid(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string trimmed = StatsNormalizeXuid(value);
            return trimmed.Length >= 12 && trimmed.All(char.IsDigit);
        }

        private static string StatsShortXuid(string xuid)
        {
            string normalized = StatsNormalizeXuid(xuid);
            return normalized.Length <= 4 ? normalized : $"...{normalized[^4..]}";
        }

        private static string StatsFormatServerRegion(string region)
        {
            if (string.IsNullOrWhiteSpace(region))
                return "—";

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

        private static string StatsFormatSkillPercentile(double? value)
        {
            if (!value.HasValue)
                return "—";

            double percentile = value.Value;
            if (percentile > 0 && percentile <= 1)
                percentile *= 100;

            return percentile.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string StatsGetSquadKey(MatchmakingPlayerPing ping)
        {
            if (!string.IsNullOrWhiteSpace(ping.SquadId))
                return ping.SquadId.Trim();

            if (!ping.AverageGroupSkillPercentile.HasValue)
                return "";

            double percentile = ping.AverageGroupSkillPercentile.Value;
            if (percentile > 0 && percentile <= 1)
                percentile *= 100;

            string roundedPercentile = percentile.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
            return $"skill:{roundedPercentile}";
        }

        private static void StatsFillMissingSkillPercentilesFromSquads(List<StatsPlayerRow> rows)
        {
            var squadPercentiles = rows
                .Where(r => StatsIsKnownSquadLabel(r.SquadLabel) && r.SkillPercentile != "—")
                .GroupBy(r => r.SquadLabel, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().SkillPercentile, StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                if (row.SkillPercentile != "—" ||
                    !StatsIsKnownSquadLabel(row.SquadLabel) ||
                    !squadPercentiles.TryGetValue(row.SquadLabel, out var percentile))
                {
                    continue;
                }

                row.SkillPercentile = percentile;
            }
        }

        private static Dictionary<string, string> StatsBuildSquadLabels(IEnumerable<string> squadIds)
        {
            return squadIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id.Trim())
                .GroupBy(id => id, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .Select((g, index) => new { SquadId = g.Key, Label = $"S{index + 1}" })
                .ToDictionary(x => x.SquadId, x => x.Label, StringComparer.OrdinalIgnoreCase);
        }

        private static string StatsFormatSquadLabel(string squadId, IReadOnlyDictionary<string, string> squadLabels) =>
            !string.IsNullOrWhiteSpace(squadId) && squadLabels.TryGetValue(squadId.Trim(), out var label)
                ? label
                : "?";

        private static bool StatsIsKnownSquadLabel(string squadLabel) =>
            !string.IsNullOrWhiteSpace(squadLabel) && squadLabel != "?";

        private static string StatsSquadSortKey(string squadId, IReadOnlyDictionary<string, string> squadLabels)
        {
            string label = StatsFormatSquadLabel(squadId, squadLabels);
            return label == "?" ? "ZZZ" : label;
        }

        // ── File monitoring ───────────────────────────────────────────────────

        private void StatsInitializeSignature()
        {
            if (!Directory.Exists(StatsWatchPath)) return;
            var f = StatsLatestCarnageFile();
            if (f != null) lock (_statsLock) { _statsLastFileSig = StatsSig(f); }
        }

        private async Task StatsMonitorLoop()
        {
            while (true)
            {
                try { StatsCheckForNewFile(); }
                catch (Exception ex) { StatsSetStatus($"Stat tracker error: {ex.Message}"); }
                await Task.Delay(1000);
            }
        }

        private void StatsCheckForNewFile()
        {
            if (!Directory.Exists(StatsWatchPath)) return;
            var f = StatsLatestCarnageFile();
            if (f == null) return;
            string sig = StatsSig(f);
            bool changed;
            lock (_statsLock) { changed = sig != _statsLastFileSig; }
            if (changed && StatsTryProcessFile(f.FullName))
            {
                // Do not acknowledge MCC's file until it has been parsed completely.
                // MCC writes this XML progressively, so an early read can be incomplete.
                lock (_statsLock) { _statsLastFileSig = sig; }
            }
        }

        private static FileInfo? StatsLatestCarnageFile() =>
            new DirectoryInfo(StatsWatchPath)
                .GetFiles("mpcarnagereport*.xml")
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

        private static string StatsSig(FileInfo f) =>
            $"{f.FullName}|{f.Length}|{f.LastWriteTime.Ticks}";

        private bool StatsTryProcessFile(string path, bool countTowardSession = true)
        {
            try
            {
                return StatsProcessFile(path, countTowardSession);
            }
            catch (Exception ex)
            {
                StatsSetStatus($"Could not load {Path.GetFileName(path)}: {ex.Message}");
                return false;
            }
        }

        private bool StatsProcessFile(string path, bool countTowardSession = true)
        {
            XDocument? doc = null;
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    doc = XDocument.Load(stream);
                    break;
                }
                catch { Thread.Sleep(300); }
            }
            if (doc == null)
            {
                StatsSetStatus($"Waiting for {Path.GetFileName(path)} to finish writing…");
                return false;
            }

            string? gId = doc.Descendants("GameUniqueId")
                .Select(e =>
                {
                    string? attr = e.Attribute("GameUniqueId")?.Value;
                    return !string.IsNullOrEmpty(attr) ? attr : e.Value;
                })
                .FirstOrDefault(v => !string.IsNullOrEmpty(v));

            if (string.IsNullOrEmpty(gId))
            {
                StatsSetStatus($"Waiting for {Path.GetFileName(path)} game ID…");
                return false;
            }

            var players = doc.Descendants("Player").ToList();
            if (players.Count == 0)
            {
                StatsSetStatus($"Waiting for {Path.GetFileName(path)} player stats…");
                return false;
            }
            bool triggerLobby = false;
            bool alreadyLogged = false;
            bool playerFound = false;
            string trackedGamertag = "";
            string gameTypeName = doc.Root?.Element("GameTypeName")?.Attribute("GameTypeName")?.Value
                ?? doc.Root?.Element("GameTypeName")?.Value
                ?? "";
            string playedAt = File.GetLastWriteTime(path).ToString("MMM d, h:mm tt");

            lock (_statsLock)
            {
                _statsLastPlayers = players;
                foreach (var player in players)
                {
                    string playerXuid = StatsNormalizeXuid(player.Attribute("mXboxUserId")?.Value ?? "");
                    StatsRememberGamertagForXuid(
                        playerXuid,
                        player.Attribute("mGamertagText")?.Value ?? "");
                }
                _statsLastCompletedLobbyRows = StatsMatchLobbyRowsToPlayers(_statsCurrentLobbySnapshotRows, players)
                    .Select(StatsClonePlayerRow)
                    .ToList();
                _statsLastGameServerText = _statsCurrentLobbyServerText;
                _statsLastGameModeText = gameTypeName;
                _statsLastGamePlayedText = playedAt;
                trackedGamertag = _statsGamertag.Trim();
                alreadyLogged = _statsSession.ProcessedGameIds.Contains(gId);
                if (countTowardSession && !alreadyLogged)
                {
                    var me = players.FirstOrDefault(p =>
                        string.Equals(
                            p.Attribute("mGamertagText")?.Value?.Trim(),
                            trackedGamertag,
                            StringComparison.OrdinalIgnoreCase));

                    if (me != null)
                    {
                        playerFound = true;
                        int.TryParse(me.Attribute("mStanding")?.Value, out int standing);
                        long.TryParse(me.Attribute("mKills")?.Value,   out long k);
                        long.TryParse(me.Attribute("mDeaths")?.Value,  out long d);
                        int.TryParse(me.Attribute("mMostKillsInARow")?.Value, out int spree);
                        var gameMedals = StatsReadMultikillCounts(me);
                        string highestMultikill = StatsHighestMultikill(gameMedals);
                        bool won = standing == 0;
                        long previousKills = _statsSession.Kills;
                        long previousDeaths = _statsSession.Deaths;
                        int previousBestSpree = _statsSession.BestSpree;
                        var previousMedals = new Dictionary<string, int>(
                            _statsSession.MultikillCounts,
                            StringComparer.OrdinalIgnoreCase);

                        _statsSession.Kills += k;
                        _statsSession.Deaths += d;
                        _statsSession.GamesPlayed++;
                        _statsSession.ProcessedGameIds.Add(gId);
                        if (won)
                        {
                            _statsSession.Wins++;
                            _statsSession.CurrentWinStreak++;
                            _statsSession.LongestWinStreak = Math.Max(
                                _statsSession.LongestWinStreak,
                                _statsSession.CurrentWinStreak);
                        }
                        else
                        {
                            _statsSession.Losses++;
                            _statsSession.CurrentWinStreak = 0;
                        }

                        _statsSession.BestSpree = Math.Max(_statsSession.BestSpree, spree);
                        double gameKd = d > 0 ? (double)k / d : k;
                        if (gameKd > _statsSession.BestGameKd)
                        {
                            _statsSession.BestGameKd = gameKd;
                            _statsSession.BestGameScore = $"{k}–{d}";
                        }

                        foreach (var medal in StatsMultikillMedals)
                            _statsSession.MultikillCounts[medal.Name] += gameMedals.GetValueOrDefault(medal.Name);

                        double previousSessionKd = previousDeaths > 0
                            ? (double)previousKills / previousDeaths
                            : previousKills;
                        double newSessionKd = _statsSession.Deaths > 0
                            ? (double)_statsSession.Kills / _statsSession.Deaths
                            : _statsSession.Kills;
                        var deltas = StatsMultikillMedals
                            .Where(m => gameMedals.GetValueOrDefault(m.Name) > 0)
                            .Select(m => new ObsMedalDelta(
                                m.Name,
                                previousMedals.GetValueOrDefault(m.Name),
                                _statsSession.MultikillCounts.GetValueOrDefault(m.Name),
                                gameMedals.GetValueOrDefault(m.Name)))
                            .ToList();
                        if (deltas.Count > 4)
                        {
                            var featured = deltas.Last();
                            deltas = deltas.Take(3).ToList();
                            if (!deltas.Any(d => d.Name == featured.Name)) deltas.Add(featured);
                        }
                        var capturedAt = DateTimeOffset.UtcNow;
                        _postGameRecap = new ObsPostGameRecap(
                            Won: won,
                            Kills: k,
                            Deaths: d,
                            GameKd: gameKd.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
                            PreviousSessionKd: previousSessionKd.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
                            SessionKd: newSessionKd.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
                            BestSpree: spree,
                            IsNewBestSpree: spree > previousBestSpree,
                            FeaturedMedal: highestMultikill,
                            MedalDeltas: deltas,
                            CapturedAtUtc: capturedAt,
                            ExpiresAtUtc: capturedAt.AddSeconds(10));

                        var sessionPlayerRows = StatsCaptureSessionLobby(
                            players,
                            me,
                            _statsSession.GamesPlayed,
                            won);
                        var lobbyTeams = StatsBuildSessionLobbyTeams(sessionPlayerRows.LobbyPlayers);
                        var gameRow = new StatsSessionGameRow
                        {
                            Game = _statsSession.GamesPlayed,
                            Result = won ? "WIN" : "LOSS",
                            KillsDeaths = $"{k}–{d}",
                            KD = gameKd.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
                            BestSpree = spree,
                            HighestMultikill = highestMultikill,
                            HighestMultikillIcon = StatsMultikillMedals
                                .FirstOrDefault(m => m.Name == highestMultikill)?.ResourcePath ?? "",
                            PlayedAt = File.GetLastWriteTime(path).ToString("h:mm tt"),
                            LobbyPlayerCount = sessionPlayerRows.LobbyPlayers.Count,
                            LobbyColumnCount = Math.Clamp(lobbyTeams.Count, 1, 2),
                            LobbyTeams = lobbyTeams,
                            IsExpanded = _statsSession.GamesPlayed == 1
                        };
                        Dispatcher.InvokeAsync(() =>
                        {
                            _statsSessionGames.Insert(0, gameRow);
                            _statsSessionPlayers.Clear();
                            foreach (var playerRow in sessionPlayerRows.PlayerHistory)
                                _statsSessionPlayers.Add(playerRow);
                            StatsSessionGameCountLabel.Text = _statsSessionGames.Count.ToString();
                            StatsSessionPlayerCountLabel.Text = _statsSessionPlayers.Count.ToString();
                        });

                        StatsSetStatus($"Game logged — K:{k}  D:{d}  Standing:{standing}");
                        triggerLobby = true;
                    }
                }
            }

            StatsRefreshSessionUI();
            StatsRebuildLobbyRows();
            Dispatcher.InvokeAsync(() => StatsLastGameServerLabel.Text = StatsFormatLastGameHeader());
            if (!countTowardSession)
                StatsSetStatus($"Loaded last game: {Path.GetFileName(path)}");
            else if (alreadyLogged)
                StatsSetStatus($"Last game is already logged: {Path.GetFileName(path)}");
            else if (!playerFound)
            {
                string available = string.Join(", ", players
                    .Select(p => p.Attribute("mGamertagText")?.Value?.Trim())
                    .Where(name => !string.IsNullOrWhiteSpace(name)));
                StatsSetStatus(string.IsNullOrWhiteSpace(trackedGamertag)
                    ? "Configure your gamertag on Dashboard before loading the last game."
                    : $"Gamertag '{trackedGamertag}' was not in the report. Players: {available}");
                return false;
            }
            if (triggerLobby) _ = StatsFetchLobbyStats();
            return true;
        }

        private (List<StatsSessionLobbyPlayerRow> LobbyPlayers, List<StatsSessionPlayerHistoryRow> PlayerHistory)
            StatsCaptureSessionLobby(
                IReadOnlyList<XElement> players,
                XElement me,
                int gameNumber,
                bool won)
        {
            string TeamOf(XElement player) =>
                player.Attribute("mTeamIndex")?.Value ??
                player.Attribute("mTeamId")?.Value ??
                "0";

            string myGamertag = me.Attribute("mGamertagText")?.Value?.Trim() ?? "";
            string myXuid = StatsNormalizeXuid(me.Attribute("mXboxUserId")?.Value ?? "");
            string myTeam = TeamOf(me);
            bool isFfa = players.Select(TeamOf).Distinct(StringComparer.OrdinalIgnoreCase).Count() <= 1;
            var lobbyPlayers = new List<StatsSessionLobbyPlayerRow>(players.Count);

            foreach (var player in players
                .OrderBy(TeamOf)
                .ThenBy(p => int.TryParse(p.Attribute("mStanding")?.Value, out int standing) ? standing : 99))
            {
                string gamertag = player.Attribute("mGamertagText")?.Value?.Trim() ?? "Unknown";
                string xuid = StatsNormalizeXuid(player.Attribute("mXboxUserId")?.Value ?? "");
                string team = TeamOf(player);
                bool isMe = (!string.IsNullOrWhiteSpace(myXuid) && xuid.Equals(myXuid, StringComparison.OrdinalIgnoreCase)) ||
                            gamertag.Equals(myGamertag, StringComparison.OrdinalIgnoreCase);
                bool isTeammate = !isMe && !isFfa && team.Equals(myTeam, StringComparison.OrdinalIgnoreCase);
                int encounters = 0;

                if (!isMe)
                {
                    string playerKey = !string.IsNullOrWhiteSpace(xuid)
                        ? $"xuid:{xuid}"
                        : $"gt:{gamertag}";
                    if (!_statsSessionPlayerHistory.TryGetValue(playerKey, out var aggregate))
                    {
                        aggregate = new StatsSessionPlayerAggregate
                        {
                            Gamertag = gamertag,
                            Xuid = xuid
                        };
                        _statsSessionPlayerHistory[playerKey] = aggregate;
                    }

                    aggregate.Gamertag = gamertag;
                    aggregate.LastGame = gameNumber;
                    aggregate.LastSeen = DateTime.Now;
                    aggregate.Matches++;
                    if (won) aggregate.Wins++; else aggregate.Losses++;
                    aggregate.WasTeammate |= isTeammate;
                    aggregate.WasOpponent |= !isTeammate;
                    aggregate.Kills += ParseInt(player.Attribute("mKills")?.Value);
                    aggregate.Deaths += ParseInt(player.Attribute("mDeaths")?.Value);
                    encounters = aggregate.Matches;
                }

                lobbyPlayers.Add(new StatsSessionLobbyPlayerRow
                {
                    Gamertag = gamertag,
                    TeamKey = isFfa ? "FFA" : team,
                    IsMe = isMe,
                    Kills = ParseInt(player.Attribute("mKills")?.Value),
                    Deaths = ParseInt(player.Attribute("mDeaths")?.Value),
                    Assists = ParseInt(player.Attribute("mAssists")?.Value),
                    EncounterLabel = isMe ? "YOU" : encounters <= 1 ? "NEW" : $"{encounters}×"
                });
            }

            var history = _statsSessionPlayerHistory.Values
                .OrderByDescending(player => player.LastGame)
                .ThenByDescending(player => player.Matches)
                .ThenBy(player => player.Gamertag, StringComparer.OrdinalIgnoreCase)
                .Select(player => new StatsSessionPlayerHistoryRow
                {
                    Gamertag = player.Gamertag,
                    LastSeen = $"GAME {player.LastGame}  {player.LastSeen:h:mm tt}",
                    Encounter = player.WasTeammate && player.WasOpponent
                        ? "BOTH"
                        : player.WasTeammate ? "TEAMMATE" : "OPPONENT",
                    Matches = player.Matches,
                    Record = $"{player.Wins}W–{player.Losses}L",
                    AverageKD = (player.Deaths > 0 ? (double)player.Kills / player.Deaths : player.Kills)
                        .ToString("F2", System.Globalization.CultureInfo.InvariantCulture)
                })
                .ToList();

            return (lobbyPlayers, history);
        }

        private static List<StatsSessionLobbyTeamRow> StatsBuildSessionLobbyTeams(
            IReadOnlyList<StatsSessionLobbyPlayerRow> lobbyPlayers)
        {
            return lobbyPlayers
                .GroupBy(player => player.TeamKey, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => StatsSessionTeamSort(group.Key))
                .Select(group => new StatsSessionLobbyTeamRow
                {
                    TeamName = StatsSessionTeamName(group.Key),
                    TeamColor = StatsSessionTeamColor(group.Key),
                    Players = group.ToList()
                })
                .ToList();
        }

        private static int StatsSessionTeamSort(string team) => team.ToUpperInvariant() switch
        {
            "0" => 0,
            "1" => 1,
            "2" => 2,
            "3" => 3,
            "FFA" => 0,
            _ => 9
        };

        private static string StatsSessionTeamName(string team) => team.ToUpperInvariant() switch
        {
            "0" => "RED TEAM",
            "1" => "BLUE TEAM",
            "2" => "GREEN TEAM",
            "3" => "YELLOW TEAM",
            "FFA" => "FFA LOBBY",
            _ => $"TEAM {team}"
        };

        private static Brush StatsSessionTeamColor(string team)
        {
            Color color = team.ToUpperInvariant() switch
            {
                "0" => Color.FromRgb(0xFF, 0x2D, 0x55),
                "1" => Color.FromRgb(0x00, 0xC8, 0xFF),
                "2" => Color.FromRgb(0x39, 0xFF, 0x14),
                "3" => Color.FromRgb(0xFF, 0xD6, 0x0A),
                _ => Color.FromRgb(0x00, 0xC8, 0xFF)
            };
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        private static Dictionary<string, int> StatsReadMultikillCounts(XElement player)
        {
            var byId = player.Descendants("Medal")
                .Select(e => new
                {
                    Id = ParseInt(e.Attribute("mId")?.Value),
                    Count = ParseInt(e.Attribute("mCount")?.Value)
                })
                .GroupBy(x => x.Id)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Count));

            return StatsMultikillMedals.ToDictionary(
                medal => medal.Name,
                medal => byId.GetValueOrDefault(medal.CarnageId),
                StringComparer.OrdinalIgnoreCase);
        }

        private static string StatsHighestMultikill(IReadOnlyDictionary<string, int> counts)
        {
            for (int i = StatsMultikillMedals.Length - 1; i >= 0; i--)
                if (counts.GetValueOrDefault(StatsMultikillMedals[i].Name) > 0)
                    return StatsMultikillMedals[i].Name;
            return "—";
        }

        private static List<StatsPlayerRow> StatsMatchLobbyRowsToPlayers(
            IEnumerable<StatsPlayerRow> lobbyRows,
            IEnumerable<XElement> players)
        {
            var playerXuids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var playerGamertags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var player in players)
            {
                string xuid = StatsNormalizeXuid(player.Attribute("mXboxUserId")?.Value ?? "");
                if (!string.IsNullOrWhiteSpace(xuid))
                    playerXuids.Add(xuid);

                string gamertag = player.Attribute("mGamertagText")?.Value ?? "";
                if (!string.IsNullOrWhiteSpace(gamertag))
                    playerGamertags.Add(gamertag);
            }

            return lobbyRows
                .Where(row =>
                {
                    string xuid = StatsNormalizeXuid(row.Xuid);
                    return (!string.IsNullOrWhiteSpace(xuid) && playerXuids.Contains(xuid)) ||
                           (!string.IsNullOrWhiteSpace(row.Gamertag) && playerGamertags.Contains(row.Gamertag));
                })
                .ToList();
        }

        // ── API orchestration ─────────────────────────────────────────────────

        private async Task StatsFetchStats(string gt)
        {
            if (_rejoinProxy.HasPopulationAuthorization)
            {
                string xuid = StatsResolveEnteredTargetToCachedXuid(gt);
                if (string.IsNullOrEmpty(xuid) && gt.Equals(_rejoinProxy.CurrentPlayerGamertag, StringComparison.OrdinalIgnoreCase))
                    xuid = _rejoinProxy.CurrentPlayerXuid;
                if (string.IsNullOrEmpty(xuid))
                {
                    string token;
                    lock (_statsLock) { token = _statsSpartanToken; }
                    xuid = await StatsResolveEnteredTargetToXuidAsync(gt, token);
                }
                var live = await _rejoinProxy.GetCareerTotalsAsync(xuid);
                if (live is { } totals)
                {
                    string kd = (totals.deaths > 0 ? (double)totals.kills / totals.deaths : totals.kills).ToString("F2");
                    lock (_statsLock)
                    {
                        _statsKd[gt] = kd;
                        _statsTotals[gt] = $"{totals.kills:N0}K / {totals.deaths:N0}D · Xbox Live";
                    }
                    StatsSetStatus($"[Xbox Live] {gt} — K/D: {kd}");
                    StatsRefreshLifetimeUI();
                    StatsRebuildCurrentLobbyRows();
                    StatsRebuildLobbyRows();
                    return;
                }
            }
            // Wort uses the same matchmaking totals; do not substitute Waypoint's different scope.
            await StatsFetchWortStats(gt);
        }

        private async Task<(bool success, bool unauthorized)> StatsFetchHaloWaypointStats(string gt, string token)
        {
            try
            {
                string url = $"https://mccapi.svc.halowaypoint.com/hmcc/users/gt({Uri.EscapeDataString(gt)})/service-record";
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.TryAddWithoutValidation("x-343-authorization-spartan", token);
                req.Headers.TryAddWithoutValidation("Accept", "application/json");

                var resp = await StatsHttp.SendAsync(req);

                if (resp.StatusCode == HttpStatusCode.Unauthorized)
                    return (false, true);

                if (!resp.IsSuccessStatusCode)
                {
                    StatsSetStatus($"[HW] API {(int)resp.StatusCode} for {gt}");
                    return (false, false);
                }

                lock (_statsLock)
                {
                    if (_statsSpartanToken == token)
                    {
                        _statsTokenLastValidatedUtc = DateTimeOffset.UtcNow;
                        _statsWaypointUnavailable = false;
                        _statsHwTokenExpired = false;
                    }
                }
                StatsUpdateHwStatus();

                string body = await resp.Content.ReadAsStringAsync();
                using var json = JsonDocument.Parse(body);
                var root = json.RootElement;

                if (!root.TryGetProperty("multiplayer", out var mp) ||
                    mp.ValueKind != JsonValueKind.Object)
                    return (false, false);

                mp.TryGetProperty("kills",       out var kEl);  kEl.TryGetInt64(out long kills);
                mp.TryGetProperty("deaths",      out var dEl);  dEl.TryGetInt64(out long deaths);
                mp.TryGetProperty("gamesPlayed", out var gpEl); gpEl.TryGetInt64(out long gamesPlayed);

                string kdVal, totals;
                if (kills == 0 && deaths == 0)
                {
                    kdVal = "N/A"; totals = "";
                }
                else
                {
                    kdVal  = deaths > 0 ? ((double)kills / deaths).ToString("F2") : kills.ToString();
                    totals = $"{kills:N0}K / {deaths:N0}D";
                }
                string gamesStr = gamesPlayed > 0 ? gamesPlayed.ToString("N0") : "";
                lock (_statsLock)
                {
                    _statsKd[gt]     = kdVal;
                    _statsTotals[gt] = totals;
                    _statsGames[gt]  = gamesStr;
                }
                if (kdVal != "N/A") StatsAddToCache(gt, kdVal, totals);
                StatsSetStatus($"[HW] {gt} — K/D: {kdVal}");
                StatsRefreshLifetimeUI();
                StatsRebuildCurrentLobbyRows();
                StatsRebuildLobbyRows();
                return (true, false);
            }
            catch (Exception ex)
            {
                StatsSetStatus($"[HW] Error for {gt}: {ex.Message}");
                return (false, false);
            }
        }

        private async Task StatsFetchRecentStatsAsync(string gt, string token)
        {
            try
            {
                var (firstMatches, maxPage) = await StatsFetchPageWithMetaAsync(gt, token, 1);
                int totalPages = Math.Clamp(maxPage > 0 ? maxPage : 1, 1, 5);

                var allMatches = new List<(DateTime date, long kills, long deaths)>(firstMatches);
                if (totalPages > 1)
                {
                    var restTasks = Enumerable.Range(2, totalPages - 1)
                        .Select(p => StatsFetchMatchPageRawAsync(gt, token, p))
                        .ToArray();
                    foreach (var page in await Task.WhenAll(restTasks))
                        allMatches.AddRange(page);
                }

                var matches = allMatches.OrderByDescending(m => m.date).ToList();
                if (!matches.Any()) return;

                long totalKills  = matches.Sum(m => m.kills);
                long totalDeaths = matches.Sum(m => m.deaths);
                double kd = totalDeaths > 0 ? (double)totalKills / totalDeaths : totalKills;

                lock (_statsLock) { _statsRecentKd[gt] = kd.ToString("F2"); }
                StatsRebuildCurrentLobbyRows();
                StatsRebuildLobbyRows();
            }
            catch { }
        }

        private async Task<(List<(DateTime date, long kills, long deaths)> matches, int maxPage)>
            StatsFetchPageWithMetaAsync(string gt, string token, int page)
        {
            try
            {
                string url = $"https://mccapi.svc.halowaypoint.com/hmcc/users/gt({Uri.EscapeDataString(gt)})/matches?page={page}&pageSize=20";
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.TryAddWithoutValidation("x-343-authorization-spartan", token);
                req.Headers.TryAddWithoutValidation("Accept", "application/json");

                var resp = await StatsHttp.SendAsync(req);
                if (!resp.IsSuccessStatusCode) return (new(), 0);

                string body = await resp.Content.ReadAsStringAsync();
                using var json = JsonDocument.Parse(body);
                var root = json.RootElement;

                int maxPage = root.TryGetProperty("maxPage", out var mpEl) &&
                              mpEl.TryGetInt32(out int mp) ? mp : 1;

                if (!root.TryGetProperty("matches", out var arr) || arr.ValueKind != JsonValueKind.Array)
                    return (new(), maxPage);

                var result = new List<(DateTime, long, long)>();
                foreach (var m in arr.EnumerateArray())
                {
                    DateTime date = m.TryGetProperty("datePlayed", out var dpEl) &&
                                    dpEl.TryGetDateTime(out var dt) ? dt : DateTime.MinValue;
                    m.TryGetProperty("kills",  out var kEl); kEl.TryGetInt64(out long kills);
                    m.TryGetProperty("deaths", out var dEl); dEl.TryGetInt64(out long deaths);
                    result.Add((date, kills, deaths));
                }
                return (result, maxPage);
            }
            catch { return (new(), 0); }
        }

        private async Task<List<(DateTime date, long kills, long deaths)>> StatsFetchMatchPageRawAsync(
            string gt, string token, int page)
        {
            try
            {
                string url = $"https://mccapi.svc.halowaypoint.com/hmcc/users/gt({Uri.EscapeDataString(gt)})/matches?page={page}&pageSize=20";
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.TryAddWithoutValidation("x-343-authorization-spartan", token);
                req.Headers.TryAddWithoutValidation("Accept", "application/json");

                var resp = await StatsHttp.SendAsync(req);
                if (!resp.IsSuccessStatusCode) return new();

                string body = await resp.Content.ReadAsStringAsync();
                using var json = JsonDocument.Parse(body);
                if (!json.RootElement.TryGetProperty("matches", out var arr) || arr.ValueKind != JsonValueKind.Array)
                    return new();

                var result = new List<(DateTime, long, long)>();
                foreach (var m in arr.EnumerateArray())
                {
                    DateTime date = m.TryGetProperty("datePlayed", out var dpEl) &&
                                    dpEl.TryGetDateTime(out var dt) ? dt : DateTime.MinValue;
                    m.TryGetProperty("kills",  out var kEl); kEl.TryGetInt64(out long kills);
                    m.TryGetProperty("deaths", out var dEl); dEl.TryGetInt64(out long deaths);
                    result.Add((date, kills, deaths));
                }
                return result;
            }
            catch { return new(); }
        }

        // ── wort.gg fallback ──────────────────────────────────────────────────

        private async Task<bool> StatsFetchWortStats(string gt)
        {
            bool success = false;
            try
            {
                string url = $"https://wort.gg/api/stats/{Uri.EscapeDataString(gt)}/multiplayer";
                using var resp = await StatsHttp.GetAsync(url);
                string body = await resp.Content.ReadAsStringAsync();
                using var json = JsonDocument.Parse(body);

                if (!resp.IsSuccessStatusCode)
                {
                    StatsShowLiveStatsUnavailable(gt);
                    StatsRefreshLifetimeUI(); StatsRebuildLobbyRows(); return false;
                }

                var (kills, deaths) = StatsExtractWortKillsDeaths(json.RootElement);
                if (kills > 0 || deaths > 0)
                {
                    string kdVal = deaths > 0 ? ((double)kills / deaths).ToString("F2") : kills.ToString();
                    string totals = $"{kills:N0}K / {deaths:N0}D · wort.gg";
                    lock (_statsLock) { _statsKd[gt] = kdVal; _statsTotals[gt] = totals; }
                    StatsAddToCache(gt, kdVal, totals);
                    StatsSetStatus($"[wort.gg] {gt} — K/D: {kdVal}");
                    success = true;
                }
                else
                {
                    StatsShowLiveStatsUnavailable(gt);
                }
            }
            catch (Exception)
            {
                StatsShowLiveStatsUnavailable(gt);
            }
            StatsRefreshLifetimeUI();
            StatsRebuildCurrentLobbyRows();
            StatsRebuildLobbyRows();
            return success;
        }

        private void StatsShowLiveStatsUnavailable(string gt)
        {
            lock (_statsLock) { _statsKd[gt] = "—"; _statsTotals[gt] = "Live stats temporarily unavailable"; }
            StatsSetStatus($"Could not retrieve live stats for {gt}. Try Sync again.");
        }
        private static (long kills, long deaths) StatsExtractWortKillsDeaths(JsonElement root)
        {
            if (root.TryGetProperty("stats", out var statsEl) &&
                statsEl.TryGetProperty("Multiplayer", out var multi) &&
                multi.TryGetProperty("Matchmaking", out var mm) &&
                mm.TryGetProperty("All", out var all) &&
                all.TryGetProperty("Stats", out var stats) &&
                stats.ValueKind == JsonValueKind.Object)
            {
                long kills  = stats.TryGetProperty("kills",  out var kEl) && kEl.ValueKind == JsonValueKind.Number ? kEl.GetInt64() : 0;
                long deaths = stats.TryGetProperty("deaths", out var dEl) && dEl.ValueKind == JsonValueKind.Number ? dEl.GetInt64() : 0;
                return (kills, deaths);
            }
            return (0, 0);
        }

        // ── Lobby scan ────────────────────────────────────────────────────────

        private async Task StatsFetchLobbyStats()
        {
            List<XElement> snapshot;
            lock (_statsLock) { snapshot = _statsLastPlayers.ToList(); }
            if (!snapshot.Any()) { StatsSetStatus("No lobby data yet — play a game first."); return; }

            StatsSetStatus("Scanning lobby…");
            _ = Dispatcher.InvokeAsync(() => StatsScanBtn.IsEnabled = false);

            string hwToken; bool hwExpired;
            lock (_statsLock) { hwToken = _statsSpartanToken; hwExpired = _statsHwTokenExpired; }
            bool useHw = !string.IsNullOrEmpty(hwToken) && !hwExpired;

            var rng = new Random();
            foreach (var p in snapshot)
            {
                string? gt = p.Attribute("mGamertagText")?.Value;
                if (string.IsNullOrEmpty(gt)) continue;

                bool skip, hasRecent;
                lock (_statsLock)
                {
                    skip      = _statsKd.TryGetValue(gt, out string? existing) &&
                                existing != "ERR" && existing != "N/A" && existing != "…";
                    hasRecent = _statsRecentKd.ContainsKey(gt);
                }
                if (skip)
                {
                    if (useHw && !hasRecent) _ = StatsFetchRecentStatsAsync(gt, hwToken);
                    continue;
                }

                if (gt.Contains('(') || gt.Contains(')'))
                {
                    lock (_statsLock) { _statsKd[gt] = "GUEST"; }
                    StatsRebuildLobbyRows();
                    continue;
                }

                if (!useHw)
                {
                    StatsCachedPlayer? cached;
                    lock (_statsLock) { _statsPersistentCache.TryGetValue(gt, out cached); }
                    if (cached != null)
                    {
                        lock (_statsLock) { _statsKd[gt] = cached.KD; _statsTotals[gt] = cached.Totals; }
                        StatsRebuildLobbyRows();
                        continue;
                    }
                }

                lock (_statsLock) { _statsKd[gt] = "…"; }
                StatsRebuildLobbyRows();
                await StatsFetchStats(gt);
                if (useHw) _ = StatsFetchRecentStatsAsync(gt, hwToken);
                await Task.Delay(useHw ? rng.Next(200, 500) : rng.Next(3500, 6000));
            }

            StatsSetStatus("Scan complete.");
            _ = Dispatcher.InvokeAsync(() => StatsScanBtn.IsEnabled = true);
        }

        private async Task StatsFetchCurrentLobbyStats()
        {
            if (!_rejoinProxy.IsRunning)
                return;

            List<string> gamertags;
            lock (_statsLock)
            {
                if (_statsCurrentLobbyScanRunning)
                    return;

                _statsCurrentLobbyScanRunning = true;
                var gamertagsByXuid = new Dictionary<string, string>(_statsGamertagsByXuid, StringComparer.OrdinalIgnoreCase);
                gamertags = _statsMatchmakingPings.Values
                    .Where(p => p.ObservedAt >= DateTime.UtcNow - TimeSpan.FromMinutes(30))
                    .Where(p => (!string.IsNullOrWhiteSpace(p.Gamertag) && !StatsLooksLikeXuid(p.Gamertag)) ||
                                gamertagsByXuid.ContainsKey(StatsNormalizeXuid(p.Xuid)))
                    .Select(p => StatsResolveGamertag(p, gamertagsByXuid))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            try
            {
                if (!gamertags.Any())
                    return;

                StatsSetStatus("Scanning current lobby...");

                string hwToken; bool hwExpired;
                lock (_statsLock) { hwToken = _statsSpartanToken; hwExpired = _statsHwTokenExpired; }
                bool useHw = !string.IsNullOrEmpty(hwToken) && !hwExpired;
                var rng = new Random();

                foreach (var gt in gamertags)
                {
                    if (!_rejoinProxy.IsRunning)
                        return;

                    bool skip, hasRecent;
                    lock (_statsLock)
                    {
                        skip = _statsKd.TryGetValue(gt, out string? existing) &&
                               existing != "ERR" && existing != "N/A" && existing != "…";
                        hasRecent = _statsRecentKd.ContainsKey(gt);
                    }

                    if (skip)
                    {
                        if (useHw && !hasRecent)
                            _ = StatsFetchRecentStatsAsync(gt, hwToken);
                        continue;
                    }

                    if (gt.Contains('(') || gt.Contains(')'))
                    {
                        lock (_statsLock) { _statsKd[gt] = "GUEST"; }
                        StatsRebuildCurrentLobbyRows();
                        continue;
                    }

                    if (!useHw)
                    {
                        StatsCachedPlayer? cached;
                        lock (_statsLock) { _statsPersistentCache.TryGetValue(gt, out cached); }
                        if (cached != null)
                        {
                            lock (_statsLock) { _statsKd[gt] = cached.KD; _statsTotals[gt] = cached.Totals; }
                            StatsRebuildCurrentLobbyRows();
                            continue;
                        }
                    }

                    lock (_statsLock) { _statsKd[gt] = "…"; }
                    StatsRebuildCurrentLobbyRows();
                    await StatsFetchStats(gt);
                    if (useHw)
                        _ = StatsFetchRecentStatsAsync(gt, hwToken);
                    await Task.Delay(useHw ? rng.Next(200, 500) : rng.Next(3500, 6000));
                }

                StatsSetStatus("Current lobby scan complete.");
            }
            finally
            {
                lock (_statsLock) { _statsCurrentLobbyScanRunning = false; }
            }
        }

        // ── Persistence ───────────────────────────────────────────────────────

        private void StatsLoadGamertag()
        {
            _statsGamertag = App.LoadPlayerGamertag();
            if (File.Exists(StatsSettingsFile))
                try { _statsGamertag = File.ReadAllText(StatsSettingsFile).Trim(); } catch { }
        }

        private void StatsLoadSpartanToken()
        {
            if (!File.Exists(StatsTokenFile)) return;
            try
            {
                string t = File.ReadAllText(StatsTokenFile).Trim();
                if (!string.IsNullOrEmpty(t))
                {
                    _statsSpartanToken = t;
                }
            }
            catch { }
        }

        private void StatsSaveToken(string token)
        {
            try
            {
                Directory.CreateDirectory(App.ToolboxDataRoot);
                File.WriteAllText(StatsTokenFile, token);
            }
            catch { }
        }

        private void StatsLoadPersistentCache()
        {
            if (!File.Exists(StatsCacheFile)) return;
            try
            {
                var loaded = JsonSerializer.Deserialize<Dictionary<string, StatsCachedPlayer>>(
                    File.ReadAllText(StatsCacheFile));
                if (loaded == null) return;
                foreach (var (k, v) in loaded) { _statsPersistentCache[k] = v; _statsCacheOrder.Enqueue(k); }
            }
            catch { }
        }

        private void StatsAddToCache(string gt, string kd, string totals)
        {
            lock (_statsLock)
            {
                if (gt.Equals(_statsGamertag, StringComparison.OrdinalIgnoreCase)) return;
                if (!_statsPersistentCache.ContainsKey(gt))
                {
                    if (_statsCacheOrder.Count >= 1000) _statsPersistentCache.Remove(_statsCacheOrder.Dequeue());
                    _statsCacheOrder.Enqueue(gt);
                }
                _statsPersistentCache[gt] = new StatsCachedPlayer { KD = kd, Totals = totals, Added = DateTime.Now };
                try
                {
                    Directory.CreateDirectory(App.ToolboxDataRoot);
                    File.WriteAllText(StatsCacheFile, JsonSerializer.Serialize(_statsPersistentCache));
                }
                catch { }
            }
        }
    }

    // ------------------------------------------
    // Data model
    // ------------------------------------------
    public class MapEntry : INotifyPropertyChanged
    {
        private bool _isEnabled = true;

        public string FileName    { get; set; } = "";
        public string BaseName    { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public bool   IsModded    { get; set; } = false;

        /// <summary>True for the "-- MODDED MAPS --" section divider row.</summary>
        public bool IsHeader { get; set; } = false;

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled == value) return;
                _isEnabled = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ------------------------------------------
    // Report Tab -- player scoreboard entry
    // ------------------------------------------
    public class PlayerEntry : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string Gamertag   { get; set; } = "";
        public string XboxUserId { get; set; } = ""; // mXboxUserId -- key for reporting
        public int    Score      { get; set; }
        public int    Kills      { get; set; }
        public int    Deaths     { get; set; }
        public int    Assists    { get; set; }
        public int    Betrayals  { get; set; }
        public int    Suicides   { get; set; }
        public string Team       { get; set; } = "";
        public bool   Completed  { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ------------------------------------------
    public sealed class MatchmakingPopulationRow
    {
        public MatchmakingPopulationRow(MatchmakingHopperDefinition hopper, HopperPopulationResult result)
        {
            HopperName = hopper.HopperName;
            DisplayName = hopper.DisplayName;
            Mode = hopper.Mode;
            Size = hopper.Size;
            Population = result.Population;
            WaitSeconds = result.WaitSeconds;
            Error = result.Error;
        }

        public string HopperName { get; }
        public string DisplayName { get; }
        public string Mode { get; }
        public string Size { get; }
        public int? Population { get; }
        public int? WaitSeconds { get; }
        public string Error { get; }
        public string PopulationDisplay => Population?.ToString() ?? "—";
        public string WaitDisplay => WaitSeconds switch
        {
            null => "—",
            < 60 => $"~{WaitSeconds} sec",
            _ => $"~{Math.Ceiling(WaitSeconds.Value / 60.0):0} min"
        };
        public string Activity => Population switch
        {
            null => string.IsNullOrWhiteSpace(Error) ? "UNKNOWN" : "UNAVAILABLE",
            0 => "QUIET",
            < 5 => "LOW",
            < 20 => "MEDIUM",
            _ => "HIGH"
        };
    }


    // Stats Tab — Player row (lobby ListView)
    // ------------------------------------------
    public class StatsPlayerRow : INotifyPropertyChanged
    {
        private string _kd = "—";
        private string _totals = "";
        private string _gamesPlayed = "";
        private string _recentKD = "";
        private string _recentTrend = "";
        private string _skillPercentile = "—";

        public string Gamertag  { get; set; } = "";
        public string Xuid      { get; set; } = "";
        public string Team      { get; set; } = "0";
        public string TeamLabel => Team switch
        {
            "" or "—" => "—",
            "FFA" => "FFA",
            "0" => "RED",
            "1" => "BLUE",
            "2" => "GREEN",
            "3" => "YELLOW",
            _ => $"T{Team}"
        };
        public string BestServer { get; set; } = "—";
        public string Ping       { get; set; } = "—";
        public string SquadId    { get; set; } = "";
        public string SquadLabel { get; set; } = "";
        public string SquadToolTip
        {
            get
            {
                if (SquadLabel == "?")
                    return "Squad unknown: MCC did not expose a party id for this player";

                if (string.IsNullOrWhiteSpace(SquadLabel))
                    return "";

                if (SquadId.StartsWith("skill:", StringComparison.OrdinalIgnoreCase))
                    return $"Squad {SquadLabel}: inferred from group skill";

                return $"Squad {SquadLabel}: {SquadId}";
            }
        }
        public bool   IsMe      { get; set; }
        public bool   IsScanning { get; set; }
        public int    Standing  { get; set; }
        public string MatchResult { get; set; } = "—";
        public int MatchScore { get; set; }
        public int MatchKills { get; set; }
        public int MatchDeaths { get; set; }
        public int MatchAssists { get; set; }
        public string MatchObjectiveStats { get; set; } = "—";
        public Brush MatchResultColor => MatchResult switch
        {
            "WIN" => new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14)),
            "LOSS" => new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55)),
            _ => new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF))
        };

        public string KD
        {
            get => _kd;
            set { _kd = value; OnPropertyChanged(nameof(KD)); OnPropertyChanged(nameof(KdColor)); }
        }

        public string Totals
        {
            get => _totals;
            set { _totals = value; OnPropertyChanged(nameof(Totals)); }
        }

        public string GamesPlayed
        {
            get => _gamesPlayed;
            set { _gamesPlayed = value; OnPropertyChanged(nameof(GamesPlayed)); }
        }

        public string RecentKD
        {
            get => _recentKD;
            set { _recentKD = value; OnPropertyChanged(nameof(RecentKD)); OnPropertyChanged(nameof(RecentKdColor)); }
        }

        public string RecentTrend
        {
            get => _recentTrend;
            set { _recentTrend = value; OnPropertyChanged(nameof(RecentTrend)); OnPropertyChanged(nameof(TrendColor)); }
        }

        public string SkillPercentile
        {
            get => _skillPercentile;
            set { _skillPercentile = value; OnPropertyChanged(nameof(SkillPercentile)); }
        }

        public Brush KdColor
        {
            get
            {
                if (!double.TryParse(_kd, out double v)) return new SolidColorBrush(Color.FromRgb(0x4A, 0x5A, 0x6A));
                if (v >= 2.0) return new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14));
                if (v >= 1.0) return new SolidColorBrush(Color.FromRgb(0xC8, 0xD8, 0xE8));
                return new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));
            }
        }

        public Brush RecentKdColor
        {
            get
            {
                if (!double.TryParse(_recentKD, out double v)) return new SolidColorBrush(Color.FromRgb(0x4A, 0x5A, 0x6A));
                if (v >= 2.0) return new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14));
                if (v >= 1.0) return new SolidColorBrush(Color.FromRgb(0xC8, 0xD8, 0xE8));
                return new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));
            }
        }

        public Brush TrendColor => _recentTrend switch
        {
            "▲" => new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14)),
            "▼" => new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55)),
            _   => new SolidColorBrush(Color.FromRgb(0x4A, 0x5A, 0x6A)),
        };

        public Uri WortUrl =>
            new($"https://wort.gg/profile/{Uri.EscapeDataString(Gamertag)}/multiplayer/all");

        public Brush GamertagColor => IsMe
            ? new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF))
            : new SolidColorBrush(Color.FromRgb(0xC8, 0xD8, 0xE8));

        public FontWeight GamertagWeight => IsMe ? FontWeights.Bold : FontWeights.Normal;

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ------------------------------------------
    // Stats Tab — Session stats tracker
    // ------------------------------------------
    class StatsSessionStats
    {
        public int  Wins        { get; set; }
        public int  Losses      { get; set; }
        public int  GamesPlayed { get; set; }
        public long Kills       { get; set; }
        public long Deaths      { get; set; }
        public int BestSpree { get; set; }
        public double BestGameKd { get; set; }
        public string BestGameScore { get; set; } = "";
        public int CurrentWinStreak { get; set; }
        public int LongestWinStreak { get; set; }
        public Dictionary<string, int> MultikillCounts { get; } = new(
            MainWindow.StatsMultikillMedals.ToDictionary(m => m.Name, _ => 0),
            StringComparer.OrdinalIgnoreCase);
        public HashSet<string> ProcessedGameIds { get; } = new(StringComparer.OrdinalIgnoreCase);

        public void Reset()
        {
            Wins = 0; Losses = 0; GamesPlayed = 0; Kills = 0; Deaths = 0;
            BestSpree = 0; BestGameKd = 0; BestGameScore = "";
            CurrentWinStreak = 0; LongestWinStreak = 0;
            foreach (var key in MultikillCounts.Keys.ToList()) MultikillCounts[key] = 0;
            ProcessedGameIds.Clear();
        }
    }

    internal sealed record StatsMedalDefinition(string Name, int CarnageId, string ResourcePath);

    public sealed class FirewallRuleRow
    {
        public string Mode { get; init; } = "";
        public string Port { get; init; } = "";
        public string Protocol { get; init; } = "";
        public string Direction { get; init; } = "";
        public string State { get; init; } = "";
        public string Action { get; init; } = "";
        public string Definition { get; init; } = "";
        public Brush StateBrush { get; init; } = Brushes.Gray;
    }

    class StatsSessionGameRow
    {
        public int Game { get; init; }
        public string Result { get; init; } = "";
        public string KillsDeaths { get; init; } = "";
        public string KD { get; init; } = "";
        public int BestSpree { get; init; }
        public string HighestMultikill { get; init; } = "";
        public string HighestMultikillIcon { get; init; } = "";
        public string PlayedAt { get; init; } = "";
        public int LobbyPlayerCount { get; init; }
        public int LobbyColumnCount { get; init; } = 1;
        public List<StatsSessionLobbyTeamRow> LobbyTeams { get; init; } = new();
        public bool IsExpanded { get; set; }
        public Brush ResultColor => Result == "WIN"
            ? new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14))
            : new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));
    }

    class StatsSessionLobbyTeamRow
    {
        public string TeamName { get; init; } = "";
        public Brush TeamColor { get; init; } = Brushes.Transparent;
        public List<StatsSessionLobbyPlayerRow> Players { get; init; } = new();
        public string PlayerCountText => $"{Players.Count} PLAYERS";
    }

    class StatsSessionLobbyPlayerRow
    {
        public string Gamertag { get; init; } = "";
        public string TeamKey { get; init; } = "";
        public bool IsMe { get; init; }
        public int Kills { get; init; }
        public int Deaths { get; init; }
        public int Assists { get; init; }
        public string EncounterLabel { get; init; } = "NEW";
        public Brush PlayerColor => IsMe
            ? new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF))
            : new SolidColorBrush(Color.FromRgb(0xD8, 0xE6, 0xEC));
        public Brush EncounterColor => EncounterLabel switch
        {
            "YOU" => new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF)),
            "NEW" => new SolidColorBrush(Color.FromRgb(0x5D, 0x76, 0x84)),
            _ => new SolidColorBrush(Color.FromRgb(0xD6, 0xB7, 0x4A))
        };
    }

    class StatsSessionPlayerHistoryRow
    {
        public string Gamertag { get; init; } = "";
        public string LastSeen { get; init; } = "";
        public string Encounter { get; init; } = "";
        public int Matches { get; init; }
        public string Record { get; init; } = "";
        public string AverageKD { get; init; } = "";
        public Brush EncounterColor => Encounter switch
        {
            "TEAMMATE" => new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF)),
            "OPPONENT" => new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55)),
            _ => new SolidColorBrush(Color.FromRgb(0xD6, 0xB7, 0x4A))
        };
        public Brush KdColor => double.TryParse(
                AverageKD,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out double kd) && kd >= 1
            ? new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14))
            : new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));
    }

    class StatsSessionPlayerAggregate
    {
        public string Gamertag { get; set; } = "";
        public string Xuid { get; init; } = "";
        public int LastGame { get; set; }
        public DateTime LastSeen { get; set; }
        public int Matches { get; set; }
        public int Wins { get; set; }
        public int Losses { get; set; }
        public long Kills { get; set; }
        public long Deaths { get; set; }
        public bool WasTeammate { get; set; }
        public bool WasOpponent { get; set; }
    }

    // ------------------------------------------
    // Stats Tab — Persistent player cache entry
    // ------------------------------------------
    class StatsCachedPlayer
    {
        public string   KD      { get; set; } = "";
        public string   Totals  { get; set; } = "";
        public DateTime Added   { get; set; }
    }

    // ------------------------------------------
    // Basic/Advanced mode toggle converter
    // ------------------------------------------
    public class IntToVisibilityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            // 0 = Basic (hide Advanced sections), 1 = Advanced (show all)
            return (int)value == 0 ? Visibility.Collapsed : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class BoolToVisibilityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            // false = Basic (hide Advanced), true = Advanced (show)
            return (bool)value ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
