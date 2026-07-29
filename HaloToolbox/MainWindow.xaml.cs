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
        private const string StatsSettingsFile = "stats_gamertag.txt";
        private const string StatsCacheFile    = "stats_cache.json";
        private const string StatsTokenFile    = "stats_token.txt";

        // ── Stats Tab — mutable state (always access under _statsLock) ───────
        private readonly object _statsLock = new();
        private string _statsGamertag = "";
        private StatsSessionStats _statsSession = new();
        private readonly ObservableCollection<StatsSessionGameRow> _statsSessionGames = new();
        private List<XElement> _statsLastPlayers = new();
        private string _statsLastFileSig = "";
        private bool _statsAutoPullLobby = true;
        private string _statsSpartanToken = "";
        private bool _statsHwTokenExpired = false;
        private bool _statsCurrentLobbyScanRunning = false;
        private string _statsCurrentLobbyServerText = "";
        private string _statsLastGameServerText = "";
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
        private readonly SemaphoreSlim _statsPopulationRefreshLock = new(1, 1);
        private string _statsPopulationSortProperty = nameof(MatchmakingPopulationRow.Population);
        private ListSortDirection _statsPopulationSortDirection = ListSortDirection.Descending;
        private readonly List<string> _sessionLogLines = new();
        private readonly ProxyService _rejoinProxy = new();
        private readonly NetworkStatsMonitor _networkStatsMonitor = new();
        private readonly GameServerConnectionMonitor _gameServerConnectionMonitor = new();
        private readonly ObsOverlayServer _obsOverlayServer = new();
        private GameNetworkStatsOverlayWindow? _gameNetworkStatsOverlay;
        private GameNetworkStatsOverlayWindow? _matchmakingWaitOverlay;
        private GameNetworkStatsOverlayWindow? _sessionStatsOverlay;
        private bool _networkStatsOverlayEnabled = true;
        private bool _matchmakingWaitOverlayEnabled = true;
        private SmartMatchWaitEstimate? _smartMatchWaitEstimate;
        private int? _smartMatchHopperPopulation;
        private string _smartMatchHopperDisplayName = "";
        private readonly System.Windows.Threading.DispatcherTimer _matchmakingPopulationTimer;
        private readonly System.Windows.Threading.DispatcherTimer _populationHistoryTimer;
        private DateTimeOffset _lastFullPopulationRefreshUtc = DateTimeOffset.MinValue;
        private bool _networkStatsOverlayMoveEnabled;
        private bool _obsBrowserOverlayEnabled;
        private bool _obsBrowserOverlaySessionStatsEnabled = true;
        private bool _networkStatsObsOnly;
        private bool _matchmakingWaitObsOnly;
        private bool _sessionStatsObsOnly;
        private NetworkStatsSnapshot? _lastNetworkStatsSnapshot;
        private NetworkTrafficSnapshot? _lastNetworkTrafficSnapshot;
        private ObsPostGameRecap? _postGameRecap;
        private Task? _supportSessionCheckTask;
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
        private bool _closeFirewallCleanupStarted;
        private TabItem? _lastMainTab;
        private bool _restoringMainTabSelection;
        private static readonly SemaphoreSlim SteamFirewallCommandLock = new(1, 1);
        private DateTime _steamFirewallAutoResumeAfterUtc = DateTime.MinValue;
        private const int SteamFirewallAutoSearchHoldSeconds = 180;
        private const int SteamFirewallAutoMatchFoundHoldSeconds = 5;
        private static readonly bool SteamFirewallFeatureEnabled = false;
        private const string RejoinFirewallCampaignLabel = "Firewall Fix (Campaign)";
        private const string RejoinFirewallMatchmakingLabel = "Firewall Fix (Matchmaking)";
        private const string RejoinFirewallDisabledSuffix = " (Disabled until Rejoin Fix is Enabled)";

        private const long MaxDiagnosticExportBytes = 25L * 1024 * 1024;
        private const string ToolboxRegistryPath = @"Software\HaloMCCToolbox";
        private const string RejoinFixProxyAddress = "127.0.0.1:19999";
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
            _populationHistoryTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(1)
            };
            _populationHistoryTimer.Tick += PopulationHistoryTimer_Tick;
            _populationHistoryTimer.Start();
            RestoreMainWindowPlacement();
            LoadSectionVisibility();
            _lastMainTab = MainTabs.SelectedItem as TabItem;
            MainTabs.SelectionChanged += MainTabs_SelectionChanged;
            _networkStatsOverlayEnabled = App.LoadGameNetworkStatsOverlayEnabled();
            ChkNetworkStatsOverlay.IsChecked = _networkStatsOverlayEnabled;
            _matchmakingWaitOverlayEnabled = App.LoadMatchmakingWaitOverlayEnabled();
            ChkMatchmakingWaitOverlay.IsChecked = _matchmakingWaitOverlayEnabled;
            string savedRejoinFirewallMode = App.LoadRejoinFirewallMode();
            SetRejoinFirewallCheckbox(
                ChkRejoinFixFirewall,
                string.Equals(savedRejoinFirewallMode, "Campaign", StringComparison.OrdinalIgnoreCase));
            SetRejoinFirewallCheckbox(
                ChkRejoinFixFirewallMatchmaking,
                string.Equals(savedRejoinFirewallMode, "Matchmaking", StringComparison.OrdinalIgnoreCase));
            TxtMccPath.Text = App.LoadMccInstallationPath();
            PlaylistsTab.SetMccInstallationPath(TxtMccPath.Text);
            TxtMccPath.TextChanged += TxtMccPath_TextChanged;
            MapList.ItemsSource = _maps;
            AppendLog("[INFO]", "Halo MCC Toolbox started. Made by The FFA Panda.", "#00C8FF");

            // Load maps asynchronously in background after window renders
            Loaded += async (_, _) =>
            {
                ThemeToggleBtn.Content = App.IsDarkTheme ? "☾" : "☀";

                // Load maps asynchronously so UI isn't blocked
                string mccPath = TxtMccPath.Text.Trim();
                var defaultMapsPath = Path.Combine(mccPath, "halo3", "maps");
                if (Directory.Exists(defaultMapsPath))
                {
                    AppendLog("[INFO]", "Loading maps in background...", "#4A5A6A");
                    await Task.Run(() => LoadMaps(mccPath));
                }

                // Start stats monitoring loop after UI is fully initialized
                _ = Task.Run(StatsMonitorLoop);
            };

            // Initialize the Stats tab (lobby monitor)
            StatsInitialize();
            StatsPopulationList.ItemsSource = _statsPopulationRows;
            _ = StatsRefreshMatchmakingPopulationAsync();
            _statsAutoPullLobby = App.LoadStatsAutoLobbyEnabled();
            StatsAutoToggle.IsChecked = _statsAutoPullLobby;
            StatsAutoToggle.Content = _statsAutoPullLobby ? "AUTO: ON" : "AUTO: OFF";
            _obsBrowserOverlayEnabled = App.LoadObsBrowserOverlayEnabled();
            _obsBrowserOverlaySessionStatsEnabled = App.LoadObsBrowserOverlaySessionStatsEnabled();
            _networkStatsObsOnly = App.LoadNetworkStatsObsOnlyEnabled();
            _matchmakingWaitObsOnly = App.LoadMatchmakingWaitObsOnlyEnabled();
            _sessionStatsObsOnly = App.LoadSessionStatsObsOnlyEnabled();
            StatsObsOverlayToggle.IsChecked = _obsBrowserOverlayEnabled;
            StatsObsSessionStatsToggle.IsChecked = _obsBrowserOverlaySessionStatsEnabled;
            NetworkStatsObsOnlyToggle.IsChecked = _networkStatsObsOnly;
            MatchmakingWaitObsOnlyToggle.IsChecked = _matchmakingWaitObsOnly;
            SessionStatsObsOnlyToggle.IsChecked = _sessionStatsObsOnly;
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
                    if (_rejoinProxy.IsRunning && pings.Count > 0)
                        _ = StatsFetchCurrentLobbyStats();
                });
            };
            _rejoinProxy.OnSmartMatchWaitEstimateChanged += (_, estimate) =>
                Dispatcher.InvokeAsync(() =>
                {
                    _smartMatchWaitEstimate = estimate;
                    _smartMatchHopperPopulation = null;
                    _smartMatchHopperDisplayName = PlaylistsTab.GetMatchmakingHoppers()
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
            _networkStatsMonitor.StatsUpdated += (_, snapshot) =>
                Dispatcher.InvokeAsync(() => UpdateNetworkStatsOverlay(snapshot));
            _gameServerConnectionMonitor.ActiveServerChanged += (_, serverInfo) =>
                Dispatcher.InvokeAsync(() => HandleNetworkStatsObservedServer(serverInfo));
            _gameServerConnectionMonitor.TrafficStatsUpdated += (_, snapshot) =>
                Dispatcher.InvokeAsync(() => UpdateNetworkTrafficOverlay(snapshot));
            _gameServerConnectionMonitor.StatusChanged += (_, status) =>
                Dispatcher.InvokeAsync(() => AppendLog("[NET]", status, "#4A5A6A"));
            Closed += (_, _) =>
            {
                ModsTab.Dispose();
                StopRejoinCrashWatcher();
                _gameServerConnectionMonitor.Dispose();
                _networkStatsMonitor.Dispose();
                _obsOverlayServer.Dispose();
                _rejoinProxy.Dispose();
                _matchmakingPopulationTimer.Stop();
                _populationHistoryTimer.Stop();
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
            Dispatcher.InvokeAsync(SynchronizeStartupFirewallStateAsync);
            Dispatcher.InvokeAsync(StartPendingRejoinFixAfterElevationAsync);

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

            bool shouldCleanFirewall = _rejoinProxy.IsRunning
                || _steamFirewallAutoEnabled
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
            if (ReferenceEquals(selectedTab, ReportSection))
                _ = EnsureSupportSessionCheckedAsync();

            if (!ReferenceEquals(selectedTab, H3ModsSection) || IsRunningAsAdministrator())
            {
                _lastMainTab = selectedTab;
                return;
            }

            // Do not leave the admin-only control active while the UAC decision is pending.
            _restoringMainTabSelection = true;
            MainTabs.SelectedItem = _lastMainTab ?? ToolsSection;
            _restoringMainTabSelection = false;

            var result = MessageBox.Show(
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
                MessageBox.Show(
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

        private void LoadSectionVisibility()
        {
            SetSectionVisibility(H3ModsSection, ShowH3ModsSection, "H3Mods", App.LoadMainSectionVisible("H3Mods"));
            SetSectionVisibility(ReportSection, ShowReportSection, "Report", App.LoadMainSectionVisible("Report"));
            SetSectionVisibility(StatsSection, ShowStatsSection, "Stats", App.LoadMainSectionVisible("Stats"));
            SetSectionVisibility(TheaterSection, ShowTheaterSection, "Theater", App.LoadMainSectionVisible("Theater"));
            SetSectionVisibility(PlaylistsSection, ShowPlaylistsSection, "Playlists", App.LoadMainSectionVisible("Playlists"));
            SetSectionVisibility(AboutSection, ShowAboutSection, "About", App.LoadMainSectionVisible("About"));
            SetSectionVisibility(LogSection, ShowLogSection, "Log", App.LoadMainSectionVisible("Log"));
            UpdateToggleAllSectionsButton();
        }

        private void SectionVisibility_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not CheckBox checkBox || checkBox.Tag is not string sectionName)
                return;

            bool visible = checkBox.IsChecked == true;
            TabItem? section = sectionName switch
            {
                "H3Mods" => H3ModsSection,
                "Report" => ReportSection,
                "Stats" => StatsSection,
                "Theater" => TheaterSection,
                "Playlists" => PlaylistsSection,
                "About" => AboutSection,
                "Log" => LogSection,
                _ => null
            };

            if (section is null)
                return;

            section.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            if (!visible && section.IsSelected)
                ToolsSection.IsSelected = true;
            App.SaveMainSectionVisible(sectionName, visible);
            UpdateToggleAllSectionsButton();
        }

        private void ToggleAllSections_Click(object sender, RoutedEventArgs e)
        {
            bool allVisible = AreAllOptionalSectionsVisible();
            bool makeVisible = !allVisible;

            SetSectionVisibility(H3ModsSection, ShowH3ModsSection, "H3Mods", makeVisible, save: true);
            SetSectionVisibility(ReportSection, ShowReportSection, "Report", makeVisible, save: true);
            SetSectionVisibility(StatsSection, ShowStatsSection, "Stats", makeVisible, save: true);
            SetSectionVisibility(TheaterSection, ShowTheaterSection, "Theater", makeVisible, save: true);
            SetSectionVisibility(PlaylistsSection, ShowPlaylistsSection, "Playlists", makeVisible, save: true);
            SetSectionVisibility(AboutSection, ShowAboutSection, "About", makeVisible, save: true);
            SetSectionVisibility(LogSection, ShowLogSection, "Log", makeVisible, save: true);

            if (!makeVisible)
                ToolsSection.IsSelected = true;
            UpdateToggleAllSectionsButton();
        }

        private bool AreAllOptionalSectionsVisible() =>
            ShowH3ModsSection.IsChecked == true &&
            ShowReportSection.IsChecked == true &&
            ShowStatsSection.IsChecked == true &&
            ShowTheaterSection.IsChecked == true &&
            ShowPlaylistsSection.IsChecked == true &&
            ShowAboutSection.IsChecked == true &&
            ShowLogSection.IsChecked == true;

        private void UpdateToggleAllSectionsButton()
        {
            ToggleAllSectionsBtn.Content = AreAllOptionalSectionsVisible() ? "HIDE ALL" : "SHOW ALL";
        }

        private static void SetSectionVisibility(TabItem section, CheckBox checkBox, string sectionName, bool visible, bool save = false)
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
            Dispatcher.Invoke(() =>
            {
                var ts = DateTime.Now.ToString("HH:mm:ss");
                var line = $"[{ts}] {tag} {message}";
                _sessionLogLines.Add(line);
                TxtLog.AppendText(line + Environment.NewLine);
                TxtLog.ScrollToEnd();
            });
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
                !_obsBrowserOverlaySessionStatsEnabled && !_obsBrowserOverlayEnabled)
            {
                _networkStatsMonitor.Stop();
                CloseGameNetworkStatsOverlay();
                ClearNetworkOverlaySnapshots();
                PublishObsOverlaySnapshot();
                return;
            }

            if (_networkStatsOverlayEnabled || _matchmakingWaitOverlayEnabled || _obsBrowserOverlaySessionStatsEnabled)
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

            if (_networkStatsOverlayEnabled)
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
            StatsLastGameServerLabel.Text = label;
            PublishObsOverlaySnapshot();
        }

        private void ClearNetworkStatsOverlayDisplay()
        {
            if ((!_networkStatsOverlayEnabled && !_matchmakingWaitOverlayEnabled && !_obsBrowserOverlayEnabled) || !_rejoinProxy.IsRunning)
                return;

            if (_networkStatsOverlayEnabled || _matchmakingWaitOverlayEnabled)
            {
                EnsureGameNetworkStatsOverlay();
                _gameNetworkStatsOverlay?.SetPreferredProcessId(TryGetMccProcessId());
                _gameNetworkStatsOverlay?.SetMoveMode(_networkStatsOverlayMoveEnabled);
            }
            _networkStatsMonitor.Stop();
            ClearNetworkOverlaySnapshots();
            if (_networkStatsOverlayEnabled)
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
            if (_networkStatsOverlayEnabled)
            {
                EnsureGameNetworkStatsOverlay();
                _gameNetworkStatsOverlay?.UpdateStats(snapshot);
            }
            PublishObsOverlaySnapshot();
        }

        private void UpdateNetworkTrafficOverlay(NetworkTrafficSnapshot snapshot)
        {
            _lastNetworkTrafficSnapshot = snapshot;
            if ((!_networkStatsOverlayEnabled && !_obsBrowserOverlayEnabled) || !_rejoinProxy.IsRunning)
                return;

            if (_networkStatsOverlayEnabled)
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
            EnsureOverlaySourceServer(logStatus: false);

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
            var overlay = new GameNetworkStatsOverlayWindow(component) { Owner = this };
            overlay.SetOverlaySource(_obsOverlayServer.ComponentUrl(component));
            overlay.RelativePlacementChanged += (_, placement) =>
                ComponentOverlay_RelativePlacementChanged(component, placement);
            overlay.Closed += (_, _) =>
            {
                if (ReferenceEquals(_gameNetworkStatsOverlay, overlay)) _gameNetworkStatsOverlay = null;
                if (ReferenceEquals(_matchmakingWaitOverlay, overlay)) _matchmakingWaitOverlay = null;
                if (ReferenceEquals(_sessionStatsOverlay, overlay)) _sessionStatsOverlay = null;
            };
            overlay.Show();
            return overlay;
        }

        private IEnumerable<GameNetworkStatsOverlayWindow> AllGameOverlays()
        {
            if (_gameNetworkStatsOverlay is not null) yield return _gameNetworkStatsOverlay;
            if (_matchmakingWaitOverlay is not null) yield return _matchmakingWaitOverlay;
            if (_sessionStatsOverlay is not null) yield return _sessionStatsOverlay;
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
                return Process.GetProcessesByName("MCC-Win64-Shipping")
                    .Concat(Process.GetProcessesByName("MCC"))
                    .OrderByDescending(p => p.StartTime)
                    .Select(p => (int?)p.Id)
                    .FirstOrDefault();
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
            _gameServerConnectionMonitor.Stop();
            _networkStatsMonitor.Stop();
            CloseComponentOverlay(ref _gameNetworkStatsOverlay);
            PublishObsOverlaySnapshot();
            if (!_matchmakingWaitOverlayEnabled && !_obsBrowserOverlaySessionStatsEnabled && !_obsBrowserOverlayEnabled)
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
            if (!_networkStatsOverlayEnabled && !_obsBrowserOverlaySessionStatsEnabled && !_obsBrowserOverlayEnabled)
                _obsOverlayServer.Stop();
            AppendLog("[MATCH]", "Matchmaking wait estimate overlay disabled.", "#C8D8E8");
            UpdateRejoinFixUi();
        }

        private void ChkNetworkStatsOverlayMove_Checked(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayMoveEnabled = true;
            if (!_mainWindowInitialized)
                return;

            if (_networkStatsOverlayEnabled && _rejoinProxy.IsRunning)
                StartNetworkStatsOverlay(GetNetworkStatsTargetIp(), GetNetworkStatsTargetServerInfo());

            foreach (var overlay in AllGameOverlays()) overlay.SetMoveMode(true);
            AppendLog("[NET]", "Overlay drag mode enabled. Drag the overlay, then turn drag mode off.", "#00C8FF");
        }

        private void ChkNetworkStatsOverlayMove_Unchecked(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayMoveEnabled = false;
            if (!_mainWindowInitialized)
                return;

            foreach (var overlay in AllGameOverlays()) overlay.SetMoveMode(false);
            AppendLog("[NET]", "Overlay drag mode disabled; overlay is click-through.", "#C8D8E8");
        }

        private void BtnNetworkStatsOverlayMove_Click(object sender, RoutedEventArgs e)
        {
            _networkStatsOverlayMoveEnabled = !_networkStatsOverlayMoveEnabled;
            foreach (var overlay in AllGameOverlays()) overlay.SetMoveMode(_networkStatsOverlayMoveEnabled);
            BtnNetworkStatsOverlayMove.Content = _networkStatsOverlayMoveEnabled
                ? "FINISH"
                : "REPOSITION";
            AppendLog(
                "[NET]",
                _networkStatsOverlayMoveEnabled
                    ? "Overlay drag mode enabled. Drag the overlay, then finish repositioning."
                    : "Overlay drag mode disabled; overlay is click-through.",
                _networkStatsOverlayMoveEnabled ? "#00C8FF" : "#C8D8E8");
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            TxtLog.Clear();
            _sessionLogLines.Clear();
            AppendLog("[INFO]", "Log cleared.", "#4A5A6A");
        }

        private void BtnCopyLog_Click(object sender, RoutedEventArgs e)
        {
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
            BtnFirewallCheck.IsEnabled = false;
            AppendLog("[FIREWALL]", "Running Firewall Check for all Toolbox firewall rules...", "#00C8FF");
            SetStatus("Checking Toolbox firewall rules...", "#00C8FF");

            try
            {
                string expectedProgram = ResolveMccExecutablePath(TxtMccPath.Text.Trim());

                foreach (var group in GetFirewallCheckRuleGroups())
                {
                    AppendLog("[FIREWALL]", group.Header, "#C8D8E8");
                    foreach (string ruleName in group.RuleNames)
                    {
                        var result = await RunNetshAsync("advfirewall", "firewall", "show", "rule", $"name={ruleName}", "verbose");
                        if (result.ExitCode != 0 || !NetshRuleExists(result.Output))
                        {
                            LogFirewallRuleMissing(ruleName);
                            continue;
                        }

                        LogFirewallRuleStatus(ruleName, result.Output, logProgram: false);
                        LogFirewallRuleDefinitionProblems(ruleName, result.Output, expectedProgram);
                    }
                }

                AppendLog("[FIREWALL]", $"Expected Program: {expectedProgram}", "#4A5A6A");

                SetStatus("Firewall Check complete.", "#39FF14");
                AppendLog("[FIREWALL]", "Firewall Check complete.", "#39FF14");
            }
            catch (Exception ex)
            {
                SetStatus("Firewall Check failed.", "#FF2D55");
                AppendLog("[ERROR]", $"Firewall Check failed: {ex.Message}", "#FF2D55");
            }
            finally
            {
                BtnFirewallCheck.IsEnabled = true;
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
                        var open = MessageBox.Show(
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
                        MessageBox.Show($"Failed to export logs:\n\n{ex.Message}",
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
            var confirm = MessageBox.Show(
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

                MessageBox.Show(
                    "Halo MCC Toolbox traces were removed.\n\nRestart the app if you want to keep using it with fresh settings.",
                    "Toolbox Traces Removed -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Toolbox trace cleanup failed: {ex.Message}", "#FF2D55");
                SetStatus("Toolbox trace cleanup failed.", "#FF2D55");
                MessageBox.Show(
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
if ($winHttp -match '127\.0\.0\.1:19999') {{
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
            try
            {
                HiddenCookieChecker.Dispose();
            }
            catch (Exception ex)
            {
                AppendLog("[WARN]", $"Could not release WebView2 before cleanup: {ex.Message}", "#FF6A00");
            }
        }

        private bool ClearToolboxWinInetProxyIfPresent()
        {
            const string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(keyPath, writable: true);
                if (key is null)
                    return false;

                var proxyServer = key.GetValue("ProxyServer") as string;
                if (!string.Equals(proxyServer, RejoinFixProxyAddress, StringComparison.OrdinalIgnoreCase))
                    return false;

                key.SetValue("ProxyEnable", 0, RegistryValueKind.DWord);
                key.DeleteValue("ProxyServer", throwOnMissingValue: false);

                var proxyOverride = key.GetValue("ProxyOverride") as string;
                if (string.Equals(proxyOverride, "localhost;127.0.0.1;<local>", StringComparison.OrdinalIgnoreCase))
                    key.DeleteValue("ProxyOverride", throwOnMissingValue: false);

                return true;
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
            var result = MessageBox.Show(
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
                        MessageBox.Show("Cleanup complete!\n\nXBL credentials and webcache have been cleared.\nRestart Halo MCC and sign in again.",
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
                MessageBox.Show(msg, "EAC Not Found -- Halo MCC Toolbox",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show(
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
                MessageBox.Show($"Could not launch EasyAntiCheat EOS setup:\n\n{ex.Message}",
                    "Error -- Halo MCC Toolbox", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ------------------------------------------
        // TOOL: Repair MCC audio device selection
        // ------------------------------------------
        private void BtnRepairAudioDevices_Click(object sender, RoutedEventArgs e)
        {
            var mccProcesses = Process.GetProcessesByName("MCC-Win64-Shipping");
            var isMccRunning = mccProcesses.Length > 0;
            foreach (var process in mccProcesses)
                process.Dispose();

            if (isMccRunning)
            {
                AppendLog("[AUDIO]", "Repair blocked because MCC is running.", "#FF6A00");
                MessageBox.Show(
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
                MessageBox.Show(
                    $"MCC's settings file was not found:\n\n{settingsPath}\n\nLaunch MCC once, close it, and try again.",
                    "Settings Not Found -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show(
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
                    MessageBox.Show(
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
                MessageBox.Show(
                    "Audio device repair complete.\n\nLaunch MCC to test the fix.",
                    "Done -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppendLog("[ERROR]", $"Audio device repair failed: {ex.Message}", "#FF2D55");
                SetStatus("Audio device repair failed.", "#FF2D55");
                MessageBox.Show(
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
            Process[] processes;
            try
            {
                processes = Process.GetProcessesByName("MCC-Win64-Shipping");
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

            BtnRejoinFix.Content = isRunning ? "STOP SERVICES" : "START SERVICES";
            ChkRejoinRecovery.IsChecked = isRunning;
            bool hasPlayerVisibleOverlay =
                (_networkStatsOverlayEnabled && !_networkStatsObsOnly) ||
                (_matchmakingWaitOverlayEnabled && !_matchmakingWaitObsOnly) ||
                (_obsBrowserOverlaySessionStatsEnabled && !_sessionStatsObsOnly);
            BtnNetworkStatsOverlayMove.Visibility =
                hasPlayerVisibleOverlay && isRunning
                ? Visibility.Visible
                : Visibility.Collapsed;

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

            UpdateRejoinFirewallOptionAvailability(isRunning);
            UpdateRejoinFirewallStatus();
            StatsUpdateCurrentLobbyVisibility(isRunning);
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
            ChkRejoinFixFirewall.IsEnabled = true;
            ChkRejoinFixFirewallMatchmaking.IsEnabled = true;
            ChkRejoinFixFirewall.Content = RejoinFirewallCampaignLabel;
            ChkRejoinFixFirewallMatchmaking.Content = RejoinFirewallMatchmakingLabel;
            string? pendingReason = isRejoinFixRunning
                ? null
                : "Selection saved. This starts only when you click Start Services.";
            ChkRejoinFixFirewall.ToolTip = pendingReason;
            ChkRejoinFixFirewallMatchmaking.ToolTip = pendingReason;
        }

        private void UpdateRejoinFirewallStatus(string? overrideText = null, string? overrideColor = null)
        {
            if (!string.IsNullOrWhiteSpace(overrideText))
            {
                TxtRejoinFirewallStatus.Text = overrideText;
                TxtRejoinFirewallStatus.Foreground = Brush(overrideColor ?? "#C8D8E8");
                return;
            }

            if (_steamFirewallAutoSuspendedForCrashRestore)
            {
                TxtRejoinFirewallStatus.Text = "FIREWALL: REJOIN RESTORE - ports are open for crash rejoin";
                TxtRejoinFirewallStatus.Foreground = Brush("#00C8FF");
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
                    return;
                }

                if (_steamFirewallUiState == SteamFirewallState.Partial)
                {
                    TxtRejoinFirewallStatus.Text = "FIREWALL: PARTIAL - rules are mixed; restart Rejoin Fix to repair";
                    TxtRejoinFirewallStatus.Foreground = Brush("#FF6A00");
                    return;
                }

                TxtRejoinFirewallStatus.Text = _steamFirewallRulesPrepared
                    ? "FIREWALL: READY - rules installed, currently off"
                    : "FIREWALL: OFF - rules will be prepared when Rejoin Fix starts as admin";
                TxtRejoinFirewallStatus.Foreground = Brush(_steamFirewallRulesPrepared ? "#C8D8E8" : "#4A5A6A");
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
                SteamFirewallCard.Visibility = Visibility.Visible;
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
                MessageBox.Show(
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

            if (!await _steamFirewallAutoLock.WaitAsync(0))
                return;

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

            return path.Contains("Party/RequestParty", StringComparison.OrdinalIgnoreCase)
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
            if (!_steamFirewallAutoEnabled || !_steamFirewallAutoPaused)
                return;

            _steamFirewallAutoResumeAfterUtc = DateTime.UtcNow.AddSeconds(holdSeconds);
            TxtSteamFirewallStatus.Text = $"AUTO - re-enabling soon ({reason})";
            TxtSteamFirewallStatus.Foreground = Brush("#00C8FF");
            UpdateRejoinFirewallStatus($"FIREWALL: MATCHMAKING PENDING - re-enabling soon ({reason})", "#00C8FF");
        }

        private void HoldSteamFirewallPausedForActiveMatch(string reason)
        {
            if (!_steamFirewallAutoEnabled || !_steamFirewallAutoPaused)
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
            if (!_steamFirewallAutoEnabled || !_steamFirewallAutoPaused || DateTime.UtcNow < _steamFirewallAutoResumeAfterUtc)
                return;

            await ResumeSteamFirewallAfterMatchmakingAsync();
        }

        private async Task ResumeSteamFirewallAfterMatchmakingAsync()
        {
            if (!await _steamFirewallAutoLock.WaitAsync(0))
                return;

            try
            {
                if (!_steamFirewallAutoEnabled || !_steamFirewallAutoPaused)
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
                MessageBox.Show(
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
                    MessageBox.Show(
                        $"{requestedFeature} needs Advanced Features. The Toolbox will relaunch as Administrator and start them automatically.\n\nIf MCC is currently open, restart MCC afterward so traffic capture can take effect.",
                        "Advanced Features -- Halo MCC Toolbox",
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
                AppendLog("[ERROR]", $"Advanced Features failed: {ex.Message}", "#FF2D55");
                SetStatus("Advanced Features failed to start.", "#FF2D55");
                MessageBox.Show(
                    $"Advanced Features could not start:\n\n{ex.Message}",
                    "Advanced Features -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
        }

        private async Task StartPendingRejoinFixAfterElevationAsync()
        {
            if (!App.ConsumePendingRejoinFixAutoStart())
                return;

            if (!IsRunningAsAdministrator() || _rejoinProxy.IsRunning)
                return;

            BtnRejoinFix.IsEnabled = false;
            try
            {
                await StartRejoinFixAsync();
            }
            catch (Exception ex)
            {
                RejoinFixDiagnostics.Error("proxy", $"Automatic activation after elevation failed: {ex.Message}");
                AppendLog("[ERROR]", $"Rejoin Fix failed: {ex.Message}", "#FF2D55");
                SetStatus("Rejoin Fix failed to start.", "#FF2D55");
                MessageBox.Show(
                    $"Rejoin Fix could not start:\n\n{ex.Message}",
                    "Rejoin Fix -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                UpdateRejoinFixUi();
                BtnRejoinFix.IsEnabled = true;
            }
        }

        private async void BtnRejoinFix_Click(object sender, RoutedEventArgs e)
        {
            BtnRejoinFix.IsEnabled = false;

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
                    AppendLog("[REJOIN]", "Advanced Features stopped.", "#C8D8E8");
                    SetStatus("Rejoin Fix stopped.", "#C8D8E8");
                }
                else
                {
                    if (!IsRunningAsAdministrator())
                    {
                        MessageBox.Show(
                            "Advanced Features need the Toolbox to run as Administrator so live MCC features can use the system proxy and firewall settings.\n\nIf MCC is currently open, restart it afterward so traffic capture can take effect.\n\nThe Toolbox will relaunch as Administrator now.",
                            "Advanced Features -- Halo MCC Toolbox",
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
                MessageBox.Show(
                    $"Rejoin Fix could not start:\n\n{ex.Message}",
                    "Rejoin Fix -- Halo MCC Toolbox",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                UpdateRejoinFixUi();
                BtnRejoinFix.IsEnabled = true;
            }
        }

        private void HandleAdministratorRelaunchCancelled()
        {
            App.SavePendingRejoinFixAutoStart(false);
            AppendLog("[INFO]", "Advanced Features cancelled at administrator prompt.", "#4A5A6A");
            SetStatus("Advanced Features require Administrator.", "#4A5A6A");
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
                TxtMccPath.Text = dlg.FolderName;
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
            PlaylistsTab.SetMccInstallationPath(TxtMccPath.Text);
            App.SaveMccInstallationPath(TxtMccPath.Text);
        }

        private void SaveMccInstallationPath()
        {
            var mccPath = TxtMccPath.Text.Trim();
            App.SaveMccInstallationPath(mccPath);
            PlaylistsTab.SetMccInstallationPath(mccPath);
        }

        private void LoadMaps(string mccPath)
        {
            var mapsPath = Path.Combine(mccPath, "halo3", "maps");

            if (!Directory.Exists(mapsPath))
            {
                Dispatcher.InvokeAsync(() =>
                {
                    TxtMapStatus.Text = $"Maps folder not found: {mapsPath}";
                    TxtMapStatus.Foreground = Brush("#FF2D55");
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
            MessageBox.Show("No maps loaded. Load your maps first.", "Halo MCC Toolbox",
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
        private void BtnLoadCarnage_Click(object sender, RoutedEventArgs e)
        {
            var tempDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "AppData", "LocalLow", "MCC", "Temporary");

            if (!Directory.Exists(tempDir))
            {
                MessageBox.Show($"MCC Temporary folder not found:\n{tempDir}",
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
                MessageBox.Show($"Failed to parse carnage report:\n\n{ex.Message}\n\nPath: {xmlPath}",
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
        /// The hidden WebView2 (HiddenCookieChecker) shares the same CoreWebView2Environment
        /// as HaloReportWindow so it reads from the same on-disk cookie store.
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
                var env = await WebViewEnvironmentManager.GetOrCreateAsync();
                await HiddenCookieChecker.EnsureCoreWebView2Async(env);

                var mgr = HiddenCookieChecker.CoreWebView2.CookieManager;

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
                MessageBox.Show("Please select the cheating player on the scoreboard first.",
                    "No Player Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var mapName = (CboReportMap.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Unknown";
            var cheatType = (CboCheatType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Not specified";
            var notes = TxtReportNotes.Text.Trim();

            if (CboCheatType.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a cheat type.", "Missing Info",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Ask where to save
            var safeTag = string.Concat(suspect.Gamertag.Select(c =>
                Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
            var dlg = new SaveFileDialog
            {
                Title            = "Save Cheat Report ZIP",
                Filter           = "ZIP Archive (*.zip)|*.zip",
                FileName         = $"CheatReport_{safeTag}_{DateTime.Now:yyyyMMdd_HHmmss}.zip",
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
                    sb.AppendLine("  HALO MCC -- CHEATER REPORT");
                    sb.AppendLine("  Generated by Halo MCC Toolbox  /  The FFA Panda");
                    sb.AppendLine($"  {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine("=======================================================");
                    sb.AppendLine();
                    sb.AppendLine("[ REPORTED PLAYER ]");
                    sb.AppendLine($"  Gamertag    : {suspect.Gamertag}");
                    sb.AppendLine($"  Xbox User ID: {suspect.XboxUserId}");
                    sb.AppendLine($"  Cheat Type  : {cheatType}");
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

                        var open = MessageBox.Show(
                            $"Report ZIP created!\n\n" +
                            $"  Suspect   : {suspect.Gamertag}\n" +
                            $"  Game      : {selectedGame}\n" +
                            $"  Map       : {mapName}\n" +
                            $"  Cheat     : {cheatType}\n" +
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
                        MessageBox.Show($"Failed to build report:\n\n{ex.Message}",
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
                MessageBox.Show("Please select the cheating player on the scoreboard first.",
                    "No Player Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (CboCheatType.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a cheat type first.",
                    "Missing Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var mapName   = (CboReportMap.SelectedItem  as ComboBoxItem)?.Content?.ToString() ?? "Unknown";
            var cheatType = (CboCheatType.SelectedItem  as ComboBoxItem)?.Content?.ToString() ?? "";
            var gameType  = TxtGameMap.Text;
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
                CheatType       = cheatType,
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
            StatsShowLobbyView();

            StatsLoadGamertag();
            StatsLoadPersistentCache();
            StatsLoadSpartanToken();

            StatsGamertagBox.Text = _statsGamertag;
            StatsInitializeSignature();
            StatsUpdateHwStatus();

            if (!string.IsNullOrWhiteSpace(_statsGamertag))
            {
                _ = StatsFetchStats(_statsGamertag);
                if (!string.IsNullOrEmpty(_statsSpartanToken))
                    _ = StatsFetchRecentStatsAsync(_statsGamertag, _statsSpartanToken);
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
            Task.Run(() => StatsProcessFile(file.FullName, countTowardSession: false));
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

        private void StatsSyncBtn_Click(object sender, RoutedEventArgs e)
        {
            string gt; lock (_statsLock) { gt = _statsGamertag; }
            if (!string.IsNullOrWhiteSpace(gt)) _ = StatsFetchStats(gt);
        }

        private void StatsResetBtn_Click(object sender, RoutedEventArgs e)
        {
            lock (_statsLock) { _statsSession.Reset(); _postGameRecap = null; }
            Dispatcher.Invoke(() => _statsSessionGames.Clear());
            StatsRefreshSessionUI();
            StatsSetStatus("Session reset.");
        }

        private void StatsLobbyView_Click(object sender, RoutedEventArgs e) => StatsShowLobbyView();

        private void StatsSessionView_Click(object sender, RoutedEventArgs e) => StatsShowSessionView();

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

        private void StatsLastGameBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(StatsWatchPath)) { StatsSetStatus("MCC folder not found."); return; }
            var f = StatsLatestCarnageFile();
            if (f == null) { StatsSetStatus("No carnage report found."); return; }
            StatsSetStatus($"Loading {f.Name}…");
            Task.Run(() => StatsProcessFile(f.FullName));
        }

        private void StatsAutoToggle_Checked(object sender, RoutedEventArgs e)
        {
            lock (_statsLock) { _statsAutoPullLobby = true; }
            StatsAutoToggle.Content = "AUTO: ON";
            App.SaveStatsAutoLobbyEnabled(true);
        }

        private void StatsAutoToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            lock (_statsLock) { _statsAutoPullLobby = false; }
            StatsAutoToggle.Content = "AUTO: OFF";
            App.SaveStatsAutoLobbyEnabled(false);
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
            if (!_networkStatsOverlayEnabled)
                _obsOverlayServer.Stop();
            if (!_networkStatsOverlayEnabled)
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

        private void StatsLobbyList_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            StatsPlayerRow? row = sender switch
            {
                ListView { SelectedItem: StatsPlayerRow selected } => selected,
                _ => null
            };
            if (row is null) return;
            string token; lock (_statsLock) { token = _statsSpartanToken; }
            if (string.IsNullOrEmpty(token))
            {
                StatsSetStatus("Connect Halo Waypoint first to open match history.");
                return;
            }
            var win = new PlayerMatchHistoryWindow(row.Gamertag, token) { Owner = this };
            win.Show();
        }

        private void BtnStatsMyHistory_Click(object sender, RoutedEventArgs e)
        {
            string gt, token;
            lock (_statsLock) { gt = _statsGamertag; token = _statsSpartanToken; }
            if (string.IsNullOrEmpty(gt))
            {
                StatsSetStatus("Apply a gamertag first.");
                return;
            }
            if (string.IsNullOrEmpty(token))
            {
                StatsSetStatus("Connect Halo Waypoint first to open match history.");
                return;
            }
            var win = new PlayerMatchHistoryWindow(gt, token) { Owner = this };
            win.Show();
        }

        private void BtnStatsBanChecker_Click(object sender, RoutedEventArgs e)
        {
            var window = new BanCheckerWindow(StatsCheckBanTargetsAsync) { Owner = this };
            window.ShowDialog();
        }

        private async Task<IReadOnlyList<BanCheckDisplayResult>> StatsCheckBanTargetsAsync(IReadOnlyList<string> targets)
        {
            if (!_rejoinProxy.TryGetLatestBanSpartanToken(out var token, out var capturedAtUtc, out _))
            {
                StatsSetStatus("Start Rejoin Fix, then let MCC make a ban summary request before using Ban Checker.");
                throw new InvalidOperationException("No fresh MCC banprocessor token is available yet.\n\nStart Rejoin Fix, let MCC make a /hmcc/bansummary request, then try Ban Checker again.");
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
                _rejoinProxy.ClearBanSpartanToken(token);
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

        private void StatsHyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
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
                    $"Session: {_obsOverlayServer.ComponentUrl("session", "obs")}");
                Clipboard.SetText(urls);
                StatsSetStatus("Three independent OBS overlay URLs copied.");
            }
            catch (Exception ex)
            {
                StatsSetStatus($"OBS overlay ready, but clipboard copy failed: {ex.Message}");
            }
        }

        private void StatsRefreshObsOverlayUi()
        {
            StatsObsOverlayToggle.Content = "OBS BROWSER OVERLAY";
            StatsObsOverlayUrlLabel.Text = _obsBrowserOverlayEnabled
                ? $"NETWORK  {_obsOverlayServer.ComponentUrl("network", "obs")}\n" +
                  $"WAIT     {_obsOverlayServer.ComponentUrl("wait", "obs")}\n" +
                  $"SESSION  {_obsOverlayServer.ComponentUrl("session", "obs")}"
                : "";
            StatsObsOverlayUrlLabel.Visibility = _obsBrowserOverlayEnabled
                ? Visibility.Visible
                : Visibility.Collapsed;
            // Session-stat visibility applies to every overlay and is intentionally
            // independent of whether the OBS browser-source link is enabled.
            StatsObsSessionStatsToggle.IsEnabled = true;
        }

        private void PublishObsOverlaySnapshot()
        {
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

            return new ObsOverlaySnapshot(
                ShowSessionStats: _obsBrowserOverlaySessionStatsEnabled,
                ShowNetworkStats: _networkStatsOverlayEnabled,
                ShowMatchmakingWait: _matchmakingWaitOverlayEnabled && _smartMatchWaitEstimate is not null,
                MatchmakingWaitSeconds: _smartMatchWaitEstimate?.WaitSeconds,
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

        private void StatsPopulationGraph_Click(object sender, RoutedEventArgs e)
        {
            var graph = new PopulationHistoryWindow(_statsPopulationHistory) { Owner = this };
            graph.ShowDialog();
        }

        private async void PopulationHistoryTimer_Tick(object? sender, EventArgs e)
        {
            // This timer deliberately runs for the lifetime of the app, independently
            // of matchmaking, so an idle session continues building graph history.
            await StatsRefreshMatchmakingPopulationAsync();
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
                var hoppers = PlaylistsTab.GetMatchmakingHoppers();
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

                var capturedAt = DateTimeOffset.Now;
                foreach (var item in results.Where(x =>
                    string.IsNullOrWhiteSpace(x.Result.Error) && x.Result.Population.HasValue))
                {
                    _statsPopulationHistory.Add(new MatchmakingPopulationSample(
                        capturedAt,
                        item.Hopper.HopperName,
                        item.Hopper.DisplayName,
                        item.Result.Population!.Value));
                }

                _statsPopulationRows.Clear();
                foreach (var item in results
                    .Select(x => new MatchmakingPopulationRow(x.Hopper, x.Result))
                    .OrderByDescending(x => x.Population ?? -1)
                    .ThenBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase))
                {
                    _statsPopulationRows.Add(item);
                }

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
            try { File.WriteAllText(StatsSettingsFile, gt); } catch { }
            StatsRefreshSessionUI();
            StatsSetStatus("Fetching stats…");
            _ = StatsFetchStats(gt);
            string tok; lock (_statsLock) { tok = _statsSpartanToken; }
            if (!string.IsNullOrEmpty(tok))
                _ = StatsFetchRecentStatsAsync(gt, tok);
        }

        private void StatsApplyCapturedToken(string token)
        {
            lock (_statsLock) { _statsSpartanToken = token; _statsHwTokenExpired = false; }
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

        private Task<bool> StatsTrySilentTokenRefreshAsync()
        {
            string gt; lock (_statsLock) { gt = _statsGamertag; }
            return Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    var win = new StatsAuthWindow(gt, silent: true);
                    bool ok = win.ShowDialog() == true && !string.IsNullOrEmpty(win.CapturedToken);
                    if (ok) StatsApplyCapturedToken(win.CapturedToken!);
                    return ok;
                }
                catch { return false; }
            }).Task;
        }

        // ── UI helpers ────────────────────────────────────────────────────────

        private void StatsSetStatus(string msg) =>
            Dispatcher.InvokeAsync(() => StatsStatusLabel.Text = msg);

        private void StatsUpdateHwStatus()
        {
            string text, buttonText; Brush color;
            lock (_statsLock)
            {
                if (string.IsNullOrEmpty(_statsSpartanToken))
                    (text, buttonText, color) = ("WAYPOINT: NOT CONNECTED", "CONNECT WAYPOINT", new SolidColorBrush(Color.FromRgb(0x4A, 0x5A, 0x6A)));
                else if (_statsHwTokenExpired)
                    (text, buttonText, color) = ("WAYPOINT: EXPIRED", "RECONNECT WAYPOINT", new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55)));
                else
                    (text, buttonText, color) = ("WAYPOINT: CONNECTED", "WAYPOINT CONNECTED", new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF)));
            }
            Dispatcher.InvokeAsync(() =>
            {
                StatsHwStatusLabel.Text = text;
                StatsHwStatusLabel.Foreground = color;
                StatsHwAuthBtn.Content = buttonText;
            });
        }

        private void StatsRefreshSessionUI()
        {
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
            Dispatcher.InvokeAsync(() =>
            {
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
            });
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
                .Select(p =>
                {
                    string gt    = p.Attribute("mGamertagText")?.Value ?? "Unknown";
                    string xuid  = p.Attribute("mXboxUserId")?.Value ?? "";
                    string normalizedXuid = StatsNormalizeXuid(xuid);
                    string team  = isFfa ? "FFA"
                        : p.Attribute("mTeamIndex")?.Value ?? p.Attribute("mTeamId")?.Value ?? "0";
                    string kd       = kdSnap.GetValueOrDefault(gt, "—");
                    string recentKd = recentKdSnap.GetValueOrDefault(gt, "");
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
                        Standing      = int.TryParse(p.Attribute("mStanding")?.Value, out int s) ? s : 99,
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

            return string.IsNullOrWhiteSpace(normalizedXuid) ? "Resolving..." : $"Resolving {StatsShortXuid(normalizedXuid)}";
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
                try { StatsCheckForNewFile(); } catch { }
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
            lock (_statsLock) { changed = sig != _statsLastFileSig; if (changed) _statsLastFileSig = sig; }
            if (changed) StatsProcessFile(f.FullName);
        }

        private static FileInfo? StatsLatestCarnageFile() =>
            new DirectoryInfo(StatsWatchPath)
                .GetFiles("mpcarnagereport*.xml")
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

        private static string StatsSig(FileInfo f) =>
            $"{f.FullName}|{f.Length}|{f.LastWriteTime.Ticks}";

        private void StatsProcessFile(string path, bool countTowardSession = true)
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
            if (doc == null) return;

            string? gId = doc.Descendants("GameUniqueId")
                .Select(e =>
                {
                    string? attr = e.Attribute("GameUniqueId")?.Value;
                    return !string.IsNullOrEmpty(attr) ? attr : e.Value;
                })
                .FirstOrDefault(v => !string.IsNullOrEmpty(v));

            if (string.IsNullOrEmpty(gId)) return;

            var players = doc.Descendants("Player").ToList();
            bool triggerLobby = false;

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
                if (countTowardSession && !_statsSession.ProcessedGameIds.Contains(gId))
                {
                    var me = players.FirstOrDefault(p =>
                        p.Attribute("mGamertagText")?.Value
                         .Equals(_statsGamertag, StringComparison.OrdinalIgnoreCase) == true);

                    if (me != null)
                    {
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

                        var gameRow = new StatsSessionGameRow
                        {
                            Game = _statsSession.GamesPlayed,
                            Result = won ? "WIN" : "LOSS",
                            KillsDeaths = $"{k}–{d}",
                            KD = gameKd.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
                            BestSpree = spree,
                            HighestMultikill = highestMultikill,
                            HighestMultikillIcon = StatsMultikillMedals
                                .FirstOrDefault(m => m.Name == highestMultikill)?.ResourcePath ?? ""
                        };
                        Dispatcher.InvokeAsync(() => _statsSessionGames.Insert(0, gameRow));

                        StatsSetStatus($"Game logged — K:{k}  D:{d}  Standing:{standing}");
                        triggerLobby = _statsAutoPullLobby;
                    }
                }
            }

            StatsRefreshSessionUI();
            StatsRebuildLobbyRows();
            Dispatcher.InvokeAsync(() => StatsLastGameServerLabel.Text = _statsLastGameServerText);
            if (!countTowardSession)
                StatsSetStatus($"Loaded last game: {Path.GetFileName(path)}");
            if (triggerLobby) _ = StatsFetchLobbyStats();
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
            string token; bool expired;
            lock (_statsLock) { token = _statsSpartanToken; expired = _statsHwTokenExpired; }

            if (!string.IsNullOrEmpty(token) && !expired)
            {
                var (success, unauthorized) = await StatsFetchHaloWaypointStats(gt, token);
                if (success) return;

                if (unauthorized)
                {
                    lock (_statsLock) { _statsHwTokenExpired = true; }
                    StatsUpdateHwStatus();
                    StatsSetStatus("HW token expired — attempting silent refresh…");
                    bool refreshed = await StatsTrySilentTokenRefreshAsync();
                    if (refreshed)
                    {
                        string newToken; lock (_statsLock) { newToken = _statsSpartanToken; }
                        var (s2, _) = await StatsFetchHaloWaypointStats(gt, newToken);
                        if (s2) return;
                    }
                    else
                    {
                        StatsSetStatus("Silent refresh failed — connect Halo Waypoint again.");
                    }
                }
            }

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

                if (resp.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
                    return (false, true);

                if (!resp.IsSuccessStatusCode)
                {
                    StatsSetStatus($"[HW] API {(int)resp.StatusCode} for {gt}");
                    return (false, false);
                }

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

        private async Task StatsFetchWortStats(string gt)
        {
            try
            {
                string url = $"https://wort.gg/api/stats/{Uri.EscapeDataString(gt)}/multiplayer";
                var resp = await StatsHttp.GetAsync(url);
                string body = await resp.Content.ReadAsStringAsync();
                using var json = JsonDocument.Parse(body);

                if (!resp.IsSuccessStatusCode)
                {
                    lock (_statsLock) { _statsKd[gt] = "N/A"; _statsTotals[gt] = ""; }
                    StatsSetStatus($"[wort.gg] API {(int)resp.StatusCode} for {gt}");
                    StatsRefreshLifetimeUI(); StatsRebuildLobbyRows(); return;
                }

                var (kills, deaths) = StatsExtractWortKillsDeaths(json.RootElement);
                if (kills > 0 || deaths > 0)
                {
                    string kdVal = deaths > 0 ? ((double)kills / deaths).ToString("F2") : kills.ToString();
                    string totals = $"{kills:N0}K / {deaths:N0}D";
                    lock (_statsLock) { _statsKd[gt] = kdVal; _statsTotals[gt] = totals; }
                    StatsAddToCache(gt, kdVal, totals);
                    StatsSetStatus($"[wort.gg] {gt} — K/D: {kdVal}");
                }
                else
                {
                    lock (_statsLock) { _statsKd[gt] = "N/A"; _statsTotals[gt] = ""; }
                    StatsSetStatus($"[wort.gg] No stats found for {gt}");
                }
            }
            catch (Exception ex)
            {
                lock (_statsLock) { _statsKd[gt] = "ERR"; _statsTotals[gt] = ""; }
                StatsSetStatus($"[wort.gg] Error for {gt}: {ex.Message}");
            }
            StatsRefreshLifetimeUI();
            StatsRebuildCurrentLobbyRows();
            StatsRebuildLobbyRows();
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
                    .Select(p => StatsResolveGamertag(p, gamertagsByXuid))
                    .Where(gt => !string.IsNullOrWhiteSpace(gt) &&
                                 !gt.StartsWith("Resolving", StringComparison.OrdinalIgnoreCase))
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
            if (File.Exists(StatsSettingsFile))
                try { _statsGamertag = File.ReadAllText(StatsSettingsFile).Trim(); } catch { }
        }

        private void StatsLoadSpartanToken()
        {
            if (!File.Exists(StatsTokenFile)) return;
            try
            {
                string t = File.ReadAllText(StatsTokenFile).Trim();
                if (!string.IsNullOrEmpty(t)) _statsSpartanToken = t;
            }
            catch { }
        }

        private void StatsSaveToken(string token)
        {
            try { File.WriteAllText(StatsTokenFile, token); } catch { }
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
                try { File.WriteAllText(StatsCacheFile, JsonSerializer.Serialize(_statsPersistentCache)); } catch { }
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

    public sealed record MatchmakingPopulationSample(
        DateTimeOffset CapturedAt,
        string HopperName,
        string DisplayName,
        int Population);

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

    class StatsSessionGameRow
    {
        public int Game { get; init; }
        public string Result { get; init; } = "";
        public string KillsDeaths { get; init; } = "";
        public string KD { get; init; } = "";
        public int BestSpree { get; init; }
        public string HighestMultikill { get; init; } = "";
        public string HighestMultikillIcon { get; init; } = "";
        public Brush ResultColor => Result == "WIN"
            ? new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14))
            : new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));
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
