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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
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
        private List<XElement> _statsLastPlayers = new();
        private string _statsLastFileSig = "";
        private bool _statsAutoPullLobby = false;
        private string _statsSpartanToken = "";
        private bool _statsHwTokenExpired = false;

        // ── Stats Tab — lookup caches ────────────────────────────────────────
        private readonly Dictionary<string, string> _statsKd =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statsTotals =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statsGames =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statsRecentKd =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _statsRecentMaxGap =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, StatsCachedPlayer> _statsPersistentCache =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Queue<string> _statsCacheOrder = new();

        // ── Stats Tab — UI collection ────────────────────────────────────────
        private readonly ObservableCollection<StatsPlayerRow> _statsLobbyRows = new();
        private readonly List<string> _sessionLogLines = new();
        private readonly ProxyService _rejoinProxy = new();
        private bool _rejoinWinHttpManualNeeded;

        private const long MaxDiagnosticExportBytes = 25L * 1024 * 1024;
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
            MapList.ItemsSource = _maps;
            AppendLog("[INFO]", "Halo MCC Toolbox started. Made by The FFA Panda.", "#00C8FF");

            // Check Halo Support login state once the window is fully rendered
            // Load maps asynchronously in background after window renders
            Loaded += async (_, _) =>
            {
                await CheckSupportSessionAsync();
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
                    AppendLog("[REJOIN]", "Captured matchmaking session and saved it to Toolbox appdata.", "#00C8FF");
                    UpdateRejoinFixUi();
                });
            _rejoinProxy.OnPlayerIdentityChanged += (_, _) =>
                Dispatcher.InvokeAsync(UpdateRejoinFixUi);
            Closed += (_, _) => _rejoinProxy.Dispose();
            UpdateRejoinFixUi();

        }

        // ------------------------------------------
        // Window chrome
        // ------------------------------------------
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }

        private void MinBtn_Click(object sender, RoutedEventArgs e) =>
            WindowState = WindowState.Minimized;

        private void CloseBtn_Click(object sender, RoutedEventArgs e) =>
            Close();

        private void ThemeToggleBtn_Click(object sender, RoutedEventArgs e)
        {
            App.ToggleTheme();
            ThemeToggleBtn.Content = App.IsDarkTheme ? "☾" : "☀";
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
                var color = Brush(colorHex);
                _sessionLogLines.Add(line);
                TxtLog.Inlines.Add(new Run($"[{ts}] ") { Foreground = Brush("#4A5A6A") });
                TxtLog.Inlines.Add(new Run($"{tag} ") { Foreground = color, FontWeight = FontWeights.Bold });
                TxtLog.Inlines.Add(new Run(message + "\n") { Foreground = Brush("#C8D8E8") });
                LogScroller.ScrollToEnd();
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

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            TxtLog.Inlines.Clear();
            _sessionLogLines.Clear();
            AppendLog("[INFO]", "Log cleared.", "#4A5A6A");
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
            var mccCrashReportPath = Path.Combine(programFilesX86,
                @"Steam\steamapps\common\Halo The Master Chief Collection\crash_report");

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
            var ext = Path.GetExtension(path);
            if (DiagnosticExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
                return true;

            var fileName = Path.GetFileName(path);
            return fileName.Contains(".log.", StringComparison.OrdinalIgnoreCase)
                || fileName.Contains("crash", StringComparison.OrdinalIgnoreCase);
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
                @"C:\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection",
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

        private void UpdateRejoinFixUi()
        {
            bool isRunning = _rejoinProxy.IsRunning;
            bool hasSavedState = File.Exists(RejoinFixPaths.LastHandleFile)
                || File.Exists(RejoinFixPaths.LastMatchSessionFile)
                || File.Exists(RejoinFixPaths.LastGameServerFile);
            string gamertagSuffix = string.IsNullOrWhiteSpace(_rejoinProxy.CurrentPlayerGamertag)
                ? ""
                : $" ({_rejoinProxy.CurrentPlayerGamertag})";

            BtnRejoinFix.Content = isRunning ? "STOP FIX" : "RUN FIX";

            if (isRunning && _rejoinWinHttpManualNeeded)
            {
                TxtRejoinFixStatus.Text = $"ACTIVE{gamertagSuffix} - proxy listening; MCC capture may still need admin proxy approval";
                TxtRejoinFixStatus.Foreground = Brush("#FF6A00");
            }
            else if (isRunning)
            {
                TxtRejoinFixStatus.Text = $"ACTIVE{gamertagSuffix} - capturing in background and saving rejoin state to Toolbox appdata";
                TxtRejoinFixStatus.Foreground = Brush("#39FF14");
            }
            else if (hasSavedState)
            {
                TxtRejoinFixStatus.Text = $"OFF{gamertagSuffix} - saved rejoin capture files are present for diagnostics export";
                TxtRejoinFixStatus.Foreground = Brush("#C8D8E8");
            }
            else
            {
                TxtRejoinFixStatus.Text = $"OFF{gamertagSuffix}";
                TxtRejoinFixStatus.Foreground = Brush("#4A5A6A");
            }
        }

        private async void BtnRejoinFix_Click(object sender, RoutedEventArgs e)
        {
            BtnRejoinFix.IsEnabled = false;

            try
            {
                if (_rejoinProxy.IsRunning)
                {
                    _rejoinProxy.Stop();
                    _rejoinWinHttpManualNeeded = false;
                    AppendLog("[REJOIN]", "Rejoin Fix proxy stopped.", "#C8D8E8");
                    SetStatus("Rejoin Fix stopped.", "#C8D8E8");
                }
                else
                {
                    var confirm = MessageBox.Show(
                        "Rejoin Fix will enable the Toolbox proxy and change the system proxy while it is running.\n\nRestart MCC after it starts.\n\nRun the fix now?",
                        "Rejoin Fix -- Halo MCC Toolbox",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (confirm != MessageBoxResult.Yes)
                    {
                        AppendLog("[INFO]", "Rejoin Fix cancelled by user.", "#4A5A6A");
                        SetStatus("Rejoin Fix cancelled.", "#4A5A6A");
                        return;
                    }

                    RejoinFixPaths.EnsureRootDirectory();
                    _rejoinWinHttpManualNeeded = false;
                    RejoinFixDiagnostics.Info("proxy", "Activation requested from Toolbox UI.");
                    AppendLog("[REJOIN]", "Starting Rejoin Fix proxy...", "#FF6A00");
                    SetStatus("Starting Rejoin Fix...", "#FF6A00");
                    await _rejoinProxy.StartAsync();
                    AppendLog("[REJOIN]", $"Rejoin Fix active on 127.0.0.1:{_rejoinProxy.Port}. Restart MCC now.", "#39FF14");
                    SetStatus("Rejoin Fix active.", "#39FF14");
                }
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
                TxtMccPath.Text = dlg.FolderName;
        }

        private void BtnLoadMaps_Click(object sender, RoutedEventArgs e) => LoadMaps(TxtMccPath.Text.Trim());

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
                map.IsEnabled = !map.IsEnabled;
        }

        private void BtnEnableAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var m in _maps.Where(m => !m.IsHeader)) m.IsEnabled = true;
            AppendLog("[INFO]", "All maps set to ENABLED.", "#39FF14");
        }

        private void BtnDisableAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var m in _maps.Where(m => !m.IsHeader)) m.IsEnabled = false;
            AppendLog("[INFO]", "All maps set to DISABLED.", "#FF2D55");
        }

        private void BtnDisable343_Click(object sender, RoutedEventArgs e)
        {
            int count = 0;
            foreach (var m in _maps.Where(m => !m.IsHeader && Map343Names.Contains(m.BaseName)))
            {
                m.IsEnabled = false;
                count++;
            }
            if (count == 0)
                AppendLog("[WARN]", "No 343 maps found. Load maps first.", "#FF6A00");
            else
                AppendLog("[INFO]", $"Disabled {count} 343 map(s): Edge, Waterfall, Icebox. Click APPLY to save.", "#FF2D55");
        }

        private void BtnApplyMaps_Click(object sender, RoutedEventArgs e)
        {
            var applyList = _maps.Where(m => !m.IsHeader).ToList();
            if (applyList.Count == 0)
            {
                MessageBox.Show("No maps loaded. Load your maps first.", "Halo MCC Toolbox",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show(
                    "Apply map changes?\n\nEnabled maps will have the REMOVED_ prefix removed.\nDisabled maps will be prefixed with REMOVED_ so MCC skips them.\n\nMake sure MCC is closed.",
                    "Confirm Apply -- Halo MCC Toolbox",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            int ok = 0, fail = 0;
            foreach (var map in applyList)
            {
                try
                {
                    var dir    = Path.GetDirectoryName(map.FileName)!;
                    var ext    = Path.GetExtension(map.FileName);
                    var target = map.IsEnabled
                        ? Path.Combine(dir, map.BaseName + ext)
                        : Path.Combine(dir, RemovedPrefix + map.BaseName + ext);

                    if (!map.FileName.Equals(target, StringComparison.OrdinalIgnoreCase))
                    {
                        File.Move(map.FileName, target);
                        AppendLog(map.IsEnabled ? "[ENABLE]" : "[REMOVE]",
                            $"{Path.GetFileName(map.FileName)}  =>  {Path.GetFileName(target)}",
                            map.IsEnabled ? "#39FF14" : "#FF2D55");
                        map.FileName = target;
                    }
                    ok++;
                }
                catch (Exception ex)
                {
                    AppendLog("[ERROR]", $"Failed to rename {map.BaseName}: {ex.Message}", "#FF2D55");
                    fail++;
                }
            }

            var col = fail > 0 ? "#FF6A00" : "#39FF14";
            SetStatus($"Applied: {ok} updated, {fail} failed.", col);
            AppendLog("[DONE]", $"Applied. {ok} updated, {fail} failed.", "#39FF14");
            MessageBox.Show($"Done!\n\n{ok} map(s) updated.\n{fail} failed.",
                "Applied -- Halo MCC Toolbox", MessageBoxButton.OK, MessageBoxImage.Information);
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
            win.Closed += (_, _) => _ = CheckSupportSessionAsync();

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
            StatsLobbyList.ItemsSource = _statsLobbyRows;

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
            lock (_statsLock) { _statsSession.Reset(); }
            StatsRefreshSessionUI();
            StatsSetStatus("Session reset.");
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
        }

        private void StatsAutoToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            lock (_statsLock) { _statsAutoPullLobby = false; }
            StatsAutoToggle.Content = "AUTO: OFF";
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
            if (StatsLobbyList.SelectedItem is not StatsPlayerRow row) return;
            string token; lock (_statsLock) { token = _statsSpartanToken; }
            if (string.IsNullOrEmpty(token))
            {
                StatsSetStatus("HW Auth required for match history — click 🔑 HW AUTH first.");
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
                StatsSetStatus("HW Auth required for match history — click 🔑 HW AUTH first.");
                return;
            }
            var win = new PlayerMatchHistoryWindow(gt, token) { Owner = this };
            win.Show();
        }

        private void StatsHyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        // ══════════════════════════════════════════════════════════════════════
        // Stats Tab — Core logic
        // ══════════════════════════════════════════════════════════════════════

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
            string text; Brush color;
            lock (_statsLock)
            {
                if (string.IsNullOrEmpty(_statsSpartanToken))
                    (text, color) = ("HW: —", new SolidColorBrush(Color.FromRgb(0x4A, 0x5A, 0x6A)));
                else if (_statsHwTokenExpired)
                    (text, color) = ("HW: ✗", new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55)));
                else
                    (text, color) = ("HW: ✓", new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF)));
            }
            Dispatcher.InvokeAsync(() =>
            {
                StatsHwStatusLabel.Text = text;
                StatsHwStatusLabel.Foreground = color;
            });
        }

        private void StatsRefreshSessionUI()
        {
            int wins, losses, games; long kills, deaths;
            lock (_statsLock)
            {
                wins = _statsSession.Wins; losses = _statsSession.Losses;
                games = _statsSession.GamesPlayed;
                kills = _statsSession.Kills; deaths = _statsSession.Deaths;
            }
            double kdr = deaths > 0 ? (double)kills / deaths : kills;
            Dispatcher.InvokeAsync(() =>
            {
                StatsWinsLabel.Text       = $"{wins}W";
                StatsLossesLabel.Text     = $"{losses}L";
                StatsGamesLabel.Text      = $"{games} game{(games == 1 ? "" : "s")}";
                StatsSessionKdLabel.Text  = kdr.ToString("F2");
            });
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

        private void StatsRebuildLobbyRows()
        {
            List<XElement> players; string myGt;
            Dictionary<string, string> kdSnap, totSnap, gamesSnap, recentKdSnap;
            Dictionary<string, int>    recentGapSnap;

            lock (_statsLock)
            {
                players      = _statsLastPlayers.ToList();
                myGt         = _statsGamertag;
                kdSnap       = new Dictionary<string, string>(_statsKd,       StringComparer.OrdinalIgnoreCase);
                totSnap      = new Dictionary<string, string>(_statsTotals,   StringComparer.OrdinalIgnoreCase);
                gamesSnap    = new Dictionary<string, string>(_statsGames,    StringComparer.OrdinalIgnoreCase);
                recentKdSnap  = new Dictionary<string, string>(_statsRecentKd, StringComparer.OrdinalIgnoreCase);
                recentGapSnap = new Dictionary<string, int>(_statsRecentMaxGap, StringComparer.OrdinalIgnoreCase);
            }

            bool isFfa = players.Select(p =>
                p.Attribute("mTeamIndex")?.Value ?? p.Attribute("mTeamId")?.Value ?? "0")
                .Distinct().Count() <= 1;

            var rows = players
                .OrderBy(p => p.Attribute("mTeamIndex")?.Value ?? p.Attribute("mTeamId")?.Value ?? "0")
                .ThenBy(p => int.TryParse(p.Attribute("mStanding")?.Value, out int s) ? s : 99)
                .Select(p =>
                {
                    string gt    = p.Attribute("mGamertagText")?.Value ?? "Unknown";
                    string team  = isFfa ? "FFA"
                        : p.Attribute("mTeamIndex")?.Value ?? p.Attribute("mTeamId")?.Value ?? "0";
                    string kd       = kdSnap.GetValueOrDefault(gt, "—");
                    string recentKd = recentKdSnap.GetValueOrDefault(gt, "");

                    string trend = "";
                    if (!string.IsNullOrEmpty(recentKd) &&
                        double.TryParse(recentKd, out double rkd) &&
                        double.TryParse(kd, out double lkd))
                    {
                        trend = rkd > lkd + 0.05 ? "▲"
                              : rkd < lkd - 0.05 ? "▼"
                              : "≈";
                    }

                    int gap = recentGapSnap.GetValueOrDefault(gt, 0);

                    return new StatsPlayerRow
                    {
                        Gamertag      = gt,
                        Team          = team,
                        KD            = kd,
                        Totals        = totSnap.GetValueOrDefault(gt, ""),
                        GamesPlayed   = gamesSnap.GetValueOrDefault(gt, ""),
                        IsMe          = gt.Equals(myGt, StringComparison.OrdinalIgnoreCase),
                        IsScanning    = kd == "…",
                        Standing      = int.TryParse(p.Attribute("mStanding")?.Value, out int s) ? s : 99,
                        RecentKD      = recentKd,
                        RecentTrend   = trend,
                        ReturnGapText = StatsFormatGap(gap),
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

        private static string StatsFormatGap(int days)
        {
            if (days < 30)   return "";
            if (days >= 365) return $"↩ ~{days / 365}yr break";
            if (days >= 60)  return $"↩ ~{days / 30}mo break";
            return $"↩ {days}d break";
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

        private void StatsProcessFile(string path)
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
                if (!_statsSession.ProcessedGameIds.Contains(gId))
                {
                    var me = players.FirstOrDefault(p =>
                        p.Attribute("mGamertagText")?.Value
                         .Equals(_statsGamertag, StringComparison.OrdinalIgnoreCase) == true);

                    if (me != null)
                    {
                        int.TryParse(me.Attribute("mStanding")?.Value, out int standing);
                        long.TryParse(me.Attribute("mKills")?.Value,   out long k);
                        long.TryParse(me.Attribute("mDeaths")?.Value,  out long d);

                        _statsSession.Kills += k;
                        _statsSession.Deaths += d;
                        _statsSession.GamesPlayed++;
                        _statsSession.ProcessedGameIds.Add(gId);
                        if (standing == 0) _statsSession.Wins++; else _statsSession.Losses++;

                        StatsSetStatus($"Game logged — K:{k}  D:{d}  Standing:{standing}");
                        triggerLobby = _statsAutoPullLobby;
                    }
                }
            }

            StatsRefreshSessionUI();
            StatsRebuildLobbyRows();
            if (triggerLobby) _ = StatsFetchLobbyStats();
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
                        StatsSetStatus("Silent refresh failed — click 🔑 HW AUTH to reconnect.");
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

                int maxGap = 0;
                for (int i = 0; i < matches.Count - 1; i++)
                {
                    int gap = (int)(matches[i].date - matches[i + 1].date).TotalDays;
                    if (gap > maxGap) maxGap = gap;
                }

                lock (_statsLock)
                {
                    _statsRecentKd[gt]     = kd.ToString("F2");
                    _statsRecentMaxGap[gt] = maxGap;
                }
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
    // Stats Tab — Player row (lobby ListView)
    // ------------------------------------------
    public class StatsPlayerRow : INotifyPropertyChanged
    {
        private string _kd = "—";
        private string _totals = "";
        private string _gamesPlayed = "";
        private string _recentKD = "";
        private string _recentTrend = "";
        private string _returnGapText = "";

        public string Gamertag  { get; set; } = "";
        public string Team      { get; set; } = "0";
        public string TeamLabel => Team == "FFA" ? "FFA" : $"T{Team}";
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

        public string ReturnGapText
        {
            get => _returnGapText;
            set { _returnGapText = value; OnPropertyChanged(nameof(ReturnGapText)); OnPropertyChanged(nameof(ReturnGapVisibility)); }
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

        public Visibility ReturnGapVisibility =>
            string.IsNullOrEmpty(_returnGapText) ? Visibility.Collapsed : Visibility.Visible;

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
        public HashSet<string> ProcessedGameIds { get; } = new(StringComparer.OrdinalIgnoreCase);

        public void Reset()
        {
            Wins = 0; Losses = 0; GamesPlayed = 0; Kills = 0; Deaths = 0;
            ProcessedGameIds.Clear();
        }
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
