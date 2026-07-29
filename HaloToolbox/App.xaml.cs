using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace HaloToolbox
{
    public partial class App : Application
    {
        private static bool _isDark = true;
        private const string SettingsRegistryPath = @"Software\HaloMCCToolbox";
        private const string DefaultMccInstallationPath = @"C:\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection";

        public static bool IsDarkTheme => _isDark;
        public static string DefaultMccPath => DefaultMccInstallationPath;
        public readonly record struct WindowPlacement(double Left, double Top, double Width, double Height, bool IsMaximized);

        public static bool LoadMainSectionVisible(string sectionName)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return (key?.GetValue($"MainSection.{sectionName}") as string) != "Hidden";
            }
            catch
            {
                return true;
            }
        }

        public static void SaveMainSectionVisible(string sectionName, bool visible)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue($"MainSection.{sectionName}", visible ? "Visible" : "Hidden");
            }
            catch { }
        }

        public static bool LoadGameNetworkStatsOverlayEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return (key?.GetValue("GameNetworkStatsOverlay") as string) != "Disabled";
            }
            catch
            {
                return true;
            }
        }

        public static void SaveGameNetworkStatsOverlayEnabled(bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("GameNetworkStatsOverlay", enabled ? "Enabled" : "Disabled");
            }
            catch { }
        }

        public static bool LoadMatchmakingWaitOverlayEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return (key?.GetValue("MatchmakingWaitOverlay") as string) != "Disabled";
            }
            catch { return true; }
        }

        public static void SaveMatchmakingWaitOverlayEnabled(bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("MatchmakingWaitOverlay", enabled ? "Enabled" : "Disabled");
            }
            catch { }
        }

        public static string LoadRejoinFirewallMode()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return key?.GetValue("RejoinFirewallMode") as string ?? "Disabled";
            }
            catch { return "Disabled"; }
        }

        public static void SaveRejoinFirewallMode(string mode)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("RejoinFirewallMode", mode);
            }
            catch { }
        }

        public static bool LoadObsBrowserOverlayEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return (key?.GetValue("ObsBrowserOverlay") as string) == "Enabled";
            }
            catch
            {
                return false;
            }
        }

        public static void SaveObsBrowserOverlayEnabled(bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("ObsBrowserOverlay", enabled ? "Enabled" : "Disabled");
            }
            catch { }
        }

        private static bool LoadFeatureObsOnlyOverlayEnabled(string valueName)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                var value = key?.GetValue(valueName) as string;
                if (value is not null)
                    return value == "Enabled";

                // Preserve the previous global choice the first time this build runs.
                return (key?.GetValue("ObsOnlyOverlay") as string) == "Enabled";
            }
            catch { return false; }
        }

        private static void SaveFeatureObsOnlyOverlayEnabled(string valueName, bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue(valueName, enabled ? "Enabled" : "Disabled");
            }
            catch { }
        }

        public static bool LoadNetworkStatsObsOnlyEnabled() =>
            LoadFeatureObsOnlyOverlayEnabled("NetworkStatsObsOnly");

        public static void SaveNetworkStatsObsOnlyEnabled(bool enabled) =>
            SaveFeatureObsOnlyOverlayEnabled("NetworkStatsObsOnly", enabled);

        public static bool LoadMatchmakingWaitObsOnlyEnabled() =>
            LoadFeatureObsOnlyOverlayEnabled("MatchmakingWaitObsOnly");

        public static void SaveMatchmakingWaitObsOnlyEnabled(bool enabled) =>
            SaveFeatureObsOnlyOverlayEnabled("MatchmakingWaitObsOnly", enabled);

        public static bool LoadSessionStatsObsOnlyEnabled() =>
            LoadFeatureObsOnlyOverlayEnabled("SessionStatsObsOnly");

        public static void SaveSessionStatsObsOnlyEnabled(bool enabled) =>
            SaveFeatureObsOnlyOverlayEnabled("SessionStatsObsOnly", enabled);

        public static bool LoadObsBrowserOverlaySessionStatsEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return (key?.GetValue("ObsBrowserOverlaySessionStats") as string) != "Disabled";
            }
            catch
            {
                return true;
            }
        }

        public static void SaveObsBrowserOverlaySessionStatsEnabled(bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("ObsBrowserOverlaySessionStats", enabled ? "Enabled" : "Disabled");
            }
            catch { }
        }

        public static bool LoadStatsAutoLobbyEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return (key?.GetValue("StatsAutoLobby") as string) != "Disabled";
            }
            catch
            {
                return true;
            }
        }

        public static void SaveStatsAutoLobbyEnabled(bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("StatsAutoLobby", enabled ? "Enabled" : "Disabled");
            }
            catch { }
        }

        public static void SavePendingRejoinFixAutoStart(bool pending)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("PendingRejoinFixAutoStart", pending ? "Enabled" : "Disabled");
            }
            catch { }
        }

        public static bool ConsumePendingRejoinFixAutoStart()
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                bool pending = (key.GetValue("PendingRejoinFixAutoStart") as string) == "Enabled";
                key.SetValue("PendingRejoinFixAutoStart", "Disabled");
                return pending;
            }
            catch
            {
                return false;
            }
        }

        public static WindowPlacement? LoadMainWindowPlacement()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                if (key is null)
                    return null;

                double? left = ReadDoubleRegistryValue(key, "MainWindowLeft");
                double? top = ReadDoubleRegistryValue(key, "MainWindowTop");
                double? width = ReadDoubleRegistryValue(key, "MainWindowWidth");
                double? height = ReadDoubleRegistryValue(key, "MainWindowHeight");
                if (left is null || top is null || width is null || height is null)
                    return null;

                bool isMaximized = (key.GetValue("MainWindowState") as string) == "Maximized";
                return new WindowPlacement(left.Value, top.Value, width.Value, height.Value, isMaximized);
            }
            catch
            {
                return null;
            }
        }

        public static void SaveMainWindowPlacement(WindowPlacement placement)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("MainWindowLeft", placement.Left.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String);
                key.SetValue("MainWindowTop", placement.Top.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String);
                key.SetValue("MainWindowWidth", placement.Width.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String);
                key.SetValue("MainWindowHeight", placement.Height.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String);
                key.SetValue("MainWindowState", placement.IsMaximized ? "Maximized" : "Normal", RegistryValueKind.String);
            }
            catch { }
        }

        private static double? ReadDoubleRegistryValue(RegistryKey key, string name)
        {
            return double.TryParse(
                key.GetValue(name) as string,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out double value)
                    ? value
                    : null;
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            DispatcherUnhandledException += (_, args) =>
            {
                try
                {
                    MessageBox.Show(
                        $"The app hit an unexpected error:\n\n{args.Exception.Message}",
                        "Halo MCC Toolbox",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
                catch { }
            };
            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                if (args.ExceptionObject is Exception ex)
                {
                    try
                    {
                        MessageBox.Show(
                            $"A fatal error occurred:\n\n{ex.Message}",
                            "Halo MCC Toolbox",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                    catch { }
                }
            };
            LoadSavedTheme();
        }

        public static void ToggleTheme()
        {
            _isDark = !_isDark;
            ApplyTheme(_isDark);
            SaveTheme(_isDark);
        }

        private static void LoadSavedTheme()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                _isDark = (key?.GetValue("Theme") as string) != "Light";
            }
            catch { _isDark = true; }
            ApplyTheme(_isDark);
        }

        private static void SaveTheme(bool dark)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("Theme", dark ? "Dark" : "Light");
            }
            catch { }
        }

        public static string LoadMccInstallationPath()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                var savedPath = key?.GetValue("MccInstallationPath") as string;
                return string.IsNullOrWhiteSpace(savedPath) ? DefaultMccInstallationPath : savedPath;
            }
            catch
            {
                return DefaultMccInstallationPath;
            }
        }

        public static void SaveMccInstallationPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("MccInstallationPath", path.Trim());
            }
            catch { }
        }

        public static string LoadDownpatchWorkspacePath()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                return key?.GetValue("DownpatchWorkspacePath") as string ?? "";
            }
            catch
            {
                return "";
            }
        }

        public static void SaveDownpatchWorkspacePath(string path)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key.SetValue("DownpatchWorkspacePath", path.Trim());
            }
            catch { }
        }

        private static void ApplyTheme(bool dark) { if (dark) ApplyDark(); else ApplyLight(); }

        private static void Set(string key, string hex)
        {
            Application.Current.Resources[key] =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        }

        private static void ApplyDark()
        {
            Set("BgBrush",            "#0A0C10");
            Set("PanelBrush",         "#0F1318");
            Set("SurfaceBrush",       "#080B0F");
            Set("BorderBrush",        "#1E2530");
            Set("TextBrush",          "#C8D8E8");
            Set("MutedBrush",         "#4A5A6A");
            Set("SubtleBrush",        "#2A3A4A");
            Set("AccentBrush",        "#00C8FF");
            Set("GreenBrush",         "#39FF14");
            Set("RedBrush",           "#FF2D55");
            Set("OrangeBrush",        "#FF6A00");
            Set("ComboHoverBrush",    "#1A2535");
            Set("ComboSelectedBrush", "#0A2040");
            Set("StatsMyRowBrush",    "#081C10");
            Set("StatsTeam0RowBrush", "#1C0808");
            Set("StatsTeam1RowBrush", "#080C1C");
            Set("StatsScanRowBrush",  "#0C1810");
            Set("StatsHoverRowBrush", "#141C28");
            Set("MatchWinRowBrush",   "#071410");
            Set("MatchLossRowBrush",  "#140808");
        }

        private static void ApplyLight()
        {
            Set("BgBrush",            "#F1F5F9");
            Set("PanelBrush",         "#E2E8F0");
            Set("SurfaceBrush",       "#FFFFFF");
            Set("BorderBrush",        "#CBD5E1");
            Set("TextBrush",          "#1E293B");
            Set("MutedBrush",         "#475569");
            Set("SubtleBrush",        "#94A3B8");
            Set("AccentBrush",        "#0284C7");
            Set("GreenBrush",         "#16A34A");
            Set("RedBrush",           "#DC2626");
            Set("OrangeBrush",        "#C2410C");
            Set("ComboHoverBrush",    "#E2E8F0");
            Set("ComboSelectedBrush", "#DBEAFE");
            Set("StatsMyRowBrush",    "#DCFCE7");
            Set("StatsTeam0RowBrush", "#FEE2E2");
            Set("StatsTeam1RowBrush", "#DBEAFE");
            Set("StatsScanRowBrush",  "#F0FDF4");
            Set("StatsHoverRowBrush", "#EFF6FF");
            Set("MatchWinRowBrush",   "#F0FDF4");
            Set("MatchLossRowBrush",  "#FEF2F2");
        }
    }
}
