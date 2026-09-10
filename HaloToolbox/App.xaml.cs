using System;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace HaloToolbox
{
    public partial class App : Application
    {
        private static bool _isDark = true;
        public static bool IsDarkTheme => _isDark;

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
                using var key = Registry.CurrentUser.OpenSubKey(@"Software\HaloMCCToolbox");
                _isDark = (key?.GetValue("Theme") as string) != "Light";
            }
            catch { _isDark = true; }
            ApplyTheme(_isDark);
        }

        private static void SaveTheme(bool dark)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(@"Software\HaloMCCToolbox");
                key.SetValue("Theme", dark ? "Dark" : "Light");
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
