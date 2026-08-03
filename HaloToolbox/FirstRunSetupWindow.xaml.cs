using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace HaloToolbox
{
    public partial class FirstRunSetupWindow : Window
    {
        private int _step = 1;
        private string? _waypointToken;

        public FirstRunSetupWindow()
        {
            InitializeComponent();
            MccPathBox.Text = App.FindMccInstallationPath();
            GamertagBox.Text = App.LoadPlayerGamertag();
            UpdatePathStatus();
            UpdateStep();
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Select your Halo MCC installation folder",
                InitialDirectory = Directory.Exists(MccPathBox.Text) ? MccPathBox.Text : App.DefaultMccPath
            };
            if (dialog.ShowDialog() == true)
                MccPathBox.Text = dialog.FolderName;
        }

        private void MccPathBox_TextChanged(object sender, TextChangedEventArgs e) => UpdatePathStatus();

        private void UpdatePathStatus()
        {
            if (PathStatus is null || NextButton is null)
                return;
            bool valid = App.IsValidMccInstallationPath(MccPathBox.Text.Trim());
            PathStatus.Text = valid ? "✓ STEAM INSTALLATION VERIFIED" : "MCC INSTALLATION NOT VERIFIED";
            PathStatus.Foreground = (System.Windows.Media.Brush)FindResource(valid ? "GreenBrush" : "OrangeBrush");
            NextButton.IsEnabled = _step != 1 || valid;
        }

        private void ConnectWaypoint_Click(object sender, RoutedEventArgs e)
        {
            string gamertag = GamertagBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(gamertag))
            {
                ToolboxDialog.Show(
                    "Enter your gamertag before connecting Halo Waypoint.",
                    "Halo MCC Toolbox Setup",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                GamertagBox.Focus();
                return;
            }

            var auth = new StatsAuthWindow(gamertag) { Owner = this };
            if (auth.ShowDialog() == true && !string.IsNullOrWhiteSpace(auth.CapturedToken))
            {
                _waypointToken = auth.CapturedToken;
                WaypointStatus.Text = "✓ WAYPOINT CONNECTED";
                WaypointStatus.Foreground = (System.Windows.Media.Brush)FindResource("GreenBrush");
                ConnectWaypointButton.Content = "WAYPOINT CONNECTED";
            }
        }

        private void Recommended_Click(object sender, RoutedEventArgs e)
        {
            StatsPreference.IsChecked = true;
            NetworkPreference.IsChecked = true;
            ModsPreference.IsChecked = true;
            TheaterPreference.IsChecked = false;
            PlaylistsPreference.IsChecked = false;
            AdvancedPreference.IsChecked = false;
            OpenLastSectionPreference.IsChecked = true;
        }

        private void FeatureCard_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border { Tag: CheckBox checkBox })
                return;

            // The CheckBox handles its own direct click. Toggle here only when
            // the rest of the card was clicked.
            if (e.OriginalSource is CheckBox ||
                FindVisualParent<CheckBox>(e.OriginalSource as DependencyObject) is not null)
                return;

            checkBox.IsChecked = checkBox.IsChecked != true;
        }

        private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
        {
            while (child is not null)
            {
                if (child is T match)
                    return match;
                child = System.Windows.Media.VisualTreeHelper.GetParent(child);
            }
            return null;
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            if (_step == 1)
            {
                DialogResult = false;
                return;
            }
            _step--;
            UpdateStep();
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            if (_step == 1 && !App.IsValidMccInstallationPath(MccPathBox.Text.Trim()))
                return;

            if (_step < 3)
            {
                _step++;
                UpdateStep();
                return;
            }

            SaveSetup();
            DialogResult = true;
        }

        private void UpdateStep()
        {
            InstallPanel.Visibility = _step == 1 ? Visibility.Visible : Visibility.Collapsed;
            PreferencesPanel.Visibility = _step == 2 ? Visibility.Visible : Visibility.Collapsed;
            ReviewPanel.Visibility = _step == 3 ? Visibility.Visible : Visibility.Collapsed;

            BackButton.Content = _step == 1 ? "EXIT" : "‹  BACK";
            NextButton.Content = _step == 3 ? "FINISH SETUP  ✓" : "CONTINUE  ›";
            RecommendedButton.Visibility = _step == 2 ? Visibility.Visible : Visibility.Collapsed;

            SetStepVisual(Step1Box, Step1Number, 1);
            SetStepVisual(Step2Box, Step2Number, 2);
            SetStepVisual(Step3Box, Step3Number, 3);
            Step2Label.Foreground = (System.Windows.Media.Brush)FindResource(_step == 2 ? "AccentBrush" : "MutedBrush");
            Step3Label.Foreground = (System.Windows.Media.Brush)FindResource(_step == 3 ? "AccentBrush" : "MutedBrush");

            if (_step == 3)
                PopulateReview();
            UpdatePathStatus();
        }

        private void SetStepVisual(Border box, TextBlock number, int step)
        {
            string brush = step < _step ? "GreenBrush" : step == _step ? "AccentBrush" : "BorderBrush";
            box.BorderBrush = (System.Windows.Media.Brush)FindResource(brush);
            number.Foreground = (System.Windows.Media.Brush)FindResource(
                step < _step ? "GreenBrush" : step == _step ? "AccentBrush" : "MutedBrush");
            number.Text = step < _step ? "✓" : step.ToString();
        }

        private IEnumerable<string> SelectedFeatures()
        {
            if (StatsPreference.IsChecked == true) yield return "✓ Stats & Lobbies";
            if (NetworkPreference.IsChecked == true) yield return "✓ Network & Region";
            if (ModsPreference.IsChecked == true) yield return "✓ Mods & Maps";
            if (TheaterPreference.IsChecked == true) yield return "✓ Theater";
            if (PlaylistsPreference.IsChecked == true) yield return "✓ Playlists";
            if (AdvancedPreference.IsChecked == true) yield return "✓ Advanced / Rejoin";
        }

        private void PopulateReview()
        {
            ReviewPath.Text = "STEAM  ✓\n" + MccPathBox.Text.Trim();
            ReviewGamertag.Text = string.IsNullOrWhiteSpace(GamertagBox.Text)
                ? "Not set — configure later from Stats"
                : GamertagBox.Text.Trim();
            ReviewWaypoint.Text = _waypointToken is null ? "Not connected — setup later" : "Connected  ✓";
            var features = SelectedFeatures().ToArray();
            ReviewFeatures.Text = features.Length == 0 ? "Tools only" : string.Join(Environment.NewLine, features);
        }

        private void SaveSetup()
        {
            App.SaveMccInstallationPath(MccPathBox.Text.Trim());
            App.SavePlayerGamertag(GamertagBox.Text.Trim());

            App.SaveMainSectionVisible("H3Mods", ModsPreference.IsChecked == true);
            App.SaveMainSectionVisible("Report", true);
            App.SaveMainSectionVisible("Stats", StatsPreference.IsChecked == true);
            App.SaveMainSectionVisible("Theater", TheaterPreference.IsChecked == true);
            App.SaveMainSectionVisible("Playlists", PlaylistsPreference.IsChecked == true);
            App.SaveMainSectionVisible("About", false);
            App.SaveMainSectionVisible("Log", AdvancedPreference.IsChecked == true);

            App.SaveSetupPreference("StatsAndLobbies", StatsPreference.IsChecked == true);
            App.SaveSetupPreference("NetworkAndRegion", NetworkPreference.IsChecked == true);
            App.SaveSetupPreference("ModsAndMaps", ModsPreference.IsChecked == true);
            App.SaveSetupPreference("Theater", TheaterPreference.IsChecked == true);
            App.SaveSetupPreference("Playlists", PlaylistsPreference.IsChecked == true);
            App.SaveSetupPreference("AdvancedRejoin", AdvancedPreference.IsChecked == true);
            App.SaveSetupPreference("OpenLastSection", OpenLastSectionPreference.IsChecked == true);

            App.SaveGameNetworkStatsOverlayEnabled(false);
            App.SaveMatchmakingWaitOverlayEnabled(false);
            App.SaveObsBrowserOverlayEnabled(false);
            App.SaveStatsAutoLobbyEnabled(StatsPreference.IsChecked == true);
            App.SaveRejoinFirewallMode("Disabled");

            Directory.CreateDirectory(App.ToolboxDataRoot);
            if (!string.IsNullOrWhiteSpace(GamertagBox.Text))
                File.WriteAllText(App.StatsGamertagPath, GamertagBox.Text.Trim());
            if (!string.IsNullOrWhiteSpace(_waypointToken))
                File.WriteAllText(App.StatsTokenPath, _waypointToken);

            App.SaveFirstLaunchSetupCompleted();
        }
    }
}
