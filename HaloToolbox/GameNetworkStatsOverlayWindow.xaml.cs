using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace HaloToolbox;

internal enum GameOverlayVisualStyle
{
    Classic,
    Modern
}

public partial class GameNetworkStatsOverlayWindow : Window
{
    private const int GwlExStyle = -20;
    private const int WsExTransparent = 0x00000020;
    private const int WsExToolWindow = 0x00000080;
    private const int WsExNoActivate = 0x08000000;
    private const int WmNcLButtonDown = 0x00A1;
    private const int HtCaption = 2;
    private const int HtBottomRight = 17;

    private readonly DispatcherTimer _positionTimer;
    private int? _preferredProcessId;
    private Point? _manualOffset;
    private Size? _manualSize;
    private bool _manualPlacementIsRelative;
    private bool _moveMode;
    private bool _isUserEditingPlacement;
    private Rect? _lastRelativePlacement;
    private int _missedGameWindowScans;
    private readonly string _component;
    private readonly string _positionFile;
    private GameOverlayVisualStyle _visualStyle;
    private ObsOverlaySnapshot _lastSnapshot = ObsOverlaySnapshot.Empty;

    internal GameNetworkStatsOverlayWindow(
        string component = "all",
        GameOverlayVisualStyle visualStyle = GameOverlayVisualStyle.Classic)
    {
        _component = component;
        _visualStyle = visualStyle;
        _positionFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HaloMCCToolbox",
            $"{component}-overlay-position.txt");
        InitializeComponent();
        (Width, Height, MinWidth, MinHeight) = component switch
        {
            "network" => (430, 132, 360, 112),
            "wait" => (360, 112, 300, 96),
            "session" => (920, 230, 620, 215),
            _ => (1280, 170, 520, 132)
        };
        var placement = LoadManualPlacement();
        _manualOffset = placement.Offset;
        _manualSize = placement.Size;
        _manualPlacementIsRelative = placement.IsRelative;
        if (_manualSize.HasValue && !_manualPlacementIsRelative)
        {
            Width = Math.Max(MinWidth, _manualSize.Value.Width);
            Height = Math.Max(MinHeight, _manualSize.Value.Height);
        }

        _positionTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(500)
        };
        _positionTimer.Tick += PositionTimer_Tick;
        ConfigureComponentPanel();
        ConfigureVisualStyle();
        RenderSnapshot(_lastSnapshot);
    }

    internal event EventHandler<Rect>? RelativePlacementChanged;

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        ApplyWindowInteractionMode();
        _positionTimer.Start();
        FollowGameWindow();
    }

    protected override void OnClosed(EventArgs e)
    {
        _positionTimer.Stop();
        _positionTimer.Tick -= PositionTimer_Tick;
        RelativePlacementChanged = null;
        base.OnClosed(e);
    }

    private void PositionTimer_Tick(object? sender, EventArgs e)
    {
        FollowGameWindow();
        RefreshTimeSensitiveDisplay();
    }

    public void UpdateStats(NetworkStatsSnapshot snapshot)
    {
        // MainWindow publishes a unified snapshot immediately after this update.
    }

    public void UpdateTrafficStats(NetworkTrafficSnapshot snapshot)
    {
        // MainWindow publishes a unified snapshot immediately after this update.
    }

    internal void UpdateSessionStats(ObsOverlaySnapshot snapshot)
    {
        _lastSnapshot = snapshot;
        RenderSnapshot(snapshot);
    }

    public void ClearStats()
    {
        _lastSnapshot = ObsOverlaySnapshot.Empty;
        RenderSnapshot(_lastSnapshot);
    }

    public void UpdateServer(GameServerInfo? serverInfo)
    {
        // MainWindow publishes a unified snapshot immediately after this update.
    }

    public void SetPreferredProcessId(int? processId)
    {
        _preferredProcessId = processId;
        FollowGameWindow();
    }

    public void SetMoveMode(bool enabled)
    {
        _moveMode = enabled;
        OverlayRoot.IsHitTestVisible = enabled;
        DragSurface.IsHitTestVisible = enabled;
        DragSurface.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        Cursor = enabled ? Cursors.SizeAll : Cursors.Arrow;
        ResizeMode = enabled ? ResizeMode.CanResize : ResizeMode.NoResize;
        ResizeThumb.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        EditToolbar.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        EditToolbarLabel.Text = $"{_component.ToUpperInvariant()}  ·  MOVE / RESIZE";
        StyleOptionsPanel.Visibility =
            enabled && (_component == "network" || _component == "session")
                ? Visibility.Visible
                : Visibility.Collapsed;
        OverlayRoot.Background = enabled ? Brush("#66081018") : Brushes.Transparent;
        OverlayRoot.BorderBrush = enabled ? Brush("#CC00C8FF") : Brushes.Transparent;
        ConfigureVisualStyle();
        ApplyWindowInteractionMode();
    }

    internal void SetVisualStyle(GameOverlayVisualStyle visualStyle)
    {
        if (_visualStyle == visualStyle)
        {
            ConfigureVisualStyle();
            return;
        }

        _visualStyle = visualStyle;
        ConfigureVisualStyle();
        RenderSnapshot(_lastSnapshot);
    }

    private void ConfigureComponentPanel()
    {
        NetworkPanel.Visibility = _component == "network" ? Visibility.Visible : Visibility.Collapsed;
        WaitPanel.Visibility = _component == "wait" ? Visibility.Visible : Visibility.Collapsed;
        SessionPanel.Visibility = _component == "session" ? Visibility.Visible : Visibility.Collapsed;
        RecapPanel.Visibility = Visibility.Collapsed;
    }

    private void ConfigureVisualStyle()
    {
        bool classic = _visualStyle == GameOverlayVisualStyle.Classic;
        NetworkClassicLayout.Visibility = classic ? Visibility.Visible : Visibility.Collapsed;
        NetworkModernLayout.Visibility = classic ? Visibility.Collapsed : Visibility.Visible;
        SessionClassicLayout.Visibility = classic ? Visibility.Visible : Visibility.Collapsed;
        SessionModernLayout.Visibility = classic ? Visibility.Collapsed : Visibility.Visible;
        ClassicStyleButton.IsChecked = classic;
        ModernStyleButton.IsChecked = !classic;
    }

    private void ClassicStyleButton_Click(object sender, RoutedEventArgs e)
    {
        SetVisualStyle(GameOverlayVisualStyle.Classic);
        App.SaveGameOverlayVisualStyle(_component, GameOverlayVisualStyle.Classic);
    }

    private void ModernStyleButton_Click(object sender, RoutedEventArgs e)
    {
        SetVisualStyle(GameOverlayVisualStyle.Modern);
        App.SaveGameOverlayVisualStyle(_component, GameOverlayVisualStyle.Modern);
    }

    private void RenderSnapshot(ObsOverlaySnapshot snapshot)
    {
        switch (_component)
        {
            case "network":
                RenderNetwork(snapshot);
                break;
            case "wait":
                RenderWait(snapshot);
                break;
            case "session":
                RenderSession(snapshot);
                break;
        }
    }

    private void RenderNetwork(ObsOverlaySnapshot snapshot)
    {
        NetworkPanel.Visibility = Visibility.Visible;
        string serverText = string.IsNullOrWhiteSpace(snapshot.ServerLabel)
            ? "SERVER: --"
            : snapshot.ServerLabel;
        NetworkServerText.Text = serverText;
        NetworkModernServerText.Text = serverText;

        bool hasNetwork = snapshot.RttMs.HasValue ||
            snapshot.UploadKilobytesPerSecond.HasValue ||
            snapshot.DownloadKilobytesPerSecond.HasValue;
        NetworkPingText.Text = snapshot.RttMs.HasValue ? $"Ping: {snapshot.RttMs.Value} ms" : "";
        NetworkJitterText.Text = snapshot.JitterMs.HasValue ? $"Jitter: {snapshot.JitterMs.Value:0.0} ms" : "";
        NetworkLossText.Text = hasNetwork ? $"Loss: {snapshot.PacketLossPercent:0}%" : "";
        NetworkModernPingText.Text = snapshot.RttMs.HasValue ? $"PING {snapshot.RttMs.Value} ms" : "";
        NetworkModernJitterText.Text = snapshot.JitterMs.HasValue ? $"JITTER {snapshot.JitterMs.Value:0.0} ms" : "";
        NetworkModernLossText.Text = hasNetwork ? $"LOSS {snapshot.PacketLossPercent:0}%" : "";
        var jitterBrush = snapshot.JitterMs switch
        {
            < 10 => Brush("#35E08D"),
            < 20 => Brush("#00C8FF"),
            null => Brush("#E7F5FF"),
            _ => Brush("#FF5D73")
        };
        NetworkJitterText.Foreground = jitterBrush;
        NetworkModernJitterText.Foreground = jitterBrush;

        NetworkUpText.Text = snapshot.UploadKilobytesPerSecond.HasValue
            ? FormatRate(snapshot.UploadKilobytesPerSecond.Value)
            : "";
        NetworkDownText.Text = snapshot.DownloadKilobytesPerSecond.HasValue
            ? FormatRate(snapshot.DownloadKilobytesPerSecond.Value)
            : "";
        NetworkUpPacketsText.Text = snapshot.UploadPacketsPerSecond.HasValue
            ? $"{snapshot.UploadPacketsPerSecond.Value:0} packets/s"
            : "";
        NetworkDownPacketsText.Text = snapshot.DownloadPacketsPerSecond.HasValue
            ? $"{snapshot.DownloadPacketsPerSecond.Value:0} packets/s"
            : "";
        NetworkModernUpText.Text = snapshot.UploadKilobytesPerSecond.HasValue
            ? $"UP   {FormatRate(snapshot.UploadKilobytesPerSecond.Value)}  ·  {snapshot.UploadPacketsPerSecond ?? 0:0} packets/s"
            : "";
        NetworkModernDownText.Text = snapshot.DownloadKilobytesPerSecond.HasValue
            ? $"DOWN {FormatRate(snapshot.DownloadKilobytesPerSecond.Value)}  ·  {snapshot.DownloadPacketsPerSecond ?? 0:0} packets/s"
            : "";

        var history = snapshot.RttHistoryMs.TakeLast(60).ToArray();
        var classicPoints = new PointCollection();
        var modernPoints = new PointCollection();
        if (history.Length > 0)
        {
            int historyMax = history.Where(x => x.HasValue).Select(x => x!.Value).DefaultIfEmpty(1).Max();
            int classicMax = Math.Max(80, historyMax);
            int modernMax = Math.Max(1, historyMax);
            for (int i = 0; i < history.Length; i++)
            {
                if (!history[i].HasValue)
                    continue;
                double classicX = history.Length == 1 ? 352 : i * 352.0 / (history.Length - 1);
                double classicY = 27 - Math.Clamp(history[i]!.Value / (double)classicMax, 0, 1) * 23;
                classicPoints.Add(new Point(classicX, classicY));

                double modernX = history.Length == 1 ? 240 : i * 240.0 / (history.Length - 1);
                double modernY = 52 - Math.Clamp(history[i]!.Value / (double)modernMax, 0, 1) * 48;
                modernPoints.Add(new Point(modernX, modernY));
            }
        }
        NetworkGraphGlow.Points = classicPoints;
        NetworkGraphLine.Points = classicPoints;
        NetworkModernGraphGlow.Points = modernPoints;
        NetworkModernGraphLine.Points = modernPoints;
    }

    private void RenderWait(ObsOverlaySnapshot snapshot)
    {
        bool visible = snapshot.ShowMatchmakingWait &&
            snapshot.MatchmakingWaitSeconds.HasValue &&
            (!snapshot.MatchmakingExpiresAtUtc.HasValue ||
             DateTimeOffset.UtcNow < snapshot.MatchmakingExpiresAtUtc.Value);
        WaitPanel.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        RefreshTimeSensitiveDisplay();
    }

    private void RenderSession(ObsOverlaySnapshot snapshot)
    {
        SessionWinsText.Text = $"{snapshot.Wins}W";
        SessionLossesText.Text = $"{snapshot.Losses}L";
        SessionGamesText.Text = $"{snapshot.GamesPlayed} game{(snapshot.GamesPlayed == 1 ? "" : "s")}";
        SessionWinRateText.Text = snapshot.GamesPlayed > 0
            ? $"{Math.Round(snapshot.Wins * 100.0 / snapshot.GamesPlayed):0}%"
            : "0%";
        SessionKdText.Text = snapshot.SessionKd;
        SessionKillsText.Text = $"{snapshot.Kills} Kills";
        SessionDeathsText.Text = $"{snapshot.Deaths} Deaths";
        SessionBestSpreeText.Text = $"Best Spree {snapshot.BestSpree}";
        SessionModernWinsText.Text = $"{snapshot.Wins}W";
        SessionModernLossesText.Text = $"{snapshot.Losses}L";
        SessionModernGamesText.Text = $"{snapshot.GamesPlayed} game{(snapshot.GamesPlayed == 1 ? "" : "s")}";
        SessionModernKdText.Text = $"{snapshot.SessionKd} K/D";
        SessionModernKillsDeathsText.Text = $"{snapshot.Kills} Kills  ·  {snapshot.Deaths} Deaths";
        SessionModernBestSpreeText.Text = $"Best Spree {snapshot.BestSpree}";

        SessionMedalsPanel.Children.Clear();
        SessionModernMedalsPanel.Children.Clear();
        PopulateSessionMedals(
            _visualStyle == GameOverlayVisualStyle.Classic
                ? SessionMedalsPanel
                : SessionModernMedalsPanel,
            snapshot);
        RefreshTimeSensitiveDisplay();
    }

    private static void PopulateSessionMedals(Panel panel, ObsOverlaySnapshot snapshot)
    {
        AddMedalCard(panel, "Double Kill", "double-kill.png", snapshot.DoubleKills, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Triple Kill", "triple-kill.png", snapshot.TripleKills, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Overkill", "overkill.png", snapshot.Overkills, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Killtacular", "killtacular.png", snapshot.Killtaculars, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Killtrocity", "killtrocity.png", snapshot.Killtrocities, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Killimanjaro", "killimanjaro.png", snapshot.Killimanjaros, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Killtastrophe", "killtastrophe.png", snapshot.Killtastrophes, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Killpocalypse", "killpocalypse.png", snapshot.Killpocalypses, scaleToAvailableSpace: true);
        AddMedalCard(panel, "Killionaire", "killionaire.png", snapshot.Killionaires, scaleToAvailableSpace: true);
    }

    private void RefreshTimeSensitiveDisplay()
    {
        var snapshot = _lastSnapshot;
        if (_component == "wait" && WaitPanel.Visibility == Visibility.Visible)
        {
            int estimate = Math.Max(0, snapshot.MatchmakingWaitSeconds ?? 0);
            WaitValueText.Text = $"EST. WAIT {FormatDuration(estimate)}";
            WaitValueText.Foreground = estimate >= 90
                ? Brush("#FF5D73")
                : estimate >= 30 ? Brush("#FFD166") : Brush("#00C8FF");

            int elapsed = snapshot.MatchmakingStartedAtUtc.HasValue
                ? Math.Max(0, (int)(DateTimeOffset.UtcNow - snapshot.MatchmakingStartedAtUtc.Value).TotalSeconds)
                : 0;
            WaitDetailText.Text = snapshot.MatchmakingPopulation.HasValue
                ? $"{snapshot.MatchmakingPopulation.Value} player{(snapshot.MatchmakingPopulation.Value == 1 ? "" : "s")} searching " +
                  $"{DefaultText(snapshot.MatchmakingPlaylistName, "this playlist")} across " +
                  DefaultText(snapshot.MatchmakingSearchScope, "all gametypes")
                : $"elapsed {elapsed / 60}:{elapsed % 60:00}";
        }

        if (_component != "session")
            return;

        var recap = snapshot.PostGameRecap;
        bool showRecap = recap is not null && DateTimeOffset.UtcNow < recap.ExpiresAtUtc;
        SessionPanel.Visibility = showRecap ? Visibility.Collapsed : Visibility.Visible;
        RecapPanel.Visibility = showRecap ? Visibility.Visible : Visibility.Collapsed;
        if (!showRecap || recap is null)
            return;

        RecapResultText.Text = recap.Won ? "VICTORY" : "DEFEAT";
        RecapResultText.Foreground = recap.Won ? Brush("#35E08D") : Brush("#FF5D73");
        RecapGameText.Text = $"{recap.Kills} KILLS   {recap.Deaths} DEATHS";
        RecapGameKdText.Text = $"{recap.GameKd} K/D";
        RecapKdText.Text = $"{recap.PreviousSessionKd} → {recap.SessionKd}";
        bool hasFeaturedMedal =
            !string.IsNullOrWhiteSpace(recap.FeaturedMedal) &&
            recap.FeaturedMedal != "—";
        RecapFeatureText.Text = !hasFeaturedMedal
            ? "SESSION UPDATE"
            : recap.FeaturedMedal.ToUpperInvariant();
        RecapFeatureDetailText.Text = hasFeaturedMedal
            ? "HIGHEST MULTIKILL THIS GAME"
            : "GAME ADDED TO SESSION";
        RecapFeatureIcon.Source = LoadMedalImage(MedalFile(recap.FeaturedMedal));
        RecapBestText.Text = recap.IsNewBestSpree
            ? $"NEW SESSION BEST  ·  {recap.BestSpree} KILL SPREE"
            : $"{recap.BestSpree} KILL SPREE";

        RecapDeltaPanel.Children.Clear();
        var medalDeltas = recap.MedalDeltas.Take(4).ToList();
        RecapDeltaPanel.Rows = medalDeltas.Count > 2 ? 2 : 1;
        RecapDeltaPanel.Columns = 2;
        foreach (var delta in medalDeltas)
        {
            AddRecapDeltaCard(RecapDeltaPanel, delta);
        }
        if (medalDeltas.Count == 0)
        {
            RecapDeltaPanel.Children.Add(new TextBlock
            {
                Text = "NO MULTIKILLS THIS GAME",
                Foreground = Brush("#8EA4B8"),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 9,
                VerticalAlignment = VerticalAlignment.Center
            });
        }

        double total = Math.Max(1, (recap.ExpiresAtUtc - recap.CapturedAtUtc).TotalMilliseconds);
        double remaining = Math.Clamp((recap.ExpiresAtUtc - DateTimeOffset.UtcNow).TotalMilliseconds / total, 0, 1);
        RecapProgressScale.ScaleX = remaining;
    }

    private static void AddRecapDeltaCard(Panel parent, ObsMedalDelta delta)
    {
        var card = new Grid
        {
            Margin = new Thickness(0, 3, 10, 3),
            VerticalAlignment = VerticalAlignment.Center
        };
        card.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(28) });
        card.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
        card.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        card.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var icon = new Image
        {
            Width = 28,
            Height = 28,
            Stretch = Stretch.Uniform,
            Source = LoadMedalImage(MedalFile(delta.Name))
        };
        Grid.SetColumn(icon, 0);
        card.Children.Add(icon);

        var name = new TextBlock
        {
            Text = delta.Name.ToUpperInvariant(),
            Foreground = Brush("#8EA4B8"),
            FontFamily = new FontFamily("Consolas"),
            FontSize = 8,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis
        };
        Grid.SetColumn(name, 2);
        card.Children.Add(name);

        var count = new TextBlock
        {
            Text = $"{delta.PreviousCount} → {delta.NewCount}",
            Foreground = Brush("#39FF14"),
            FontFamily = new FontFamily("Consolas"),
            FontSize = 12,
            FontWeight = FontWeights.Bold,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 0, 0)
        };
        Grid.SetColumn(count, 3);
        card.Children.Add(count);
        parent.Children.Add(card);
    }

    private static ImageSource? LoadMedalImage(string fileName)
    {
        try
        {
            var image = new BitmapImage();
            image.BeginInit();
            image.UriSource = new Uri(
                $"pack://application:,,,/Resources/Medals/{fileName}",
                UriKind.Absolute);
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.EndInit();
            image.Freeze();
            return image;
        }
        catch
        {
            return null;
        }
    }

    private static void AddMedalCard(
        Panel parent,
        string name,
        string fileName,
        int count,
        string? detail = null,
        bool scaleToAvailableSpace = false)
    {
        var panel = new StackPanel
        {
            Width = scaleToAvailableSpace ? 80 : 68,
            Margin = scaleToAvailableSpace ? new Thickness(2) : new Thickness(4),
            HorizontalAlignment = HorizontalAlignment.Center,
            Opacity = count > 0 ? 1 : 0.34
        };
        try
        {
            panel.Children.Add(new Image
            {
                Width = 34,
                Height = 34,
                Stretch = Stretch.Uniform,
                Source = new BitmapImage(new Uri(
                    $"pack://application:,,,/Resources/Medals/{fileName}",
                    UriKind.Absolute))
            });
        }
        catch
        {
            // A missing optional medal image must not disrupt the overlay.
        }
        panel.Children.Add(new TextBlock
        {
            Text = name.ToUpperInvariant(),
            Foreground = Brush("#8EA4B8"),
            FontFamily = new FontFamily("Consolas"),
            FontSize = 8,
            TextAlignment = TextAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis
        });
        panel.Children.Add(new TextBlock
        {
            Text = detail ?? $"×{count}",
            Foreground = Brush("#00C8FF"),
            FontFamily = new FontFamily("Consolas"),
            FontSize = scaleToAvailableSpace ? 12 : 10,
            FontWeight = FontWeights.Bold,
            TextAlignment = TextAlignment.Center
        });
        if (scaleToAvailableSpace)
        {
            parent.Children.Add(new Viewbox
            {
                Stretch = Stretch.Uniform,
                StretchDirection = StretchDirection.Both,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Child = panel
            });
        }
        else
        {
            parent.Children.Add(panel);
        }
    }

    private static string MedalFile(string name) => name.ToLowerInvariant() switch
    {
        "double kill" => "double-kill.png",
        "triple kill" => "triple-kill.png",
        "overkill" => "overkill.png",
        "killtacular" => "killtacular.png",
        "killtrocity" => "killtrocity.png",
        "killimanjaro" => "killimanjaro.png",
        "killtastrophe" => "killtastrophe.png",
        "killpocalypse" => "killpocalypse.png",
        "killionaire" => "killionaire.png",
        _ => "double-kill.png"
    };

    private static string FormatRate(double kilobytesPerSecond) =>
        kilobytesPerSecond >= 1024
            ? $"{kilobytesPerSecond / 1024:0.0} MB/s"
            : $"{kilobytesPerSecond:0} KB/s";

    private static string FormatDuration(int seconds) =>
        seconds < 60 ? $"~{seconds} SEC" :
        seconds < 120 ? "~1 MIN" :
        $"~{Math.Ceiling(seconds / 60.0):0} MIN";

    private static string DefaultText(string value, string fallback) =>
        string.IsNullOrWhiteSpace(value) ? fallback : value;

    private void ApplyWindowInteractionMode()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero)
            return;

        int exStyle = GetWindowLong(hwnd, GwlExStyle);
        exStyle |= WsExToolWindow;

        if (_moveMode)
            exStyle &= ~(WsExTransparent | WsExNoActivate);
        else
            exStyle |= WsExTransparent | WsExNoActivate;

        SetWindowLong(hwnd, GwlExStyle, exStyle);
    }

    private void FollowGameWindow()
    {
        if (_isUserEditingPlacement)
            return;

        var hwnd = FindMccWindow(_preferredProcessId);
        if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out var rect))
        {
            // MCC can briefly fail window enumeration while changing menus,
            // focus, or presentation state. Preserve the last valid overlay
            // placement so a single missed scan cannot make it flash.
            _missedGameWindowScans++;
            if (_missedGameWindowScans >= 8 && Visibility != Visibility.Collapsed)
                Visibility = Visibility.Collapsed;
            return;
        }

        _missedGameWindowScans = 0;

        var dipRect = ToDipRect(rect);
        ApplyRelativeManualSize(dipRect);
        var size = CoerceOverlaySizeToGameRect(dipRect);
        double width = size.Width;
        double height = size.Height;
        double maxLeft = Math.Max(dipRect.Left, dipRect.Right - width);
        double maxTop = Math.Max(dipRect.Top, dipRect.Bottom - height);

        if (_manualOffset.HasValue)
        {
            double offsetX = _manualPlacementIsRelative
                ? _manualOffset.Value.X * dipRect.Width
                : _manualOffset.Value.X;
            double offsetY = _manualPlacementIsRelative
                ? _manualOffset.Value.Y * dipRect.Height
                : _manualOffset.Value.Y;
            SetPositionIfChanged(
                Math.Clamp(dipRect.Left + offsetX, dipRect.Left, maxLeft),
                Math.Clamp(dipRect.Top + offsetY, dipRect.Top, maxTop));
        }
        else
        {
            const double margin = 22;
            const double hudOffsetY = 122;

            SetPositionIfChanged(
                Math.Clamp(dipRect.Right - width - margin, dipRect.Left, maxLeft),
                Math.Clamp(Math.Min(
                    dipRect.Bottom - height - margin,
                    dipRect.Top + hudOffsetY), dipRect.Top, maxTop));
        }

        // Reassigning Visible on every polling tick can alter z-order between
        // the three topmost component windows, making overlapping transparent
        // surfaces flash. Only transition when state changes.
        if (Visibility != Visibility.Visible)
            Visibility = Visibility.Visible;
        PublishRelativePlacement(dipRect, width, height);
    }

    private void SetPositionIfChanged(double left, double top)
    {
        if (double.IsNaN(Left) || Math.Abs(Left - left) >= 0.25)
            Left = left;
        if (double.IsNaN(Top) || Math.Abs(Top - top) >= 0.25)
            Top = top;
    }

    private void DragSurface_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!_moveMode || e.LeftButton != MouseButtonState.Pressed)
            return;

        BeginPlacementEdit();
        try
        {
            ReleaseCapture();
            SendMessage(new WindowInteropHelper(this).Handle, WmNcLButtonDown, HtCaption, 0);
            SaveCurrentManualPlacement();
        }
        finally
        {
            EndPlacementEdit();
            e.Handled = true;
        }
    }

    private void ResizeThumb_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!_moveMode || e.LeftButton != MouseButtonState.Pressed)
            return;

        BeginPlacementEdit();
        try
        {
            ReleaseCapture();
            SendMessage(new WindowInteropHelper(this).Handle, WmNcLButtonDown, HtBottomRight, 0);
            SaveCurrentManualPlacement();
        }
        finally
        {
            EndPlacementEdit();
            e.Handled = true;
        }
    }

    private void BeginPlacementEdit()
    {
        _isUserEditingPlacement = true;
        if (_positionTimer.IsEnabled)
            _positionTimer.Stop();
    }

    private void EndPlacementEdit()
    {
        _isUserEditingPlacement = false;
        FollowGameWindow();
        if (!_positionTimer.IsEnabled)
            _positionTimer.Start();
    }

    private void SaveCurrentManualPlacement()
    {
        var hwnd = FindMccWindow(_preferredProcessId);
        if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out var rect))
            return;

        var dipRect = ToDipRect(rect);
        var size = CoerceOverlaySizeToGameRect(dipRect);
        double width = size.Width;
        double height = size.Height;
        double maxLeft = Math.Max(dipRect.Left, dipRect.Right - width);
        double maxTop = Math.Max(dipRect.Top, dipRect.Bottom - height);
        double x = Math.Clamp(Left, dipRect.Left, maxLeft) - dipRect.Left;
        double y = Math.Clamp(Top, dipRect.Top, maxTop) - dipRect.Top;
        _manualOffset = new Point(x / dipRect.Width, y / dipRect.Height);
        _manualSize = new Size(width / dipRect.Width, height / dipRect.Height);
        _manualPlacementIsRelative = true;
        PublishRelativePlacement(dipRect, width, height);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_positionFile)!);
            File.WriteAllText(
                _positionFile,
                string.Format(
                    CultureInfo.InvariantCulture,
                    "relative,{0:0.######},{1:0.######},{2:0.######},{3:0.######}",
                    _manualOffset.Value.X,
                    _manualOffset.Value.Y,
                    _manualSize.Value.Width,
                    _manualSize.Value.Height));
        }
        catch
        {
            // Placement persistence is best-effort; dragging/resizing should still work.
        }
    }

    private void PublishRelativePlacement(Rect gameRect, double width, double height)
    {
        if (gameRect.Width <= 0 || gameRect.Height <= 0)
            return;

        var relative = new Rect(
            Math.Clamp((Left - gameRect.Left) / gameRect.Width, 0, 1),
            Math.Clamp((Top - gameRect.Top) / gameRect.Height, 0, 1),
            Math.Clamp(width / gameRect.Width, 0, 1),
            Math.Clamp(height / gameRect.Height, 0, 1));

        if (_lastRelativePlacement.HasValue &&
            Math.Abs(_lastRelativePlacement.Value.X - relative.X) < 0.0005 &&
            Math.Abs(_lastRelativePlacement.Value.Y - relative.Y) < 0.0005 &&
            Math.Abs(_lastRelativePlacement.Value.Width - relative.Width) < 0.0005 &&
            Math.Abs(_lastRelativePlacement.Value.Height - relative.Height) < 0.0005)
        {
            return;
        }

        _lastRelativePlacement = relative;
        RelativePlacementChanged?.Invoke(this, relative);
    }

    private Size CoerceOverlaySizeToGameRect(Rect gameRect)
    {
        double width = ActualWidth > 0 ? ActualWidth : Width;
        double height = ActualHeight > 0 ? ActualHeight : Height;

        double maxWidth = Math.Max(MinWidth, gameRect.Width);
        double maxHeight = Math.Max(MinHeight, gameRect.Height);
        double coercedWidth = Math.Clamp(width, MinWidth, maxWidth);
        double coercedHeight = Math.Clamp(height, MinHeight, maxHeight);

        if (Math.Abs(coercedWidth - Width) > 0.5)
            Width = coercedWidth;
        if (Math.Abs(coercedHeight - Height) > 0.5)
            Height = coercedHeight;

        return new Size(coercedWidth, coercedHeight);
    }

    private void ApplyRelativeManualSize(Rect gameRect)
    {
        if (!_manualPlacementIsRelative || !_manualSize.HasValue)
            return;

        Width = Math.Max(MinWidth, _manualSize.Value.Width * gameRect.Width);
        Height = Math.Max(MinHeight, _manualSize.Value.Height * gameRect.Height);
    }

    private OverlayPlacement LoadManualPlacement()
    {
        try
        {
            if (!File.Exists(_positionFile))
                return new OverlayPlacement(null, null, IsRelative: false);

            var parts = File.ReadAllText(_positionFile).Split(',');
            if (parts.Length >= 5 &&
                string.Equals(parts[0], "relative", StringComparison.OrdinalIgnoreCase) &&
                double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double relativeX) &&
                double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double relativeY) &&
                double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double relativeWidth) &&
                double.TryParse(parts[4], NumberStyles.Float, CultureInfo.InvariantCulture, out double relativeHeight))
            {
                return new OverlayPlacement(
                    new Point(relativeX, relativeY),
                    new Size(relativeWidth, relativeHeight),
                    IsRelative: true);
            }

            // Older Session Stats builds stored absolute desktop pixels. Those values
            // become invalid after a resolution, DPI, or MCC window-mode change.
            if (string.Equals(_component, "session", StringComparison.OrdinalIgnoreCase))
                return new OverlayPlacement(null, null, IsRelative: false);

            if (parts.Length == 2 &&
                double.TryParse(parts[0], out double x) &&
                double.TryParse(parts[1], out double y))
            {
                return new OverlayPlacement(new Point(x, y), null, IsRelative: false);
            }

            if (parts.Length >= 4 &&
                double.TryParse(parts[0], out x) &&
                double.TryParse(parts[1], out y) &&
                double.TryParse(parts[2], out double width) &&
                double.TryParse(parts[3], out double height))
            {
                return new OverlayPlacement(new Point(x, y), new Size(width, height), IsRelative: false);
            }
        }
        catch
        {
            // Ignore malformed or inaccessible placement files.
        }

        return new OverlayPlacement(null, null, IsRelative: false);
    }

    private Rect ToDipRect(WindowRect rect)
    {
        var source = PresentationSource.FromVisual(this);
        var transform = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
        var topLeft = transform.Transform(new Point(rect.Left, rect.Top));
        var bottomRight = transform.Transform(new Point(rect.Right, rect.Bottom));
        return new Rect(topLeft, bottomRight);
    }

    private static IntPtr FindMccWindow(int? preferredProcessId)
    {
        if (preferredProcessId.HasValue)
        {
            try
            {
                var preferred = Process.GetProcessById(preferredProcessId.Value);
                if (preferred.MainWindowHandle != IntPtr.Zero &&
                    IsWindowVisible(preferred.MainWindowHandle))
                {
                    return preferred.MainWindowHandle;
                }
            }
            catch
            {
                // Fall through to top-level window enumeration.
            }

            var hwnd = FindWindowForProcessId(preferredProcessId.Value);
            if (hwnd != IntPtr.Zero)
                return hwnd;
        }

        var processes = Process.GetProcessesByName("MCC-Win64-Shipping")
            .Concat(Process.GetProcessesByName("MCC"));

        foreach (var process in processes)
        {
            try
            {
                if (process.MainWindowHandle != IntPtr.Zero)
                    return process.MainWindowHandle;
            }
            catch
            {
                // Process may exit while enumerating.
            }

            var hwnd = FindWindowForProcessId(process.Id);
            if (hwnd != IntPtr.Zero)
                return hwnd;
        }

        return IntPtr.Zero;
    }

    private static IntPtr FindWindowForProcessId(int processId)
    {
        IntPtr found = IntPtr.Zero;
        long largestArea = 0;

        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out int windowProcessId);
            if (windowProcessId != processId || !IsWindowVisible(hwnd) ||
                !GetWindowRect(hwnd, out var rect))
                return true;

            long width = Math.Max(0, rect.Right - rect.Left);
            long height = Math.Max(0, rect.Bottom - rect.Top);
            long area = width * height;
            if (area > largestArea)
            {
                largestArea = area;
                found = hwnd;
            }
            return true;
        }, IntPtr.Zero);

        return found;
    }

    private static SolidColorBrush Brush(string hex) =>
        new((Color)ColorConverter.ConvertFromString(hex));

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out WindowRect rect);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hwnd, int message, int wParam, int lParam);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    private sealed record OverlayPlacement(Point? Offset, Size? Size, bool IsRelative);

    [StructLayout(LayoutKind.Sequential)]
    private struct WindowRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

}
