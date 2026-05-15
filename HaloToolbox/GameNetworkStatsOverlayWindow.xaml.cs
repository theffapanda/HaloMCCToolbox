using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace HaloToolbox;

public partial class GameNetworkStatsOverlayWindow : Window
{
    private const int GwlExStyle = -20;
    private const int WsExTransparent = 0x00000020;
    private const int WsExToolWindow = 0x00000080;
    private const int WsExNoActivate = 0x08000000;

    private readonly DispatcherTimer _positionTimer;
    private int? _preferredProcessId;
    private Point? _manualOffset;
    private Size? _manualSize;
    private bool _moveMode;
    private bool _isUserEditingPlacement;
    private string _serverLabel = "SERVER: --";
    private NetworkStatsSnapshot? _lastSnapshot;
    private static readonly string OverlayPositionFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox",
        "network-overlay-position.txt");

    public GameNetworkStatsOverlayWindow()
    {
        InitializeComponent();
        var placement = LoadManualPlacement();
        _manualOffset = placement.Offset;
        _manualSize = placement.Size;
        if (_manualSize.HasValue)
        {
            Width = Math.Max(MinWidth, _manualSize.Value.Width);
            Height = Math.Max(MinHeight, _manualSize.Value.Height);
        }

        _positionTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _positionTimer.Tick += (_, _) => FollowGameWindow();
        SizeChanged += (_, _) =>
        {
            if (!_isUserEditingPlacement && _lastSnapshot is not null)
                DrawGraph(_lastSnapshot);
        };
    }

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
        base.OnClosed(e);
    }

    public void UpdateStats(NetworkStatsSnapshot snapshot)
    {
        TxtNetworkServer.Text = string.IsNullOrWhiteSpace(_serverLabel)
            ? "ACTIVE SERVER"
            : _serverLabel;

        if (snapshot.RttMs.HasValue)
        {
            TxtNetworkPing.Text = $"Ping: {snapshot.RttMs.Value} ms";
            TxtNetworkPing.Foreground = GetPingBrush(snapshot.RttMs.Value);
        }
        else
        {
            TxtNetworkPing.Text = "Ping: timeout";
            TxtNetworkPing.Foreground = Brush("#FF2D55");
        }

        TxtNetworkPacketLoss.Text = $"Loss: {snapshot.PacketLossPercent:0}%";
        TxtNetworkPacketLoss.Foreground = snapshot.PacketLossPercent > 0
            ? Brush("#FF6A00")
            : Brush("#C8D8E8");

        _lastSnapshot = snapshot;
        DrawGraph(snapshot);
    }

    public void UpdateTrafficStats(NetworkTrafficSnapshot snapshot)
    {
        TxtNetworkUpload.Text = $"{FormatKilobytes(snapshot.UploadKilobytesPerSecond)}/s";
        TxtNetworkDownload.Text = $"{FormatKilobytes(snapshot.DownloadKilobytesPerSecond)}/s";
        TxtNetworkUploadPackets.Text = $"{snapshot.UploadPacketsPerSecond:0} packets/s";
        TxtNetworkDownloadPackets.Text = $"{snapshot.DownloadPacketsPerSecond:0} packets/s";
    }

    public void ClearStats()
    {
        _serverLabel = "";
        TxtNetworkServer.Text = "";
        TxtNetworkPing.Text = "";
        TxtNetworkPacketLoss.Foreground = Brush("#C8D8E8");
        TxtNetworkPacketLoss.Text = "";
        TxtNetworkUpload.Text = "";
        TxtNetworkDownload.Text = "";
        TxtNetworkUploadPackets.Text = "";
        TxtNetworkDownloadPackets.Text = "";
        _lastSnapshot = null;
        PingGraphLine.Points.Clear();
        PingGraphGlow.Points.Clear();
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
        Cursor = enabled ? Cursors.SizeAll : Cursors.Arrow;
        ResizeMode = enabled ? ResizeMode.CanResizeWithGrip : ResizeMode.NoResize;
        ResizeThumb.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        OverlayRoot.Background = enabled ? Brush("#B0081018") : Brushes.Transparent;
        OverlayRoot.BorderBrush = enabled ? Brush("#8800C8FF") : Brushes.Transparent;
        GraphBorder.Background = enabled ? Brush("#33040A10") : Brushes.Transparent;
        GraphBorder.BorderBrush = enabled ? Brush("#3300C8FF") : Brushes.Transparent;
        ApplyWindowInteractionMode();
    }

    public void UpdateServer(GameServerInfo? serverInfo)
    {
        if (serverInfo is null)
        {
            ClearStats();
            return;
        }

        var regionLabel = GameServerRegionResolver.GetRegionLabel(serverInfo);
        var port = serverInfo.Ports.FirstOrDefault()?.Num;
        _serverLabel = !string.IsNullOrWhiteSpace(regionLabel)
            ? regionLabel
            : "ACTIVE SERVER";
        TxtNetworkServer.Text = _serverLabel;
    }

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
            Visibility = Visibility.Collapsed;
            return;
        }

        var dipRect = ToDipRect(rect);
        var size = CoerceOverlaySizeToGameRect(dipRect);
        double width = size.Width;
        double height = size.Height;
        double maxLeft = Math.Max(dipRect.Left, dipRect.Right - width);
        double maxTop = Math.Max(dipRect.Top, dipRect.Bottom - height);

        if (_manualOffset.HasValue)
        {
            Left = Math.Clamp(dipRect.Left + _manualOffset.Value.X, dipRect.Left, maxLeft);
            Top = Math.Clamp(dipRect.Top + _manualOffset.Value.Y, dipRect.Top, maxTop);
        }
        else
        {
            const double margin = 22;
            const double hudOffsetY = 122;

            Left = Math.Clamp(dipRect.Right - width - margin, dipRect.Left, maxLeft);
            Top = Math.Clamp(Math.Min(
                dipRect.Bottom - height - margin,
                dipRect.Top + hudOffsetY), dipRect.Top, maxTop);
        }

        Visibility = Visibility.Visible;
    }

    private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!_moveMode || e.LeftButton != MouseButtonState.Pressed)
            return;

        try
        {
            BeginPlacementEdit();
            DragMove();
            SaveCurrentManualPlacement();
        }
        catch
        {
            // DragMove can throw if Windows cancels the mouse capture.
        }
        finally
        {
            EndPlacementEdit();
        }
    }

    private void ResizeThumb_DragStarted(object sender, System.Windows.Controls.Primitives.DragStartedEventArgs e)
    {
        BeginPlacementEdit();
    }

    private void ResizeThumb_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
    {
        if (!_moveMode)
            return;

        double maxWidth = double.PositiveInfinity;
        double maxHeight = double.PositiveInfinity;

        var hwnd = FindMccWindow(_preferredProcessId);
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out var rect))
        {
            var dipRect = ToDipRect(rect);
            maxWidth = Math.Max(MinWidth, dipRect.Width);
            maxHeight = Math.Max(MinHeight, dipRect.Height);
        }

        Width = Math.Clamp(Width + e.HorizontalChange, MinWidth, maxWidth);
        Height = Math.Clamp(Height + e.VerticalChange, MinHeight, maxHeight);
    }

    private void ResizeThumb_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
    {
        SaveCurrentManualPlacement();
        EndPlacementEdit();
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
        if (_lastSnapshot is not null)
            DrawGraph(_lastSnapshot);
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
        _manualOffset = new Point(x, y);
        _manualSize = new Size(width, height);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(OverlayPositionFile)!);
            File.WriteAllText(OverlayPositionFile, $"{x:0.###},{y:0.###},{width:0.###},{height:0.###}");
        }
        catch
        {
            // Placement persistence is best-effort; dragging/resizing should still work.
        }
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

    private static OverlayPlacement LoadManualPlacement()
    {
        try
        {
            if (!File.Exists(OverlayPositionFile))
                return new OverlayPlacement(null, null);

            var parts = File.ReadAllText(OverlayPositionFile).Split(',');
            if (parts.Length == 2 &&
                double.TryParse(parts[0], out double x) &&
                double.TryParse(parts[1], out double y))
            {
                return new OverlayPlacement(new Point(x, y), null);
            }

            if (parts.Length >= 4 &&
                double.TryParse(parts[0], out x) &&
                double.TryParse(parts[1], out y) &&
                double.TryParse(parts[2], out double width) &&
                double.TryParse(parts[3], out double height))
            {
                return new OverlayPlacement(new Point(x, y), new Size(width, height));
            }
        }
        catch
        {
            // Ignore malformed or inaccessible placement files.
        }

        return new OverlayPlacement(null, null);
    }

    private Rect ToDipRect(WindowRect rect)
    {
        var source = PresentationSource.FromVisual(this);
        var transform = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
        var topLeft = transform.Transform(new Point(rect.Left, rect.Top));
        var bottomRight = transform.Transform(new Point(rect.Right, rect.Bottom));
        return new Rect(topLeft, bottomRight);
    }

    private void DrawGraph(NetworkStatsSnapshot snapshot)
    {
        var history = snapshot.RttHistory;
        if (history.Count == 0)
        {
            PingGraphLine.Points.Clear();
            PingGraphGlow.Points.Clear();
            return;
        }

        double width = Math.Max(24, GraphCanvas.ActualWidth);
        double height = Math.Max(24, GraphCanvas.ActualHeight);
        const double minMs = 0;
        const double maxMs = 180;
        double step = history.Count <= 1 ? width : width / (history.Count - 1);
        var points = new PointCollection();

        GraphGuideTop.X2 = width;
        GraphGuideTop.Y1 = GraphGuideTop.Y2 = height / 3;
        GraphGuideBottom.X2 = width;
        GraphGuideBottom.Y1 = GraphGuideBottom.Y2 = height * 2 / 3;

        for (int i = 0; i < history.Count; i++)
        {
            long value = history[i] ?? (long)maxMs;
            double clamped = Math.Clamp(value, minMs, maxMs);
            double x = i * step;
            double y = height - ((clamped - minMs) / (maxMs - minMs) * (height - 4)) - 2;
            points.Add(new Point(x, y));
        }

        PingGraphLine.Points = points;
        PingGraphGlow.Points = points.Clone();
        PingGraphLine.Stroke = snapshot.RttMs.HasValue
            ? GetPingBrush(snapshot.RttMs.Value)
            : Brush("#FF2D55");
    }

    private static IntPtr FindMccWindow(int? preferredProcessId)
    {
        if (preferredProcessId.HasValue)
        {
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

        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out int windowProcessId);
            if (windowProcessId != processId || !IsWindowVisible(hwnd))
                return true;

            found = hwnd;
            return false;
        }, IntPtr.Zero);

        return found;
    }

    private static SolidColorBrush GetPingBrush(long rttMs)
    {
        if (rttMs < 60)
            return Brush("#39FF14");

        if (rttMs > 150)
            return Brush("#FF2D55");

        return Brush("#FFD700");
    }

    private static string FormatKilobytes(double kilobytesPerSecond)
    {
        if (kilobytesPerSecond >= 1024)
            return $"{kilobytesPerSecond / 1024.0:0.0} MB";

        return $"{kilobytesPerSecond:0.0} KB";
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

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    private sealed record OverlayPlacement(Point? Offset, Size? Size);

    [StructLayout(LayoutKind.Sequential)]
    private struct WindowRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
