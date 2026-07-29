using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;

namespace HaloToolbox;

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
    private bool _browserInitialized;
    private string _overlaySource = "";
    private Rect? _lastRelativePlacement;
    private int _missedGameWindowScans;
    private readonly string _component;
    private readonly string _positionFile;

    private static readonly string OverlayWebViewUserDataFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox",
        "OverlayWebView2",
        Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture));

    public GameNetworkStatsOverlayWindow(string component = "all")
    {
        _component = component;
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
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _positionTimer.Tick += (_, _) => FollowGameWindow();
        Loaded += OnLoadedAsync;
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
        base.OnClosed(e);
    }

    public void SetOverlaySource(string url)
    {
        if (string.Equals(_overlaySource, url, StringComparison.Ordinal))
            return;

        _overlaySource = url;
        NavigateOverlayBrowser();
    }

    public void UpdateStats(NetworkStatsSnapshot snapshot)
    {
        // The WebView overlay polls the shared local state endpoint.
    }

    public void UpdateTrafficStats(NetworkTrafficSnapshot snapshot)
    {
        // The WebView overlay polls the shared local state endpoint.
    }

    internal void UpdateSessionStats(ObsOverlaySnapshot snapshot)
    {
        // The WebView overlay polls the shared local state endpoint.
    }

    public void ClearStats()
    {
        // The WebView overlay polls the shared local state endpoint.
    }

    public void UpdateServer(GameServerInfo? serverInfo)
    {
        // The WebView overlay polls the shared local state endpoint.
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
        OverlayBrowser.IsHitTestVisible = !enabled;
        OverlayBrowser.IsEnabled = !enabled;
        DragSurface.IsHitTestVisible = enabled;
        DragSurface.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        Cursor = enabled ? Cursors.SizeAll : Cursors.Arrow;
        ResizeMode = enabled ? ResizeMode.CanResize : ResizeMode.NoResize;
        ResizeThumb.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        OverlayRoot.Background = enabled ? Brush("#66081018") : Brushes.Transparent;
        OverlayRoot.BorderBrush = enabled ? Brush("#CC00C8FF") : Brushes.Transparent;
        ApplyWindowInteractionMode();
    }

    private async void OnLoadedAsync(object sender, RoutedEventArgs e)
    {
        try
        {
            var env = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: OverlayWebViewUserDataFolder);
            OverlayBrowser.DefaultBackgroundColor = System.Drawing.Color.Transparent;
            await OverlayBrowser.EnsureCoreWebView2Async(env);
            OverlayBrowser.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            OverlayBrowser.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _browserInitialized = true;
            NavigateOverlayBrowser();
        }
        catch
        {
            // If WebView2 is unavailable, keep the transparent overlay window harmless.
        }
    }

    private void NavigateOverlayBrowser()
    {
        if (!_browserInitialized || string.IsNullOrWhiteSpace(_overlaySource))
            return;

        try
        {
            OverlayBrowser.CoreWebView2.Navigate(_overlaySource);
        }
        catch
        {
            // Navigation is best-effort; the next SetOverlaySource can retry.
        }
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
        // the three topmost WebView2 component windows, making overlapping
        // transparent surfaces flash. Only transition when state changes.
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
