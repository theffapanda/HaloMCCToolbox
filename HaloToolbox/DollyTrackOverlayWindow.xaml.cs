using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace HaloToolbox;

public sealed record DollyTrackSnapshot(
    float CameraX,
    float CameraY,
    float CameraZ,
    float Yaw,
    float Pitch,
    float Roll,
    double HorizontalFovDegrees,
    IReadOnlyList<H3DollyMarker> Markers,
    int SelectedIndex);

public partial class DollyTrackOverlayWindow : Window
{
    private const int GwlExStyle = -20;
    private const int WsExTransparent = 0x00000020;
    private const int WsExToolWindow = 0x00000080;
    private const int WsExNoActivate = 0x08000000;

    private readonly Func<DollyTrackSnapshot?> _snapshotProvider;
    private readonly DispatcherTimer _renderTimer;
    private int? _preferredProcessId;

    public DollyTrackOverlayWindow(Func<DollyTrackSnapshot?> snapshotProvider)
    {
        InitializeComponent();
        _snapshotProvider = snapshotProvider;
        _renderTimer = new DispatcherTimer(DispatcherPriority.Render)
        {
            Interval = TimeSpan.FromMilliseconds(33)
        };
        _renderTimer.Tick += (_, _) => RenderFrame();
    }

    public void SetPreferredProcessId(int? processId)
    {
        _preferredProcessId = processId;
        FollowGameWindow();
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        var hwnd = new WindowInteropHelper(this).Handle;
        var style = GetWindowLong(hwnd, GwlExStyle);
        SetWindowLong(hwnd, GwlExStyle, style | WsExTransparent | WsExToolWindow | WsExNoActivate);
        _renderTimer.Start();
        RenderFrame();
    }

    protected override void OnClosed(EventArgs e)
    {
        _renderTimer.Stop();
        base.OnClosed(e);
    }

    private void RenderFrame()
    {
        if (!FollowGameWindow())
            return;

        TrackVisual.Snapshot = _snapshotProvider();
        TrackVisual.InvalidateVisual();
    }

    private bool FollowGameWindow()
    {
        var hwnd = FindMccWindow(_preferredProcessId);
        if (hwnd == IntPtr.Zero || !GetClientRect(hwnd, out var clientRect))
        {
            Visibility = Visibility.Collapsed;
            return false;
        }

        var topLeft = new NativePoint { X = clientRect.Left, Y = clientRect.Top };
        if (!ClientToScreen(hwnd, ref topLeft))
        {
            Visibility = Visibility.Collapsed;
            return false;
        }

        var source = PresentationSource.FromVisual(this);
        var transform = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
        var dipTopLeft = transform.Transform(new Point(topLeft.X, topLeft.Y));
        var dipBottomRight = transform.Transform(new Point(
            topLeft.X + clientRect.Right - clientRect.Left,
            topLeft.Y + clientRect.Bottom - clientRect.Top));

        Left = dipTopLeft.X;
        Top = dipTopLeft.Y;
        Width = Math.Max(1, dipBottomRight.X - dipTopLeft.X);
        Height = Math.Max(1, dipBottomRight.Y - dipTopLeft.Y);
        Visibility = Visibility.Visible;
        return true;
    }

    private static IntPtr FindMccWindow(int? preferredProcessId)
    {
        if (preferredProcessId.HasValue)
        {
            var preferred = FindWindowForProcessId(preferredProcessId.Value);
            if (preferred != IntPtr.Zero)
                return preferred;
        }

        foreach (var process in Process.GetProcessesByName("MCC-Win64-Shipping").Concat(Process.GetProcessesByName("MCC")))
        {
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
            GetWindowThreadProcessId(hwnd, out var id);
            if (id != processId || !IsWindowVisible(hwnd))
                return true;
            found = hwnd;
            return false;
        }, IntPtr.Zero);
        return found;
    }

    [DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr hwnd, out NativeRect rect);

    [DllImport("user32.dll")]
    private static extern bool ClientToScreen(IntPtr hwnd, ref NativePoint point);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hwnd, int index);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hwnd, int index, int value);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int processId);

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr parameter);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint { public int X, Y; }
}

public sealed class DollyTrackVisual : FrameworkElement
{
    private static readonly Pen PathPen = MakePen(Color.FromArgb(220, 0, 210, 255), 3);
    private static readonly Pen HiddenPathPen = MakePen(Color.FromArgb(80, 0, 210, 255), 1.5);
    private static readonly Pen MarkerPen = MakePen(Color.FromArgb(255, 255, 255, 255), 2);
    private static readonly Pen SelectedPen = MakePen(Color.FromArgb(255, 255, 190, 0), 3);
    private static readonly Brush MarkerBrush = new SolidColorBrush(Color.FromArgb(185, 0, 120, 190));
    private static readonly Brush SelectedBrush = new SolidColorBrush(Color.FromArgb(220, 255, 120, 0));
    private static readonly Typeface LabelTypeface = new("Segoe UI Semibold");

    public DollyTrackSnapshot? Snapshot { get; set; }

    protected override void OnRender(DrawingContext dc)
    {
        base.OnRender(dc);
        var snapshot = Snapshot;
        if (snapshot is null || snapshot.Markers.Count == 0 || ActualWidth < 2 || ActualHeight < 2)
            return;

        var camera = new Vector3(snapshot.CameraX, snapshot.CameraY, snapshot.CameraZ);

        var samples = SamplePath(snapshot.Markers, 20);
        Point? previous = null;
        foreach (var sample in samples)
        {
            if (TryProject(sample, camera, snapshot.Yaw, snapshot.Pitch, snapshot.Roll, snapshot.HorizontalFovDegrees, out var point))
            {
                if (previous.HasValue)
                    dc.DrawLine(PathPen, previous.Value, point);
                previous = point;
            }
            else
            {
                previous = null;
            }
        }

        for (var i = 0; i < snapshot.Markers.Count; i++)
        {
            var marker = snapshot.Markers[i];
            var position = new Vector3(marker.X, marker.Y, marker.Z);
            if (!TryProject(position, camera, snapshot.Yaw, snapshot.Pitch, snapshot.Roll, snapshot.HorizontalFovDegrees, out var point))
                continue;

            var selected = i == snapshot.SelectedIndex;
            var distance = (position - camera).Length;
            var radius = Math.Clamp(13.0 - distance * 0.035, 5, 13);
            dc.DrawEllipse(selected ? SelectedBrush : MarkerBrush, selected ? SelectedPen : MarkerPen, point, radius, radius);

            var label = new FormattedText(
                marker.Index.ToString(),
                System.Globalization.CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                LabelTypeface,
                11,
                Brushes.White,
                VisualTreeHelper.GetDpi(this).PixelsPerDip);
            dc.DrawText(label, new Point(point.X - label.Width / 2, point.Y - label.Height / 2));

            var (look, _, _) = CameraBasis(marker.A, marker.B, marker.C);
            if (TryProject(position + look * Math.Max(1.5, distance * 0.025), camera, snapshot.Yaw, snapshot.Pitch, snapshot.Roll, snapshot.HorizontalFovDegrees, out var lookPoint))
            {
                dc.DrawLine(selected ? SelectedPen : HiddenPathPen, point, lookPoint);
            }
        }
    }

    private bool TryProject(Vector3 world, Vector3 camera, double cameraYaw, double cameraPitch, double cameraRoll, double horizontalFovDegrees, out Point point)
    {
        point = default;
        var relative = world - camera;
        var horizontalDistance = Math.Sqrt(relative.X * relative.X + relative.Y * relative.Y);
        if (horizontalDistance < 0.0001 && Math.Abs(relative.Z) < 0.0001)
            return false;

        var targetYaw = Math.Atan2(relative.Y, relative.X);
        var targetPitch = Math.Atan2(relative.Z, horizontalDistance);
        var deltaYaw = WrapRadians(targetYaw - cameraYaw);
        var deltaPitch = targetPitch - cameraPitch;
        var horizontalFov = Math.Clamp(horizontalFovDegrees, 1, 150) * Math.PI / 180.0;
        var verticalFov = horizontalFov * 9.0 / 16.0;
        var xOffset = -deltaYaw / horizontalFov * ActualWidth;
        var yOffset = -deltaPitch / verticalFov * ActualHeight;
        var cosRoll = Math.Cos(cameraRoll);
        var sinRoll = Math.Sin(cameraRoll);
        var x = ActualWidth * 0.5 + xOffset * cosRoll + yOffset * sinRoll;
        var y = ActualHeight * 0.5 + yOffset * cosRoll - xOffset * sinRoll;
        if (x < -ActualWidth * 0.25 || x > ActualWidth * 1.25 ||
            y < -ActualHeight * 0.25 || y > ActualHeight * 1.25)
            return false;

        point = new Point(x, y);
        return true;
    }

    private static double WrapRadians(double angle)
    {
        while (angle > Math.PI) angle -= Math.PI * 2;
        while (angle < -Math.PI) angle += Math.PI * 2;
        return angle;
    }

    private static List<Vector3> SamplePath(IReadOnlyList<H3DollyMarker> markers, int samplesPerSegment)
    {
        var result = new List<Vector3>();
        if (markers.Count == 1)
        {
            result.Add(new Vector3(markers[0].X, markers[0].Y, markers[0].Z));
            return result;
        }

        for (var segment = 0; segment < markers.Count - 1; segment++)
        {
            var p0 = markers[Math.Max(0, segment - 1)];
            var p1 = markers[segment];
            var p2 = markers[segment + 1];
            var p3 = markers[Math.Min(markers.Count - 1, segment + 2)];
            for (var step = segment == 0 ? 0 : 1; step <= samplesPerSegment; step++)
            {
                var t = step / (double)samplesPerSegment;
                result.Add(markers.Count > 2 ? CatmullRom(p0, p1, p2, p3, t) : Lerp(p1, p2, t));
            }
        }
        return result;
    }

    private static Vector3 CatmullRom(H3DollyMarker p0, H3DollyMarker p1, H3DollyMarker p2, H3DollyMarker p3, double t)
    {
        var t2 = t * t;
        var t3 = t2 * t;
        return new Vector3(
            0.5 * ((2 * p1.X) + (-p0.X + p2.X) * t + (2 * p0.X - 5 * p1.X + 4 * p2.X - p3.X) * t2 + (-p0.X + 3 * p1.X - 3 * p2.X + p3.X) * t3),
            0.5 * ((2 * p1.Y) + (-p0.Y + p2.Y) * t + (2 * p0.Y - 5 * p1.Y + 4 * p2.Y - p3.Y) * t2 + (-p0.Y + 3 * p1.Y - 3 * p2.Y + p3.Y) * t3),
            0.5 * ((2 * p1.Z) + (-p0.Z + p2.Z) * t + (2 * p0.Z - 5 * p1.Z + 4 * p2.Z - p3.Z) * t2 + (-p0.Z + 3 * p1.Z - 3 * p2.Z + p3.Z) * t3));
    }

    private static Vector3 Lerp(H3DollyMarker a, H3DollyMarker b, double t) =>
        new(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t, a.Z + (b.Z - a.Z) * t);

    private static (Vector3 Forward, Vector3 Right, Vector3 Up) CameraBasis(double yaw, double pitch, double roll)
    {
        var cosYaw = Math.Cos(yaw);
        var sinYaw = Math.Sin(yaw);
        var cosPitch = Math.Cos(pitch);
        var sinPitch = Math.Sin(pitch);
        var forward = new Vector3(cosYaw * cosPitch, sinYaw * cosPitch, sinPitch);
        var levelRight = new Vector3(-sinYaw, cosYaw, 0);
        var levelUp = new Vector3(-cosYaw * sinPitch, -sinYaw * sinPitch, cosPitch);
        var cosRoll = Math.Cos(roll);
        var sinRoll = Math.Sin(roll);
        var right = levelRight * cosRoll + levelUp * sinRoll;
        var up = levelUp * cosRoll - levelRight * sinRoll;
        return (forward, right, up);
    }
    private static Pen MakePen(Color color, double thickness) { var pen = new Pen(new SolidColorBrush(color), thickness); pen.Freeze(); return pen; }

    private readonly record struct Vector3(double X, double Y, double Z)
    {
        public double LengthSquared => X * X + Y * Y + Z * Z;
        public double Length => Math.Sqrt(LengthSquared);
        public Vector3 Normalized() => LengthSquared < 0.000001 ? default : this * (1.0 / Length);
        public static Vector3 operator +(Vector3 a, Vector3 b) => new(a.X + b.X, a.Y + b.Y, a.Z + b.Z);
        public static Vector3 operator -(Vector3 a, Vector3 b) => new(a.X - b.X, a.Y - b.Y, a.Z - b.Z);
        public static Vector3 operator *(Vector3 a, double scale) => new(a.X * scale, a.Y * scale, a.Z * scale);
    }
}
