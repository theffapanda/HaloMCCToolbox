using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using ShapePath = System.Windows.Shapes.Path;

namespace HaloToolbox;

public partial class VpnRegionMap : UserControl
{
    private const double WorldWidth = 1000;
    private const double WorldHeight = 500;
    private const double MinScale = 0.45;
    private const double MaxScale = 6.0;
    private const double LocalMarkerRadius = 9;
    private const double SelectedRegionMarkerRadius = 8;
    private static readonly Point LocalCoordinates = Project(40.7128, -74.0060);

    private readonly Dictionary<string, HaloVpnRegion> _regions = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Button> _buttons = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<FrameworkElement> _fixedSizeElements = [];
    private readonly Dictionary<string, Vector> _markerOffsets = new(StringComparer.OrdinalIgnoreCase)
    {
        ["East US"] = new(-10, 0),
        ["East US 2"] = new(10, 0)
    };

    private Border? _selectedLabel;
    private TextBlock? _selectedLabelText;
    private string? _selectedLabelRegion;
    private bool? _selectedLabelPlaceLeft;
    private int? _selectedLatencyMs;
    private bool _selectedLatencyPending;
    private ShapePath? _routePath;
    private ShapePath? _routeGlowPath;
    private FrameworkElement? _localMarker;
    private bool _showWorld;
    private bool _focusSelectedRegion;
    private bool _initialized;
    private bool _dragging;
    private Point _dragStart;
    private double _dragTranslateX;
    private double _dragTranslateY;

    internal event Action<string>? RegionSelected;

    public VpnRegionMap()
    {
        InitializeComponent();
        Loaded += (_, _) =>
        {
            BuildBaseMap();
            if (!_initialized)
            {
                _initialized = true;
                FitCurrentView();
            }
        };
    }

    internal string SelectedRegion { get; private set; } = "East US 2";

    internal void LoadRegions(IEnumerable<HaloVpnRegion> regions)
    {
        _regions.Clear();
        foreach (HaloVpnRegion region in regions)
            _regions[region.Name] = region;

        BuildRegionMarkers();
        SelectRegion(SelectedRegion, notify: false);
    }

    internal void SelectRegion(string regionName, bool notify = false, bool focus = false)
    {
        if (!_regions.TryGetValue(regionName, out HaloVpnRegion? region))
            return;

        bool changed = !SelectedRegion.Equals(region.Name, StringComparison.OrdinalIgnoreCase);
        SelectedRegion = region.Name;
        if (changed)
        {
            _selectedLatencyMs = null;
            _selectedLatencyPending = false;
        }
        foreach ((string name, Button button) in _buttons)
            button.Tag = name.Equals(region.Name, StringComparison.OrdinalIgnoreCase) ? "Selected" : null;

        bool viewportChanged = false;
        if (focus)
        {
            _focusSelectedRegion = true;
            FocusRegion(region, updateVisuals: false);
            viewportChanged = true;
        }
        else if (notify)
        {
            _focusSelectedRegion = false;
        }

        DrawRoute(region);
        DrawSelectedLabel(region);
        if (viewportChanged)
            UpdateFixedElementScale();

        if (notify)
            RegionSelected?.Invoke(region.Name);
    }

    internal void SetRegionLatency(string regionName, int? latencyMs, bool testing = false)
    {
        if (!SelectedRegion.Equals(regionName, StringComparison.OrdinalIgnoreCase) ||
            !_regions.TryGetValue(regionName, out HaloVpnRegion? region))
        {
            return;
        }

        if (_selectedLatencyMs == latencyMs && _selectedLatencyPending == testing)
            return;

        _selectedLatencyMs = latencyMs;
        _selectedLatencyPending = testing;
        DrawSelectedLabel(region);
    }

    internal void SetInteractionEnabled(bool enabled)
    {
        foreach (Button button in _buttons.Values)
            button.IsEnabled = enabled;
    }

    private void BuildBaseMap()
    {
        if (LandCanvas.Children.Count > 0)
            return;

        DrawGrid();
        PathGeometry? geometry = LoadWorldGeometry();
        if (geometry is not null)
        {
            var land = new ShapePath
            {
                Data = geometry,
                StrokeThickness = 0.8,
                SnapsToDevicePixels = true
            };
            land.SetResourceReference(Shape.FillProperty, "PanelBrush");
            land.SetResourceReference(Shape.StrokeProperty, "BorderBrush");
            LandCanvas.Children.Add(land);
        }

        BuildRegionMarkers();
    }

    private void DrawGrid()
    {
        GridCanvas.Children.Clear();
        for (int longitude = -150; longitude <= 150; longitude += 30)
        {
            Point top = Project(90, longitude);
            Point bottom = Project(-90, longitude);
            AddGridLine(top, bottom);
        }

        for (int latitude = -60; latitude <= 60; latitude += 30)
        {
            Point left = Project(latitude, -180);
            Point right = Project(latitude, 180);
            AddGridLine(left, right);
        }
    }

    private void AddGridLine(Point start, Point end)
    {
        var line = new Line
        {
            X1 = start.X,
            Y1 = start.Y,
            X2 = end.X,
            Y2 = end.Y,
            StrokeThickness = 0.65,
            Opacity = 0.34
        };
        line.SetResourceReference(Shape.StrokeProperty, "BorderBrush");
        GridCanvas.Children.Add(line);
    }

    private static PathGeometry? LoadWorldGeometry()
    {
        const string resourceName = "HaloToolbox.Resources.Vpn.world-countries-lowres.geojson";
        using Stream? stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
        if (stream is null)
            return null;

        using JsonDocument document = JsonDocument.Parse(stream);
        var geometry = new PathGeometry { FillRule = FillRule.EvenOdd };

        foreach (JsonElement feature in document.RootElement.GetProperty("features").EnumerateArray())
        {
            JsonElement source = feature.GetProperty("geometry");
            string? type = source.GetProperty("type").GetString();
            JsonElement coordinates = source.GetProperty("coordinates");
            if (type == "Polygon")
            {
                AppendPolygon(geometry, coordinates);
            }
            else if (type == "MultiPolygon")
            {
                foreach (JsonElement polygon in coordinates.EnumerateArray())
                    AppendPolygon(geometry, polygon);
            }
        }

        geometry.Freeze();
        return geometry;
    }

    private static void AppendPolygon(PathGeometry geometry, JsonElement rings)
    {
        foreach (JsonElement ring in rings.EnumerateArray())
        {
            JsonElement.ArrayEnumerator points = ring.EnumerateArray();
            if (!points.MoveNext())
                continue;

            JsonElement first = points.Current;
            var figure = new PathFigure
            {
                StartPoint = Project(first[1].GetDouble(), first[0].GetDouble()),
                IsClosed = true,
                IsFilled = true
            };
            var segment = new PolyLineSegment();
            while (points.MoveNext())
            {
                JsonElement point = points.Current;
                segment.Points.Add(Project(point[1].GetDouble(), point[0].GetDouble()));
            }

            figure.Segments.Add(segment);
            geometry.Figures.Add(figure);
        }
    }

    private void BuildRegionMarkers()
    {
        RegionCanvas.Children.Clear();
        _buttons.Clear();
        _fixedSizeElements.Clear();
        _selectedLabel = null;
        _selectedLabelText = null;
        _selectedLabelRegion = null;
        _selectedLabelPlaceLeft = null;

        if (_regions.Count == 0)
            return;

        _localMarker = CreateLocalMarker();
        RegionCanvas.Children.Add(_localMarker);
        _fixedSizeElements.Add(_localMarker);
        Canvas.SetLeft(_localMarker, LocalCoordinates.X - 9);
        Canvas.SetTop(_localMarker, LocalCoordinates.Y - 9);

        foreach (HaloVpnRegion region in _regions.Values)
        {
            var button = new Button
            {
                Style = (Style)FindResource("MapRegionButton"),
                ToolTip = $"{region.Name} · click to select",
                Tag = region.Name.Equals(SelectedRegion, StringComparison.OrdinalIgnoreCase) ? "Selected" : null,
                RenderTransformOrigin = new Point(0.5, 0.5)
            };
            button.Click += (_, _) => SelectRegion(region.Name, notify: true);
            PositionMarker(button, region);
            Canvas.SetZIndex(button, 5);
            RegionCanvas.Children.Add(button);
            _buttons[region.Name] = button;
            _fixedSizeElements.Add(button);
        }

        if (_regions.TryGetValue(SelectedRegion, out HaloVpnRegion? selected))
        {
            DrawRoute(selected);
            DrawSelectedLabel(selected);
        }

        UpdateFixedElementScale();
    }

    private FrameworkElement CreateLocalMarker()
    {
        var diamond = new Border
        {
            Width = 12,
            Height = 12,
            BorderThickness = new Thickness(2),
            RenderTransformOrigin = new Point(0.5, 0.5),
            RenderTransform = new RotateTransform(45),
            ToolTip = "Your PC · approximate map origin"
        };
        diamond.SetResourceReference(Border.BackgroundProperty, "TextBrush");
        diamond.SetResourceReference(Border.BorderBrushProperty, "SurfaceBrush");

        var host = new Grid
        {
            Width = 18,
            Height = 18,
            RenderTransformOrigin = new Point(0.5, 0.5),
            IsHitTestVisible = false
        };
        host.Children.Add(diamond);
        Canvas.SetZIndex(host, 6);
        return host;
    }

    private void DrawRoute(HaloVpnRegion region)
    {
        if (_routePath is null || _routeGlowPath is null)
        {
            RouteCanvas.Children.Clear();
            _routeGlowPath = new ShapePath
            {
                StrokeThickness = 9 / Math.Max(MapScale.ScaleX, 0.01),
                Opacity = 0.08,
                IsHitTestVisible = false
            };
            _routeGlowPath.SetResourceReference(Shape.StrokeProperty, "AccentBrush");
            RouteCanvas.Children.Add(_routeGlowPath);

            _routePath = new ShapePath
            {
                StrokeThickness = 2 / Math.Max(MapScale.ScaleX, 0.01),
                StrokeDashArray = [4, 3],
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round,
                IsHitTestVisible = false
            };
            _routePath.SetResourceReference(Shape.StrokeProperty, "AccentBrush");
            RouteCanvas.Children.Add(_routePath);

            if (SystemParameters.ClientAreaAnimation)
            {
                _routePath.BeginAnimation(
                    Shape.StrokeDashOffsetProperty,
                    new DoubleAnimation(0, -7, TimeSpan.FromSeconds(1.25))
                    {
                        RepeatBehavior = RepeatBehavior.Forever
                    });
            }
        }

        UpdateRouteGeometry(region);
    }

    private void UpdateRouteGeometry(HaloVpnRegion region)
    {
        if (_routePath is null || _routeGlowPath is null)
            return;

        double scale = Math.Max(MapScale.ScaleX, 0.01);
        Point destination = MarkerCenter(region, scale);
        double arc = Math.Min(70, Math.Max(8, Math.Abs(destination.X - LocalCoordinates.X) * 0.16));
        var control = new Point(
            (LocalCoordinates.X + destination.X) / 2,
            Math.Min(LocalCoordinates.Y, destination.Y) - arc);
        Point start = MoveTowards(LocalCoordinates, control, LocalMarkerRadius / scale);
        Point end = MoveTowards(destination, control, SelectedRegionMarkerRadius / scale);

        var figure = new PathFigure { StartPoint = start };
        figure.Segments.Add(new QuadraticBezierSegment(control, end, true));
        var geometry = new PathGeometry([figure]);
        _routeGlowPath.Data = geometry;
        _routePath.Data = geometry.Clone();
    }

    private static Point MoveTowards(Point start, Point target, double distance)
    {
        Vector direction = target - start;
        if (direction.LengthSquared < 0.0001)
            return start;

        direction.Normalize();
        return start + direction * distance;
    }

    private void DrawSelectedLabel(HaloVpnRegion region)
    {
        bool regionChanged = !region.Name.Equals(_selectedLabelRegion, StringComparison.OrdinalIgnoreCase);
        if (_selectedLabel is null || _selectedLabelText is null)
        {
            _selectedLabelText = new TextBlock
            {
                FontFamily = new FontFamily("Consolas"),
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center
            };
            _selectedLabelText.SetResourceReference(TextBlock.ForegroundProperty, "AccentBrush");

            _selectedLabel = new Border
            {
                Child = _selectedLabelText,
                Padding = new Thickness(7, 4, 7, 4),
                CornerRadius = new CornerRadius(3),
                BorderThickness = new Thickness(1),
                RenderTransformOrigin = new Point(0, 0),
                IsHitTestVisible = false
            };
            _selectedLabel.SetResourceReference(Border.BackgroundProperty, "PanelBrush");
            _selectedLabel.SetResourceReference(Border.BorderBrushProperty, "AccentBrush");
            Canvas.SetZIndex(_selectedLabel, 4);
            RegionCanvas.Children.Add(_selectedLabel);
            _fixedSizeElements.Add(_selectedLabel);
        }

        if (regionChanged)
        {
            _selectedLabelRegion = region.Name;
            _selectedLabelPlaceLeft = null;

            // Reserve enough width for every state so TESTING and latency updates
            // cannot resize the label or make it jump to the opposite side.
            _selectedLabel.Width = double.NaN;
            _selectedLabelText.Text = $"{region.Name.ToUpperInvariant()}  ·  TESTING…";
            _selectedLabel.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            _selectedLabel.Width = Math.Ceiling(_selectedLabel.DesiredSize.Width);
        }

        string text = _selectedLatencyMs is int latency
            ? $"{region.Name.ToUpperInvariant()}  ·  {latency} MS"
            : _selectedLatencyPending
                ? $"{region.Name.ToUpperInvariant()}  ·  TESTING…"
                : region.Name.ToUpperInvariant();
        if (!_selectedLabelText.Text.Equals(text, StringComparison.Ordinal))
            _selectedLabelText.Text = text;

        if (!regionChanged)
            return;

        _selectedLabel.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        double inverse = 1 / Math.Max(MapScale.ScaleX, 0.01);
        _selectedLabel.RenderTransform = new ScaleTransform(inverse, inverse);
        PositionSelectedLabel(region);
    }

    private Point RegionPoint(HaloVpnRegion region)
        => Project(region.Latitude, region.Longitude);

    private Point MarkerCenter(HaloVpnRegion region, double scale)
    {
        Point point = RegionPoint(region);
        Vector offset = _markerOffsets.GetValueOrDefault(region.Name);
        return point + offset / scale;
    }

    private void PositionMarker(Button button, HaloVpnRegion region)
    {
        double scale = Math.Max(MapScale.ScaleX, 0.01);
        Point center = MarkerCenter(region, scale);
        Canvas.SetLeft(button, center.X - 16);
        Canvas.SetTop(button, center.Y - 16);
    }

    private void PositionSelectedLabel(HaloVpnRegion region)
    {
        if (_selectedLabel is null)
            return;

        double scale = Math.Max(MapScale.ScaleX, 0.01);
        Point anchor = MarkerCenter(region, scale);
        double width = _selectedLabel.DesiredSize.Width / scale;
        double height = _selectedLabel.DesiredSize.Height / scale;
        double gap = 19 / scale;
        double top = anchor.Y - height / 2;
        var left = new Rect(anchor.X - width - gap, top, width, height);
        var right = new Rect(anchor.X + gap, top, width, height);

        if (_selectedLabelPlaceLeft is null)
        {
            // Prefer the side away from the incoming route, unless that would put the
            // label over another datacenter marker or outside the visible viewport.
            bool preferLeft = anchor.X <= LocalCoordinates.X;
            double leftScore = ScoreLabelPlacement(left, region, scale) + (preferLeft ? 0 : 1);
            double rightScore = ScoreLabelPlacement(right, region, scale) + (preferLeft ? 1 : 0);
            _selectedLabelPlaceLeft = leftScore <= rightScore;
        }

        Rect placement = _selectedLabelPlaceLeft.Value ? left : right;
        Canvas.SetLeft(_selectedLabel, placement.X);
        Canvas.SetTop(_selectedLabel, placement.Y);
    }

    private double ScoreLabelPlacement(Rect placement, HaloVpnRegion selected, double scale)
    {
        double score = 0;
        var markerClearance = placement;
        markerClearance.Inflate(12 / scale, 10 / scale);
        foreach (HaloVpnRegion region in _regions.Values)
        {
            if (region.Name.Equals(selected.Name, StringComparison.OrdinalIgnoreCase))
                continue;

            if (markerClearance.Contains(MarkerCenter(region, scale)))
                score += 10_000;
        }

        if (Viewport.ActualWidth > 0 && Viewport.ActualHeight > 0)
        {
            double screenLeft = placement.Left * scale + MapTranslate.X;
            double screenTop = placement.Top * scale + MapTranslate.Y;
            double screenRight = placement.Right * scale + MapTranslate.X;
            double screenBottom = placement.Bottom * scale + MapTranslate.Y;
            score += Math.Max(0, -screenLeft) + Math.Max(0, screenRight - Viewport.ActualWidth);
            score += Math.Max(0, -screenTop) + Math.Max(0, screenBottom - Viewport.ActualHeight);
        }

        return score;
    }

    private static Point Project(double latitude, double longitude) =>
        new((longitude + 180) / 360 * WorldWidth, (90 - latitude) / 180 * WorldHeight);

    private void Viewport_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (_initialized && e.NewSize.Width > 0 && e.NewSize.Height > 0)
        {
            if (_focusSelectedRegion && _regions.TryGetValue(SelectedRegion, out HaloVpnRegion? selected))
                FocusRegion(selected);
            else
                FitCurrentView();
        }
    }

    private void Viewport_MouseWheel(object sender, MouseWheelEventArgs e)
    {
        double factor = e.Delta > 0 ? 1.18 : 1 / 1.18;
        ZoomAt(factor, e.GetPosition(Viewport));
        e.Handled = true;
    }

    private void Viewport_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.OriginalSource is DependencyObject source && FindAncestor<Button>(source) is not null)
            return;

        _dragging = true;
        _dragStart = e.GetPosition(Viewport);
        _dragTranslateX = MapTranslate.X;
        _dragTranslateY = MapTranslate.Y;
        Viewport.Cursor = Cursors.SizeAll;
        Viewport.CaptureMouse();
        e.Handled = true;
    }

    private void Viewport_MouseMove(object sender, MouseEventArgs e)
    {
        if (!_dragging)
            return;

        Point current = e.GetPosition(Viewport);
        MapTranslate.X = _dragTranslateX + current.X - _dragStart.X;
        MapTranslate.Y = _dragTranslateY + current.Y - _dragStart.Y;
    }

    private void Viewport_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) => EndDrag();

    private void Viewport_MouseLeave(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed)
            EndDrag();
    }

    private void EndDrag()
    {
        if (!_dragging)
            return;
        _dragging = false;
        Viewport.ReleaseMouseCapture();
        Viewport.Cursor = Cursors.Arrow;
    }

    private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
    {
        while (current is not null)
        {
            if (current is T match)
                return match;
            current = VisualTreeHelper.GetParent(current);
        }
        return null;
    }

    private void ZoomIn_Click(object sender, RoutedEventArgs e) =>
        ZoomAt(1.25, new Point(Viewport.ActualWidth / 2, Viewport.ActualHeight / 2));

    private void ZoomOut_Click(object sender, RoutedEventArgs e) =>
        ZoomAt(0.8, new Point(Viewport.ActualWidth / 2, Viewport.ActualHeight / 2));

    private void MapView_Click(object sender, RoutedEventArgs e)
    {
        _focusSelectedRegion = false;
        _showWorld = !_showWorld;
        BtnMapView.Content = _showWorld ? "US + EU" : "WORLD";
        FitCurrentView();
    }

    private void ZoomAt(double factor, Point pivot)
    {
        double oldScale = MapScale.ScaleX;
        double newScale = Math.Clamp(oldScale * factor, MinScale, MaxScale);
        if (Math.Abs(newScale - oldScale) < 0.0001)
            return;

        double mapX = (pivot.X - MapTranslate.X) / oldScale;
        double mapY = (pivot.Y - MapTranslate.Y) / oldScale;
        MapScale.ScaleX = newScale;
        MapScale.ScaleY = newScale;
        MapTranslate.X = pivot.X - mapX * newScale;
        MapTranslate.Y = pivot.Y - mapY * newScale;
        UpdateFixedElementScale();
    }

    private void FitCurrentView()
    {
        if (Viewport.ActualWidth <= 0 || Viewport.ActualHeight <= 0)
            return;

        if (_showWorld)
        {
            FitBounds(0, 0, WorldWidth, WorldHeight, 16);
            return;
        }

        Point northWest = Project(66, -132);
        Point southEast = Project(18, 47);
        FitBounds(northWest.X, northWest.Y, southEast.X, southEast.Y, 14);
    }

    private void FocusRegion(HaloVpnRegion region, bool updateVisuals = true)
    {
        if (Viewport.ActualWidth <= 0 || Viewport.ActualHeight <= 0)
            return;

        Point destination = RegionPoint(region);
        double arc = Math.Min(70, Math.Max(8, Math.Abs(destination.X - LocalCoordinates.X) * 0.16));
        double left = Math.Min(LocalCoordinates.X, destination.X);
        double right = Math.Max(LocalCoordinates.X, destination.X);
        double top = Math.Min(LocalCoordinates.Y, destination.Y) - arc;
        double bottom = Math.Max(LocalCoordinates.Y, destination.Y);

        // Keep nearby East US selections at a useful regional scale while centering
        // longer routes across the full map card.
        FitBounds(left, top, right, bottom, 64, 5.25, updateVisuals);
    }

    private void FitBounds(
        double left,
        double top,
        double right,
        double bottom,
        double margin,
        double maximumScale = MaxScale,
        bool updateVisuals = true)
    {
        double width = Math.Max(1, right - left);
        double height = Math.Max(1, bottom - top);
        double scale = Math.Min(
            Math.Max(1, Viewport.ActualWidth - margin * 2) / width,
            Math.Max(1, Viewport.ActualHeight - margin * 2) / height);
        scale = Math.Clamp(scale, MinScale, maximumScale);
        MapScale.ScaleX = scale;
        MapScale.ScaleY = scale;
        MapTranslate.X = Viewport.ActualWidth / 2 - (left + right) / 2 * scale;
        MapTranslate.Y = Viewport.ActualHeight / 2 - (top + bottom) / 2 * scale;
        if (updateVisuals)
            UpdateFixedElementScale();
    }

    private void UpdateFixedElementScale()
    {
        double inverse = 1 / Math.Max(MapScale.ScaleX, 0.01);
        foreach (FrameworkElement element in _fixedSizeElements)
            element.RenderTransform = new ScaleTransform(inverse, inverse);

        foreach ((string name, Button button) in _buttons)
        {
            if (_regions.TryGetValue(name, out HaloVpnRegion? region))
                PositionMarker(button, region);
        }

        if (_regions.TryGetValue(SelectedRegion, out HaloVpnRegion? selected))
        {
            PositionSelectedLabel(selected);
            UpdateRouteGeometry(selected);
        }

        foreach (ShapePath route in RouteCanvas.Children.OfType<ShapePath>())
            route.StrokeThickness = (route == _routePath ? 2 : 9) * inverse;
    }
}
