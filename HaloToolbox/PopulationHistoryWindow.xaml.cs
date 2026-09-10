using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace HaloToolbox;

public partial class PopulationHistoryWindow : Window
{
    private const string OverallHopperName = "__overall__";
    private readonly List<MatchmakingPopulationSample> _samples;

    public PopulationHistoryWindow(IEnumerable<MatchmakingPopulationSample> samples)
    {
        InitializeComponent();
        _samples = samples.ToList();
        var hoppers = _samples.GroupBy(x => x.HopperName)
            .Select(g => new HopperChoice(g.Key, g.Last().DisplayName, g.Count()))
            .OrderByDescending(x => x.SampleCount).ThenBy(x => x.DisplayName)
            .ToList();
        if (hoppers.Count > 0)
            hoppers.Insert(0, new HopperChoice(OverallHopperName, "Overall — all queues", _samples.Count));
        HopperBox.ItemsSource = hoppers;
        HopperBox.DisplayMemberPath = nameof(HopperChoice.Label);
        if (hoppers.Count > 0)
            HopperBox.SelectedIndex = 0;
        else
            SummaryLabel.Text = "No population samples yet. Refresh the Population page, then open the graph again.";
    }

    private void HopperBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (HopperBox.SelectedItem is not HopperChoice choice) return;
        var points = _samples.Where(x => choice.HopperName == OverallHopperName || x.HopperName == choice.HopperName)
            .OrderBy(x => x.CapturedAt).ToList();
        Chart.Samples = points;
        if (choice.HopperName == OverallHopperName)
        {
            int queueCount = points.Select(x => x.HopperName).Distinct().Count();
            int refreshCount = points.Select(x => x.CapturedAt).Distinct().Count();
            SummaryLabel.Text = $"{queueCount} queues · {refreshCount} session refresh{(refreshCount == 1 ? "" : "es")} · each line is a queue (counts are not summed)";
        }
        else
        {
            var latest = points[^1];
            SummaryLabel.Text = $"{points.Count} session sample{(points.Count == 1 ? "" : "s")} · latest {latest.Population} players at {latest.CapturedAt:h:mm:ss tt}";
        }
    }

    private void ExportCsv_Click(object sender, RoutedEventArgs e)
    {
        if (_samples.Count == 0)
        {
            ToolboxDialog.Show(this, "There are no population samples to export yet.", "Population Export",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        bool overall = HopperBox.SelectedItem is HopperChoice choice && choice.HopperName == OverallHopperName;
        IEnumerable<MatchmakingPopulationSample> export = overall || HopperBox.SelectedItem is not HopperChoice selected
            ? _samples
            : _samples.Where(x => x.HopperName == selected.HopperName);
        var dialog = new SaveFileDialog
        {
            Title = "Export population history",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            DefaultExt = ".csv",
            AddExtension = true,
            FileName = $"mcc-population-{DateTime.Now:yyyyMMdd-HHmmss}.csv"
        };
        if (dialog.ShowDialog(this) != true) return;

        static string Csv(string value) => $"\"{value.Replace("\"", "\"\"")}\"";
        var csv = new StringBuilder("timestamp_local,timestamp_utc,hopper_name,queue,population\r\n");
        foreach (var sample in export.OrderBy(x => x.CapturedAt).ThenBy(x => x.DisplayName))
            csv.Append(Csv(sample.CapturedAt.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss zzz"))).Append(',')
               .Append(Csv(sample.CapturedAt.UtcDateTime.ToString("O"))).Append(',')
               .Append(Csv(sample.HopperName)).Append(',').Append(Csv(sample.DisplayName)).Append(',')
               .Append(sample.Population).Append("\r\n");
        File.WriteAllText(dialog.FileName, csv.ToString(), new UTF8Encoding(true));
        SummaryLabel.Text = $"Exported {export.Count()} samples to {Path.GetFileName(dialog.FileName)}";
    }

    private sealed record HopperChoice(string HopperName, string DisplayName, int SampleCount)
    {
        public string Label => $"{DisplayName} ({SampleCount})";
    }
}

public sealed class PopulationHistoryChart : FrameworkElement
{
    private static readonly Brush[] SeriesBrushes =
    {
        Brushes.DeepSkyBlue, Brushes.Orange, Brushes.LimeGreen, Brushes.MediumPurple,
        Brushes.HotPink, Brushes.Gold, Brushes.Turquoise, Brushes.Coral,
        Brushes.DodgerBlue, Brushes.YellowGreen, Brushes.Orchid, Brushes.SandyBrown
    };
    private IReadOnlyList<MatchmakingPopulationSample> _samples = Array.Empty<MatchmakingPopulationSample>();
    private DateTimeOffset? _hoveredAt;
    private IReadOnlyList<MatchmakingPopulationSample> _hoveredValues = Array.Empty<MatchmakingPopulationSample>();
    private Point _hoverPoint;
    private double _plotLeft, _plotRight, _plotTop, _plotBottom;

    public PopulationHistoryChart()
    {
        MouseMove += Chart_MouseMove;
        MouseLeave += (_, _) => { _hoveredAt = null; _hoveredValues = Array.Empty<MatchmakingPopulationSample>(); InvalidateVisual(); };
    }
    public IReadOnlyList<MatchmakingPopulationSample> Samples
    {
        get => _samples;
        set { _samples = value; _hoveredAt = null; _hoveredValues = Array.Empty<MatchmakingPopulationSample>(); InvalidateVisual(); }
    }

    private void Chart_MouseMove(object sender, MouseEventArgs e)
    {
        if (_samples.Count == 0 || _plotRight <= _plotLeft) return;
        Point mouse = e.GetPosition(this);
        if (mouse.X < _plotLeft || mouse.X > _plotRight || mouse.Y < _plotTop || mouse.Y > _plotBottom)
        {
            _hoveredAt = null; _hoveredValues = Array.Empty<MatchmakingPopulationSample>(); InvalidateVisual(); return;
        }
        DateTimeOffset first = _samples.Min(x => x.CapturedAt), last = _samples.Max(x => x.CapturedAt);
        double ratio = Math.Clamp((mouse.X - _plotLeft) / (_plotRight - _plotLeft), 0, 1);
        var target = first + TimeSpan.FromTicks((long)((last - first).Ticks * ratio));
        var nearest = _samples.OrderBy(x => Math.Abs((x.CapturedAt - target).Ticks)).First();
        _hoveredAt = nearest.CapturedAt;
        _hoveredValues = _samples.Where(x => x.CapturedAt == nearest.CapturedAt)
            .OrderByDescending(x => x.Population).ToList();
        _hoverPoint = mouse;
        InvalidateVisual();
    }

    protected override void OnRender(DrawingContext dc)
    {
        base.OnRender(dc);
        var series = _samples.GroupBy(x => x.HopperName).OrderBy(x => x.Last().DisplayName).ToList();
        bool overall = series.Count > 1;
        double legendWidth = overall ? Math.Min(250, Math.Max(170, ActualWidth * .28)) : 0;
        double left = 54, top = 18, right = Math.Max(left + 1, ActualWidth - 18 - legendWidth), bottom = Math.Max(top + 1, ActualHeight - 38);
        (_plotLeft, _plotRight, _plotTop, _plotBottom) = (left, right, top, bottom);
        var muted = (TryFindResource("MutedBrush") as Brush) ?? Brushes.Gray;
        var border = (TryFindResource("BorderBrush") as Brush) ?? Brushes.DimGray;
        var accent = (TryFindResource("AccentBrush") as Brush) ?? Brushes.DeepSkyBlue;
        var text = (TryFindResource("TextBrush") as Brush) ?? Brushes.White;
        var background = (TryFindResource("BgBrush") as Brush) ?? Brushes.Black;
        var axisPen = new Pen(border, 1);

        int max = Math.Max(1, _samples.Count == 0 ? 1 : _samples.Max(x => x.Population));
        max = (int)Math.Ceiling(max / 10.0) * 10;
        for (int i = 0; i <= 4; i++)
        {
            double y = bottom - (bottom - top) * i / 4.0;
            dc.DrawLine(axisPen, new Point(left, y), new Point(right, y));
            DrawText(dc, (max * i / 4).ToString(), muted, 10, 2, y - 7);
        }

        if (_samples.Count == 0) { DrawText(dc, "Refresh population data to begin the session graph.", muted, 12, left + 15, top + 20); return; }
        DateTimeOffset first = _samples.Min(x => x.CapturedAt), last = _samples.Max(x => x.CapturedAt);
        double seconds = Math.Max(1, (last - first).TotalSeconds);
        Point Map(MatchmakingPopulationSample s) => new(
            left + (right - left) * (s.CapturedAt - first).TotalSeconds / seconds,
            bottom - (bottom - top) * s.Population / max);

        if (_hoveredAt.HasValue)
        {
            double hoverX = left + (right - left) * (_hoveredAt.Value - first).TotalSeconds / seconds;
            dc.DrawLine(new Pen(muted, 1) { DashStyle = DashStyles.Dash },
                new Point(hoverX, top), new Point(hoverX, bottom));
        }

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            var points = series[seriesIndex].OrderBy(x => x.CapturedAt).ToList();
            Brush seriesBrush = overall ? SeriesBrushes[seriesIndex % SeriesBrushes.Length] : accent;
            var geometry = new StreamGeometry();
            using (var ctx = geometry.Open())
            {
                ctx.BeginFigure(Map(points[0]), false, false);
                if (points.Count == 1) ctx.LineTo(new Point(right, Map(points[0]).Y), true, false);
                else foreach (var sample in points.Skip(1)) ctx.LineTo(Map(sample), true, false);
            }
            geometry.Freeze();
            dc.DrawGeometry(null, new Pen(seriesBrush, overall ? 1.5 : 2), geometry);
            foreach (var sample in points) dc.DrawEllipse(seriesBrush, null, Map(sample), overall ? 2 : 3, overall ? 2 : 3);

            if (overall)
            {
                double legendY = top + seriesIndex * 19;
                if (legendY + 14 <= bottom)
                {
                    dc.DrawLine(new Pen(seriesBrush, 2), new Point(right + 16, legendY + 6), new Point(right + 34, legendY + 6));
                    string label = points[^1].DisplayName;
                    if (label.Length > 27) label = label[..26] + "…";
                    DrawText(dc, label, text, 10, right + 40, legendY);
                }
            }
        }
        if (_hoveredAt.HasValue && _hoveredValues.Count > 0)
        {
            const double cardWidth = 280;
            double cardHeight = 31 + _hoveredValues.Count * 17;
            double cardX = _hoverPoint.X + 14;
            if (cardX + cardWidth > ActualWidth - 8) cardX = _hoverPoint.X - cardWidth - 14;
            cardX = Math.Max(8, cardX);
            double cardY = Math.Clamp(_hoverPoint.Y - cardHeight / 2, 8, Math.Max(8, ActualHeight - cardHeight - 8));
            dc.DrawRoundedRectangle(background, new Pen(accent, 1),
                new Rect(cardX, cardY, cardWidth, cardHeight), 4, 4);
            DrawText(dc, _hoveredAt.Value.ToLocalTime().ToString("MMM d, yyyy · h:mm:ss tt"),
                text, 11, cardX + 10, cardY + 7);
            for (int i = 0; i < _hoveredValues.Count; i++)
            {
                var value = _hoveredValues[i];
                int colorIndex = series.FindIndex(x => x.Key == value.HopperName);
                Brush color = overall && colorIndex >= 0 ? SeriesBrushes[colorIndex % SeriesBrushes.Length] : accent;
                double rowY = cardY + 28 + i * 17;
                dc.DrawEllipse(color, null, new Point(cardX + 13, rowY + 6), 3, 3);
                string label = value.DisplayName.Length > 24 ? value.DisplayName[..23] + "…" : value.DisplayName;
                DrawText(dc, label, text, 10, cardX + 22, rowY);
                var countText = MakeText(value.Population.ToString(), text, 10);
                dc.DrawText(countText, new Point(cardX + cardWidth - countText.Width - 10, rowY));
            }
        }
        DrawText(dc, first.ToLocalTime().ToString("h:mm:ss tt"), text, 10, left, bottom + 8);
        string end = last.ToLocalTime().ToString("h:mm:ss tt");
        var endText = MakeText(end, text, 10);
        dc.DrawText(endText, new Point(right - endText.Width, bottom + 8));
    }

    private static FormattedText MakeText(string value, Brush brush, double size) =>
        new(value, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface("Consolas"), size, brush, 1.0);

    private static void DrawText(DrawingContext dc, string value, Brush brush, double size, double x, double y) =>
        dc.DrawText(MakeText(value, brush, size), new Point(x, y));
}
