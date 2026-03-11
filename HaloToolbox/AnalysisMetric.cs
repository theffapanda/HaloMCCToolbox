using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace HaloToolbox;

/// <summary>A metric tile shown in the analysis bar (value + label stacked).</summary>
public class AnalysisMetric : UserControl
{
    private readonly TextBlock _val;
    private readonly TextBlock _lbl;

    public static readonly DependencyProperty LabelProperty =
        DependencyProperty.Register(nameof(Label), typeof(string), typeof(AnalysisMetric),
            new PropertyMetadata(string.Empty));

    public string Label
    {
        get => (string)GetValue(LabelProperty);
        set
        {
            SetValue(LabelProperty, value);
            _lbl.Text = value;
        }
    }

    public AnalysisMetric()
    {
        var mono = new FontFamily("Cascadia Code, Consolas, Courier New");

        _val = new TextBlock
        {
            FontFamily        = mono,
            FontSize          = 20,
            FontWeight        = FontWeights.Bold,
            Foreground        = new SolidColorBrush(Color.FromRgb(0x39, 0xD0, 0xC8)),
            HorizontalAlignment = HorizontalAlignment.Center,
            Text              = "—"
        };

        _lbl = new TextBlock
        {
            FontFamily        = mono,
            FontSize          = 8,
            Foreground        = new SolidColorBrush(Color.FromRgb(0x7D, 0x85, 0x90)),
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin            = new Thickness(0, 2, 0, 0)
        };

        var sp = new StackPanel
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(10, 0, 10, 0)
        };
        sp.Children.Add(_val);
        sp.Children.Add(_lbl);
        Content = sp;
    }

    public void SetValue(string value, SolidColorBrush? color = null)
    {
        _val.Text       = value;
        _val.Foreground = color ?? new SolidColorBrush(Color.FromRgb(0x39, 0xD0, 0xC8));
        _lbl.Text       = Label;
    }
}
