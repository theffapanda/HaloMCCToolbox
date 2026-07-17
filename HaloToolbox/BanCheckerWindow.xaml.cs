using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace HaloToolbox;

public partial class BanCheckerWindow : Window
{
    public const int MaxTargets = 15;
    private readonly Func<IReadOnlyList<string>, Task<IReadOnlyList<BanCheckDisplayResult>>> _checkAsync;
    private readonly ObservableCollection<BanCheckDisplayResult> _results = new();

    public BanCheckerWindow(Func<IReadOnlyList<string>, Task<IReadOnlyList<BanCheckDisplayResult>>> checkAsync)
    {
        InitializeComponent();
        _checkAsync = checkAsync;
        ResultsGrid.ItemsSource = _results;
        TargetsBox.Focus();
    }

    private List<string> ReadTargets() => TargetsBox.Text
        .Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(value => value.Trim())
        .Where(value => value.Length > 0)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

    private void TargetsBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        int count = ReadTargets().Count;
        CountText.Text = $"{count} / {MaxTargets}";
        CountText.Foreground = count > MaxTargets
            ? System.Windows.Media.Brushes.OrangeRed
            : (System.Windows.Media.Brush)FindResource("MutedBrush");
        CheckButton.IsEnabled = count is > 0 and <= MaxTargets;
    }

    private async void Check_Click(object sender, RoutedEventArgs e)
    {
        var targets = ReadTargets();
        if (targets.Count == 0 || targets.Count > MaxTargets)
            return;

        CheckButton.IsEnabled = false;
        TargetsBox.IsEnabled = false;
        StatusText.Text = $"Checking {targets.Count} player(s)...";
        _results.Clear();
        try
        {
            foreach (var result in await _checkAsync(targets))
                _results.Add(result);
            StatusText.Text = $"Checked {_results.Count} player(s)";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Check failed";
            MessageBox.Show(this, ex.Message, "Ban Checker Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TargetsBox.IsEnabled = true;
            CheckButton.IsEnabled = ReadTargets().Count is > 0 and <= MaxTargets;
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}

public sealed class BanCheckDisplayResult
{
    public string Target { get; init; } = "";
    public string Result { get; init; } = "";
    public string Details { get; init; } = "";
}
