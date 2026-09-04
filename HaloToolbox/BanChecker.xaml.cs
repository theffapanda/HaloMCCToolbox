using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace HaloToolbox;

public partial class BanChecker : UserControl
{
    public const int MaxTargets = 15;
    private readonly Func<IReadOnlyList<string>, Task<IReadOnlyList<BanCheckDisplayResult>>> _checkAsync;
    private readonly Func<BanCheckerAuthorizationState> _getAuthorizationState;
    private readonly Action _openLiveFeatures;
    private readonly ObservableCollection<BanCheckDisplayResult> _results = new();

    public BanChecker(
        Func<IReadOnlyList<string>, Task<IReadOnlyList<BanCheckDisplayResult>>> checkAsync,
        Func<BanCheckerAuthorizationState> getAuthorizationState,
        Action openLiveFeatures)
    {
        InitializeComponent();
        _checkAsync = checkAsync;
        _getAuthorizationState = getAuthorizationState;
        _openLiveFeatures = openLiveFeatures;
        ResultsGrid.ItemsSource = _results;
        Loaded += (_, _) =>
        {
            RefreshAuthorizationStatus();
            TargetsBox.Focus();
        };
    }

    public void RefreshAuthorizationStatus()
    {
        BanCheckerAuthorizationState state = _getAuthorizationState();
        AuthStatusText.Text = state.Status;
        AuthStatusText.Foreground = (Brush)FindResource(state.IsAvailable ? "GreenBrush" : "OrangeBrush");
        AuthDetailText.Text = state.Detail;
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
        CountText.Text = $"{count} / {MaxTargets} PLAYERS";
        CountText.Foreground = count > MaxTargets
            ? Brushes.OrangeRed
            : (Brush)FindResource("MutedBrush");
        CheckButton.IsEnabled = count is > 0 and <= MaxTargets && TargetsBox.IsEnabled;
    }

    private async void Check_Click(object sender, RoutedEventArgs e)
    {
        var targets = ReadTargets();
        if (targets.Count == 0 || targets.Count > MaxTargets)
            return;

        CheckButton.IsEnabled = false;
        TargetsBox.IsEnabled = false;
        StatusText.Text = $"Checking {targets.Count} player(s)...";
        ErrorCard.Visibility = Visibility.Collapsed;
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
            ErrorText.Text = ex.Message;
            ErrorCard.Visibility = Visibility.Visible;
        }
        finally
        {
            TargetsBox.IsEnabled = true;
            CheckButton.IsEnabled = ReadTargets().Count is > 0 and <= MaxTargets;
            RefreshAuthorizationStatus();
        }
    }

    private void RefreshAuth_Click(object sender, RoutedEventArgs e) => RefreshAuthorizationStatus();

    private void OpenFeatures_Click(object sender, RoutedEventArgs e) => _openLiveFeatures();
}

public sealed record BanCheckerAuthorizationState(bool IsAvailable, string Status, string Detail);

public sealed class BanCheckDisplayResult
{
    public string Target { get; init; } = "";
    public string Result { get; init; } = "";
    public string Details { get; init; } = "";
}
