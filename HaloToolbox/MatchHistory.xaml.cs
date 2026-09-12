using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HaloToolbox
{
    internal sealed class MatchHistoryRow
    {
        public string Date { get; init; } = "";
        public string Game { get; init; } = "";
        public string Duration { get; init; } = "";
        public string Result { get; init; } = "";
        public bool Won { get; init; }
        public string Kills { get; init; } = "";
        public string Deaths { get; init; } = "";
        public string KD { get; init; } = "";
        public string Assists { get; init; } = "";
        public string Headshots { get; init; } = "";
        public string Medals { get; init; } = "";
        public string Score { get; init; } = "";

        public Brush ResultColor => Won
            ? new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14))
            : new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));

        public Brush KdColor
        {
            get
            {
                if (!double.TryParse(KD, out double value))
                    return new SolidColorBrush(Colors.Gray);
                if (value >= 2.0) return new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14));
                if (value >= 1.0) return new SolidColorBrush(Color.FromRgb(0xC8, 0xD8, 0xE8));
                return new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));
            }
        }
    }

    internal sealed record MatchHistoryData(
        DateTime DatePlayed,
        int DurationSeconds,
        bool Won,
        long Kills,
        long Deaths,
        long Assists,
        long Headshots,
        long Medals,
        long Score,
        string HaloTitleId);

    internal sealed record MatchHistoryPage(List<MatchHistoryData> Matches, int MaximumPage);

    public partial class MatchHistory : UserControl
    {
        private const int PageSize = 20;
        private const int MaximumMatches = 100;
        private const int MaximumPages = MaximumMatches / PageSize;

        private readonly Func<string> _accessTokenProvider;
        private CancellationTokenSource? _loadCancellation;

        public MatchHistory(string initialGamertag, Func<string> accessTokenProvider)
        {
            _accessTokenProvider = accessTokenProvider;
            InitializeComponent();
            GamertagBox.Text = initialGamertag;
        }

        private async void LoadButton_Click(object sender, RoutedEventArgs e) =>
            await LoadAsync(GamertagBox.Text);

        private async void GamertagBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            e.Handled = true;
            await LoadAsync(GamertagBox.Text);
        }

        public async Task LoadAsync(string gamertag)
        {
            gamertag = gamertag.Trim();
            GamertagBox.Text = gamertag;

            if (string.IsNullOrWhiteSpace(gamertag))
            {
                StatusText.Text = "Enter a gamertag first.";
                GamertagBox.Focus();
                return;
            }

            string token = _accessTokenProvider().Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                StatusText.Text = "Connect Halo Waypoint on Home before loading match history.";
                return;
            }

            _loadCancellation?.Cancel();
            _loadCancellation?.Dispose();
            _loadCancellation = new CancellationTokenSource();
            CancellationToken cancellationToken = _loadCancellation.Token;

            LoadButton.IsEnabled = false;
            GamertagBox.IsEnabled = false;
            StatusText.Text = $"Fetching up to {MaximumMatches} games for {gamertag}…";

            try
            {
                MatchHistoryPage firstPage = await FetchPageAsync(gamertag, token, 1, cancellationToken);
                int pageCount = Math.Clamp(firstPage.MaximumPage, 1, MaximumPages);
                var pages = new List<MatchHistoryPage> { firstPage };
                if (pageCount > 1)
                {
                    Task<MatchHistoryPage>[] remainingPageTasks = Enumerable.Range(2, pageCount - 1)
                        .Select(page => FetchPageAsync(gamertag, token, page, cancellationToken))
                        .ToArray();
                    pages.AddRange(await Task.WhenAll(remainingPageTasks));
                }

                List<MatchHistoryData> matches = pages
                    .SelectMany(page => page.Matches)
                    .OrderByDescending(match => match.DatePlayed)
                    .Take(MaximumMatches)
                    .ToList();

                cancellationToken.ThrowIfCancellationRequested();
                MatchList.ItemsSource = matches.Select(BuildRow).ToList();

                if (matches.Count == 0)
                {
                    ResetSummary();
                    StatusText.Text = $"No match history is available for {gamertag}.";
                    return;
                }

                PopulateSummary(matches);
                StatusText.Text = matches.Count == MaximumMatches
                    ? $"Loaded the latest {MaximumMatches} games for {gamertag}."
                    : $"Loaded all {matches.Count} available games for {gamertag}.";
            }
            catch (OperationCanceledException)
            {
                // A newer search replaced this one.
            }
            catch (UnauthorizedAccessException)
            {
                MatchList.ItemsSource = null;
                ResetSummary();
                StatusText.Text = "Halo Waypoint authorization expired. Reconnect on Home and try again.";
            }
            catch (Exception ex)
            {
                MatchList.ItemsSource = null;
                ResetSummary();
                StatusText.Text = $"Unable to load match history: {ex.Message}";
            }
            finally
            {
                if (!cancellationToken.IsCancellationRequested)
                {
                    LoadButton.IsEnabled = true;
                    GamertagBox.IsEnabled = true;
                }
            }
        }

        private static async Task<MatchHistoryPage> FetchPageAsync(
            string gamertag,
            string token,
            int page,
            CancellationToken cancellationToken)
        {
            string url =
                $"https://mccapi.svc.halowaypoint.com/hmcc/users/gt({Uri.EscapeDataString(gamertag)})" +
                $"/matches?page={page}&pageSize={PageSize}";

            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.TryAddWithoutValidation("x-343-authorization-spartan", token);
            request.Headers.TryAddWithoutValidation("Accept", "application/json");

            using HttpResponseMessage response = await MainWindow.StatsHttp.SendAsync(request, cancellationToken);
            if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
                throw new UnauthorizedAccessException();

            response.EnsureSuccessStatusCode();
            string body = await response.Content.ReadAsStringAsync(cancellationToken);
            using JsonDocument json = JsonDocument.Parse(body);

            int maximumPage = json.RootElement.TryGetProperty("maxPage", out JsonElement maximumPageElement) &&
                              maximumPageElement.TryGetInt32(out int parsedMaximumPage)
                ? parsedMaximumPage
                : 1;

            if (!json.RootElement.TryGetProperty("matches", out JsonElement matchArray) ||
                matchArray.ValueKind != JsonValueKind.Array)
                return new MatchHistoryPage(new List<MatchHistoryData>(), maximumPage);

            var matches = new List<MatchHistoryData>();
            foreach (JsonElement match in matchArray.EnumerateArray())
            {
                DateTime datePlayed = match.TryGetProperty("datePlayed", out JsonElement dateElement) &&
                                      dateElement.TryGetDateTime(out DateTime parsedDate)
                    ? parsedDate.ToLocalTime()
                    : DateTime.MinValue;

                match.TryGetProperty("durationSeconds", out JsonElement durationElement);
                durationElement.TryGetInt32(out int durationSeconds);
                bool won = match.TryGetProperty("won", out JsonElement wonElement) &&
                           wonElement.ValueKind is JsonValueKind.True or JsonValueKind.False &&
                           wonElement.GetBoolean();

                matches.Add(new MatchHistoryData(
                    datePlayed,
                    durationSeconds,
                    won,
                    ReadInt64(match, "kills"),
                    ReadInt64(match, "deaths"),
                    ReadInt64(match, "assists"),
                    ReadInt64(match, "headshots"),
                    ReadInt64(match, "medals"),
                    ReadInt64(match, "score"),
                    match.TryGetProperty("haloTitleId", out JsonElement titleElement)
                        ? titleElement.GetString() ?? ""
                        : ""));
            }

            return new MatchHistoryPage(matches, maximumPage);
        }

        private static long ReadInt64(JsonElement source, string propertyName) =>
            source.TryGetProperty(propertyName, out JsonElement element) && element.TryGetInt64(out long value)
                ? value
                : 0;

        private void PopulateSummary(IReadOnlyCollection<MatchHistoryData> matches)
        {
            int total = matches.Count;
            int wins = matches.Count(match => match.Won);
            long totalKills = matches.Sum(match => match.Kills);
            long totalDeaths = matches.Sum(match => match.Deaths);
            long totalScore = matches.Sum(match => match.Score);

            MatchesLabel.Text = total.ToString();
            WinRateLabel.Text = $"{(double)wins / total * 100:F0}%";
            KdLabel.Text = (totalDeaths > 0 ? (double)totalKills / totalDeaths : totalKills).ToString("F2");
            TotalKillsLabel.Text = totalKills.ToString("N0");
            TotalDeathsLabel.Text = totalDeaths.ToString("N0");
            AvgKillsLabel.Text = ((double)totalKills / total).ToString("F1");
            AvgDeathsLabel.Text = ((double)totalDeaths / total).ToString("F1");
            AvgScoreLabel.Text = ((double)totalScore / total).ToString("F1");
        }

        private void ResetSummary()
        {
            MatchesLabel.Text = "—";
            WinRateLabel.Text = "—";
            KdLabel.Text = "—";
            TotalKillsLabel.Text = "—";
            TotalDeathsLabel.Text = "—";
            AvgKillsLabel.Text = "—";
            AvgDeathsLabel.Text = "—";
            AvgScoreLabel.Text = "—";
        }

        private static MatchHistoryRow BuildRow(MatchHistoryData match)
        {
            double kd = match.Deaths > 0 ? (double)match.Kills / match.Deaths : match.Kills;
            int minutes = match.DurationSeconds / 60;
            int seconds = match.DurationSeconds % 60;

            return new MatchHistoryRow
            {
                Date = match.DatePlayed == DateTime.MinValue
                    ? "—"
                    : match.DatePlayed.ToString("MMM dd yyyy h:mm tt"),
                Game = GameTitle(match.HaloTitleId),
                Duration = $"{minutes}:{seconds:D2}",
                Result = match.Won ? "WIN" : "LOSS",
                Won = match.Won,
                Kills = match.Kills.ToString(),
                Deaths = match.Deaths.ToString(),
                KD = kd.ToString("F2"),
                Assists = match.Assists.ToString(),
                Headshots = match.Headshots.ToString(),
                Medals = match.Medals.ToString(),
                Score = match.Score.ToString()
            };
        }

        private static string GameTitle(string id) => id switch
        {
            "HaloCEA" => "H:CE",
            "Halo2Anniversary" => "H2A",
            "Halo3" => "H3",
            "HaloReach" => "Reach",
            "Halo4" => "H4",
            "Halo5Forge" => "H5",
            _ => id
        };
    }
}
