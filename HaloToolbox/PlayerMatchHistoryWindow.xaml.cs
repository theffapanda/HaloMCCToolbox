using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace HaloToolbox
{
    // ── Per-match display model ───────────────────────────────────────────────
    class MatchRow
    {
        public string Date      { get; set; } = "";
        public string Game      { get; set; } = "";
        public string Duration  { get; set; } = "";
        public string Result    { get; set; } = "";
        public bool   Won       { get; set; }
        public string Kills     { get; set; } = "";
        public string Deaths    { get; set; } = "";
        public string KD        { get; set; } = "";
        public string Assists   { get; set; } = "";
        public string Headshots { get; set; } = "";
        public string Medals    { get; set; } = "";
        public string Score     { get; set; } = "";

        public Brush ResultColor => Won
            ? new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14))
            : new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));

        public Brush KdColor
        {
            get
            {
                if (!double.TryParse(KD, out double v))
                    return new SolidColorBrush(Colors.Gray);
                if (v >= 2.0) return new SolidColorBrush(Color.FromRgb(0x39, 0xFF, 0x14));
                if (v >= 1.0) return new SolidColorBrush(Color.FromRgb(0xC8, 0xD8, 0xE8));
                return new SolidColorBrush(Color.FromRgb(0xFF, 0x2D, 0x55));
            }
        }
    }

    // ── Raw data fetched from the API ─────────────────────────────────────────
    record MatchData(
        DateTime DatePlayed,
        int      DurationSeconds,
        bool     Won,
        long     Kills,
        long     Deaths,
        long     Assists,
        long     Headshots,
        long     Medals,
        long     Score,
        string   HaloTitleId);

    // ── Window ────────────────────────────────────────────────────────────────
    public partial class PlayerMatchHistoryWindow : Window
    {
        private readonly string _gamertag;
        private readonly string _token;

        public PlayerMatchHistoryWindow(string gamertag, string token)
        {
            _gamertag = gamertag;
            _token    = token;
            InitializeComponent();
            Title = $"Match History — {gamertag}";
            Loaded += OnLoadedAsync;
        }

        private async void OnLoadedAsync(object sender, RoutedEventArgs e)
        {
            StatusText.Text = "Fetching match history…";
            try
            {
                var pageTasks = Enumerable.Range(1, 5)
                    .Select(p => FetchPageAsync(p))
                    .ToArray();
                var pages = await Task.WhenAll(pageTasks);

                var matches = pages
                    .SelectMany(p => p)
                    .OrderByDescending(m => m.DatePlayed)
                    .ToList();

                if (!matches.Any())
                {
                    StatusText.Text = "No match data returned.";
                    return;
                }

                PopulateSummary(matches);
                MatchList.ItemsSource = matches.Select(BuildRow).ToList();
                StatusText.Text = $"{matches.Count} matches loaded.";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Error: {ex.Message}";
            }
        }

        private void PopulateSummary(List<MatchData> matches)
        {
            int  total = matches.Count;
            int  wins  = matches.Count(m => m.Won);
            long totalKills  = matches.Sum(m => m.Kills);
            long totalDeaths = matches.Sum(m => m.Deaths);
            long totalScore  = matches.Sum(m => m.Score);

            double winRate   = total > 0 ? (double)wins / total * 100 : 0;
            double kd        = totalDeaths > 0 ? (double)totalKills / totalDeaths : totalKills;
            double avgKills  = total > 0 ? (double)totalKills  / total : 0;
            double avgDeaths = total > 0 ? (double)totalDeaths / total : 0;
            double avgScore  = total > 0 ? (double)totalScore  / total : 0;

            MatchesLabel.Text   = total.ToString();
            WinRateLabel.Text   = $"{winRate:F0}%";
            KdLabel.Text        = kd.ToString("F2");
            TotalKillsLabel.Text = totalKills.ToString("N0");
            TotalDeathsLabel.Text = totalDeaths.ToString("N0");
            AvgKillsLabel.Text  = avgKills.ToString("F1");
            AvgDeathsLabel.Text = avgDeaths.ToString("F1");
            AvgScoreLabel.Text  = avgScore.ToString("F1");
        }

        private async Task<List<MatchData>> FetchPageAsync(int page)
        {
            try
            {
                string url =
                    $"https://mccapi.svc.halowaypoint.com/hmcc/users/gt({Uri.EscapeDataString(_gamertag)})" +
                    $"/matches?page={page}&pageSize=20";

                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.TryAddWithoutValidation("x-343-authorization-spartan", _token);
                req.Headers.TryAddWithoutValidation("Accept", "application/json");

                var resp = await MainWindow.StatsHttp.SendAsync(req);
                if (!resp.IsSuccessStatusCode) return new();

                string body = await resp.Content.ReadAsStringAsync();
                using var json = JsonDocument.Parse(body);
                var root = json.RootElement;

                if (!root.TryGetProperty("matches", out var arr) ||
                    arr.ValueKind != JsonValueKind.Array)
                    return new();

                var result = new List<MatchData>();
                foreach (var m in arr.EnumerateArray())
                {
                    DateTime date = m.TryGetProperty("datePlayed", out var dpEl) &&
                                    dpEl.TryGetDateTime(out var dt)
                                    ? dt.ToLocalTime() : DateTime.MinValue;

                    m.TryGetProperty("durationSeconds", out var durEl);
                    durEl.TryGetInt32(out int dur);

                    bool won = m.TryGetProperty("won", out var wonEl) && wonEl.GetBoolean();

                    m.TryGetProperty("kills",     out var kEl);  kEl.TryGetInt64(out long kills);
                    m.TryGetProperty("deaths",    out var dEl);  dEl.TryGetInt64(out long deaths);
                    m.TryGetProperty("assists",   out var aEl);  aEl.TryGetInt64(out long assists);
                    m.TryGetProperty("headshots", out var hsEl); hsEl.TryGetInt64(out long headshots);
                    m.TryGetProperty("medals",    out var mdEl); mdEl.TryGetInt64(out long medals);
                    m.TryGetProperty("score",     out var scEl); scEl.TryGetInt64(out long score);

                    string titleId = m.TryGetProperty("haloTitleId", out var tidEl)
                                     ? tidEl.GetString() ?? "" : "";

                    result.Add(new MatchData(date, dur, won, kills, deaths,
                                             assists, headshots, medals, score, titleId));
                }
                return result;
            }
            catch { return new(); }
        }

        private static MatchRow BuildRow(MatchData m)
        {
            double kd   = m.Deaths > 0 ? (double)m.Kills / m.Deaths : m.Kills;
            int    mins = m.DurationSeconds / 60;
            int    secs = m.DurationSeconds % 60;

            return new MatchRow
            {
                Date      = m.DatePlayed == DateTime.MinValue
                            ? "—" : m.DatePlayed.ToString("MMM dd yyyy h:mm tt"),
                Game      = GameTitle(m.HaloTitleId),
                Duration  = $"{mins}:{secs:D2}",
                Result    = m.Won ? "WIN" : "LOSS",
                Won       = m.Won,
                Kills     = m.Kills.ToString(),
                Deaths    = m.Deaths.ToString(),
                KD        = kd.ToString("F2"),
                Assists   = m.Assists.ToString(),
                Headshots = m.Headshots.ToString(),
                Medals    = m.Medals.ToString(),
                Score     = m.Score.ToString(),
            };
        }

        private static string GameTitle(string id) => id switch
        {
            "HaloCEA"          => "H:CE",
            "Halo2Anniversary" => "H2A",
            "Halo3"            => "H3",
            "HaloReach"        => "Reach",
            "Halo4"            => "H4",
            "Halo5Forge"       => "H5",
            _                  => id,
        };
    }
}
