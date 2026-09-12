using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace HaloToolbox
{
    public partial class MainWindow
    {
        private readonly DispatcherTimer _lobbyIdentityTimer = new() { Interval = TimeSpan.FromSeconds(30) };
        private readonly Dictionary<string, DateTime> _lobbyIdentityRetryAt = new();
        private bool _lobbyIdentityRunning;

        private async Task StatsResolveLobbyIdentitiesAsync()
        {
            if (_lobbyIdentityRunning || !_rejoinProxy.IsRunning) return;
            _lobbyIdentityRunning = true;
            try
            {
                List<string> xuids;
                lock (_statsLock)
                    xuids = _statsMatchmakingPings.Values
                        .Where(p => p.ObservedAt >= DateTime.UtcNow.AddMinutes(-30) &&
                            (string.IsNullOrWhiteSpace(p.Gamertag) || StatsLooksLikeXuid(p.Gamertag)))
                        .Select(p => StatsNormalizeXuid(p.Xuid))
                        .Where(x => StatsLooksLikeXuid(x) && !_statsGamertagsByXuid.ContainsKey(x))
                        .Distinct().ToList();

                foreach (string xuid in xuids)
                {
                    if (!_rejoinProxy.IsRunning) break;
                    if (_lobbyIdentityRetryAt.TryGetValue(xuid, out var retryAt) && retryAt > DateTime.UtcNow)
                        continue;
                    // Bound each request and back off failures; unchanged lobby traffic
                    // does not emit another observation event to trigger a retry.
                    _lobbyIdentityRetryAt[xuid] = DateTime.UtcNow.AddMinutes(2);
                    try
                    {
                        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                        using var request = new HttpRequestMessage(HttpMethod.Get,
                            $"https://playerdb.co/api/player/xbox/{Uri.EscapeDataString(xuid)}");
                        request.Headers.TryAddWithoutValidation("User-Agent", "HaloMCCToolbox/1.0");
                        using var response = await StatsHttp.SendAsync(request, timeout.Token);
                        response.EnsureSuccessStatusCode();
                        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync(timeout.Token));
                        var player = json.RootElement.GetProperty("data").GetProperty("player");
                        string id = player.GetProperty("id").ToString();
                        string name = player.GetProperty("username").GetString() ?? "";
                        if (StatsNormalizeXuid(id) == xuid && !string.IsNullOrWhiteSpace(name) && !StatsLooksLikeXuid(name))
                        {
                            StatsRememberGamertagForXuid(xuid, name);
                            _lobbyIdentityRetryAt.Remove(xuid);
                            StatsRebuildCurrentLobbyRows();
                        }
                    }
                    catch (Exception ex)
                    {
                        RejoinFixDiagnostics.Warn("stats", $"Lobby name lookup failed for {StatsShortXuid(xuid)}: {ex.Message}");
                    }
                }
            }
            finally
            {
                _lobbyIdentityRunning = false;
                StatsRebuildCurrentLobbyRows();
                if (_rejoinProxy.IsRunning) _ = StatsFetchCurrentLobbyStats();
            }
        }
    }
}
