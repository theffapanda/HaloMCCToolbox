using System.Globalization;
using System.Text.Json;

namespace HaloToolbox;

internal static class XboxCareerStats
{
    internal const string Scid = "77290100-225e-4768-9373-98164430a9f8";
    internal const string Kills = "EnemyDefeats.GameplayModeId.3.MatchMade.1";
    internal const string Deaths = "Deaths.GameplayModeId.3.MatchMade.1";

    internal static string CreateRequest(string xuid) => JsonSerializer.Serialize(new
    {
        requestedusers = new[] { xuid },
        requestedscids = new[] { new { scid = Scid, requestedstats = new[] { Kills, Deaths } } }
    });

    internal static (long kills, long deaths)? Parse(string body, string xuid)
    {
        using var json = JsonDocument.Parse(body);
        if (!json.RootElement.TryGetProperty("users", out var users) || users.ValueKind != JsonValueKind.Array) return null;
        foreach (var user in users.EnumerateArray())
        {
            if (!user.TryGetProperty("xuid", out var id) || id.ToString() != xuid ||
                !user.TryGetProperty("scids", out var scids) || scids.ValueKind != JsonValueKind.Array) continue;
            foreach (var config in scids.EnumerateArray())
            {
                if (!config.TryGetProperty("scid", out var scid) || !string.Equals(scid.ToString(), Scid, StringComparison.OrdinalIgnoreCase) ||
                    !config.TryGetProperty("stats", out var stats) || stats.ValueKind != JsonValueKind.Array) continue;
                long? kills = null, deaths = null;
                foreach (var stat in stats.EnumerateArray())
                {
                    if (!stat.TryGetProperty("statname", out var name) || !stat.TryGetProperty("value", out var value) ||
                        !long.TryParse(value.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out long count) || count < 0) continue;
                    if (name.ToString() == Kills) kills = count;
                    if (name.ToString() == Deaths) deaths = count;
                }
                if (kills.HasValue && deaths.HasValue) return (kills.Value, deaths.Value);
            }
        }
        return null;
    }
}
