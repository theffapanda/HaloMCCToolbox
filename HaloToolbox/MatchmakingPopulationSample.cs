namespace HaloToolbox;

public sealed record MatchmakingPopulationSample(
    DateTimeOffset CapturedAt,
    string HopperName,
    string DisplayName,
    int Population)
{
    public string SessionId { get; init; } = "";
}
