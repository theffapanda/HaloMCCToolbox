using System.ComponentModel;
using System.Text.Json.Serialization;
using System.Windows.Media;

namespace HaloToolbox;

public class ServiceRecord
{
    [JsonPropertyName("xuid")]
    public string Xuid { get; set; } = "";

    [JsonPropertyName("xp")]
    public long Xp { get; set; }

    [JsonPropertyName("skillRank")]
    public int SkillRank { get; set; }

    [JsonPropertyName("timePlayedSeconds")]
    public long TimePlayedSeconds { get; set; }

    [JsonPropertyName("multiplayer")]
    public MultiplayerStats Multiplayer { get; set; } = new();

    [JsonPropertyName("campaign")]
    public CampaignStats? Campaign { get; set; }
}

public class MultiplayerStats
{
    [JsonPropertyName("gamesPlayed")]
    public int GamesPlayed { get; set; }

    [JsonPropertyName("wins")]
    public int Wins { get; set; }

    [JsonPropertyName("losses")]
    public int Losses { get; set; }

    [JsonPropertyName("kills")]
    public long Kills { get; set; }

    [JsonPropertyName("deaths")]
    public long Deaths { get; set; }

    [JsonPropertyName("assists")]
    public long Assists { get; set; }
}

public class CampaignStats
{
    [JsonPropertyName("missionsCompleted")]
    public int MissionsCompleted { get; set; }

    [JsonPropertyName("missionKills")]
    public long MissionKills { get; set; }
}

// ── Rejoin Guard ──────────────────────────────────────────────────────────────
public class SavedHandleInfo
{
    [JsonPropertyName("scid")]
    public string Scid { get; set; } = "";

    [JsonPropertyName("templateName")]
    public string TemplateName { get; set; } = "";

    [JsonPropertyName("sessionName")]
    public string SessionName { get; set; } = "";

    [JsonPropertyName("savedAt")]
    public DateTime SavedAt { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("requestHeaders")]
    public Dictionary<string, string> RequestHeaders { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    [JsonPropertyName("connectionGuid")]
    public string ConnectionGuid { get; set; } = "";  // Captured from player's connection GUID when joining

    [JsonPropertyName("subscriptionId")]
    public string SubscriptionId { get; set; } = "";  // RTA subscription ID for connection validation

    [JsonPropertyName("changeTypes")]
    public List<string> ChangeTypes { get; set; } = new() { "everything" };  // RTA change types

    [JsonPropertyName("playerXuid")]
    public string PlayerXuid { get; set; } = "";  // Player's Xbox Live User ID - CRITICAL for member injection

    // Display helpers
    public string SessionShort => SessionName.Length > 13 ? SessionName[..13] + "…" : SessionName;
    public string SavedAtStr   => SavedAt.ToLocalTime().ToString("HH:mm:ss");
    public string SessionUrl   =>
        $"https://sessiondirectory.xboxlive.com/serviceconfigs/{Scid}" +
        $"/sessionTemplates/{TemplateName}/sessions/{SessionName}";
    public string MemberUrl    => SessionUrl + "/members/me";

    /// <summary>True if this is a solo squad (cascadesquadsession with 1 member).</summary>
    public bool IsSoloSquad => TemplateName?.Equals("cascadesquadsession", StringComparison.OrdinalIgnoreCase) == true;
}

// ── Game Server Redirection ───────────────────────────────────────────────────
public class GameServerInfo
{
    [JsonPropertyName("partyId")]
    public string PartyId { get; set; } = "";

    [JsonPropertyName("serverId")]
    public string ServerId { get; set; } = "";

    [JsonPropertyName("vmId")]
    public string VmId { get; set; } = "";

    [JsonPropertyName("ipv4Address")]
    public string IPv4Address { get; set; } = "";

    [JsonPropertyName("fqdn")]
    public string FQDN { get; set; } = "";

    [JsonPropertyName("ports")]
    public List<GameServerPort> Ports { get; set; } = new();

    [JsonPropertyName("region")]
    public string Region { get; set; } = "";

    [JsonPropertyName("state")]
    public string State { get; set; } = "";

    [JsonPropertyName("buildId")]
    public string BuildId { get; set; } = "";

    [JsonPropertyName("dtlsCertificateSHA2Thumbprint")]
    public string DTLSCertificateSHA2Thumbprint { get; set; } = "";

    [JsonPropertyName("cachedAt")]
    public DateTime CachedAt { get; set; } = DateTime.UtcNow;

    // Display helpers
    public string ServerShort => $"{IPv4Address}:{(Ports.FirstOrDefault()?.Num ?? 0)}";
    public string CachedAtStr => CachedAt.ToLocalTime().ToString("HH:mm:ss");
}

public class GameServerPort
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("num")]
    public int Num { get; set; }

    [JsonPropertyName("protocol")]
    public string Protocol { get; set; } = "";
}

// ── Theater Backups ───────────────────────────────────────────────────────────
public class TheaterClip : INotifyPropertyChanged
{
    private bool    _isSelected;
    private bool    _sourcePresent;
    private bool    _isBackedUp;
    private string? _customName;
    private bool    _isRenaming;
    private string  _pendingName = "";

    public string Game          { get; init; } = "";   // "Halo 3", "Halo 2: Anniv.", etc.
    public string GameKey       { get; init; } = "";   // "Halo3", "Halo2A", etc.
    public string FileName      { get; init; } = "";   // e.g. "asq_guardia_abc123.mov"
    public string MapName       { get; init; } = "";   // raw filename without extension
    public string MapDisplayName { get; init; } = "";  // resolved human-readable map name
    public long   FileSizeBytes { get; init; }
    public DateTime RecordedAt  { get; init; }
    public string SourcePath    { get; init; } = "";
    public string BackupPath    { get; init; } = "";

    public SolidColorBrush GameBrush { get; init; } = new(Colors.Gray);

    // ── Computed display strings ───────────────────────────────────────────────
    /// <summary>The name shown in the UI: custom override if set, else resolved map name.</summary>
    public string DisplayName => _customName ?? MapDisplayName;

    public string FileSizeStr => FileSizeBytes >= 1_048_576
        ? $"{FileSizeBytes / 1_048_576.0:F1} MB"
        : $"{FileSizeBytes / 1024.0:F0} KB";

    public string DateShort => RecordedAt.ToString("MM-dd HH:mm");

    public string SrcIndicator => _sourcePresent ? "✓" : "—";

    public SolidColorBrush SrcBrush => _sourcePresent
        ? new SolidColorBrush(Color.FromRgb(0x3F, 0xB9, 0x50))
        : new SolidColorBrush(Color.FromRgb(0x48, 0x4F, 0x58));

    // ── Notifying properties ───────────────────────────────────────────────────
    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; Notify(nameof(IsSelected)); }
    }

    public bool SourcePresent
    {
        get => _sourcePresent;
        set
        {
            _sourcePresent = value;
            Notify(nameof(SourcePresent));
            Notify(nameof(SrcIndicator));
            Notify(nameof(SrcBrush));
        }
    }

    public bool IsBackedUp
    {
        get => _isBackedUp;
        set { _isBackedUp = value; Notify(nameof(IsBackedUp)); }
    }

    /// <summary>User-defined display name override. Null clears the override (reverts to MapDisplayName).</summary>
    public string? CustomName
    {
        get => _customName;
        set
        {
            _customName = string.IsNullOrWhiteSpace(value) ? null : value.Trim();
            Notify(nameof(CustomName));
            Notify(nameof(DisplayName));
        }
    }

    /// <summary>True while the inline rename TextBox is visible for this row.</summary>
    public bool IsRenaming
    {
        get => _isRenaming;
        set
        {
            if (value) PendingName = DisplayName; // pre-fill before showing box
            _isRenaming = value;
            Notify(nameof(IsRenaming));
        }
    }

    /// <summary>Temporary edit value bound to the inline rename TextBox.</summary>
    public string PendingName
    {
        get => _pendingName;
        set { _pendingName = value; Notify(nameof(PendingName)); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void Notify(string name) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

// ── Session Intel ─────────────────────────────────────────────────────────────
public class SessionMember
{
    public string Gamertag    { get; set; } = "";
    public string Xuid        { get; set; } = "";
    public string InitialTeam { get; set; } = "";
    public string PartyId     { get; set; } = "";

    public double MinSkillPct { get; set; }
    public double MaxSkillPct { get; set; }
    public double AvgSkillPct { get; set; }

    // Set after grouping analysis
    public int PartyIndex { get; set; } = -1;   // -1 = solo
    public int PartySize  { get; set; } = 1;

    // ── Display strings ───────────────────────────────────────────────────────
    public string AvgPctStr    => AvgSkillPct > 0 ? $"{AvgSkillPct:F1}%" : "—";
    public string RangeStr     => AvgSkillPct > 0 ? $"{MinSkillPct:F0}–{MaxSkillPct:F0}%" : "—";
    public string PartySizeStr => PartySize > 1 ? $"×{PartySize}" : "solo";

    // Skill bar width: column is 100px, map 0-100% → 0-90px
    public double SkillBarWidth => Math.Clamp(AvgSkillPct, 0, 100) * 0.90;

    // ── Brush palette for parties ─────────────────────────────────────────────
    private static readonly SolidColorBrush[] _partyPalette =
    [
        new SolidColorBrush(Color.FromRgb(0x39, 0xD0, 0xC8)), // teal
        new SolidColorBrush(Color.FromRgb(0xD2, 0x99, 0x22)), // yellow
        new SolidColorBrush(Color.FromRgb(0x58, 0xA6, 0xFF)), // blue
        new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49)), // red
        new SolidColorBrush(Color.FromRgb(0x3F, 0xB9, 0x50)), // green
        new SolidColorBrush(Color.FromRgb(0xBC, 0x8C, 0xF9)), // purple
        new SolidColorBrush(Color.FromRgb(0xF0, 0x88, 0x3E)), // orange
        new SolidColorBrush(Color.FromRgb(0xFF, 0x7E, 0xC6)), // pink
    ];

    private static readonly SolidColorBrush _soloBrush =
        new SolidColorBrush(Color.FromRgb(0x48, 0x4F, 0x58));

    public SolidColorBrush PartyBrush =>
        PartySize > 1 && PartyIndex >= 0 && PartyIndex < _partyPalette.Length
            ? _partyPalette[PartyIndex]
            : _soloBrush;

    public SolidColorBrush TeamBrush => InitialTeam switch
    {
        "1" or "Eagle" => new SolidColorBrush(Color.FromRgb(0x58, 0xA6, 0xFF)),
        "2" or "Cobra" => new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49)),
        _ when InitialTeam.Contains("blue", StringComparison.OrdinalIgnoreCase)
               => new SolidColorBrush(Color.FromRgb(0x58, 0xA6, 0xFF)),
        _ when InitialTeam.Contains("red",  StringComparison.OrdinalIgnoreCase)
               => new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49)),
        _      => new SolidColorBrush(Color.FromRgb(0x7D, 0x85, 0x90)),
    };

    public SolidColorBrush SkillBrush
    {
        get
        {
            double s = AvgSkillPct;
            if (s >= 90) return new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49)); // red
            if (s >= 75) return new SolidColorBrush(Color.FromRgb(0xF9, 0x73, 0x16)); // orange
            if (s >= 50) return new SolidColorBrush(Color.FromRgb(0xD2, 0x99, 0x22)); // yellow
            return new SolidColorBrush(Color.FromRgb(0x3F, 0xB9, 0x50));               // green
        }
    }
}

public class PlaylistSummary
{
    public string Id { get; init; } = "";
    public string RawName { get; init; } = "";
    public string HopperName { get; init; } = "";
    public bool IsMix { get; init; }
    public string MinPlayers { get; init; } = "";
    public string MaxPlayers { get; init; } = "";
    public string MaxPartySize { get; init; } = "";
    public List<PlaylistTagGroup> TagGroups { get; init; } = new();

    public string DisplayName
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(HopperName))
                return HopperName;
            if (!string.IsNullOrWhiteSpace(Id))
                return Id;
            return RawName;
        }
    }

    public string PlayerSummary => $"{MinPlayers}-{MaxPlayers} players  •  party up to {MaxPartySize}";

    public override string ToString() => DisplayName;
}

public class PlaylistTagGroup
{
    public string TagId { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Subtitle { get; init; } = "";
    public int TotalWeight { get; init; }
    public List<PlaylistEntry> Entries { get; init; } = new();

    public string EntrySummary => $"{Entries.Count} entries";
}

public class PlaylistEntry
{
    public string TagId { get; init; } = "";
    public string Category { get; init; } = "";
    public int Weight { get; init; }
    public string MapId { get; init; } = "";
    public string MapDisplay { get; init; } = "";
    public string MapVariant { get; init; } = "";
    public string GameVariant { get; init; } = "";
    public string GameKey { get; init; } = "";
    public double WeightShare { get; set; }

    public string WeightShareDisplay => $"{WeightShare * 100:F1}%";

    public string GameDisplay => GameKey switch
    {
        "halo1" => "Halo CE",
        "halo2" => "Halo 2",
        "halo2a" => "Halo 2A",
        "halo3" => "Halo 3",
        "halo3odst" => "Halo 3: ODST",
        "halo4" => "Halo 4",
        "haloreach" => "Halo: Reach",
        _ => "Unknown"
    };

    public string GameVariantDisplay => string.IsNullOrWhiteSpace(GameVariant) ? "Default" : GameVariant;
    public string MapVariantDisplay => string.IsNullOrWhiteSpace(MapVariant) ? "-" : MapVariant;
    public string CategoryDisplay => string.IsNullOrWhiteSpace(Category) ? "-" : Category.Replace("GameCategory_", "");

    public SolidColorBrush GameBrush => GameKey switch
    {
        "halo1" => new SolidColorBrush(Color.FromRgb(0x8C, 0xC8, 0xFF)),
        "halo2" => new SolidColorBrush(Color.FromRgb(0x7A, 0xD1, 0xFF)),
        "halo2a" => new SolidColorBrush(Color.FromRgb(0x39, 0xD0, 0xC8)),
        "halo3" => new SolidColorBrush(Color.FromRgb(0x58, 0xA6, 0xFF)),
        "halo3odst" => new SolidColorBrush(Color.FromRgb(0xD2, 0x99, 0x22)),
        "halo4" => new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49)),
        "haloreach" => new SolidColorBrush(Color.FromRgb(0xBC, 0x8C, 0xF9)),
        _ => new SolidColorBrush(Color.FromRgb(0x7D, 0x85, 0x90))
    };
}
