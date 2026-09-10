using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace HaloToolbox;

public enum MccInstallationKind
{
    Unknown,
    Steam,
    MicrosoftStore
}

public sealed record MccInstallationInfo(string RootPath, MccInstallationKind Kind)
{
    public string DisplayName => Kind switch
    {
        MccInstallationKind.Steam => "STEAM",
        MccInstallationKind.MicrosoftStore => "MICROSOFT STORE",
        _ => "UNKNOWN"
    };

    public string SelectionLabel => $"{DisplayName}  -  {RootPath}";

    public override string ToString() => SelectionLabel;
}

public static class MccInstallationResolver
{
    private const string MicrosoftStoreIdentity = "Microsoft.Chelan";
    private const string SteamFolderName = "Halo The Master Chief Collection";

    public static string NormalizeRoot(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return "";

        string candidate;
        try { candidate = Path.GetFullPath(path.Trim().Trim('"')); }
        catch { return path.Trim(); }

        if (DetectKindAtRoot(candidate) != MccInstallationKind.Unknown)
            return Path.TrimEndingDirectorySeparator(candidate);

        var content = Path.Combine(candidate, "Content");
        return DetectKindAtRoot(content) == MccInstallationKind.MicrosoftStore
            ? Path.TrimEndingDirectorySeparator(content)
            : Path.TrimEndingDirectorySeparator(candidate);
    }

    public static MccInstallationKind DetectKind(string? path)
    {
        var normalized = NormalizeRoot(path);
        return DetectKindAtRoot(normalized);
    }

    public static MccInstallationInfo? Inspect(string? path)
    {
        var normalized = NormalizeRoot(path);
        var kind = DetectKindAtRoot(normalized);
        return kind == MccInstallationKind.Unknown ? null : new MccInstallationInfo(normalized, kind);
    }

    public static IReadOnlyList<MccInstallationInfo> Discover(string? savedPath, string defaultSteamPath)
    {
        var candidates = new List<string>();
        AddCandidate(candidates, savedPath);
        AddCandidate(candidates, defaultSteamPath);
        AddSteamCandidates(candidates);
        AddMicrosoftStoreCandidates(candidates);

        return candidates
            .Select(Inspect)
            .Where(info => info is not null)
            .Select(info => info!)
            .GroupBy(info => info.RootPath, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToArray();
    }

    private static MccInstallationKind DetectKindAtRoot(string? root)
    {
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root) ||
            !Directory.Exists(Path.Combine(root, "halo3", "maps")))
        {
            return MccInstallationKind.Unknown;
        }

        if (File.Exists(Path.Combine(root, "MCC", "Binaries", "Win64", "MCCWinStore-Win64-Shipping.exe")) &&
            IsMicrosoftStoreManifest(root!))
        {
            return MccInstallationKind.MicrosoftStore;
        }

        if (File.Exists(Path.Combine(root, "MCC", "Binaries", "Win64", "MCC-Win64-Shipping.exe")))
            return MccInstallationKind.Steam;

        return MccInstallationKind.Unknown;
    }

    private static bool IsMicrosoftStoreManifest(string root)
    {
        try
        {
            var manifestPath = Path.Combine(root, "MicrosoftGame.config");
            if (!File.Exists(manifestPath))
                return false;

            var identity = XDocument.Load(manifestPath).Root?
                .Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Identity", StringComparison.Ordinal));
            return string.Equals(
                identity?.Attribute("Name")?.Value,
                MicrosoftStoreIdentity,
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static void AddSteamCandidates(List<string> candidates)
    {
        try
        {
            using var steamKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Valve\Steam");
            var steamPath = steamKey?.GetValue("SteamPath") as string;
            if (string.IsNullOrWhiteSpace(steamPath))
                return;

            AddCandidate(candidates, Path.Combine(steamPath, "steamapps", "common", SteamFolderName));
            var libraries = Path.Combine(steamPath, "steamapps", "libraryfolders.vdf");
            if (!File.Exists(libraries))
                return;

            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(
                File.ReadAllText(libraries),
                "\"path\"\\s+\"([^\"]+)\"",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                var library = match.Groups[1].Value.Replace(@"\\", @"\");
                AddCandidate(candidates, Path.Combine(library, "steamapps", "common", SteamFolderName));
            }
        }
        catch { }
    }

    private static void AddMicrosoftStoreCandidates(List<string> candidates)
    {
        var gamingRoots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        AddGamingRoot(gamingRoots, @"C:\XboxGames");

        for (char drive = 'A'; drive <= 'Z'; drive++)
        {
            try
            {
                var driveRoot = $@"{drive}:\";
                var markerPath = Path.Combine(driveRoot, ".GamingRoot");
                if (!File.Exists(markerPath))
                    continue;

                var marker = File.ReadAllBytes(markerPath);
                if (marker.Length <= 8)
                    continue;

                var relativeRoot = Encoding.Unicode.GetString(marker, 8, marker.Length - 8).TrimEnd('\0');
                if (!string.IsNullOrWhiteSpace(relativeRoot))
                    AddGamingRoot(gamingRoots, Path.Combine(driveRoot, relativeRoot));
            }
            catch { }
        }

        foreach (var gamingRoot in gamingRoots)
        {
            try
            {
                foreach (var gameFolder in Directory.EnumerateDirectories(gamingRoot))
                    AddCandidate(candidates, Path.Combine(gameFolder, "Content"));
            }
            catch { }
        }
    }

    private static void AddGamingRoot(HashSet<string> roots, string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        try
        {
            var fullPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
            if (Directory.Exists(fullPath))
                roots.Add(fullPath);
        }
        catch { }
    }

    private static void AddCandidate(List<string> candidates, string? path)
    {
        if (!string.IsNullOrWhiteSpace(path))
            candidates.Add(path);
    }
}

internal static class MccProcessLocator
{
    private static readonly string[] RuntimeProcessNames =
    [
        "MCC-Win64-Shipping",
        "MCCWinStore-Win64-Shipping",
        "MCC"
    ];

    public static IEnumerable<Process> GetRuntimeProcesses(string? preferredRoot = null)
    {
        var processes = RuntimeProcessNames.SelectMany(name =>
        {
            try { return Process.GetProcessesByName(name); }
            catch { return []; }
        });

        var preferredKind = MccInstallationResolver.DetectKind(preferredRoot);
        return processes
            .OrderByDescending(process => GetPreference(process, preferredKind))
            .ThenByDescending(GetStartTime)
            .ToArray();
    }

    public static Process? GetLatestRuntimeProcess(string? preferredRoot = null)
    {
        var processes = GetRuntimeProcesses(preferredRoot).ToArray();
        if (processes.Length == 0)
            return null;

        var selected = processes[0];
        for (int index = 1; index < processes.Length; index++)
            processes[index].Dispose();

        return selected;
    }

    private static int GetPreference(Process process, MccInstallationKind preferredKind)
    {
        string processName;
        try { processName = process.ProcessName; }
        catch { return 0; }

        var isSteamRuntime = processName.Equals("MCC-Win64-Shipping", StringComparison.OrdinalIgnoreCase);
        var isStoreRuntime = processName.Equals("MCCWinStore-Win64-Shipping", StringComparison.OrdinalIgnoreCase);
        var isLegacyRuntime = processName.Equals("MCC", StringComparison.OrdinalIgnoreCase);

        return preferredKind switch
        {
            MccInstallationKind.Steam when isSteamRuntime => 2,
            MccInstallationKind.MicrosoftStore when isStoreRuntime => 2,
            MccInstallationKind.Steam or MccInstallationKind.MicrosoftStore when isLegacyRuntime => 1,
            MccInstallationKind.Unknown when isSteamRuntime || isStoreRuntime || isLegacyRuntime => 1,
            _ => 0
        };
    }

    private static DateTime GetStartTime(Process process)
    {
        try { return process.StartTime; }
        catch { return DateTime.MinValue; }
    }
}
