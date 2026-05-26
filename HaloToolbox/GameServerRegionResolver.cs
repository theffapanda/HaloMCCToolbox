using System.Collections.Concurrent;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace HaloToolbox;

public static class GameServerRegionResolver
{
    private const string AzurePublicServiceTagsPageUrl =
        "https://www.microsoft.com/en-us/download/details.aspx?id=56519";
    private const string AzurePublicServiceTagsFallbackUrl =
        "https://download.microsoft.com/download/7/1/d/71d86715-5596-4529-9b13-da13a5de5b63/ServiceTags_Public_20260511.json";
    private static readonly TimeSpan AzureServiceTagsMaxAge = TimeSpan.FromDays(8);
    private static readonly string AzureServiceTagsCacheFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox",
        "ServiceTags_Public.json");
    private static readonly HttpClient Http = new()
    {
        Timeout = TimeSpan.FromSeconds(10)
    };

    private static readonly ConcurrentDictionary<string, string> IpRegionCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, byte> PendingIpLookups = new(StringComparer.OrdinalIgnoreCase);
    private static readonly object AzureRangeLock = new();
    private static IReadOnlyList<AzureRegionRange> AzureRegionRanges = Array.Empty<AzureRegionRange>();
    private static bool AzureRangesLoaded;
    private static Task? AzureRangesLoadTask;

    private static readonly Dictionary<string, string> RegionLabels = new(StringComparer.OrdinalIgnoreCase)
    {
        ["WestUs"] = "West US",
        ["SouthCentralUs"] = "South Central US",
        ["CentralUs"] = "Central US",
        ["NorthCentralUs"] = "North Central US",
        ["EastUs"] = "East US",
        ["EastUs2"] = "East US 2",
        ["BrazilSouth"] = "Brazil South",
        ["NorthEurope"] = "North Europe",
        ["WestEurope"] = "West Europe",
        ["SoutheastAsia"] = "Southeast Asia",
        ["EastAsia"] = "East Asia",
        ["JapanWest"] = "Japan West",
        ["JapanEast"] = "Japan East",
        ["AustraliaSoutheast"] = "Australia Southeast",
        ["AustraliaEast"] = "Australia East"
    };

    private static readonly Dictionary<string, string> AzureHostRegionLabels = new(StringComparer.OrdinalIgnoreCase)
    {
        ["westus"] = "West US",
        ["southcentralus"] = "South Central US",
        ["centralus"] = "Central US",
        ["northcentralus"] = "North Central US",
        ["eastus"] = "East US",
        ["eastus2"] = "East US 2",
        ["brazilsouth"] = "Brazil South",
        ["northeurope"] = "North Europe",
        ["westeurope"] = "West Europe",
        ["southeastasia"] = "Southeast Asia",
        ["eastasia"] = "East Asia",
        ["japanwest"] = "Japan West",
        ["japaneast"] = "Japan East",
        ["australiasoutheast"] = "Australia Southeast",
        ["australiaeast"] = "Australia East"
    };

    public static string GetRegionLabel(GameServerInfo? serverInfo)
    {
        if (serverInfo is null)
            return "";

        if (TryNormalizeRegion(serverInfo.Region, out var label))
            return label;

        if (TryExtractVmRegion(serverInfo.VmId, out label))
            return label;

        if (TryExtractHostRegion(serverInfo.FQDN, out label))
            return label;

        if (!string.IsNullOrWhiteSpace(serverInfo.IPv4Address) &&
            IpRegionCache.TryGetValue(serverInfo.IPv4Address, out label))
        {
            return label;
        }

        return "";
    }

    public static string GetCachedOrQueueIpRegion(string ip)
    {
        if (string.IsNullOrWhiteSpace(ip))
            return "";

        if (IpRegionCache.TryGetValue(ip, out var label))
            return label;

        if (TryResolveAzureRegion(ip, out label))
        {
            IpRegionCache[ip] = label;
            return label;
        }

        if (PendingIpLookups.TryAdd(ip, 0))
            _ = ResolveIpRegionAsync(ip);

        return "";
    }

    private static async Task ResolveIpRegionAsync(string ip)
    {
        try
        {
            await EnsureAzureRangesLoadedAsync();
            if (TryResolveAzureRegion(ip, out var azureLabel))
            {
                IpRegionCache[ip] = azureLabel;
                return;
            }

            var entry = await Dns.GetHostEntryAsync(ip);
            foreach (var host in entry.Aliases.Prepend(entry.HostName))
            {
                if (TryExtractHostRegion(host, out var label))
                {
                    IpRegionCache[ip] = label;
                    return;
                }
            }
        }
        catch
        {
            // Reverse DNS is best-effort; observed UDP still works without a region label.
        }
        finally
        {
            PendingIpLookups.TryRemove(ip, out _);
        }
    }

    private static async Task EnsureAzureRangesLoadedAsync()
    {
        lock (AzureRangeLock)
        {
            if (AzureRangesLoaded)
                return;

            AzureRangesLoadTask ??= LoadAzureRangesAsync();
        }

        await AzureRangesLoadTask;
    }

    private static async Task LoadAzureRangesAsync()
    {
        try
        {
            string json = "";
            if (IsFreshCache(AzureServiceTagsCacheFile))
            {
                json = await File.ReadAllTextAsync(AzureServiceTagsCacheFile);
            }
            else
            {
                json = await DownloadAzureServiceTagsJsonAsync();
                Directory.CreateDirectory(Path.GetDirectoryName(AzureServiceTagsCacheFile)!);
                await File.WriteAllTextAsync(AzureServiceTagsCacheFile, json);
            }

            var ranges = ParseAzureRegionRanges(json);
            lock (AzureRangeLock)
            {
                AzureRegionRanges = ranges;
                AzureRangesLoaded = true;
            }

            RejoinFixDiagnostics.Info("network", $"Loaded {ranges.Count} Azure service tag ranges for server region labels.");
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("network", $"Failed to load Azure service tags for server region labels: {ex.Message}");
            if (await TryLoadCachedAzureRangesAsync())
                return;

            lock (AzureRangeLock)
            {
                AzureRangesLoaded = true;
            }
        }
    }

    private static async Task<string> DownloadAzureServiceTagsJsonAsync()
    {
        var urls = new List<string>();

        try
        {
            var pageHtml = await Http.GetStringAsync(AzurePublicServiceTagsPageUrl);
            urls.AddRange(Regex.Matches(
                    pageHtml,
                    @"https://download\.microsoft\.com/download/[^""'<>\s]+/ServiceTags_Public_\d{8}\.json",
                    RegexOptions.IgnoreCase)
                .Select(match => match.Value));
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn("network", $"Failed to discover current Azure service tags URL: {ex.Message}");
        }

        urls.Add(AzurePublicServiceTagsFallbackUrl);

        Exception? lastError = null;
        foreach (var url in urls.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var json = await Http.GetStringAsync(url);
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("values", out _))
                    return json;
            }
            catch (Exception ex)
            {
                lastError = ex;
            }
        }

        throw new InvalidOperationException("No Azure service tags JSON download succeeded.", lastError);
    }

    private static async Task<bool> TryLoadCachedAzureRangesAsync()
    {
        try
        {
            if (!File.Exists(AzureServiceTagsCacheFile))
                return false;

            var json = await File.ReadAllTextAsync(AzureServiceTagsCacheFile);
            var ranges = ParseAzureRegionRanges(json);
            lock (AzureRangeLock)
            {
                AzureRegionRanges = ranges;
                AzureRangesLoaded = true;
            }

            return ranges.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsFreshCache(string path)
    {
        try
        {
            return File.Exists(path) &&
                DateTime.UtcNow - File.GetLastWriteTimeUtc(path) <= AzureServiceTagsMaxAge;
        }
        catch
        {
            return false;
        }
    }

    private static IReadOnlyList<AzureRegionRange> ParseAzureRegionRanges(string json)
    {
        var ranges = new List<AzureRegionRange>();
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("values", out var values) || values.ValueKind != JsonValueKind.Array)
            return ranges;

        foreach (var value in values.EnumerateArray())
        {
            if (!value.TryGetProperty("properties", out var properties))
                continue;

            var region = GetJsonString(properties, "region");
            if (string.IsNullOrWhiteSpace(region) || !TryNormalizeRegion(region, out var label))
                continue;

            var systemService = GetJsonString(properties, "systemService");
            var name = GetJsonString(value, "name");
            if (!string.Equals(systemService, "AzureCloud", StringComparison.OrdinalIgnoreCase) &&
                !name.StartsWith("AzureCloud.", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!properties.TryGetProperty("addressPrefixes", out var prefixes) ||
                prefixes.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            foreach (var prefix in prefixes.EnumerateArray())
            {
                var cidr = prefix.GetString();
                if (TryParseCidr(cidr, out var range))
                    ranges.Add(range with { Label = label });
            }
        }

        return ranges;
    }

    private static string GetJsonString(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    private static bool TryResolveAzureRegion(string ip, out string label)
    {
        label = "";
        if (!IPAddress.TryParse(ip, out var address) || address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork)
            return false;

        var ipValue = ToUInt32(address);
        IReadOnlyList<AzureRegionRange> ranges;
        lock (AzureRangeLock)
            ranges = AzureRegionRanges;

        foreach (var range in ranges)
        {
            if ((ipValue & range.Mask) == range.Network)
            {
                label = range.Label;
                return true;
            }
        }

        return false;
    }

    private static bool TryParseCidr(string? cidr, out AzureRegionRange range)
    {
        range = default;
        if (string.IsNullOrWhiteSpace(cidr))
            return false;

        var parts = cidr.Split('/');
        if (parts.Length != 2 ||
            !IPAddress.TryParse(parts[0], out var address) ||
            address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork ||
            !int.TryParse(parts[1], out int prefixLength) ||
            prefixLength is < 0 or > 32)
        {
            return false;
        }

        uint mask = prefixLength == 0 ? 0 : uint.MaxValue << (32 - prefixLength);
        uint network = ToUInt32(address) & mask;
        range = new AzureRegionRange(network, mask, "");
        return true;
    }

    private static uint ToUInt32(IPAddress address)
    {
        var bytes = address.GetAddressBytes();
        return ((uint)bytes[0] << 24) |
            ((uint)bytes[1] << 16) |
            ((uint)bytes[2] << 8) |
            bytes[3];
    }

    private static bool TryNormalizeRegion(string region, out string label)
    {
        label = "";
        if (string.IsNullOrWhiteSpace(region) ||
            region.Equals("ACTIVE UDP", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (RegionLabels.TryGetValue(region.Trim(), out var knownLabel))
        {
            label = knownLabel;
            return true;
        }

        label = region.Trim();
        return true;
    }

    private static bool TryExtractVmRegion(string vmId, out string label)
    {
        label = "";
        if (string.IsNullOrWhiteSpace(vmId))
            return false;

        var parts = vmId.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return parts.Length >= 2 && TryNormalizeRegion(parts[1], out label);
    }

    private static bool TryExtractHostRegion(string host, out string label)
    {
        label = "";
        if (string.IsNullOrWhiteSpace(host))
            return false;

        var normalizedHost = host.Trim().TrimEnd('.');
        foreach (var (azureRegion, display) in AzureHostRegionLabels)
        {
            if (normalizedHost.Contains($".{azureRegion}.", StringComparison.OrdinalIgnoreCase) ||
                normalizedHost.EndsWith($".{azureRegion}.cloudapp.azure.com", StringComparison.OrdinalIgnoreCase))
            {
                label = display;
                return true;
            }
        }

        return false;
    }

    private readonly record struct AzureRegionRange(uint Network, uint Mask, string Label);
}
