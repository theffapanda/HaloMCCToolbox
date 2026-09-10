using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace HaloToolbox;

internal enum VpnConnectionState
{
    NotInstalled,
    Ready,
    Installing,
    Connecting,
    Connected,
    Disconnecting,
    Error
}

internal sealed record NordVpnServer(
    string Hostname,
    string Station,
    string City,
    string Country,
    int Load,
    double Latitude,
    double Longitude)
{
    public double DistanceKm { get; init; }
    public int? LatencyMs { get; init; }
    public string LatencyText => LatencyMs is int latency ? $"{latency} ms" : "test unavailable";
    public string DisplayName => $"{LatencyText}  ·  {City}  ·  {DistanceKm:0} km  ·  {Hostname}  ·  {Load}% load";
    public override string ToString() => DisplayName;
}

internal sealed record HaloVpnRegion(
    string Name,
    string Country,
    double Latitude,
    double Longitude)
{
    public override string ToString() => Name;
}

internal static class VpnConnectionPresence
{
    public static event EventHandler? Changed;

    public static string ConnectedRegion { get; private set; } = "";

    public static void SetConnected(string region)
    {
        string normalized = region.Trim();
        if (string.Equals(ConnectedRegion, normalized, StringComparison.Ordinal))
            return;
        ConnectedRegion = normalized;
        Changed?.Invoke(null, EventArgs.Empty);
    }

    public static void Clear()
    {
        if (ConnectedRegion.Length == 0)
            return;
        ConnectedRegion = "";
        Changed?.Invoke(null, EventArgs.Empty);
    }
}

internal sealed class VpnUserSettings
{
    public string Region { get; set; } = "East US";
    public string Username { get; set; } = "";
    public string PasswordEncrypted { get; set; } = "";
    public bool RememberCredentials { get; set; } = true;
    public bool TemporaryCredentials { get; set; }
    public string LastServerHostname { get; set; } = "";
    public string ImportedProfilePath { get; set; } = "";

    [JsonIgnore]
    public string Password
    {
        get => VpnSecretProtection.Decrypt(PasswordEncrypted);
        set => PasswordEncrypted = VpnSecretProtection.Encrypt(value);
    }
}

internal static class VpnSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static string VpnRoot => Path.Combine(App.ToolboxDataRoot, "Vpn");
    public static string SettingsPath => Path.Combine(VpnRoot, "settings.json");
    public static string RouterPath => Path.Combine(VpnRoot, "router.json");
    public static string ProfilesRoot => Path.Combine(VpnRoot, "profiles");
    public static string RuntimeRoot => Path.Combine(VpnRoot, "runtime");

    public static VpnUserSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath))
                return new VpnUserSettings();

            VpnUserSettings settings = JsonSerializer.Deserialize<VpnUserSettings>(
                                           File.ReadAllText(SettingsPath),
                                           JsonOptions)
                                       ?? new VpnUserSettings();
            if (settings.TemporaryCredentials)
            {
                // Return the one-time secret to this process, but remove it
                // from disk immediately. The elevated relaunch already loads
                // settings during control construction, before connecting.
                string encrypted = settings.PasswordEncrypted;
                settings.PasswordEncrypted = "";
                settings.TemporaryCredentials = false;
                Save(settings);
                settings.PasswordEncrypted = encrypted;
            }
            return settings;
        }
        catch
        {
            return new VpnUserSettings();
        }
    }

    public static void Save(VpnUserSettings settings)
    {
        Directory.CreateDirectory(VpnRoot);
        File.WriteAllText(SettingsPath, JsonSerializer.Serialize(settings, JsonOptions));
    }

    public static void WriteRouterConfiguration()
    {
        Directory.CreateDirectory(VpnRoot);
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        string mccRoot = App.LoadMccInstallationPath();
        AddIfPresent(paths, Path.Combine(mccRoot, "mcclauncher.exe"));
        AddIfPresent(paths, Path.Combine(mccRoot, "MCC", "Binaries", "Win64", "MCC-Win64-Shipping.exe"));
        AddIfPresent(paths, Path.Combine(mccRoot, "MCC", "Binaries", "Win64", "MCCWinStore-Win64-Shipping.exe"));

        foreach (string processName in new[] { "MCC-Win64-Shipping", "MCCWinStore-Win64-Shipping", "mcclauncher" })
        {
            foreach (Process process in Process.GetProcessesByName(processName))
            {
                try { AddIfPresent(paths, process.MainModule?.FileName); }
                catch { }
                finally { process.Dispose(); }
            }
        }

        var payload = new
        {
            tunneledApps = paths.Select(path => new { exePath = path }).ToArray()
        };
        File.WriteAllText(
            RouterPath,
            JsonSerializer.Serialize(payload, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true }));
    }

    private static void AddIfPresent(HashSet<string> paths, string? path)
    {
        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            paths.Add(Path.GetFullPath(path));
    }
}

internal static class VpnSecretProtection
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("HaloMCCToolbox.NordVpn.v1");

    public static string Encrypt(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        byte[] clear = Encoding.UTF8.GetBytes(value);
        byte[] protectedBytes = ProtectedData.Protect(clear, Entropy, DataProtectionScope.CurrentUser);
        CryptographicOperations.ZeroMemory(clear);
        return Convert.ToBase64String(protectedBytes);
    }

    public static string Decrypt(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "";

        try
        {
            byte[] protectedBytes = Convert.FromBase64String(value);
            byte[] clear = ProtectedData.Unprotect(protectedBytes, Entropy, DataProtectionScope.CurrentUser);
            string result = Encoding.UTF8.GetString(clear);
            CryptographicOperations.ZeroMemory(clear);
            return result;
        }
        catch
        {
            return "";
        }
    }
}

internal sealed class VpnService : IDisposable
{
    private const string OpenVpnInstallerUrl =
        "https://swupdate.openvpn.org/community/releases/OpenVPN-2.7.6-I001-amd64.msi";
    private const string OpenVpnInstallerSha256 =
        "48C96AC092A81C6303F059EACE51C7BF878794E9575B329E82894848E0E529FC";
    private const string ServerCatalogUrl =
        "https://api.nordvpn.com/v1/servers?limit=10000" +
        "&filters%5Bservers.status%5D=online" +
        "&filters%5Bservers_technologies%5D%5Bidentifier%5D=openvpn_udp" +
        "&filters%5Bservers_groups%5D%5Bidentifier%5D=legacy_standard" +
        "&fields%5Bservers.hostname%5D" +
        "&fields%5Bservers.station%5D" +
        "&fields%5Bservers.load%5D" +
        "&fields%5Bservers.status%5D" +
        "&fields%5Bservers.locations.country.name%5D" +
        "&fields%5Bservers.locations.country.city.name%5D" +
        "&fields%5Bservers.locations.country.city.latitude%5D" +
        "&fields%5Bservers.locations.country.city.longitude%5D";

    private static readonly HttpClient Http = CreateHttpClient();
    private static readonly SemaphoreSlim CatalogLock = new(1, 1);
    private const int LatencyCandidateCount = 12;
    private const int LatencyProbeCount = 3;
    private static IReadOnlyList<NordVpnServer>? CachedCatalog;
    private static DateTime CatalogExpiresUtc;
    private readonly SemaphoreSlim _operationLock = new(1, 1);
    private readonly object _routerLogLock = new();
    private Process? _openVpnProcess;
    private Process? _routerProcess;
    private string? _authPath;
    private bool _disposed;
    private volatile bool _stopping;

    public event EventHandler? StateChanged;

    public VpnConnectionState State { get; private set; }
    public string Status { get; private set; } = "Ready.";
    public string TunnelIp { get; private set; } = "—";
    public NordVpnServer? ConnectedServer { get; private set; }
    public bool IsConnected => State == VpnConnectionState.Connected;
    public bool IsOpenVpnInstalled => FindOpenVpnExecutable() is not null;
    public bool IsRouterBundled => VpnNativeResources.IsAvailable;

    public static IReadOnlyList<HaloVpnRegion> HaloRegions { get; } =
    [
        new("East US", "United States", 37.5407, -77.4360),       // Richmond, Virginia
        new("East US 2", "United States", 37.5407, -77.4360),     // Richmond, Virginia
        new("South Central US", "United States", 29.4241, -98.4936), // San Antonio, Texas
        new("Central US", "United States", 41.5868, -93.6250),    // Des Moines, Iowa
        new("North Central US", "United States", 41.8781, -87.6298), // Chicago, Illinois
        new("West US", "United States", 37.7749, -122.4194),      // San Francisco, California
        new("Brazil South", "Brazil", -23.5505, -46.6333),
        new("North Europe", "Ireland", 53.3498, -6.2603),
        new("West Europe", "Netherlands", 52.3676, 4.9041),
        new("Southeast Asia", "Singapore", 1.3521, 103.8198),
        new("East Asia", "Hong Kong", 22.3193, 114.1694),
        new("Japan West", "Japan", 34.6937, 135.5023),
        new("Japan East", "Japan", 35.6762, 139.6503),
        new("Australia Southeast", "Australia", -37.8136, 144.9631),
        new("Australia East", "Australia", -33.8688, 151.2093)
    ];

    public VpnService()
    {
        State = IsOpenVpnInstalled && IsRouterBundled
            ? VpnConnectionState.Ready
            : VpnConnectionState.NotInstalled;
        Status = State == VpnConnectionState.Ready
            ? "Ready to connect Halo MCC."
            : "VPN support needs one-time setup.";
    }

    public static bool IsAdministrator()
    {
        using var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
        var principal = new System.Security.Principal.WindowsPrincipal(identity);
        return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
    }

    public static bool IsNordDesktopAppInstalled()
    {
        string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (Directory.Exists(Path.Combine(programFiles, "NordVPN")))
            return true;

        foreach (Process process in Process.GetProcesses())
        {
            try
            {
                if (process.ProcessName.Contains("nordvpn", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }
            finally { process.Dispose(); }
        }
        return false;
    }

    public async Task<IReadOnlyList<NordVpnServer>> GetNearbyServersAsync(
        string region,
        CancellationToken cancellationToken = default)
    {
        HaloVpnRegion target = HaloRegions.FirstOrDefault(item =>
                                   item.Name.Equals(region, StringComparison.OrdinalIgnoreCase))
                               ?? HaloRegions[0];
        IReadOnlyList<NordVpnServer> catalog = await GetServerCatalogAsync(cancellationToken)
            .ConfigureAwait(false);
        return await Task.Run<IReadOnlyList<NordVpnServer>>(
                () => catalog
                    .Where(server => server.Country.Equals(target.Country, StringComparison.OrdinalIgnoreCase))
                    .Select(server => server with
                    {
                        DistanceKm = DistanceKm(
                            target.Latitude,
                            target.Longitude,
                            server.Latitude,
                            server.Longitude)
                    })
                    .OrderBy(server => server.DistanceKm)
                    .ThenBy(server => server.Load)
                    .ThenBy(server => server.Hostname, StringComparer.OrdinalIgnoreCase)
                    .Take(LatencyCandidateCount)
                    .ToArray(),
                cancellationToken)
            .ConfigureAwait(false);
    }

    public Task<NordVpnServer> MeasureServerLatencyAsync(
        NordVpnServer server,
        bool quick,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => MeasureLatencyAsync(server, quick ? 1 : LatencyProbeCount, cancellationToken),
            cancellationToken);

    public static IReadOnlyList<NordVpnServer> RankServers(IEnumerable<NordVpnServer> servers) =>
        servers
            .OrderBy(server => server.LatencyMs.HasValue ? 0 : 1)
            .ThenBy(server => server.LatencyMs ?? int.MaxValue)
            .ThenBy(server => server.DistanceKm)
            .ThenBy(server => server.Load)
            .ThenBy(server => server.Hostname, StringComparer.OrdinalIgnoreCase)
            .ToArray();

    private static async Task<NordVpnServer> MeasureLatencyAsync(
        NordVpnServer server,
        int probeCount,
        CancellationToken cancellationToken)
    {
        string destination = IPAddress.TryParse(server.Station, out _)
            ? server.Station
            : server.Hostname;
        var samples = new List<long>(probeCount);
        string curlPath = Path.Combine(Environment.SystemDirectory, "curl.exe");
        if (!File.Exists(curlPath))
            return server;

        for (int attempt = 0; attempt < probeCount; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = curlPath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                }
            };
            process.StartInfo.ArgumentList.Add("--ipv4");
            process.StartInfo.ArgumentList.Add("--noproxy");
            process.StartInfo.ArgumentList.Add("*");
            process.StartInfo.ArgumentList.Add("--connect-timeout");
            process.StartInfo.ArgumentList.Add("1.2");
            process.StartInfo.ArgumentList.Add("--max-time");
            process.StartInfo.ArgumentList.Add("2");
            process.StartInfo.ArgumentList.Add("--insecure");
            process.StartInfo.ArgumentList.Add("--silent");
            process.StartInfo.ArgumentList.Add("--output");
            process.StartInfo.ArgumentList.Add("NUL");
            process.StartInfo.ArgumentList.Add("--write-out");
            process.StartInfo.ArgumentList.Add("%{time_connect}");
            process.StartInfo.ArgumentList.Add($"https://{destination}:443/");

            try
            {
                if (!process.Start())
                    break;
                VpnJobObject.Assign(process);
                string output = await process.StandardOutput.ReadToEndAsync(cancellationToken)
                    .ConfigureAwait(false);
                await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
                if (double.TryParse(output, NumberStyles.Float, CultureInfo.InvariantCulture, out double seconds) &&
                    seconds > 0)
                {
                    samples.Add(Math.Max(1, (long)Math.Round(seconds * 1000, MidpointRounding.AwayFromZero)));
                }
            }
            catch (OperationCanceledException)
            {
                TryKill(process);
                throw;
            }
            catch
            {
                break;
            }
        }

        if (samples.Count == 0)
            return server;

        samples.Sort();
        return server with { LatencyMs = (int)samples[samples.Count / 2] };
    }

    private static async Task<IReadOnlyList<NordVpnServer>> GetServerCatalogAsync(
        CancellationToken cancellationToken)
    {
        if (CachedCatalog is not null && DateTime.UtcNow < CatalogExpiresUtc)
            return CachedCatalog;

        await CatalogLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (CachedCatalog is not null && DateTime.UtcNow < CatalogExpiresUtc)
                return CachedCatalog;

            using HttpResponseMessage response = await Http.GetAsync(ServerCatalogUrl, cancellationToken)
                .ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
            await using Stream stream = await response.Content.ReadAsStreamAsync(cancellationToken)
                .ConfigureAwait(false);
            var records = await JsonSerializer.DeserializeAsync<List<NordServerResponse>>(
                              stream,
                              new JsonSerializerOptions { PropertyNameCaseInsensitive = true },
                              cancellationToken)
                              .ConfigureAwait(false)
                          ?? new List<NordServerResponse>();
            CachedCatalog = records
                .Where(record => string.Equals(record.Status, "online", StringComparison.OrdinalIgnoreCase))
                .Select(ToServer)
                .Where(server => server is not null)
                .Cast<NordVpnServer>()
                .ToArray();
            CatalogExpiresUtc = DateTime.UtcNow.AddMinutes(10);
            return CachedCatalog;
        }
        finally
        {
            CatalogLock.Release();
        }
    }

    private static double DistanceKm(double latitude1, double longitude1, double latitude2, double longitude2)
    {
        const double EarthRadiusKm = 6371;
        double lat1 = latitude1 * Math.PI / 180;
        double lat2 = latitude2 * Math.PI / 180;
        double deltaLat = (latitude2 - latitude1) * Math.PI / 180;
        double deltaLon = (longitude2 - longitude1) * Math.PI / 180;
        double a = Math.Sin(deltaLat / 2) * Math.Sin(deltaLat / 2) +
                   Math.Cos(lat1) * Math.Cos(lat2) *
                   Math.Sin(deltaLon / 2) * Math.Sin(deltaLon / 2);
        return EarthRadiusKm * 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));
    }

    public async Task InstallSupportAsync(IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        await _operationLock.WaitAsync(cancellationToken);
        try
        {
            if (IsOpenVpnInstalled && IsRouterBundled)
            {
                SetState(VpnConnectionState.Ready, "VPN support is already installed.");
                return;
            }
            if (!IsRouterBundled)
                throw new InvalidOperationException("HaloVpnRouter is not bundled in this build. Build the release with tools/build-vpn-router.ps1 first.");

            SetState(VpnConnectionState.Installing, "Downloading signed OpenVPN support…");
            Directory.CreateDirectory(VpnSettingsStore.RuntimeRoot);
            string installerPath = Path.Combine(VpnSettingsStore.RuntimeRoot, "OpenVPN-2.7.6-I001-amd64.msi");

            if (!File.Exists(installerPath) || !HasExpectedSha256(installerPath))
            {
                using HttpResponseMessage response = await Http.GetAsync(
                    OpenVpnInstallerUrl,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken);
                response.EnsureSuccessStatusCode();
                long total = response.Content.Headers.ContentLength ?? 0;
                await using Stream source = await response.Content.ReadAsStreamAsync(cancellationToken);
                await using var destination = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None);
                var buffer = new byte[81920];
                long copied = 0;
                int read;
                while ((read = await source.ReadAsync(buffer, cancellationToken)) > 0)
                {
                    await destination.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
                    copied += read;
                    if (total > 0)
                        progress?.Report((int)Math.Clamp(copied * 100 / total, 0, 100));
                }
            }

            if (!HasExpectedSha256(installerPath))
                throw new InvalidDataException("The OpenVPN installer failed its SHA-256 integrity check.");

            SetState(VpnConnectionState.Installing, "Installing the OpenVPN engine and signed driver…");
            var startInfo = new ProcessStartInfo
            {
                FileName = "msiexec.exe",
                Arguments = $"/i \"{installerPath}\" /passive /norestart ADDLOCAL=OpenVPN,OpenVPN.Service,Drivers,Drivers.OvpnDco,Drivers.TAPWindows6,OpenSSL",
                UseShellExecute = true,
                Verb = "runas",
                WindowStyle = ProcessWindowStyle.Normal
            };
            using Process installer = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Windows Installer did not start.");
            await installer.WaitForExitAsync(cancellationToken);
            if (installer.ExitCode is not 0 and not 3010)
                throw new InvalidOperationException($"OpenVPN setup ended with code {installer.ExitCode}.");
            if (!IsOpenVpnInstalled)
                throw new FileNotFoundException("OpenVPN setup completed, but openvpn.exe was not found.");

            SetState(VpnConnectionState.Ready, "VPN support installed. Enter your Nord service credentials.");
        }
        catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            SetState(VpnConnectionState.NotInstalled, "Setup was cancelled. No changes were made.");
            throw new OperationCanceledException("VPN support setup was cancelled.", ex);
        }
        catch (Exception ex)
        {
            SetState(VpnConnectionState.Error, ex.Message);
            throw;
        }
        finally
        {
            _operationLock.Release();
        }
    }

    public async Task ConnectAsync(
        VpnUserSettings settings,
        NordVpnServer? server,
        CancellationToken cancellationToken = default)
    {
        await _operationLock.WaitAsync(cancellationToken);
        try
        {
            if (!IsAdministrator())
                throw new UnauthorizedAccessException("Administrator permission is required to route MCC traffic.");
            if (!IsOpenVpnInstalled)
                throw new InvalidOperationException("Install VPN support first.");
            if (!IsRouterBundled)
                throw new InvalidOperationException("This build does not contain HaloVpnRouter.");
            if (IsNordDesktopAppInstalled())
                throw new InvalidOperationException("Uninstall the NordVPN desktop app before using Toolbox VPN mode; Nord blocks third-party OpenVPN adapters while its app is installed.");
            if (string.IsNullOrWhiteSpace(settings.Username) || string.IsNullOrWhiteSpace(settings.Password))
                throw new InvalidOperationException("Enter your NordVPN service username and password.");

            await StopProcessesAsync(waitForRouteCleanup: false);
            EnsureNoActiveVpnAdapter();
            uint normalInternetInterface = GetBestInternetInterfaceIndex();
            EnsureInterfaceIsNotVpn(normalInternetInterface);
            SetState(VpnConnectionState.Connecting, "Preparing Halo-only routing…");
            VpnSettingsStore.WriteRouterConfiguration();
            string profilePath = await ResolveProfileAsync(settings, server, cancellationToken);
            VpnNativeResources.EnsureExtracted();

            string openVpn = FindOpenVpnExecutable()
                ?? throw new FileNotFoundException("openvpn.exe was not found.");
            Directory.CreateDirectory(VpnSettingsStore.RuntimeRoot);
            _authPath = Path.Combine(VpnSettingsStore.RuntimeRoot, $"auth-{Environment.ProcessId}.txt");
            await File.WriteAllTextAsync(_authPath, $"{settings.Username}\n{settings.Password}\n", cancellationToken);
            TryRestrictToCurrentUser(_authPath);

            string logPath = Path.Combine(VpnSettingsStore.RuntimeRoot, "openvpn.log");
            string routerLogPath = Path.Combine(VpnSettingsStore.RuntimeRoot, "router.log");
            try { File.Delete(logPath); } catch { }
            try { File.Delete(routerLogPath); } catch { }

            _routerProcess = StartChild(
                VpnNativeResources.RouterPath,
                "observe",
                VpnNativeResources.BinRoot,
                routerLogPath);
            _routerProcess.EnableRaisingEvents = true;
            _routerProcess.Exited += RouterProcess_Exited;
            await WaitForRouterReadyAsync(routerLogPath, cancellationToken);

            StartOpenVpnProcess(openVpn, profilePath, logPath);

            SetState(VpnConnectionState.Connecting, "Connecting to NordVPN and verifying isolation…");
            await WaitForOpenVpnAsync(logPath, cancellationToken);
            if (_routerProcess.HasExited)
                throw new InvalidOperationException("Halo-only router stopped unexpectedly. Check antivirus or administrator permissions.");

            TunnelIp = await WaitForVpnAdapterIpAsync(
                TimeSpan.FromSeconds(15),
                "OpenVPN connected, but its Windows tunnel adapter was not available.",
                cancellationToken);
            VerifyNormalInternetRoute(normalInternetInterface);
            ConnectedServer = server;
            SetState(VpnConnectionState.Connected, "MCC-only routing verified. Windows and other apps remain direct.");
            VpnConnectionPresence.SetConnected(settings.Region);
        }
        catch (Exception ex)
        {
            await StopProcessesAsync(waitForRouteCleanup: false);
            VpnConnectionPresence.Clear();
            SetState(VpnConnectionState.Error, ex.Message);
            throw;
        }
        finally
        {
            DeleteAuthFile();
            _operationLock.Release();
        }
    }

    public async Task SwitchAsync(
        VpnUserSettings settings,
        NordVpnServer? server,
        CancellationToken cancellationToken = default)
    {
        await _operationLock.WaitAsync(cancellationToken);
        bool oldTunnelStopped = false;
        string previousRegion = VpnConnectionPresence.ConnectedRegion;
        NordVpnServer? previousServer = ConnectedServer;
        string previousTunnelIp = TunnelIp;
        try
        {
            if (!IsAdministrator())
                throw new UnauthorizedAccessException("Administrator permission is required to switch the MCC tunnel.");
            if (_openVpnProcess is null || _openVpnProcess.HasExited ||
                _routerProcess is null || _routerProcess.HasExited)
            {
                throw new InvalidOperationException("The current Halo-only tunnel is not running.");
            }
            if (string.IsNullOrWhiteSpace(settings.Username) || string.IsNullOrWhiteSpace(settings.Password))
                throw new InvalidOperationException("Enter your NordVPN service username and password.");

            uint normalInternetInterface = GetBestInternetInterfaceIndex();
            EnsureInterfaceIsNotVpn(normalInternetInterface);

            SetState(VpnConnectionState.Connecting, $"Preparing switch to {settings.Region}…");

            // Resolve and download the new profile while the old protected
            // tunnel is still healthy. MCC is not interrupted by this step.
            string profilePath = await ResolveProfileAsync(settings, server, cancellationToken);
            string openVpn = FindOpenVpnExecutable()
                ?? throw new FileNotFoundException("openvpn.exe was not found.");
            Directory.CreateDirectory(VpnSettingsStore.RuntimeRoot);
            _authPath = Path.Combine(VpnSettingsStore.RuntimeRoot, $"auth-{Environment.ProcessId}.txt");
            await File.WriteAllTextAsync(_authPath, $"{settings.Username}\n{settings.Password}\n", cancellationToken);
            TryRestrictToCurrentUser(_authPath);
            string logPath = Path.Combine(VpnSettingsStore.RuntimeRoot, "openvpn-switch.log");
            try { File.Delete(logPath); } catch { }

            SetState(VpnConnectionState.Connecting, $"Switching Halo tunnel to {settings.Region}…");
            StopOpenVpnProcess();
            oldTunnelStopped = true;
            VpnConnectionPresence.Clear();
            TunnelIp = "—";

            // The router remains alive throughout this gap. With no active
            // VPN adapter it consumes selected MCC packets instead of letting
            // them fall back to the normal Windows route.
            await WaitForVpnAdapterReleaseAsync(cancellationToken);
            TunnelIp = await StartReplacementOpenVpnAsync(openVpn, profilePath, logPath, cancellationToken);
            VerifyNormalInternetRoute(normalInternetInterface);
            if (_routerProcess is null || _routerProcess.HasExited)
                throw new InvalidOperationException("Halo-only router stopped during the VPN switch.");

            ConnectedServer = server;
            SetState(VpnConnectionState.Connected, $"Halo traffic switched to {settings.Region}.");
            VpnConnectionPresence.SetConnected(settings.Region);
        }
        catch (Exception ex)
        {
            if (!oldTunnelStopped)
            {
                ConnectedServer = previousServer;
                TunnelIp = previousTunnelIp;
                SetState(VpnConnectionState.Connected, "The VPN switch was cancelled; the original tunnel is still connected.");
                VpnConnectionPresence.SetConnected(previousRegion);
            }
            else
            {
                StopOpenVpnProcess();
                ConnectedServer = null;
                TunnelIp = "—";
                VpnConnectionPresence.Clear();
                SetState(VpnConnectionState.Error, $"VPN switch failed; Halo traffic remains blocked. {ex.Message}");
            }
            throw;
        }
        finally
        {
            DeleteAuthFile();
            _operationLock.Release();
        }
    }

    public async Task DisconnectAsync()
    {
        await _operationLock.WaitAsync();
        try
        {
            if (_openVpnProcess is null && _routerProcess is null)
            {
                VpnConnectionPresence.Clear();
                SetState(IsOpenVpnInstalled ? VpnConnectionState.Ready : VpnConnectionState.NotInstalled, "VPN is disconnected.");
                return;
            }
            SetState(VpnConnectionState.Disconnecting, "Closing the tunnel and restoring normal routing…");
            await StopProcessesAsync(waitForRouteCleanup: true);
            TunnelIp = "—";
            ConnectedServer = null;
            VpnConnectionPresence.Clear();
            SetState(VpnConnectionState.Ready, "VPN is disconnected. Normal networking is unchanged.");
        }
        finally
        {
            DeleteAuthFile();
            _operationLock.Release();
        }
    }

    private async Task<string> ResolveProfileAsync(
        VpnUserSettings settings,
        NordVpnServer? server,
        CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(settings.ImportedProfilePath) && File.Exists(settings.ImportedProfilePath))
            return settings.ImportedProfilePath;
        if (server is null)
            throw new InvalidOperationException("Choose a NordVPN server or import an OpenVPN profile.");
        if (!server.Hostname.EndsWith(".nordvpn.com", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("NordVPN returned an unexpected server hostname.");

        Directory.CreateDirectory(VpnSettingsStore.ProfilesRoot);
        string profilePath = Path.Combine(VpnSettingsStore.ProfilesRoot, $"{server.Hostname}.udp.ovpn");
        bool hasValidCachedProfile = TryValidateCachedProfile(profilePath);
        if (hasValidCachedProfile &&
            File.GetLastWriteTimeUtc(profilePath) >= DateTime.UtcNow.AddHours(-24))
        {
            return profilePath;
        }

        // A connected switch already has a Nord-signed profile containing the
        // shared Nord CA and tls-auth material. Build the target profile from
        // that validated template and the live catalog hostname/station rather
        // than depending on Nord's config CDN through the active tunnel.
        if (IsConnected && TryBuildProfileFromCachedNordTemplate(server, profilePath))
            return profilePath;

        string url = $"https://downloads.nordcdn.com/configs/files/ovpn_udp/servers/{Uri.EscapeDataString(server.Hostname)}.udp.ovpn";
        try
        {
            string profile = await Http.GetStringAsync(url, cancellationToken);
            ValidateProfile(profile);
            await File.WriteAllTextAsync(profilePath, profile, cancellationToken);
            return profilePath;
        }
        catch (HttpRequestException ex) when (hasValidCachedProfile)
        {
            RejoinFixDiagnostics.Warn(
                "vpn",
                $"Could not refresh {server.Hostname} profile; using validated cache: {ex.GetBaseException().Message}");
            return profilePath;
        }
        catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested && hasValidCachedProfile)
        {
            RejoinFixDiagnostics.Warn(
                "vpn",
                $"Timed out refreshing {server.Hostname} profile; using validated cache: {ex.GetBaseException().Message}");
            return profilePath;
        }
        catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Timed out downloading the NordVPN profile for {settings.Region}.",
                ex);
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException(
                $"Could not download the NordVPN profile for {settings.Region}: {ex.GetBaseException().Message}",
                ex);
        }
    }

    private static bool TryValidateCachedProfile(string profilePath)
    {
        try
        {
            if (!File.Exists(profilePath))
                return false;
            ValidateProfile(File.ReadAllText(profilePath));
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryBuildProfileFromCachedNordTemplate(
        NordVpnServer server,
        string destinationPath)
    {
        try
        {
            string? templatePath = Directory
                .EnumerateFiles(VpnSettingsStore.ProfilesRoot, "*.nordvpn.com.udp.ovpn")
                .Where(path => !path.Equals(destinationPath, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault(path => TryValidateCachedProfile(path));
            if (templatePath is null)
                return false;

            string template = File.ReadAllText(templatePath);
            if (!template.Contains("<ca>", StringComparison.OrdinalIgnoreCase) ||
                !template.Contains("<tls-auth>", StringComparison.OrdinalIgnoreCase) ||
                !template.Contains("verify-x509-name CN=", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string remoteHost = IPAddress.TryParse(server.Station, out _)
                ? server.Station
                : server.Hostname;
            string newline = template.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
            string[] lines = template.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
            bool replacedRemote = false;
            bool replacedCertificateName = false;

            for (int index = 0; index < lines.Length; index++)
            {
                string trimmed = lines[index].Trim();
                if (trimmed.StartsWith("remote ", StringComparison.OrdinalIgnoreCase))
                {
                    string[] parts = trimmed.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    string port = parts.Length >= 3 ? parts[2] : "1194";
                    lines[index] = $"remote {remoteHost} {port}";
                    replacedRemote = true;
                }
                else if (trimmed.StartsWith("verify-x509-name ", StringComparison.OrdinalIgnoreCase))
                {
                    lines[index] = $"verify-x509-name CN={server.Hostname}";
                    replacedCertificateName = true;
                }
            }

            if (!replacedRemote || !replacedCertificateName)
                return false;

            string profile = string.Join(newline, lines);
            ValidateProfile(profile);
            File.WriteAllText(destinationPath, profile);
            RejoinFixDiagnostics.Info(
                "vpn",
                $"Prepared {server.Hostname} profile from validated Nord template; no live CDN request was needed.");
            return true;
        }
        catch (Exception ex)
        {
            RejoinFixDiagnostics.Warn(
                "vpn",
                $"Could not prepare cached profile for {server.Hostname}: {ex.Message}");
            return false;
        }
    }

    public static void ValidateProfile(string profile)
    {
        if (!profile.Contains("client", StringComparison.OrdinalIgnoreCase) ||
            !profile.Contains("remote ", StringComparison.OrdinalIgnoreCase) ||
            !profile.Contains("auth-user-pass", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException("The selected file is not a supported client OpenVPN profile.");
        }

        var allowedDirectives = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "auth", "auth-nocache", "auth-user-pass", "cipher", "client", "comp-lzo",
            "data-ciphers", "data-ciphers-fallback", "dev", "explicit-exit-notify", "fast-io",
            "key-direction", "mssfix", "nobind", "persist-key", "persist-tun", "ping",
            "ping-restart", "ping-timer-rem", "proto", "pull", "remote", "remote-cert-tls",
            "remote-random", "reneg-sec", "resolv-retry", "sndbuf", "rcvbuf", "tls-cipher",
            "tls-version-min", "tun-mtu", "tun-mtu-extra", "verb", "verify-x509-name"
        };
        var allowedInlineBlocks = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ca", "tls-auth", "tls-crypt"
        };
        string? inlineBlock = null;

        foreach (string rawLine in profile.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n'))
        {
            string line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith('#') || line.StartsWith(';'))
                continue;

            if (inlineBlock is not null)
            {
                if (line.Equals($"</{inlineBlock}>", StringComparison.OrdinalIgnoreCase))
                    inlineBlock = null;
                continue;
            }

            if (line.StartsWith('<') && line.EndsWith('>'))
            {
                string block = line[1..^1].Trim();
                if (block.StartsWith('/'))
                    throw new InvalidDataException($"The VPN profile contains an unexpected closing block '{line}'.");
                if (!allowedInlineBlocks.Contains(block))
                    throw new InvalidDataException($"The VPN profile contains unsupported inline block '<{block}>'.");
                inlineBlock = block;
                continue;
            }

            string[] parts = line.Split(' ', '\t', StringSplitOptions.RemoveEmptyEntries);
            string directive = parts[0];
            if (!allowedDirectives.Contains(directive))
            {
                throw new InvalidDataException(
                    $"The VPN profile contains unsupported directive '{directive}'. " +
                    "Only passive NordVPN client options are accepted; scripts, plugins, management, routing, DNS, and output-file directives are blocked.");
            }
            if (directive.Equals("auth-user-pass", StringComparison.OrdinalIgnoreCase) && parts.Length != 1)
                throw new InvalidDataException("The VPN profile must not supply its own credential-file path.");
        }

        if (inlineBlock is not null)
            throw new InvalidDataException($"The VPN profile has an unterminated '<{inlineBlock}>' block.");
    }

    private static NordVpnServer? ToServer(NordServerResponse response)
    {
        NordCityResponse? city = response.Locations?
            .Select(location => location.Country?.City)
            .FirstOrDefault(value => value is not null);
        if (string.IsNullOrWhiteSpace(response.Hostname) || city is null)
            return null;
        string? country = response.Locations?
            .Select(location => location.Country?.Name)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        return new NordVpnServer(
            response.Hostname,
            response.Station,
            city.Name ?? "Unknown",
            country ?? "Unknown",
            response.Load,
            city.Latitude,
            city.Longitude);
    }

    private static HttpClient CreateHttpClient()
    {
        // Advanced Features installs a local system proxy for MCC observation.
        // Nord catalog/profile/setup traffic is Toolbox control-plane traffic,
        // so it bypasses that proxy and stays on the normal Windows route.
        var handler = new SocketsHttpHandler
        {
            UseProxy = false,
            AutomaticDecompression = DecompressionMethods.GZip |
                                     DecompressionMethods.Deflate |
                                     DecompressionMethods.Brotli,
            ConnectTimeout = TimeSpan.FromSeconds(15),
            PooledConnectionLifetime = TimeSpan.FromMinutes(5)
        };
        var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(30) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("HaloMCCToolbox/1.0 VPN-Setup");
        return client;
    }

    private static bool HasExpectedSha256(string path)
    {
        try
        {
            using Stream stream = File.OpenRead(path);
            return Convert.ToHexString(SHA256.HashData(stream))
                .Equals(OpenVpnInstallerSha256, StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static string? FindOpenVpnExecutable()
    {
        foreach (string path in new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OpenVPN", "bin", "openvpn.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "OpenVPN", "bin", "openvpn.exe")
        })
        {
            if (File.Exists(path))
                return path;
        }
        return null;
    }

    private Process StartChild(
        string fileName,
        string arguments,
        string workingDirectory,
        string logPath)
    {
        var process = Process.Start(new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            WorkingDirectory = workingDirectory,
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        }) ?? throw new InvalidOperationException($"{Path.GetFileName(fileName)} did not start.");
        process.OutputDataReceived += (_, args) => AppendRouterLog(logPath, args.Data);
        process.ErrorDataReceived += (_, args) => AppendRouterLog(logPath, args.Data);
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        VpnJobObject.Assign(process);
        return process;
    }

    private void AppendRouterLog(string path, string? line)
    {
        if (line is null)
            return;
        try
        {
            lock (_routerLogLock)
                File.AppendAllText(path, line + Environment.NewLine);
        }
        catch { }
    }

    private void StartOpenVpnProcess(string openVpn, string profilePath, string logPath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = openVpn,
            WorkingDirectory = Path.GetDirectoryName(profilePath)!,
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };
        foreach (string argument in new[]
        {
            "--config", profilePath,
            "--auth-user-pass", _authPath!,
            "--auth-nocache",
            "--route-nopull",
            "--pull-filter", "ignore", "redirect-gateway",
            "--pull-filter", "ignore", "block-outside-dns",
            "--pull-filter", "ignore", "dhcp-option DNS",
            "--pull-filter", "ignore", "dns",
            "--route", "0.0.0.0", "0.0.0.0", "vpn_gateway", "9999",
            "--log", logPath,
            "--verb", "3"
        })
        {
            startInfo.ArgumentList.Add(argument);
        }

        _openVpnProcess = Process.Start(startInfo)
            ?? throw new InvalidOperationException("OpenVPN did not start.");
        VpnJobObject.Assign(_openVpnProcess);
        _openVpnProcess.EnableRaisingEvents = true;
        _openVpnProcess.Exited += OpenVpnProcess_Exited;
    }

    private static uint GetBestInternetInterfaceIndex()
    {
        // Resolve the Windows route without opening a socket or sending data.
        uint destination = BitConverter.ToUInt32(IPAddress.Parse("1.1.1.1").GetAddressBytes(), 0);
        uint result = GetBestInterface(destination, out uint interfaceIndex);
        if (result != 0 || interfaceIndex == 0)
            throw new InvalidOperationException("Could not identify the normal Windows internet route. VPN was not started.");
        return interfaceIndex;
    }

    private static void EnsureInterfaceIsNotVpn(uint interfaceIndex)
    {
        NetworkInterface? adapter = FindInterface(interfaceIndex);
        if (adapter is null)
            throw new InvalidOperationException("Could not verify the normal Windows network adapter. VPN was not started.");

        if (LooksLikeVpnAdapter(adapter))
        {
            throw new InvalidOperationException(
                $"Windows already routes normal internet traffic through '{adapter.Name}'. " +
                "Disconnect the other VPN before starting MCC VPN.");
        }
    }

    private static void EnsureNoActiveVpnAdapter()
    {
        string[] active = NetworkInterface.GetAllNetworkInterfaces()
            .Where(adapter => adapter.OperationalStatus == OperationalStatus.Up &&
                              LooksLikeVpnAdapter(adapter) &&
                              HasRoutableIpv4Address(adapter))
            .Select(adapter => adapter.Name)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (active.Length > 0)
        {
            throw new InvalidOperationException(
                $"Another VPN adapter is already active ({string.Join(", ", active)}). " +
                "Disconnect other VPN connections before starting MCC VPN.");
        }
    }

    private static bool HasRoutableIpv4Address(NetworkInterface adapter)
    {
        try
        {
            return adapter.GetIPProperties().UnicastAddresses.Any(address =>
                address.Address.AddressFamily == AddressFamily.InterNetwork &&
                !IPAddress.IsLoopback(address.Address) &&
                !address.Address.ToString().StartsWith("169.254.", StringComparison.Ordinal));
        }
        catch { return false; }
    }

    private static void VerifyNormalInternetRoute(uint expectedInterfaceIndex)
    {
        uint currentInterfaceIndex = GetBestInternetInterfaceIndex();
        NetworkInterface? currentAdapter = FindInterface(currentInterfaceIndex);
        if (currentInterfaceIndex == expectedInterfaceIndex &&
            currentAdapter is not null &&
            !LooksLikeVpnAdapter(currentAdapter))
        {
            RejoinFixDiagnostics.Info(
                "vpn",
                $"MCC-only route guard passed; normal internet remains on {currentAdapter.Name} (ifIndex={currentInterfaceIndex}).");
            return;
        }

        string expectedName = FindInterface(expectedInterfaceIndex)?.Name ?? $"interface {expectedInterfaceIndex}";
        string currentName = currentAdapter?.Name ?? $"interface {currentInterfaceIndex}";
        throw new InvalidOperationException(
            $"MCC-only safety check failed: Windows internet routing changed from '{expectedName}' to '{currentName}'. " +
            "The VPN was disconnected before it could be used.");
    }

    private static NetworkInterface? FindInterface(uint interfaceIndex) =>
        NetworkInterface.GetAllNetworkInterfaces().FirstOrDefault(adapter =>
        {
            try { return adapter.GetIPProperties().GetIPv4Properties()?.Index == interfaceIndex; }
            catch { return false; }
        });

    private static bool LooksLikeVpnAdapter(NetworkInterface adapter)
    {
        string identity = $"{adapter.Name} {adapter.Description}";
        return identity.Contains("OpenVPN", StringComparison.OrdinalIgnoreCase) ||
               identity.Contains("TAP-Windows", StringComparison.OrdinalIgnoreCase) ||
               identity.Contains("Wintun", StringComparison.OrdinalIgnoreCase) ||
               identity.Contains("NordVPN", StringComparison.OrdinalIgnoreCase);
    }

    [DllImport("iphlpapi.dll")]
    private static extern uint GetBestInterface(uint destinationAddress, out uint bestInterfaceIndex);

    private void StopOpenVpnProcess()
    {
        _stopping = true;
        try
        {
            if (_openVpnProcess is not { } openVpn)
                return;
            openVpn.Exited -= OpenVpnProcess_Exited;
            TryKill(openVpn);
            _openVpnProcess = null;
        }
        finally
        {
            _stopping = false;
        }
    }

    private async Task WaitForOpenVpnAsync(string logPath, CancellationToken cancellationToken)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(45);
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_openVpnProcess is null || _openVpnProcess.HasExited)
                throw new InvalidOperationException(ReadOpenVpnFailure(logPath) ?? "OpenVPN exited before connecting.");

            string log = ReadSharedText(logPath);
            if (log.Contains("Initialization Sequence Completed", StringComparison.OrdinalIgnoreCase))
                return;
            string? failure = FindFailure(log);
            if (failure is not null)
                throw new InvalidOperationException(failure);
            await Task.Delay(250, cancellationToken);
        }
        throw new TimeoutException("NordVPN did not connect within 45 seconds. Check the service credentials and profile.");
    }

    private async Task WaitForRouterReadyAsync(string logPath, CancellationToken cancellationToken)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(10);
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_routerProcess is null || _routerProcess.HasExited)
            {
                string detail = ReadSharedText(logPath).Trim();
                throw new InvalidOperationException(
                    string.IsNullOrEmpty(detail)
                        ? "Halo-only router exited before packet capture was ready."
                        : $"Halo-only router failed to start: {detail}");
            }

            if (ReadSharedText(logPath).Contains(
                    "divert: capture+inject open",
                    StringComparison.OrdinalIgnoreCase))
                return;

            await Task.Delay(100, cancellationToken);
        }

        throw new TimeoutException(
            $"Halo-only router did not become ready. See {logPath} for details.");
    }

    private async Task<string> StartReplacementOpenVpnAsync(
        string openVpn,
        string profilePath,
        string logPath,
        CancellationToken cancellationToken)
    {
        for (int attempt = 1; attempt <= 2; attempt++)
        {
            try { File.Delete(logPath); } catch { }
            StartOpenVpnProcess(openVpn, profilePath, logPath);
            try
            {
                await WaitForOpenVpnAsync(logPath, cancellationToken);
                return await WaitForVpnAdapterIpAsync(
                    TimeSpan.FromSeconds(15),
                    "The replacement OpenVPN adapter did not become ready.",
                    cancellationToken);
            }
            catch (Exception) when (attempt == 1 && IsRetryableAdapterHandoffFailure(logPath))
            {
                PreserveFailedSwitchLog(logPath);
                StopOpenVpnProcess();
                SetState(
                    VpnConnectionState.Connecting,
                    "Windows is releasing the previous VPN adapter; retrying automatically...");
                await WaitForVpnAdapterReleaseAsync(cancellationToken);
                await Task.Delay(TimeSpan.FromSeconds(1.5), cancellationToken);
            }
        }

        throw new InvalidOperationException("The replacement OpenVPN tunnel did not start.");
    }

    private static bool IsRetryableAdapterHandoffFailure(string logPath)
    {
        string log = ReadSharedText(logPath);
        if (log.Contains("AUTH_FAILED", StringComparison.OrdinalIgnoreCase) ||
            log.Contains("VERIFY ERROR", StringComparison.OrdinalIgnoreCase) ||
            log.Contains("Options error", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string[] adapterFailures =
        [
            "adapter",
            "TAP-Windows",
            "open_tun",
            "CreateFile failed",
            "device is in use",
            "cannot open TUN/TAP"
        ];
        return adapterFailures.Any(value => log.Contains(value, StringComparison.OrdinalIgnoreCase));
    }

    private static void PreserveFailedSwitchLog(string logPath)
    {
        try
        {
            string diagnosticPath = Path.Combine(
                Path.GetDirectoryName(logPath)!,
                "openvpn-switch-first-attempt.log");
            File.Copy(logPath, diagnosticPath, overwrite: true);
        }
        catch { }
    }

    private static async Task WaitForVpnAdapterReleaseAsync(CancellationToken cancellationToken)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(15);
        DateTime? releasedSince = null;
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (FindVpnAdapterIp() is null)
            {
                releasedSince ??= DateTime.UtcNow;
                if (DateTime.UtcNow - releasedSince >= TimeSpan.FromSeconds(1.5))
                    return;
            }
            else
            {
                releasedSince = null;
            }
            await Task.Delay(250, cancellationToken);
        }

        throw new InvalidOperationException("Windows did not release the previous OpenVPN adapter in time.");
    }

    private static async Task<string> WaitForVpnAdapterIpAsync(
        TimeSpan timeout,
        string failureMessage,
        CancellationToken cancellationToken)
    {
        DateTime deadline = DateTime.UtcNow.Add(timeout);
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? address = FindVpnAdapterIp();
            if (address is not null)
                return address;
            await Task.Delay(250, cancellationToken);
        }

        throw new InvalidOperationException(failureMessage);
    }

    private static string? ReadOpenVpnFailure(string logPath) => FindFailure(ReadSharedText(logPath));

    private static string? FindFailure(string log)
    {
        if (log.Contains("AUTH_FAILED", StringComparison.OrdinalIgnoreCase))
            return "NordVPN rejected the service credentials.";
        if (log.Contains("All ovpn-dco adapters on this system are currently in use", StringComparison.OrdinalIgnoreCase))
            return "The OpenVPN adapter is busy. Close other VPN applications and try again.";
        string? fatal = log.Split('\n').LastOrDefault(line => line.Contains("fatal", StringComparison.OrdinalIgnoreCase));
        return string.IsNullOrWhiteSpace(fatal) ? null : fatal.Trim();
    }

    private static string ReadSharedText(string path)
    {
        try
        {
            if (!File.Exists(path))
                return "";
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch { return ""; }
    }

    private static string? FindVpnAdapterIp()
    {
        try
        {
            return NetworkInterface.GetAllNetworkInterfaces()
                .Where(adapter => adapter.OperationalStatus == OperationalStatus.Up)
                .Where(adapter =>
                    adapter.Description.Contains("OpenVPN", StringComparison.OrdinalIgnoreCase) ||
                    adapter.Description.Contains("TAP-Windows", StringComparison.OrdinalIgnoreCase) ||
                    adapter.Name.Contains("OpenVPN", StringComparison.OrdinalIgnoreCase))
                .SelectMany(adapter => adapter.GetIPProperties().UnicastAddresses)
                .FirstOrDefault(address =>
                    address.Address.AddressFamily == AddressFamily.InterNetwork &&
                    !address.Address.ToString().StartsWith("169.254.", StringComparison.Ordinal))
                ?.Address.ToString();
        }
        catch { return null; }
    }

    private async Task StopProcessesAsync(bool waitForRouteCleanup)
    {
        _stopping = true;
        try
        {
            if (_openVpnProcess is { } openVpn)
            {
                openVpn.Exited -= OpenVpnProcess_Exited;
                TryKill(openVpn);
                _openVpnProcess = null;
            }
            if (waitForRouteCleanup && _routerProcess is { HasExited: false })
                await Task.Delay(TimeSpan.FromSeconds(2.5));
            if (_routerProcess is { } router)
            {
                router.Exited -= RouterProcess_Exited;
                TryKill(router);
                _routerProcess = null;
            }
        }
        finally { _stopping = false; }
    }

    private void OpenVpnProcess_Exited(object? sender, EventArgs e)
    {
        if (_stopping || _disposed)
            return;
        TunnelIp = "—";
        VpnConnectionPresence.Clear();
        SetState(VpnConnectionState.Error, "The VPN disconnected. Halo traffic is blocked until you reconnect or disconnect VPN mode.");
    }

    private void RouterProcess_Exited(object? sender, EventArgs e)
    {
        if (_stopping || _disposed)
            return;

        // A dead redirector cannot enforce the kill switch. Close MCC first,
        // then tear down OpenVPN so game traffic never intentionally resumes
        // on the normal interface.
        TryTerminateMccProcesses();
        if (_openVpnProcess is { } openVpn)
        {
            openVpn.Exited -= OpenVpnProcess_Exited;
            TryKill(openVpn);
            _openVpnProcess = null;
        }
        TunnelIp = "—";
        VpnConnectionPresence.Clear();
        SetState(VpnConnectionState.Error, "The Halo-only router stopped. MCC was closed to prevent an IP leak.");
    }

    private static void TryTerminateMccProcesses()
    {
        foreach (string name in new[] { "MCC-Win64-Shipping", "MCCWinStore-Win64-Shipping", "mcclauncher" })
        {
            foreach (Process process in Process.GetProcessesByName(name))
            {
                try
                {
                    if (!process.HasExited)
                        process.Kill(entireProcessTree: true);
                }
                catch { }
                finally { process.Dispose(); }
            }
        }
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit(2500);
            }
        }
        catch { }
        finally { process.Dispose(); }
    }

    private void DeleteAuthFile()
    {
        if (_authPath is null)
            return;
        try { File.Delete(_authPath); } catch { }
        _authPath = null;
    }

    private static void TryRestrictToCurrentUser(string path)
    {
        try
        {
            var security = new System.Security.AccessControl.FileSecurity();
            var identity = System.Security.Principal.WindowsIdentity.GetCurrent().User;
            if (identity is null)
                return;
            security.SetOwner(identity);
            security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
            security.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule(
                identity,
                System.Security.AccessControl.FileSystemRights.FullControl,
                System.Security.AccessControl.AccessControlType.Allow));
            new FileInfo(path).SetAccessControl(security);
        }
        catch
        {
            // The credential file is deleted immediately after OpenVPN reads it.
        }
    }

    private void SetState(VpnConnectionState state, string status)
    {
        State = state;
        Status = status;
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        try { StopProcessesAsync(waitForRouteCleanup: false).GetAwaiter().GetResult(); } catch { }
        VpnConnectionPresence.Clear();
        DeleteAuthFile();
        _operationLock.Dispose();
    }

    private sealed class NordServerResponse
    {
        public string Hostname { get; set; } = "";
        public string Station { get; set; } = "";
        public int Load { get; set; }
        public string Status { get; set; } = "online";
        public List<NordLocationResponse>? Locations { get; set; }
    }

    private sealed class NordLocationResponse
    {
        public NordCountryResponse? Country { get; set; }
    }

    private sealed class NordCountryResponse
    {
        public string? Name { get; set; }
        public NordCityResponse? City { get; set; }
    }

    private sealed class NordCityResponse
    {
        public string? Name { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
    }
}

internal static class VpnNativeResources
{
    private const string ResourcePrefix = "HaloToolbox.Native.Vpn.";
    public static string BinRoot => Path.Combine(VpnSettingsStore.VpnRoot, "bin");
    public static string RouterPath => Path.Combine(BinRoot, "HaloVpnRouter.exe");
    public static bool IsAvailable =>
        HasResource(ResourcePrefix + "HaloVpnRouter.exe") || File.Exists(FindDevelopmentRouter());

    public static void EnsureExtracted()
    {
        Directory.CreateDirectory(BinRoot);
        ExtractOrCopy("HaloVpnRouter.exe", FindDevelopmentRouter());
        ExtractOrCopy("WinDivert.dll", FindDevelopmentSibling("WinDivert.dll"));
        ExtractOrCopy("WinDivert64.sys", FindDevelopmentSibling("WinDivert64.sys"));
        ExtractOrCopy("HaloVpnRouter-LICENSE.txt", "");
        ExtractOrCopy("WinDivert-LICENSE.txt", "");
        if (!File.Exists(RouterPath) ||
            !File.Exists(Path.Combine(BinRoot, "WinDivert.dll")) ||
            !File.Exists(Path.Combine(BinRoot, "WinDivert64.sys")))
        {
            throw new FileNotFoundException("The Halo-only routing helper is missing from this build.");
        }
    }

    private static void ExtractOrCopy(string name, string developmentPath)
    {
        string destination = Path.Combine(BinRoot, name);
        using Stream? resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourcePrefix + name);
        if (resource is not null)
        {
            using var memory = new MemoryStream();
            resource.CopyTo(memory);
            byte[] bytes = memory.ToArray();
            if (!File.Exists(destination) || !SHA256.HashData(File.ReadAllBytes(destination)).SequenceEqual(SHA256.HashData(bytes)))
                File.WriteAllBytes(destination, bytes);
            return;
        }

        if (File.Exists(developmentPath))
            File.Copy(developmentPath, destination, overwrite: true);
    }

    private static bool HasResource(string resourceName) =>
        Assembly.GetExecutingAssembly().GetManifestResourceNames().Contains(resourceName, StringComparer.Ordinal);

    private static string FindDevelopmentRouter()
    {
        string repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        return Path.Combine(repoRoot, "HaloVpnRouter", "target", "release", "HaloVpnRouter.exe");
    }

    private static string FindDevelopmentSibling(string name) =>
        Path.Combine(Path.GetDirectoryName(FindDevelopmentRouter()) ?? "", name);
}

internal static class VpnJobObject
{
    private static readonly Lazy<IntPtr> Job = new(CreateJob);

    public static void Assign(Process process)
    {
        if (!process.HasExited)
            AssignProcessToJobObject(Job.Value, process.Handle);
    }

    private static IntPtr CreateJob()
    {
        IntPtr handle = CreateJobObject(IntPtr.Zero, null);
        if (handle == IntPtr.Zero)
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());

        var info = new JobObjectExtendedLimitInformation
        {
            BasicLimitInformation = new JobObjectBasicLimitInformation { LimitFlags = 0x2000 }
        };
        int size = Marshal.SizeOf<JobObjectExtendedLimitInformation>();
        IntPtr buffer = Marshal.AllocHGlobal(size);
        try
        {
            Marshal.StructureToPtr(info, buffer, false);
            if (!SetInformationJobObject(handle, 9, buffer, (uint)size))
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        }
        finally { Marshal.FreeHGlobal(buffer); }
        return handle;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobObjectBasicLimitInformation
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public UIntPtr MinimumWorkingSetSize;
        public UIntPtr MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public UIntPtr Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IoCounters
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobObjectExtendedLimitInformation
    {
        public JobObjectBasicLimitInformation BasicLimitInformation;
        public IoCounters IoInfo;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryUsed;
        public UIntPtr PeakJobMemoryUsed;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateJobObject(IntPtr attributes, string? name);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(IntPtr job, int infoClass, IntPtr info, uint length);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
}
