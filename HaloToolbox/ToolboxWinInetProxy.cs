using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using Microsoft.Win32;

namespace HaloToolbox;

internal enum StaleProxyRecoveryResult
{
    None,
    RestoredSavedSettings,
    DisabledLegacyProxy,
    ActiveListener
}

internal static class ToolboxWinInetProxy
{
    private const string InternetSettingsKey = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";
    private const string LeaseMutexName = @"Local\HaloMCCToolbox.WinInetProxyLease";
    private const int InternetOptionRefresh = 37;
    private const int InternetOptionSettingsChanged = 39;

    private static readonly string[] KnownToolboxProxyAddresses =
    [
        ProxyService.DefaultProxyAddress,
        "127.0.0.1:19999"
    ];

    private static string LeasePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HaloMCCToolbox",
        "RejoinFix",
        "wininet-proxy-lease.json");

    public static void Enable(string proxyAddress, string proxyOverride, Action startWatchdog)
    {
        ExecuteWithLeaseLock(() =>
        {
            using var key = Registry.CurrentUser.OpenSubKey(InternetSettingsKey, writable: true)
                ?? throw new InvalidOperationException("Windows Internet Settings could not be opened.");

            var lease = CaptureLease(key, proxyAddress, proxyOverride);
            WriteLease(lease);

            try
            {
                // Start the independent owner watcher before changing Windows.
                // If the Toolbox dies on the next instruction, the watcher still restores the lease.
                startWatchdog();

                key.SetValue("ProxyEnable", 1, RegistryValueKind.DWord);
                key.SetValue("ProxyServer", proxyAddress, RegistryValueKind.String);
                key.SetValue("ProxyOverride", proxyOverride, RegistryValueKind.String);
                NotifyWindows();
            }
            catch
            {
                try
                {
                    RestoreOriginalValues(key, lease);
                    NotifyWindows();
                }
                catch { }
                DeleteLease();
                throw;
            }
        });
    }

    public static bool RestoreForCurrentProcess() => RestoreForOwner(Environment.ProcessId);

    public static bool RestoreForOwner(int ownerProcessId)
    {
        return ExecuteWithLeaseLock(() =>
        {
            var lease = ReadLease();
            if (lease is null || lease.OwnerProcessId != ownerProcessId)
                return false;

            return RestoreLease(lease);
        });
    }

    public static StaleProxyRecoveryResult RecoverStaleProxy()
    {
        return ExecuteWithLeaseLock(() =>
        {
            var lease = ReadLease();
            if (lease is not null)
            {
                if (IsLeaseOwnerActive(lease) && IsLoopbackListenerActive(lease.ProxyAddress))
                    return StaleProxyRecoveryResult.ActiveListener;

                return RestoreLease(lease)
                    ? StaleProxyRecoveryResult.RestoredSavedSettings
                    : StaleProxyRecoveryResult.None;
            }

            using var key = Registry.CurrentUser.OpenSubKey(InternetSettingsKey, writable: true);
            if (key is null || ReadProxyEnable(key) != 1)
                return StaleProxyRecoveryResult.None;

            var proxyServer = key.GetValue("ProxyServer") as string;
            if (!KnownToolboxProxyAddresses.Contains(proxyServer, StringComparer.OrdinalIgnoreCase) ||
                IsLoopbackListenerActive(proxyServer!))
            {
                return StaleProxyRecoveryResult.None;
            }

            // Older builds did not persist the user's original values. Disabling only the
            // exact, inactive Toolbox endpoint restores connectivity without guessing at
            // or deleting the rest of the user's proxy configuration.
            key.SetValue("ProxyEnable", 0, RegistryValueKind.DWord);
            NotifyWindows();
            return StaleProxyRecoveryResult.DisabledLegacyProxy;
        });
    }

    public static bool LeaseBelongsTo(int ownerProcessId)
    {
        return ExecuteWithLeaseLock(() => ReadLease()?.OwnerProcessId == ownerProcessId);
    }

    private static WinInetProxyLease CaptureLease(RegistryKey key, string proxyAddress, string proxyOverride)
    {
        return new WinInetProxyLease(
            Environment.ProcessId,
            proxyAddress,
            proxyOverride,
            DateTimeOffset.UtcNow,
            HasValue(key, "ProxyEnable"),
            ReadProxyEnable(key),
            HasValue(key, "ProxyServer"),
            key.GetValue("ProxyServer") as string,
            HasValue(key, "ProxyOverride"),
            key.GetValue("ProxyOverride") as string);
    }

    private static bool RestoreLease(WinInetProxyLease lease)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(InternetSettingsKey, writable: true);
            if (key is null)
            {
                DeleteLease();
                return false;
            }

            bool enableIsOriginal = ValueMatches(key, "ProxyEnable", lease.ProxyEnableExisted, lease.ProxyEnable);
            bool serverIsOriginal = ValueMatches(key, "ProxyServer", lease.ProxyServerExisted, lease.ProxyServer);
            bool overrideIsOriginal = ValueMatches(key, "ProxyOverride", lease.ProxyOverrideExisted, lease.ProxyOverride);
            bool enableIsApplied = ReadProxyEnable(key) == 1;
            bool serverIsApplied = string.Equals(
                key.GetValue("ProxyServer") as string,
                lease.ProxyAddress,
                StringComparison.OrdinalIgnoreCase);
            bool overrideIsApplied = string.Equals(
                key.GetValue("ProxyOverride") as string,
                lease.AppliedProxyOverride,
                StringComparison.Ordinal);

            bool restoredAnyValue = false;

            if (serverIsApplied)
            {
                // The dead loopback endpoint is the setting that can take the machine
                // offline. Always remove our endpoint, even if another program changed
                // only the bypass list or disabled the proxy while Toolbox was active.
                RestoreValue(key, "ProxyServer", lease.ProxyServerExisted, lease.ProxyServer, RegistryValueKind.String);
                restoredAnyValue = true;

                if (enableIsApplied)
                {
                    RestoreValue(key, "ProxyEnable", lease.ProxyEnableExisted, lease.ProxyEnable, RegistryValueKind.DWord);
                    restoredAnyValue = true;
                }

                if (overrideIsOriginal || overrideIsApplied)
                {
                    RestoreValue(key, "ProxyOverride", lease.ProxyOverrideExisted, lease.ProxyOverride, RegistryValueKind.String);
                    restoredAnyValue = true;
                }
            }
            else if (serverIsOriginal)
            {
                // Startup can be interrupted between individual registry writes. Roll
                // back only the components that reached our applied values.
                if (enableIsApplied && !enableIsOriginal)
                {
                    RestoreValue(key, "ProxyEnable", lease.ProxyEnableExisted, lease.ProxyEnable, RegistryValueKind.DWord);
                    restoredAnyValue = true;
                }

                if (overrideIsApplied && !overrideIsOriginal)
                {
                    RestoreValue(key, "ProxyOverride", lease.ProxyOverrideExisted, lease.ProxyOverride, RegistryValueKind.String);
                    restoredAnyValue = true;
                }
            }
            // A different ProxyServer means the user or another application replaced
            // our endpoint. Preserve that newer configuration and retire our lease.

            DeleteLease();
            if (restoredAnyValue)
                NotifyWindows();
            return restoredAnyValue;
        }
        catch
        {
            // Keep the lease so startup recovery or the watchdog can retry.
            return false;
        }
    }

    private static void RestoreValue(
        RegistryKey key,
        string name,
        bool existed,
        object? value,
        RegistryValueKind kind)
    {
        if (existed)
            key.SetValue(name, value ?? (kind == RegistryValueKind.DWord ? 0 : ""), kind);
        else
            key.DeleteValue(name, throwOnMissingValue: false);
    }

    private static void RestoreOriginalValues(RegistryKey key, WinInetProxyLease lease)
    {
        RestoreValue(key, "ProxyEnable", lease.ProxyEnableExisted, lease.ProxyEnable, RegistryValueKind.DWord);
        RestoreValue(key, "ProxyServer", lease.ProxyServerExisted, lease.ProxyServer, RegistryValueKind.String);
        RestoreValue(key, "ProxyOverride", lease.ProxyOverrideExisted, lease.ProxyOverride, RegistryValueKind.String);
    }

    private static bool ValueMatches(RegistryKey key, string name, bool expectedToExist, object? expectedValue)
    {
        bool exists = HasValue(key, name);
        if (exists != expectedToExist)
            return false;
        if (!exists)
            return true;

        object? currentValue = key.GetValue(name);
        if (currentValue is string currentString && expectedValue is string expectedString)
            return string.Equals(currentString, expectedString, StringComparison.Ordinal);

        try { return Convert.ToInt64(currentValue) == Convert.ToInt64(expectedValue); }
        catch { return Equals(currentValue, expectedValue); }
    }

    private static bool IsLoopbackListenerActive(string proxyAddress)
    {
        int separator = proxyAddress.LastIndexOf(':');
        if (separator < 0 || !int.TryParse(proxyAddress[(separator + 1)..], out int port))
            return false;

        try
        {
            return IPGlobalProperties.GetIPGlobalProperties()
                .GetActiveTcpListeners()
                .Any(endpoint =>
                    endpoint.Port == port &&
                    (IPAddress.IsLoopback(endpoint.Address) ||
                     endpoint.Address.Equals(IPAddress.Any) ||
                     endpoint.Address.Equals(IPAddress.IPv6Any)));
        }
        catch
        {
            // If listener discovery is unavailable, avoid deleting a possibly live proxy.
            return true;
        }
    }

    private static bool IsLeaseOwnerActive(WinInetProxyLease lease)
    {
        try
        {
            using var owner = Process.GetProcessById(lease.OwnerProcessId);
            if (owner.HasExited)
                return false;

            // Guard against the OS reusing a dead Toolbox process ID.
            return owner.StartTime.ToUniversalTime() <= lease.CreatedAtUtc.UtcDateTime;
        }
        catch
        {
            return false;
        }
    }

    private static int ReadProxyEnable(RegistryKey key)
    {
        try { return Convert.ToInt32(key.GetValue("ProxyEnable") ?? 0); }
        catch { return 0; }
    }

    private static bool HasValue(RegistryKey key, string name) =>
        key.GetValueNames().Contains(name, StringComparer.OrdinalIgnoreCase);

    private static WinInetProxyLease? ReadLease()
    {
        try
        {
            return File.Exists(LeasePath)
                ? JsonSerializer.Deserialize<WinInetProxyLease>(File.ReadAllText(LeasePath))
                : null;
        }
        catch
        {
            DeleteLease();
            return null;
        }
    }

    private static void WriteLease(WinInetProxyLease lease)
    {
        string directory = Path.GetDirectoryName(LeasePath)!;
        Directory.CreateDirectory(directory);
        string temporaryPath = LeasePath + $".{Environment.ProcessId}.tmp";
        File.WriteAllText(temporaryPath, JsonSerializer.Serialize(lease));
        File.Move(temporaryPath, LeasePath, overwrite: true);
    }

    private static void DeleteLease()
    {
        try
        {
            if (File.Exists(LeasePath))
                File.Delete(LeasePath);
        }
        catch { }
    }

    private static T ExecuteWithLeaseLock<T>(Func<T> action)
    {
        using var mutex = new Mutex(false, LeaseMutexName);
        bool lockTaken = false;
        try
        {
            try { lockTaken = mutex.WaitOne(TimeSpan.FromSeconds(5)); }
            catch (AbandonedMutexException) { lockTaken = true; }

            if (!lockTaken)
                throw new TimeoutException("Timed out waiting for the Toolbox proxy settings lock.");

            return action();
        }
        finally
        {
            if (lockTaken)
                mutex.ReleaseMutex();
        }
    }

    private static void ExecuteWithLeaseLock(Action action) =>
        ExecuteWithLeaseLock(() =>
        {
            action();
            return true;
        });

    private static void NotifyWindows()
    {
        InternetSetOption(IntPtr.Zero, InternetOptionSettingsChanged, IntPtr.Zero, 0);
        InternetSetOption(IntPtr.Zero, InternetOptionRefresh, IntPtr.Zero, 0);
    }

    [DllImport("wininet.dll", SetLastError = true)]
    private static extern bool InternetSetOption(
        IntPtr internetHandle,
        int option,
        IntPtr buffer,
        int bufferLength);

    private sealed record WinInetProxyLease(
        int OwnerProcessId,
        string ProxyAddress,
        string AppliedProxyOverride,
        DateTimeOffset CreatedAtUtc,
        bool ProxyEnableExisted,
        int ProxyEnable,
        bool ProxyServerExisted,
        string? ProxyServer,
        bool ProxyOverrideExisted,
        string? ProxyOverride);
}
