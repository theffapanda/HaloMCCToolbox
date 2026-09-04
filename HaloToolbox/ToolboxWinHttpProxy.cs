using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace HaloToolbox;

internal static class ToolboxWinHttpProxy
{
    private const string LeaseKeyPath = @"SOFTWARE\HaloMCCToolbox\WinHttpProxyLease";
    private const uint NamedProxyAccessType = 3;

    public static bool Enable(string proxyAddress, string proxyBypass, int ownerProcessId)
    {
        if (!TryGetConfiguration(out var original))
            return false;

        try
        {
            using var lease = Registry.LocalMachine.CreateSubKey(LeaseKeyPath, writable: true);
            if (lease is null)
                return false;

            lease.SetValue("OriginalAccessType", original.AccessType, RegistryValueKind.DWord);
            lease.SetValue("OriginalProxy", original.Proxy ?? "", RegistryValueKind.String);
            lease.SetValue("OriginalBypass", original.Bypass ?? "", RegistryValueKind.String);
            lease.SetValue("AppliedProxy", proxyAddress, RegistryValueKind.String);
            // Publish the owner last. Recovery ignores and removes an incomplete
            // lease, so it can never restore default values from a partial write.
            lease.SetValue("OwnerProcessId", ownerProcessId, RegistryValueKind.DWord);

            if (TrySetConfiguration(new WinHttpConfiguration(NamedProxyAccessType, proxyAddress, proxyBypass)))
                return true;
        }
        catch
        {
        }

        DeleteLease();
        return false;
    }

    public static bool LeaseBelongsTo(int ownerProcessId)
    {
        try
        {
            using var lease = Registry.LocalMachine.OpenSubKey(LeaseKeyPath);
            return lease is not null &&
                   Convert.ToInt32(lease.GetValue("OwnerProcessId") ?? 0) == ownerProcessId;
        }
        catch
        {
            return false;
        }
    }

    public static bool RestoreForOwner(int ownerProcessId)
    {
        try
        {
            bool restored = false;
            {
                using var lease = Registry.LocalMachine.OpenSubKey(LeaseKeyPath);
                if (lease is null || Convert.ToInt32(lease.GetValue("OwnerProcessId") ?? 0) != ownerProcessId)
                    return false;

                var original = new WinHttpConfiguration(
                    Convert.ToUInt32(lease.GetValue("OriginalAccessType") ?? 0),
                    EmptyToNull(lease.GetValue("OriginalProxy") as string),
                    EmptyToNull(lease.GetValue("OriginalBypass") as string));
                string appliedProxy = lease.GetValue("AppliedProxy") as string ?? "";

                if (TryGetConfiguration(out var current))
                {
                    bool settingsAreStillOwned =
                        current.AccessType == NamedProxyAccessType &&
                        string.Equals(current.Proxy, appliedProxy, StringComparison.OrdinalIgnoreCase);

                    // WinHTTP may normalize the bypass string. The exact Toolbox endpoint
                    // is the ownership signal; it must not survive after that endpoint dies.
                    restored = !settingsAreStillOwned || TrySetConfiguration(original);
                }
            }

            if (restored)
                DeleteLease();
            return restored;
        }
        catch
        {
            return false;
        }
    }

    public static bool RecoverStaleProxy()
    {
        try
        {
            int ownerProcessId;
            {
                using var lease = Registry.LocalMachine.OpenSubKey(LeaseKeyPath);
                if (lease is null)
                    return true;

                ownerProcessId = Convert.ToInt32(lease.GetValue("OwnerProcessId") ?? 0);
            }

            if (ownerProcessId <= 0)
            {
                DeleteLease();
                return true;
            }

            if (ToolboxWinInetProxy.LeaseBelongsTo(ownerProcessId))
                return true;

            return RestoreForOwner(ownerProcessId);
        }
        catch
        {
            return false;
        }
    }

    private static bool TryGetConfiguration(out WinHttpConfiguration configuration)
    {
        configuration = new WinHttpConfiguration(0, null, null);
        if (!WinHttpGetDefaultProxyConfiguration(out var native))
            return false;

        try
        {
            configuration = new WinHttpConfiguration(
                native.AccessType,
                Marshal.PtrToStringUni(native.Proxy),
                Marshal.PtrToStringUni(native.ProxyBypass));
            return true;
        }
        finally
        {
            if (native.Proxy != IntPtr.Zero)
                GlobalFree(native.Proxy);
            if (native.ProxyBypass != IntPtr.Zero)
                GlobalFree(native.ProxyBypass);
        }
    }

    private static bool TrySetConfiguration(WinHttpConfiguration configuration)
    {
        var native = new WinHttpProxyInfo
        {
            AccessType = configuration.AccessType,
            Proxy = StringToNative(configuration.Proxy),
            ProxyBypass = StringToNative(configuration.Bypass)
        };

        try
        {
            return WinHttpSetDefaultProxyConfiguration(ref native);
        }
        finally
        {
            if (native.Proxy != IntPtr.Zero)
                Marshal.FreeHGlobal(native.Proxy);
            if (native.ProxyBypass != IntPtr.Zero)
                Marshal.FreeHGlobal(native.ProxyBypass);
        }
    }

    private static IntPtr StringToNative(string? value) =>
        string.IsNullOrEmpty(value) ? IntPtr.Zero : Marshal.StringToHGlobalUni(value);

    private static string? EmptyToNull(string? value) => string.IsNullOrEmpty(value) ? null : value;

    private static void DeleteLease()
    {
        try { Registry.LocalMachine.DeleteSubKeyTree(LeaseKeyPath, throwOnMissingSubKey: false); }
        catch { }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WinHttpProxyInfo
    {
        public uint AccessType;
        public IntPtr Proxy;
        public IntPtr ProxyBypass;
    }

    private sealed record WinHttpConfiguration(uint AccessType, string? Proxy, string? Bypass);

    [DllImport("winhttp.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WinHttpGetDefaultProxyConfiguration(out WinHttpProxyInfo proxyInfo);

    [DllImport("winhttp.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WinHttpSetDefaultProxyConfiguration(ref WinHttpProxyInfo proxyInfo);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalFree(IntPtr memory);
}
