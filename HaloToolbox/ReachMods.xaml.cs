using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace HaloToolbox;

public partial class ReachMods : UserControl, IDisposable
{
    private const string BloomHookResourceName = "HaloToolbox.Native.ReachBloomHook.dll";
    private readonly ReachBloomSession _session = new();
    private readonly DispatcherTimer _timer;
    private bool _updatingToggle;
    private bool _disposed;

    public ReachMods()
    {
        InitializeComponent();
        _timer = new DispatcherTimer(TimeSpan.FromMilliseconds(500), DispatcherPriority.Background,
            (_, _) => RefreshStatus(), Dispatcher);
    }

    private void ReachMods_Loaded(object sender, RoutedEventArgs e)
    {
        if (!_disposed)
        {
            _timer.Start();
            RefreshStatus();
        }
    }

    private void ReachMods_Unloaded(object sender, RoutedEventArgs e) => _timer.Stop();

    private async void BtnArm_Click(object sender, RoutedEventArgs e)
    {
        var process = H3MemorySession.FindMccProcess();
        if (process is null)
        {
            SetDetail("MCC is not running. Start MCC with anti-cheat disabled and stop at the main menu.", true);
            return;
        }
        if (H3MemorySession.IsEasyAntiCheatLikelyLoaded(process))
        {
            SetDetail("Refused: Easy Anti-Cheat appears to be active. Restart MCC with anti-cheat disabled.", true);
            return;
        }

        BtnArm.IsEnabled = false;
        SetDetail("Arming the Reach lighting hook…", false);
        try
        {
            var payload = EnsureBloomHookPayload();
            await Task.Run(() => _session.ConnectOrInject(process, payload));
            SetDetail("Hook armed. Load Halo: Reach, then use the lighting toggle at any time.", false);
            RefreshStatus();
        }
        catch (Exception ex)
        {
            SetDetail(ex.Message, true);
        }
        finally
        {
            BtnArm.IsEnabled = true;
        }
    }

    private static string EnsureBloomHookPayload()
    {
        using var stream = typeof(ReachMods).Assembly.GetManifestResourceStream(BloomHookResourceName)
            ?? throw new InvalidOperationException("The embedded Reach bloom hook is missing.");
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        var payloadBytes = buffer.ToArray();
        var payloadHash = Convert.ToHexString(SHA256.HashData(payloadBytes))[..16];

        var payloadDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HaloMCCToolbox",
            "Native",
            payloadHash);
        Directory.CreateDirectory(payloadDirectory);

        var payloadPath = Path.Combine(payloadDirectory, "ReachBloomHook.dll");
        if (File.Exists(payloadPath)
            && File.ReadAllBytes(payloadPath).AsSpan().SequenceEqual(payloadBytes))
        {
            return payloadPath;
        }

        var temporaryPath = payloadPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllBytes(temporaryPath, payloadBytes);
            File.Move(temporaryPath, payloadPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);
        }

        return payloadPath;
    }

    private void ChkDisableBloom_Changed(object sender, RoutedEventArgs e)
    {
        if (_updatingToggle || !ChkDisableBloom.IsEnabled)
            return;
        if (!_session.SetBloomDisabled(ChkDisableBloom.IsChecked == true))
            SetDetail("Could not signal the Reach hook. MCC may have restarted; arm it again from the menu.", true);
    }

    private void RefreshStatus()
    {
        Process? process = null;
        try { process = H3MemorySession.FindMccProcess(); } catch { }

        var running = process is { HasExited: false };
        SetPill(TxtMccStatus, running ? $"MCC: PID {process!.Id}" : "MCC: NOT RUNNING", running);
        var eac = running && H3MemorySession.IsEasyAntiCheatLikelyLoaded(process!);
        SetPill(TxtEacStatus, eac ? "EAC: DETECTED" : running ? "EAC: NOT DETECTED" : "EAC: UNKNOWN", running && !eac);

        var state = _session.ReadState();
        var connected = _session.IsConnected && state.InstallOk;
        SetPill(TxtHookStatus, connected ? "HOOK: ARMED" : "HOOK: NOT ARMED", connected);
        SetPill(TxtShaderStatus, state.ShaderFound ? "SHADER: FOUND" : "SHADER: WAITING", state.ShaderFound);
        ChkDisableBloom.IsEnabled = connected && !eac;
        TxtBlockedDraws.Text = $"BLOCKED DRAWS: {state.BlockedDraws:N0}";

        if (connected && ChkDisableBloom.IsChecked != state.BloomDisabled)
        {
            _updatingToggle = true;
            ChkDisableBloom.IsChecked = state.BloomDisabled;
            _updatingToggle = false;
        }
    }

    private static void SetPill(TextBlock text, string value, bool good)
    {
        text.Text = value;
        text.Foreground = (Brush)Application.Current.FindResource(good ? "AccentBrush" : "MutedBrush");
    }

    private void SetDetail(string value, bool error)
    {
        TxtDetail.Text = value;
        TxtDetail.Foreground = (Brush)Application.Current.FindResource(error ? "OrangeBrush" : "MutedBrush");
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _timer.Stop();
        _session.Dispose();
    }
}
