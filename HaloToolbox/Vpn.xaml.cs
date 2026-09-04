using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;

namespace HaloToolbox;

public partial class Vpn : UserControl, IDisposable
{
    private const int LazyLatencyProbeConcurrency = 2;
    private readonly VpnService _service = new();
    private readonly Func<Task<bool>>? _ensureCompanionServicesRunningAsync;
    private readonly DispatcherTimer _processTimer;
    private readonly VpnUserSettings _settings;
    private CancellationTokenSource? _regionRefreshCts;
    private bool _loaded;
    private bool _busy;
    private bool _updatingServerList;
    private bool _serverSelectionExplicit;
    private bool _disposed;

    public Vpn() : this(null)
    {
    }

    public Vpn(Func<Task<bool>>? ensureCompanionServicesRunningAsync)
    {
        InitializeComponent();
        _ensureCompanionServicesRunningAsync = ensureCompanionServicesRunningAsync;
        _settings = VpnSettingsStore.Load();
        TxtUsername.Password = _settings.Username;
        TxtPassword.Password = _settings.Password;
        ChkRemember.IsChecked = _settings.RememberCredentials;
        CmbRegion.ItemsSource = VpnService.HaloRegions.Select(region => region.Name).ToArray();
        SelectRegion(_settings.Region);
        RegionMap.LoadRegions(VpnService.HaloRegions);
        RegionMap.SelectRegion(SelectedRegion());
        RegionMap.RegionSelected += RegionMap_RegionSelected;
        _service.StateChanged += Service_StateChanged;
        _processTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _processTimer.Tick += (_, _) => RefreshMccProcessState();
        Loaded += Vpn_Loaded;
        RefreshUi();
    }

    private async void Vpn_Loaded(object sender, RoutedEventArgs e)
    {
        if (_loaded)
            return;
        _loaded = true;
        _processTimer.Start();
        if (string.IsNullOrWhiteSpace(_settings.ImportedProfilePath))
            await RefreshServersAsync(showErrors: false);
    }

    public async Task ConnectFromRelaunchAsync()
    {
        if (!_loaded)
        {
            _loaded = true;
            _processTimer.Start();
        }
        while (_busy)
            await Task.Delay(100);
        await ConnectCoreAsync();
    }

    private async void BtnInstallSupport_Click(object sender, RoutedEventArgs e) =>
        await InstallSupportAsync();

    private void BtnOpenNordCredentials_Click(object sender, RoutedEventArgs e) =>
        VpnCredentialHelpWindow.OpenNordCredentialsPage();

    private void BtnShowCredentialGuide_Click(object sender, RoutedEventArgs e)
    {
        var guide = new VpnCredentialHelpWindow
        {
            Owner = Window.GetWindow(this)
        };
        guide.ShowDialog();
    }

    private async Task<bool> InstallSupportAsync()
    {
        if (_service.IsOpenVpnInstalled && _service.IsRouterBundled)
            return true;
        try
        {
            SetBusy(true);
            var progress = new Progress<int>(value => TxtActionStatus.Text = $"Downloading OpenVPN support… {value}%");
            await _service.InstallSupportAsync(progress);
            return true;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
            return false;
        }
        finally
        {
            SetBusy(false);
            RefreshUi();
        }
    }

    private async void BtnRefreshServers_Click(object sender, RoutedEventArgs e)
    {
        CancelPendingRegionRefresh();
        await RefreshServersAsync(showErrors: true);
    }

    private async void CmbRegion_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        string region = SelectedRegion();
        if (!string.Equals(RegionMap.SelectedRegion, region, StringComparison.OrdinalIgnoreCase))
            RegionMap.SelectRegion(region, focus: true);

        if (!_loaded || _busy)
        {
            RefreshUi();
            return;
        }

        CancelPendingRegionRefresh();
        RegionMap.SetRegionLatency(region, null, testing: true);

        // Let WPF render the new marker, route, and label before rebuilding the
        // server controls below the map on the shared UI dispatcher.
        await Dispatcher.Yield(DispatcherPriority.Background);
        if (_disposed || _busy ||
            !string.Equals(region, SelectedRegion(), StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(region, RegionMap.SelectedRegion, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        RefreshUi();
        CmbServer.ItemsSource = null;
        TxtServerHint.Text = $"{region} selected · preparing nearby Nord exits…";
        _regionRefreshCts = new CancellationTokenSource();
        CancellationToken cancellationToken = _regionRefreshCts.Token;
        try
        {
            await Task.Delay(180, cancellationToken);
            await RefreshServersAsync(
                showErrors: false,
                cancellationToken,
                setBusy: false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
    }

    private async void RegionMap_RegionSelected(string region)
    {
        if (_busy)
        {
            RegionMap.SelectRegion(SelectedRegion());
            return;
        }

        CancelPendingRegionRefresh();
        RegionMap.SetRegionLatency(region, null, testing: true);
        await Dispatcher.Yield(DispatcherPriority.Background);
        if (_disposed || _busy ||
            !string.Equals(region, RegionMap.SelectedRegion, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        SelectRegion(region);
    }

    private void CmbServer_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingServerList)
            return;

        if (_loaded && !_busy)
        {
            _serverSelectionExplicit = true;
            if (CmbServer.SelectedItem is NordVpnServer server)
                RegionMap.SetRegionLatency(SelectedRegion(), server.LatencyMs);
            RefreshUi();
        }
    }

    private async Task RefreshServersAsync(
        bool showErrors,
        CancellationToken cancellationToken = default,
        bool setBusy = true)
    {
        bool busyHeld = false;
        string region = SelectedRegion();
        try
        {
            if (!cancellationToken.CanBeCanceled)
            {
                CancelPendingRegionRefresh();
                _regionRefreshCts = new CancellationTokenSource();
                cancellationToken = _regionRefreshCts.Token;
            }

            if (setBusy)
            {
                SetBusy(true);
                busyHeld = true;
            }

            _serverSelectionExplicit = false;
            if (CmbServer.ItemsSource is not null)
                CmbServer.ItemsSource = null;
            TxtServerHint.Text = $"Testing the closest NordVPN exit near {region}…";
            RegionMap.SetRegionLatency(region, null, testing: true);
            IReadOnlyList<NordVpnServer> candidates = await _service.GetNearbyServersAsync(
                region,
                cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(region, SelectedRegion(), StringComparison.OrdinalIgnoreCase))
                return;

            if (candidates.Count == 0)
            {
                RegionMap.SetRegionLatency(region, null);
                TxtServerHint.Text = "No Nord exit matched this MCC region. Refresh or import an .ovpn file.";
                return;
            }

            NordVpnServer firstServer = await _service.MeasureServerLatencyAsync(
                candidates[0],
                quick: true,
                cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(region, SelectedRegion(), StringComparison.OrdinalIgnoreCase))
                return;

            var measuredServers = new List<NordVpnServer> { firstServer };
            ApplyServerResults(region, measuredServers, completed: 1, total: candidates.Count);
            _settings.ImportedProfilePath = "";

            if (busyHeld)
            {
                SetBusy(false);
                busyHeld = false;
            }

            if (candidates.Count > 1)
            {
                _ = ContinueLazyServerTestsAsync(
                    region,
                    candidates.Skip(1).ToArray(),
                    measuredServers,
                    candidates.Count,
                    cancellationToken);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
        catch (Exception ex)
        {
            RegionMap.SetRegionLatency(region, null);
            TxtServerHint.Text = "Could not refresh servers. You can import an .ovpn profile instead.";
            if (showErrors)
                ShowError(ex.Message);
        }
        finally
        {
            if (busyHeld)
                SetBusy(false);
        }
    }

    private async Task ContinueLazyServerTestsAsync(
        string region,
        IReadOnlyList<NordVpnServer> remainingServers,
        List<NordVpnServer> measuredServers,
        int totalCount,
        CancellationToken cancellationToken)
    {
        var queue = new Queue<NordVpnServer>(remainingServers);
        var active = new List<Task<NordVpnServer>>(LazyLatencyProbeConcurrency);

        void StartNext()
        {
            if (queue.Count > 0)
            {
                active.Add(_service.MeasureServerLatencyAsync(
                    queue.Dequeue(),
                    quick: false,
                    cancellationToken));
            }
        }

        try
        {
            while (active.Count < LazyLatencyProbeConcurrency && queue.Count > 0)
                StartNext();

            while (active.Count > 0)
            {
                Task<NordVpnServer> completedTask = await Task.WhenAny(active);
                active.Remove(completedTask);
                try
                {
                    NordVpnServer measured = await completedTask;
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.Equals(region, SelectedRegion(), StringComparison.OrdinalIgnoreCase))
                        return;

                    measuredServers.Add(measured);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    throw;
                }
                catch
                {
                    // Keep testing the remaining candidates when one probe fails.
                }
                StartNext();
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (string.Equals(region, SelectedRegion(), StringComparison.OrdinalIgnoreCase))
                ApplyServerResults(region, measuredServers, totalCount, totalCount);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
        catch
        {
            // The immediately usable result remains selected if a background probe fails.
        }
    }

    private void ApplyServerResults(
        string region,
        IReadOnlyList<NordVpnServer> measuredServers,
        int completed,
        int total)
    {
        IReadOnlyList<NordVpnServer> ranked = VpnService.RankServers(measuredServers);
        string? selectedHostname = _serverSelectionExplicit
            ? (CmbServer.SelectedItem as NordVpnServer)?.Hostname
            : null;

        _updatingServerList = true;
        try
        {
            CmbServer.ItemsSource = ranked;
            CmbServer.SelectedItem = selectedHostname is null
                ? ranked.FirstOrDefault()
                : ranked.FirstOrDefault(server => server.Hostname.Equals(
                      selectedHostname,
                      StringComparison.OrdinalIgnoreCase))
                  ?? ranked.FirstOrDefault();
        }
        finally
        {
            _updatingServerList = false;
        }

        NordVpnServer? selected = CmbServer.SelectedItem as NordVpnServer;
        RegionMap.SetRegionLatency(region, selected?.LatencyMs);
        RefreshUi();

        if (selected is null)
            return;

        string latencyText = selected.LatencyMs is int latency ? $"{latency} ms" : "latency unavailable";
        TxtServerHint.Text = completed < total
            ? $"Ready now: {selected.City} · {latencyText}. Testing {total - completed} more nearby exits quietly…"
            : selected.LatencyMs is int finalLatency
                ? $"Fastest nearby exit: {selected.City} · {finalLatency} ms to Nord · {selected.DistanceKm:0} km from {region}."
                : $"Latency tests were unavailable. Using the closest exit to {region}; load breaks ties.";
    }

    private void CancelPendingRegionRefresh()
    {
        if (_regionRefreshCts is null)
            return;
        _regionRefreshCts.Cancel();
        _regionRefreshCts.Dispose();
        _regionRefreshCts = null;
    }

    private void BtnImportProfile_Click(object sender, RoutedEventArgs e)
    {
        CancelPendingRegionRefresh();
        var dialog = new OpenFileDialog
        {
            Title = "Select a NordVPN OpenVPN profile",
            Filter = "OpenVPN profiles (*.ovpn)|*.ovpn|All files (*.*)|*.*"
        };
        if (dialog.ShowDialog() != true)
            return;
        try
        {
            string profile = File.ReadAllText(dialog.FileName);
            VpnService.ValidateProfile(profile);
            Directory.CreateDirectory(VpnSettingsStore.ProfilesRoot);
            string destination = Path.Combine(VpnSettingsStore.ProfilesRoot, Path.GetFileName(dialog.FileName));
            File.Copy(dialog.FileName, destination, overwrite: true);
            _settings.ImportedProfilePath = destination;
            CmbServer.ItemsSource = null;
            TxtServerHint.Text = $"Using imported profile: {Path.GetFileName(destination)}";
            SaveSettings();
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
    }

    private async void BtnConnect_Click(object sender, RoutedEventArgs e)
    {
        if (_service.IsConnected)
        {
            if (HasPendingVpnSwitch())
            {
                await SwitchCoreAsync();
                return;
            }

            try
            {
                SetBusy(true);
                await _service.DisconnectAsync();
            }
            catch (Exception ex) { ShowError(ex.Message); }
            finally { SetBusy(false); }
            return;
        }

        SaveSettings(allowTemporaryPassword: true);
        if (!await InstallSupportAsync())
            return;
        if (!VpnService.IsAdministrator())
        {
            RelaunchElevatedForConnection();
            return;
        }
        await ConnectCoreAsync();
    }

    private async Task SwitchCoreAsync()
    {
        try
        {
            CancelPendingRegionRefresh();
            SetBusy(true);
            SaveSettings(allowTemporaryPassword: true);
            NordVpnServer? server = CmbServer.SelectedItem as NordVpnServer;
            if (server is null)
                throw new InvalidOperationException("No NordVPN server is available for the selected MCC region.");

            await _service.SwitchAsync(_settings, server);
            _settings.LastServerHostname = server.Hostname;
            VpnSettingsStore.Save(_settings);
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
        finally
        {
            if (_settings.RememberCredentials == false)
            {
                _settings.PasswordEncrypted = "";
                _settings.TemporaryCredentials = false;
                VpnSettingsStore.Save(_settings);
            }
            SetBusy(false);
            RefreshUi();
        }
    }

    private async Task ConnectCoreAsync()
    {
        if (_busy || _service.IsConnected)
            return;
        try
        {
            CancelPendingRegionRefresh();
            SetBusy(true);
            NordVpnServer? server = CmbServer.SelectedItem as NordVpnServer;
            if (string.IsNullOrWhiteSpace(_settings.ImportedProfilePath) && server is null)
            {
                await RefreshServersAsync(showErrors: true);
                server = CmbServer.SelectedItem as NordVpnServer;
                SetBusy(true);
            }
            await _service.ConnectAsync(_settings, server);
            if (_ensureCompanionServicesRunningAsync is not null)
            {
                TxtActionStatus.Text = "VPN connected. Starting Advanced Features and Rejoin Recovery…";
                bool servicesReady = await _ensureCompanionServicesRunningAsync();
                if (!servicesReady)
                {
                    TxtActionStatus.Text = "VPN connected, but Advanced Features need attention before MCC starts.";
                    return;
                }
            }
            if (ChkLaunchMcc.IsChecked == true)
                LaunchMcc();
        }
        catch (Exception ex)
        {
            ShowError(ex.Message);
        }
        finally
        {
            if (_settings.RememberCredentials == false)
            {
                _settings.PasswordEncrypted = "";
                _settings.TemporaryCredentials = false;
                VpnSettingsStore.Save(_settings);
            }
            SetBusy(false);
            RefreshUi();
        }
    }

    private void RelaunchElevatedForConnection()
    {
        try
        {
            string executable = Environment.ProcessPath
                ?? throw new InvalidOperationException("Could not locate the Toolbox executable.");
            Process.Start(new ProcessStartInfo
            {
                FileName = executable,
                Arguments = "--vpn-connect",
                UseShellExecute = true,
                Verb = "runas"
            });
            Application.Current.MainWindow?.Close();
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            ClearTemporaryCredentials();
            TxtActionStatus.Text = "Administrator permission was cancelled. No connection was started.";
        }
        catch (Exception ex)
        {
            ClearTemporaryCredentials();
            ShowError(ex.Message);
        }
    }

    private void LaunchMcc()
    {
        string root = App.LoadMccInstallationPath();
        string launcher = Path.Combine(root, "mcclauncher.exe");
        if (!File.Exists(launcher))
        {
            TxtActionStatus.Text = "Tunnel connected. MCC launcher was not found at the saved installation path.";
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo { FileName = launcher, WorkingDirectory = root, UseShellExecute = true });
            TxtActionStatus.Text = "Tunnel connected. MCC launched.";
        }
        catch (Exception ex)
        {
            TxtActionStatus.Text = $"Tunnel connected, but MCC could not launch: {ex.Message}";
        }
    }

    private void SaveSettings(bool allowTemporaryPassword = false)
    {
        _settings.Region = SelectedRegion();
        _settings.Username = TxtUsername.Password.Trim();
        _settings.RememberCredentials = ChkRemember.IsChecked == true;
        _settings.LastServerHostname = (CmbServer.SelectedItem as NordVpnServer)?.Hostname ?? "";
        if (_settings.RememberCredentials || allowTemporaryPassword)
        {
            _settings.Password = TxtPassword.Password;
            _settings.TemporaryCredentials = !_settings.RememberCredentials;
        }
        else
        {
            _settings.PasswordEncrypted = "";
            _settings.TemporaryCredentials = false;
        }
        VpnSettingsStore.Save(_settings);
    }

    private void ClearTemporaryCredentials()
    {
        if (_settings.RememberCredentials)
            return;
        _settings.PasswordEncrypted = "";
        _settings.TemporaryCredentials = false;
        VpnSettingsStore.Save(_settings);
    }

    private string SelectedRegion() =>
        CmbRegion.SelectedItem as string ?? "East US";

    private bool HasPendingVpnSwitch()
    {
        if (!string.Equals(
                SelectedRegion(),
                VpnConnectionPresence.ConnectedRegion,
                StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        NordVpnServer? selectedServer = CmbServer.SelectedItem as NordVpnServer;
        NordVpnServer? connectedServer = _service.ConnectedServer;
        return selectedServer is not null &&
               connectedServer is not null &&
               !selectedServer.Hostname.Equals(
                   connectedServer.Hostname,
                   StringComparison.OrdinalIgnoreCase);
    }

    private void SelectRegion(string region)
    {
        CmbRegion.SelectedItem = CmbRegion.Items
            .OfType<string>()
            .FirstOrDefault(item => string.Equals(item, region, StringComparison.OrdinalIgnoreCase))
            ?? CmbRegion.Items[0];
    }

    private void Service_StateChanged(object? sender, EventArgs e) =>
        Dispatcher.InvokeAsync(RefreshUi);

    private void RefreshUi()
    {
        TxtSetupState.Text = _service.IsOpenVpnInstalled
            ? "OPENVPN ENGINE READY · ROUTER BUNDLED"
            : "ONE-TIME VPN SUPPORT NOT INSTALLED";
        BtnInstallSupport.Visibility = _service.IsOpenVpnInstalled ? Visibility.Collapsed : Visibility.Visible;
        TxtConnectionDetail.Text = _service.Status;
        TxtTunnelIp.Text = _service.TunnelIp;
        TxtConnectedServer.Text = _service.ConnectedServer?.Hostname
                                  ?? (CmbServer.SelectedItem as NordVpnServer)?.Hostname
                                  ?? "—";

        string title;
        string brush;
        switch (_service.State)
        {
            case VpnConnectionState.Connected:
                title = "PROTECTED AND READY";
                brush = "GreenBrush";
                bool regionChanged = !string.Equals(
                    SelectedRegion(),
                    VpnConnectionPresence.ConnectedRegion,
                    StringComparison.OrdinalIgnoreCase);
                bool serverChanged = HasPendingVpnSwitch() && !regionChanged;
                BtnConnect.Content = regionChanged
                    ? $"SWITCH TO {SelectedRegion().ToUpperInvariant()}"
                    : serverChanged
                        ? "SWITCH SERVER"
                        : "DISCONNECT";
                BtnConnect.Style = (Style)FindResource(regionChanged || serverChanged ? "GreenButton" : "RedButton");
                break;
            case VpnConnectionState.Connecting:
            case VpnConnectionState.Installing:
            case VpnConnectionState.Disconnecting:
                title = _service.State.ToString().ToUpperInvariant() + "…";
                brush = "AccentBrush";
                break;
            case VpnConnectionState.Error:
                title = "ACTION NEEDED";
                brush = "RedBrush";
                BtnConnect.Content = "TRY AGAIN";
                BtnConnect.Style = (Style)FindResource("GreenButton");
                break;
            default:
                title = "NOT CONNECTED";
                brush = "MutedBrush";
                BtnConnect.Content = "CONNECT MCC ONLY";
                BtnConnect.Style = (Style)FindResource("GreenButton");
                break;
        }
        TxtConnectionTitle.Text = title;
        StatusDot.Background = (Brush)FindResource(brush);
        TxtActionStatus.Text = _service.Status;
        BtnConnect.IsEnabled = !_busy;
    }

    private void SetBusy(bool busy)
    {
        _busy = busy;
        BtnInstallSupport.IsEnabled = !busy;
        BtnRefreshServers.IsEnabled = !busy;
        BtnImportProfile.IsEnabled = !busy;
        CmbRegion.IsEnabled = !busy;
        CmbServer.IsEnabled = !busy;
        TxtUsername.IsEnabled = !busy;
        TxtPassword.IsEnabled = !busy;
        RegionMap.SetInteractionEnabled(!busy);
        BtnConnect.IsEnabled = !busy;
        if (!busy)
            RefreshUi();
    }

    private void RefreshMccProcessState()
    {
        bool running = new[] { "MCC-Win64-Shipping", "MCCWinStore-Win64-Shipping", "mcclauncher" }
            .Any(name =>
            {
                Process[] processes = Process.GetProcessesByName(name);
                bool found = processes.Length > 0;
                foreach (Process process in processes) process.Dispose();
                return found;
            });
        TxtMccProcess.Text = running ? "MCC process detected" : "Waiting for game";
        TxtMccProcess.Foreground = (Brush)FindResource(running ? "GreenBrush" : "MutedBrush");
    }

    private void ShowError(string message)
    {
        TxtActionStatus.Text = message;
        ToolboxDialog.Show(message, "MCC VPN", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _processTimer.Stop();
        CancelPendingRegionRefresh();
        RegionMap.RegionSelected -= RegionMap_RegionSelected;
        _service.StateChanged -= Service_StateChanged;
        _service.Dispose();
    }
}
