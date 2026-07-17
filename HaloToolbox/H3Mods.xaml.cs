using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace HaloToolbox;

public partial class H3Mods : UserControl, IDisposable
{
    private static readonly bool EnableExperimentalDollyPlayback = true;

    private readonly H3MemorySession _session = new();
    private readonly DispatcherTimer _attachTimer;
    private readonly DispatcherTimer _coordDisplayTimer;
    private readonly DispatcherTimer _dollyTimer;
    private readonly DispatcherTimer _dollyRecordTimer;
    private readonly DispatcherTimer _swivelTimer;
    private readonly Stopwatch _dollyClock = new();
    private readonly Stopwatch _dollyRecordClock = new();
    private readonly Dictionary<string, byte[]> _originalBytes = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, H3DiscoverySnapshot> _discoveryBaselines = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, float> _liveProbeLastValues = new(StringComparer.OrdinalIgnoreCase);
    private H3DiscoverySnapshot? _cameraArchitectureBaseline;
    private H3DiscoverySnapshot? _cameraMovementBaseline;
    private List<H3CameraScanCandidate> _cameraScanCandidates = new();
    private int _cameraScanCandidateIndex = -1;
    private List<H3CameraStructCandidate> _cameraStructCandidates = new();
    private H3CameraStructCandidate? _ownedCameraCandidate;
    private Dictionary<long, float> _vitalScanBaseline = new();
    private DateTime? _vitalScanBaselineAt;
    private CancellationTokenSource? _dollyPlaybackCts;
    private DollyTrackOverlayWindow? _dollyTrackOverlay;
    private readonly List<H3RuntimeMod> _mods;
    private H3CameraAddressSet? _activeCameraAddress;
    private bool _dollyPlaying;
    private bool _dollyRecording;
    private bool _dollyEditing;
    private bool _swivelCamActive;
    private bool _swivelEnabledFreecam;
    private (float X, float Y, float Z) _swivelLocalDirection;
    private (float X, float Y, float Z) _swivelLastAppliedPosition;
    private bool _swivelLastAppliedReady;
    private bool _swivelFacingReady;
    private bool _samplingPlayerFacing;
    private int _dollyEditIndex = -1;
    private bool _suppressDollySelectionNavigation;
    private bool _loadingDollyEditor;
    private bool _dollyEditorDirty;
    private bool _dollyTrackWanted = true;
    private (float X, float Y, float Z, float A, float B, float C)? _lastRecordedCameraTransform;
    private double _lastRecordedTimeSeconds;
    private double _dollyRecordBaseTimeSeconds;
    private bool _coordinateWriteBurstActive;
    private bool _xamlReady;
    private bool _isLoaded;
    private int _sessionRefreshInProgress;
    private int _sessionRefreshGeneration;
    private bool _writesAllowed;
    private bool _disposed;

    public ObservableCollection<H3ModRow> SkullMods { get; } = new();
    public ObservableCollection<H3ModRow> CameraMods { get; } = new();
    public ObservableCollection<H3ModRow> GameplayMods { get; } = new();
    public ObservableCollection<H3DollyMarker> DollyMarkers { get; } = new();

    public H3Mods()
    {
        InitializeComponent();
        _xamlReady = true;
        DataContext = this;

        _mods = BuildModCatalog();
        foreach (var mod in _mods)
        {
            var row = new H3ModRow(mod);
            mod.Row = row;
            GetCollection(mod.Category).Add(row);
        }

        _attachTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _attachTimer.Tick += (_, _) => RefreshSession(autoAttach: ChkAutoAttach.IsChecked == true);

        _coordDisplayTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(100)
        };
        _coordDisplayTimer.Tick += (_, _) => RefreshCoordinateDisplay();

        _dollyTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(8)
        };
        _dollyTimer.Tick += DollyTimer_Tick;
        DollyMarkers.CollectionChanged += (_, _) =>
        {
            RefreshDollyButtons();
            EnsureDollyTrackOverlayShown();
        };

        _dollyRecordTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _dollyRecordTimer.Tick += DollyRecordTimer_Tick;

        _swivelTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(8) };
        _swivelTimer.Tick += SwivelTimer_Tick;
    }

    private void H3Mods_Loaded(object sender, RoutedEventArgs e)
    {
        _isLoaded = true;
        _attachTimer.Start();
        _coordDisplayTimer.Start();
        RefreshSession(autoAttach: ChkAutoAttach.IsChecked == true);
        Dispatcher.BeginInvoke(EnsureDollyTrackOverlayShown);
    }

    private void H3Mods_Unloaded(object sender, RoutedEventArgs e)
    {
        _isLoaded = false;
        Interlocked.Increment(ref _sessionRefreshGeneration);
        StopSwivelCam("Swivel Cam stopped because the H3 Mods tab unloaded.");
        StopDollyPlayback("Dolly playback stopped because the H3 Mods tab unloaded.");
        StopDollyRecording("Dolly recording stopped because the H3 Mods tab unloaded.");
        _attachTimer.Stop();
        _coordDisplayTimer.Stop();
        CloseDollyTrackOverlay();
        _session.Detach();
    }

    private ObservableCollection<H3ModRow> GetCollection(H3ModCategory category)
        => category switch
        {
            H3ModCategory.Skull => SkullMods,
            H3ModCategory.Camera => CameraMods,
            _ => GameplayMods
        };

    private async void RefreshSession(bool autoAttach)
    {
        if (_disposed || !_xamlReady || !_isLoaded ||
            Interlocked.Exchange(ref _sessionRefreshInProgress, 1) != 0)
            return;

        var generation = _sessionRefreshGeneration;
        Process? process = null;
        bool eacLoaded = false;
        bool attached = false;
        bool halo3Loaded = false;

        try
        {
            await Task.Run(() =>
            {
                process = H3MemorySession.FindMccProcess();
                if (process is null)
                {
                    _session.Detach();
                    return;
                }

                eacLoaded = H3MemorySession.IsEasyAntiCheatLikelyLoaded(process);
                if (autoAttach && !eacLoaded)
                {
                    // Attach refreshes the module snapshot itself. Do not enumerate it
                    // a second time in this refresh cycle.
                    _session.Attach(process);
                }
                else if (_session.ProcessId != process.Id)
                {
                    _session.Detach();
                }
                else if (_session.IsAttached)
                {
                    _session.RefreshModules();
                }

                attached = _session.IsAttached;
                halo3Loaded = attached && _session.HasHalo3Module;
            });
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"H3 session refresh failed: {ex}");
        }
        finally
        {
            Interlocked.Exchange(ref _sessionRefreshInProgress, 0);
        }

        if (_disposed || !_isLoaded || generation != _sessionRefreshGeneration)
        {
            _session.Detach();
            return;
        }

        if (process is null)
        {
            _writesAllowed = false;
            _session.Detach();
            SetStatus(TxtMccStatus, "MCC: SEARCHING", "MutedBrush");
            SetStatus(TxtEacStatus, "EAC: UNKNOWN", "MutedBrush");
            SetStatus(TxtGameStatus, "HALO 3: WAITING", "MutedBrush");
            SetStatus(TxtAccessStatus, "ACCESS: DETACHED", "MutedBrush");
            TxtSessionDetail.Text = "Launch MCC with Easy Anti-Cheat disabled, then load Halo 3. Auto attach will enable controls when the runtime looks safe.";
            TxtFooter.Text = "Auto attach is on. Controls stay disabled until MCC is detected, EAC is not loaded, and halo3.dll is present.";
            RefreshRows();
            return;
        }

        SetStatus(TxtMccStatus, $"MCC: PID {process.Id}", "GreenBrush");
        SetStatus(TxtEacStatus, eacLoaded ? "EAC: LOADED" : "EAC: NOT DETECTED", eacLoaded ? "RedBrush" : "GreenBrush");
        _writesAllowed = attached && halo3Loaded && !eacLoaded;

        SetStatus(TxtGameStatus, halo3Loaded ? "HALO 3: LOADED" : "HALO 3: WAITING", halo3Loaded ? "GreenBrush" : "OrangeBrush");
        SetStatus(TxtAccessStatus, _writesAllowed ? "ACCESS: READY" : attached ? "ACCESS: GATED" : "ACCESS: DETACHED", _writesAllowed ? "GreenBrush" : "MutedBrush");

        TxtSessionDetail.Text = eacLoaded
            ? "Easy Anti-Cheat appears to be loaded. Runtime mod controls are disabled."
            : halo3Loaded
                ? "Auto attach is ready. Writes are enabled for Halo 3 while EAC is not detected."
                : "MCC is running without detected EAC. Load Halo 3 to enable runtime mod controls.";

        TxtFooter.Text = _writesAllowed
            ? "Ready. Any write made by this tab is cached and can be restored before detaching."
            : "Controls are gated until MCC is attached, EAC is not detected, and halo3.dll is loaded.";

        ApplyHeldMods();
        RefreshRows();
        RefreshCoordinateButtons();
    }

    private void ApplyHeldMods()
    {
        if (!_writesAllowed)
            return;

        foreach (var mod in _mods.OfType<H3FloatHoldMod>())
            mod.ApplyHold(_session, RememberOriginalBytes);
    }

    private void RefreshRows()
    {
        foreach (var mod in _mods)
        {
            bool active = false;
            bool readable = false;
            if (_session.IsAttached && _session.HasHalo3Module && mod.IsMapped)
            {
                readable = mod.TryReadActive(_session, out active);
            }

            mod.Row!.IsActive = active;
            mod.Row.IsReadable = readable;
            mod.Row.CanToggle = _writesAllowed && mod.IsMapped && readable;
            mod.Row.StateText = !mod.IsMapped
                ? "LOCKED"
                : !readable
                    ? "WAIT"
                    : active ? "ON" : "OFF";
            mod.Row.StateBrush = ResolveBrush(!mod.IsMapped ? "SubtleBrush" : active ? "GreenBrush" : "MutedBrush");
        }

        TxtPatchDetail.Text = _originalBytes.Count == 0
            ? "No patches applied by this Toolbox session."
            : $"{_originalBytes.Count} memory location(s) cached for restore in this session.";
    }

    private void RefreshCoordinateButtons()
    {
        bool enabled = _writesAllowed;
        BtnDiscoveryBase.IsEnabled = enabled;
        BtnDiscoveryDiff.IsEnabled = enabled;
        BtnDiscoveryCopy.IsEnabled = enabled;
        BtnDiscoveryClear.IsEnabled = enabled;
        BtnSkullByteRead.IsEnabled = enabled;
        BtnSkullByteRestore.IsEnabled = enabled;
        BtnCameraArchBase.IsEnabled = enabled;
        BtnCameraArchFreecamDiff.IsEnabled = enabled;
        BtnCameraArchThirdPersonDiff.IsEnabled = enabled;
        BtnCameraArchTimingDiff.IsEnabled = enabled;
        BtnCameraMoveBase.IsEnabled = enabled;
        BtnCameraMoveDiff.IsEnabled = enabled;
        BtnCameraStructScan.IsEnabled = enabled;
        BtnCameraStructTest.IsEnabled = enabled && _cameraStructCandidates.Count > 0;
        BtnCameraOwnedProbe.IsEnabled = enabled && _ownedCameraCandidate is not null;
        RefreshDollyButtons();
    }

    private void RefreshCoordinateDisplay()
    {
        if (!_xamlReady || !_writesAllowed || Volatile.Read(ref _sessionRefreshInProgress) != 0)
            return;

        if (TryReadPlayerPosition(out var playerX, out var playerY, out var playerZ))
            SetPlayerCoordinateText(playerX, playerY, playerZ);

        if (_coordinateWriteBurstActive ||
            TxtCamX.IsKeyboardFocusWithin ||
            TxtCamY.IsKeyboardFocusWithin ||
            TxtCamZ.IsKeyboardFocusWithin)
            return;

        if (TryReadCameraPosition(out var x, out var y, out var z))
            SetCoordinateText(x, y, z);
    }

    private void SetStatus(TextBlock target, string text, string brushKey)
    {
        target.Text = text;
        target.Foreground = ResolveBrush(brushKey);
    }

    private static Brush ResolveBrush(string key)
        => Application.Current.Resources[key] as Brush ?? Brushes.White;

    private void ChkAutoAttach_Changed(object sender, RoutedEventArgs e)
    {
        if (_xamlReady)
            RefreshSession(autoAttach: ChkAutoAttach.IsChecked == true);
    }

    private void BtnManualAttach_Click(object sender, RoutedEventArgs e)
    {
        var process = H3MemorySession.FindMccProcess();
        if (process is null)
        {
            TxtFooter.Text = "Manual attach failed: MCC process was not found.";
            RefreshSession(autoAttach: false);
            return;
        }

        if (H3MemorySession.IsEasyAntiCheatLikelyLoaded(process))
        {
            TxtFooter.Text = "Manual attach blocked: Easy Anti-Cheat appears to be loaded.";
            RefreshSession(autoAttach: false);
            return;
        }

        TxtFooter.Text = _session.Attach(process)
            ? "Manual attach succeeded. Load Halo 3 if controls are still waiting."
            : "Manual attach failed: Toolbox could not open MCC with read/write access.";
        RefreshSession(autoAttach: false);
    }

    private void BtnRestoreAll_Click(object sender, RoutedEventArgs e)
        => RestoreAll();

    private void ModButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: H3ModRow row })
            return;

        if (!_writesAllowed)
        {
            TxtFooter.Text = "Runtime mod controls are gated until MCC is attached with EAC not detected and halo3.dll loaded.";
            return;
        }

        var mod = row.Definition;
        if (mod.Id == "freecam" && _swivelCamActive)
            StopSwivelCam("Swivel Cam stopped for Freecam.");
        if (!mod.IsMapped)
        {
            TxtFooter.Text = $"{mod.Name} is intentionally locked until its memory location is verified.";
            return;
        }

        if (mod.Toggle(_session, RememberOriginalBytes, out var message))
        {
            if (mod.Id == "freecam" && mod.TryReadActive(_session, out var freecamActive))
            {
                if (freecamActive && !_session.IsCameraCaptureHookInstalled)
                {
                    var installed = _session.InstallCameraCaptureHook(out var hookDetail);
                    message = installed
                        ? $"{message} Camera hook enabled automatically."
                        : $"{message} Camera hook failed: {hookDetail}";
                }
                else if (!freecamActive && _session.IsCameraCaptureHookInstalled)
                {
                    message = _session.UninstallCameraCaptureHook()
                        ? $"{message} Camera hook restored."
                        : $"{message} Camera hook restore failed; detach from MCC before continuing.";
                }
            }

            TxtFooter.Text = message;
            RefreshRows();
            RefreshCoordinateButtons();
        }
        else
        {
            TxtFooter.Text = message;
        }
    }

    private void BtnDollyCapture_Click(object sender, RoutedEventArgs e)
    {
        if (!EnableExperimentalDollyPlayback)
        {
            TxtFooter.Text = "Dolly marker capture is parked until the real camera-control target is found.";
            return;
        }

        if (!EnsureReadyForCoordinates())
            return;

        var usedTypedFallback = false;
        if (!TryReadCameraPosition(out var x, out var y, out var z))
        {
            if (!TryGetTypedNonZeroCameraPosition(out x, out y, out z))
            {
                TxtFooter.Text = "Dolly marker capture failed: could not read the coordinate buffer.";
                return;
            }

            usedTypedFallback = true;
        }

        if (IsZeroVector(x, y, z))
        {
            if (!TryGetTypedNonZeroCameraPosition(out x, out y, out z))
            {
                TxtFooter.Text = "Dolly marker capture blocked: 0/0/0 is not a valid camera marker.";
                return;
            }

            usedTypedFallback = true;
        }

        var segmentSeconds = ParseDollySegmentSeconds();
        var timeSeconds = DollyMarkers.Count == 0
            ? 0
            : DollyMarkers[DollyMarkers.Count - 1].TimeSeconds + segmentSeconds;

        if (!_session.TryReadCapturedCameraOrientation(out var a, out var b, out var c))
        {
            TxtFooter.Text = "Dolly marker capture failed: the live camera rotation is unavailable.";
            return;
        }
        var marker = new H3DollyMarker(DollyMarkers.Count + 1, timeSeconds, x, y, z, a, b, c);
        DollyMarkers.Add(marker);
        SelectDollyMarkerWithoutNavigation(marker);
        SetCoordinateText(x, y, z);
        var source = usedTypedFallback ? "typed coordinates" : "coordinate buffer";
        TxtFooter.Text = $"Dolly marker {marker.Index} captured from {source} at {marker.TimeText}.";
        RefreshDollyButtons();
    }

    private async void BtnSwivelCam_Click(object sender, RoutedEventArgs e)
    {
        LogSwivelDiagnostic($"click active={_swivelCamActive} writes={_writesAllowed} hook={_session.IsCameraCaptureHookInstalled} dolly={_dollyPlaying}/{_dollyRecording}/{_dollyEditing}");
        if (_swivelCamActive)
        {
            StopSwivelCam("Swivel Cam disabled. Normal camera position restored.");
            return;
        }

        if (!_writesAllowed || _dollyPlaying || _dollyRecording || _dollyEditing)
        {
            TxtFooter.Text = "Swivel Cam needs ready Halo 3 memory and cannot run during Dolly Cam activity.";
            return;
        }

        var freecam = _mods.First(mod => mod.Id == "freecam");
        if (!freecam.TryReadActive(_session, out var freecamActive))
        {
            TxtFooter.Text = "Swivel Cam could not read the Freecam state.";
            return;
        }
        if (!freecamActive)
        {
            if (!freecam.Toggle(_session, RememberOriginalBytes, out var freecamMessage))
            {
                TxtFooter.Text = freecamMessage;
                return;
            }
            _swivelEnabledFreecam = true;
        }

        if (!_session.IsCameraCaptureHookInstalled && !_session.InstallCameraCaptureHook(out var hookDetail))
        {
            TxtFooter.Text = $"Swivel Cam could not install its verified camera hook: {hookDetail}";
            return;
        }

        BtnSwivelCam.IsEnabled = false;
        TxtFooter.Text = "Swivel Cam is capturing the current viewing angle...";
        float playerX = 0, playerY = 0, playerZ = 0;
        var poseReady = false;
        for (var attempt = 0; attempt < 20; attempt++)
        {
            if (TryReadPlayerPosition(out playerX, out playerY, out playerZ) &&
                _session.TryReadRawCapturedCameraTransform(out _, out _, out _, out _, out _, out _))
            {
                poseReady = true;
                break;
            }
            await Task.Delay(50);
        }
        if (!poseReady)
        {
            TxtFooter.Text = "Swivel Cam could not capture a normal gameplay camera frame.";
            RefreshDollyButtons();
            return;
        }

        _swivelCamActive = true;
        _swivelFacingReady = false;
        if (ChkSwivelFollowFacing.IsChecked == true &&
            TryReadPlayerFacingYaw(playerX, playerY, playerZ, out var facingYaw) &&
            _session.TryReadRawCapturedCameraTransform(
                out var rawX, out var rawY, out var rawZ, out _, out _, out _))
        {
            var dx = rawX - playerX;
            var dy = rawY - playerY;
            var dz = rawZ - playerZ;
            var length = Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz));
            if (length > 0.001)
            {
                var cos = Math.Cos(-facingYaw);
                var sin = Math.Sin(-facingYaw);
                _swivelLocalDirection = (
                    (float)(((dx / length) * cos) - ((dy / length) * sin)),
                    (float)(((dx / length) * sin) + ((dy / length) * cos)),
                    (float)(dz / length));
                _swivelFacingReady = true;
                _swivelLastAppliedReady = false;
            }
        }
        _swivelTimer.Start();
        var firstWrite = WriteSwivelTransform(playerX, playerY, playerZ);
        LogSwivelDiagnostic($"enabled freecamSphere distance={SldSwivelDistance.Value:R} firstWrite={firstWrite}");
        if (!firstWrite)
        {
            StopSwivelCam("Swivel Cam could not write its initial camera position.");
            return;
        }
        TxtFooter.Text = "Swivel Cam is orbit-constraining Freecam around the player. Distance is live-adjustable.";
        RefreshDollyButtons();
    }

    private void SwivelTimer_Tick(object? sender, EventArgs e)
    {
        if (!_swivelCamActive)
            return;
        if (!_writesAllowed ||
            !TryReadPlayerPosition(out var x, out var y, out var z) ||
            !WriteSwivelTransform(x, y, z))
        {
            StopSwivelCam("Swivel Cam stopped because the player or camera source became unavailable.");
        }
    }

    private bool WriteSwivelTransform(float playerX, float playerY, float playerZ)
    {
        if (!_session.TryReadRawCapturedCameraTransform(
                out var rawX, out var rawY, out var rawZ,
                out var yaw, out var pitch, out _))
            return false;

        var dx = rawX - playerX;
        var dy = rawY - playerY;
        var dz = rawZ - playerZ;
        var length = Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz));
        if (ChkSwivelFollowFacing.IsChecked == true && _swivelFacingReady &&
            TryReadPlayerFacingYaw(playerX, playerY, playerZ, out var currentFacingYaw))
        {
            var cos = Math.Cos(currentFacingYaw);
            var sin = Math.Sin(currentFacingYaw);
            var worldX = (float)((_swivelLocalDirection.X * cos) - (_swivelLocalDirection.Y * sin));
            var worldY = (float)((_swivelLocalDirection.X * sin) + (_swivelLocalDirection.Y * cos));
            var worldZ = _swivelLocalDirection.Z;
            var distanceNow = (float)SldSwivelDistance.Value;

            // Freecam's raw request differs from our last committed position only by
            // the user's movement input. Apply that delta tangentially to the sphere.
            if (_swivelLastAppliedReady)
            {
                worldX = (worldX * distanceNow) + (rawX - _swivelLastAppliedPosition.X);
                worldY = (worldY * distanceNow) + (rawY - _swivelLastAppliedPosition.Y);
                worldZ = (worldZ * distanceNow) + (rawZ - _swivelLastAppliedPosition.Z);
            }
            var worldLength = Math.Sqrt((worldX * worldX) + (worldY * worldY) + (worldZ * worldZ));
            if (worldLength > 0.001)
            {
                dx = (float)(worldX / worldLength);
                dy = (float)(worldY / worldLength);
                dz = (float)(worldZ / worldLength);
                length = 1;

                // Persist the repositioned satellite in player-local coordinates.
                var inverseCos = Math.Cos(-currentFacingYaw);
                var inverseSin = Math.Sin(-currentFacingYaw);
                _swivelLocalDirection = (
                    (float)((dx * inverseCos) - (dy * inverseSin)),
                    (float)((dx * inverseSin) + (dy * inverseCos)),
                    dz);
            }
        }
        if (length < 0.001)
        {
            var forward = CameraDirection(yaw, pitch);
            dx = (float)-forward.X;
            dy = (float)-forward.Y;
            dz = (float)-forward.Z;
            length = 1;
        }

        var distance = (float)SldSwivelDistance.Value;
        var x = playerX + ((float)(dx / length) * distance);
        var y = playerY + ((float)(dy / length) * distance);
        var z = playerZ + ((float)(dz / length) * distance);
        var lookX = playerX - x;
        var lookY = playerY - y;
        var lookZ = playerZ - z;
        yaw = (float)Math.Atan2(lookY, lookX);
        pitch = (float)Math.Atan2(lookZ, Math.Sqrt((lookX * lookX) + (lookY * lookY)));
        var wrote = _session.TryWriteCapturedCameraTransform(x, y, z, yaw, pitch, 0);
        if (wrote)
        {
            _swivelLastAppliedPosition = (x, y, z);
            _swivelLastAppliedReady = true;
        }
        return wrote;
    }

    private static bool IsFiniteCameraPosition(float x, float y, float z)
        => float.IsFinite(x) && float.IsFinite(y) && float.IsFinite(z) &&
           (Math.Abs(x) + Math.Abs(y) + Math.Abs(z)) > 0.001f;

    private static void LogSwivelDiagnostic(string message)
    {
        try
        {
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox");
            Directory.CreateDirectory(root);
            File.AppendAllText(Path.Combine(root, "swivel-cam.log"), $"{DateTime.Now:O} {message}{Environment.NewLine}");
        }
        catch { }
    }

    private void StopSwivelCam(string message)
    {
        LogSwivelDiagnostic($"stop active={_swivelCamActive}: {message}");
        if (_swivelCamActive)
        {
            _swivelCamActive = false;
            _swivelTimer.Stop();
            _session.DisableCameraTransformOverride();
            _swivelFacingReady = false;
            _swivelLastAppliedReady = false;
        }
        if (_swivelEnabledFreecam)
        {
            _mods.First(mod => mod.Id == "freecam").Toggle(_session, RememberOriginalBytes, out _);
            _swivelEnabledFreecam = false;
        }
        if (_xamlReady)
        {
            TxtFooter.Text = message;
            RefreshDollyButtons();
        }
    }

    private bool TryReadPlayerFacingYaw(float playerX, float playerY, float playerZ, out float yaw)
    {
        yaw = 0;
        if (!_session.TryReadFloat(H3KnownAddresses.PlayerFacingX, out var x) ||
            !_session.TryReadFloat(H3KnownAddresses.PlayerFacingY, out var y) ||
            !_session.TryReadFloat(H3KnownAddresses.PlayerFacingZ, out var z))
            return false;

        var directLength = Math.Sqrt((x * x) + (y * y));
        var dx = x;
        var dy = y;
        if (directLength < 0.8 || directLength > 1.2)
        {
            dx = x - playerX;
            dy = y - playerY;
        }
        if (!float.IsFinite(dx) || !float.IsFinite(dy) || (Math.Abs(dx) + Math.Abs(dy)) < 0.0001)
            return false;
        yaw = (float)Math.Atan2(dy, dx);
        return true;
    }

    private async void BtnSamplePlayerFacing_Click(object sender, RoutedEventArgs e)
    {
        if (_samplingPlayerFacing)
            return;
        if (!_writesAllowed || !_session.IsAttached || !_session.HasHalo3Module)
        {
            TxtFooter.Text = "Facing sampler needs Halo 3 attached and ready.";
            return;
        }

        const int firstOffset = 0x10680;
        const int lastOffset = 0x10880;
        const int intervalMilliseconds = 50;
        const int sampleCount = 300;
        var offsets = Enumerable.Range(0, ((lastOffset - firstOffset) / 4) + 1)
            .Select(index => firstOffset + (index * 4))
            .ToArray();
        var minimums = Enumerable.Repeat(float.PositiveInfinity, offsets.Length).ToArray();
        var maximums = Enumerable.Repeat(float.NegativeInfinity, offsets.Length).ToArray();
        var csv = new StringBuilder("elapsed_ms");
        foreach (var offset in offsets)
            csv.Append($",0x{offset:X}");
        csv.AppendLine();

        _samplingPlayerFacing = true;
        BtnSamplePlayerFacing.IsEnabled = false;
        BtnSamplePlayerFacing.Content = "ROTATE NOW — RECORDING";
        TxtFooter.Text = "Facing sampler recording for 15 seconds. Rotate the Spartan left and right, including one full turn.";
        var stopwatch = Stopwatch.StartNew();
        try
        {
            for (var sample = 0; sample < sampleCount; sample++)
            {
                csv.Append(stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                for (var index = 0; index < offsets.Length; index++)
                {
                    var address = new H3Address("halo3.dll", 0x2030288, offsets[index]);
                    if (_session.TryReadFloat(address, out var value) && float.IsFinite(value))
                    {
                        minimums[index] = Math.Min(minimums[index], value);
                        maximums[index] = Math.Max(maximums[index], value);
                        csv.Append(',').Append(value.ToString("R", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        csv.Append(',');
                    }
                }
                csv.AppendLine();
                await Task.Delay(intervalMilliseconds);
            }

            var ranked = offsets
                .Select((offset, index) => new
                {
                    Offset = offset,
                    Minimum = minimums[index],
                    Maximum = maximums[index],
                    Range = maximums[index] - minimums[index]
                })
                .Where(candidate => float.IsFinite(candidate.Range) && candidate.Range > 0.00001f)
                .OrderByDescending(candidate => candidate.Range)
                .Take(40)
                .ToArray();
            csv.AppendLine();
            csv.AppendLine("ranked_offset,min,max,range");
            foreach (var candidate in ranked)
            {
                csv.Append("0x").Append(candidate.Offset.ToString("X"))
                    .Append(',').Append(candidate.Minimum.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(candidate.Maximum.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(candidate.Range.ToString("R", CultureInfo.InvariantCulture))
                    .AppendLine();
            }

            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox");
            Directory.CreateDirectory(root);
            var path = Path.Combine(root, $"h3-facing-sample-{DateTime.Now:yyyyMMdd-HHmmss}.csv");
            await File.WriteAllTextAsync(path, csv.ToString());
            TxtFooter.Text = $"Facing sample complete: {path}";
            LogSwivelDiagnostic($"facing sample saved {path}");
        }
        catch (Exception ex)
        {
            TxtFooter.Text = $"Facing sampler failed: {ex.Message}";
            LogSwivelDiagnostic($"facing sample failed: {ex}");
        }
        finally
        {
            _samplingPlayerFacing = false;
            BtnSamplePlayerFacing.Content = "SAMPLE PLAYER FACING (15S)";
            BtnSamplePlayerFacing.IsEnabled = _writesAllowed;
        }
    }

    private async void BtnLocatePlayerObject_Click(object sender, RoutedEventArgs e)
    {
        if (_samplingPlayerFacing)
            return;
        if (!_writesAllowed || !TryReadPlayerPosition(out var playerX, out var playerY, out var playerZ))
        {
            TxtFooter.Text = "Player-object probe needs Halo 3 attached and a valid player position.";
            return;
        }

        _samplingPlayerFacing = true;
        BtnLocatePlayerObject.IsEnabled = false;
        BtnSamplePlayerFacing.IsEnabled = false;
        BtnLocatePlayerObject.Content = "LOCATING XYZ COPIES...";
        TxtFooter.Text = "Searching writable MCC memory for the verified player position. This is read-only.";
        try
        {
            var scan = await Task.Run(() =>
                _session.ScanWritableFloatTriples(playerX, playerY, playerZ, 0.075f, maxMatches: 64));
            var matches = scan.Matches.Take(32).ToArray();
            if (matches.Length == 0)
            {
                TxtFooter.Text = "No writable player-position copies were found; no facing fields were sampled.";
                return;
            }

            const int radius = 0x180;
            const int intervalMilliseconds = 75;
            const int sampleCount = 200;
            var candidates = matches
                .SelectMany(match => Enumerable.Range(-radius / 4, (radius * 2 / 4) + 1)
                    .Select(index => match.Address + (index * 4L)))
                .Distinct()
                .ToArray();
            var minimums = Enumerable.Repeat(float.PositiveInfinity, candidates.Length).ToArray();
            var maximums = Enumerable.Repeat(float.NegativeInfinity, candidates.Length).ToArray();
            var sums = new double[candidates.Length];
            var counts = new int[candidates.Length];

            BtnLocatePlayerObject.Content = "ROTATE NOW — OBJECT PROBE";
            TxtFooter.Text = $"Found {matches.Length} XYZ copies. Rotate in place continuously for 15 seconds; do not walk.";
            for (var sample = 0; sample < sampleCount; sample++)
            {
                for (var index = 0; index < candidates.Length; index++)
                {
                    if (!_session.TryReadFloatAbsolute(candidates[index], out var value) ||
                        !float.IsFinite(value) || Math.Abs(value) > 100000f)
                        continue;
                    minimums[index] = Math.Min(minimums[index], value);
                    maximums[index] = Math.Max(maximums[index], value);
                    sums[index] += value;
                    counts[index]++;
                }
                await Task.Delay(intervalMilliseconds);
            }

            var rows = candidates.Select((address, index) => new
                {
                    Address = address,
                    Minimum = minimums[index],
                    Maximum = maximums[index],
                    Range = maximums[index] - minimums[index],
                    Mean = counts[index] == 0 ? double.NaN : sums[index] / counts[index]
                })
                .Where(row => float.IsFinite(row.Range) && row.Range > 0.00001f)
                .OrderByDescending(row => row.Range)
                .ToArray();
            var report = new StringBuilder();
            report.AppendLine($"player={playerX:R},{playerY:R},{playerZ:R}");
            report.AppendLine($"xyz_matches={matches.Length}; scanned_bytes={scan.ScannedBytes}");
            foreach (var match in matches)
                report.AppendLine($"position_match=0x{match.Address:X}; value={match.X:R},{match.Y:R},{match.Z:R}");
            report.AppendLine("address,min,max,range,mean,nearest_position,relative_offset");
            foreach (var row in rows.Take(500))
            {
                var nearest = matches.OrderBy(match => Math.Abs(match.Address - row.Address)).First();
                report.Append("0x").Append(row.Address.ToString("X"))
                    .Append(',').Append(row.Minimum.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.Maximum.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.Range.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.Mean.ToString("R", CultureInfo.InvariantCulture))
                    .Append(",0x").Append(nearest.Address.ToString("X"))
                    .Append(',').Append(row.Address - nearest.Address)
                    .AppendLine();
            }

            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox");
            Directory.CreateDirectory(root);
            var path = Path.Combine(root, $"h3-player-object-facing-{DateTime.Now:yyyyMMdd-HHmmss}.csv");
            await File.WriteAllTextAsync(path, report.ToString());
            TxtFooter.Text = $"Player-object facing probe complete: {path}";
            LogSwivelDiagnostic($"player-object facing probe saved {path}; matches={matches.Length}; changing={rows.Length}");
        }
        catch (Exception ex)
        {
            TxtFooter.Text = $"Player-object probe failed: {ex.Message}";
            LogSwivelDiagnostic($"player-object probe failed: {ex}");
        }
        finally
        {
            _samplingPlayerFacing = false;
            BtnLocatePlayerObject.Content = "LOCATE PLAYER OBJECT + FACING";
            BtnLocatePlayerObject.IsEnabled = _writesAllowed;
            BtnSamplePlayerFacing.IsEnabled = _writesAllowed;
        }
    }

    private async void BtnScanGlobalFacing_Click(object sender, RoutedEventArgs e)
    {
        if (_samplingPlayerFacing)
            return;
        if (!_writesAllowed || !_session.IsAttached)
        {
            TxtFooter.Text = "Global facing probe needs Halo 3 attached and ready.";
            return;
        }

        _samplingPlayerFacing = true;
        BtnScanGlobalFacing.IsEnabled = false;
        BtnScanGlobalFacing.Content = "SCANNING UNIT VECTORS...";
        TxtFooter.Text = "Scanning writable MCC memory for normalized XY direction vectors. This is read-only.";
        try
        {
            var scan = await Task.Run(() => _session.ScanWritableUnitVectorPairs());
            var candidates = scan.Matches.Select(match => match.Address).Distinct().ToArray();
            if (candidates.Length == 0)
            {
                TxtFooter.Text = "No writable normalized direction vectors were found.";
                return;
            }

            var minimumX = Enumerable.Repeat(float.PositiveInfinity, candidates.Length).ToArray();
            var maximumX = Enumerable.Repeat(float.NegativeInfinity, candidates.Length).ToArray();
            var minimumY = Enumerable.Repeat(float.PositiveInfinity, candidates.Length).ToArray();
            var maximumY = Enumerable.Repeat(float.NegativeInfinity, candidates.Length).ToArray();
            var maxNormError = new float[candidates.Length];
            var validCounts = new int[candidates.Length];
            BtnScanGlobalFacing.Content = "ROTATE NOW — GLOBAL FILTER";
            TxtFooter.Text = $"Monitoring {candidates.Length:N0} direction candidates for 15 seconds. Rotate continuously; do not walk.";
            for (var sample = 0; sample < 150; sample++)
            {
                for (var index = 0; index < candidates.Length; index++)
                {
                    if (!_session.TryReadFloatAbsolute(candidates[index], out var x) ||
                        !_session.TryReadFloatAbsolute(candidates[index] + 4, out var y) ||
                        !float.IsFinite(x) || !float.IsFinite(y))
                        continue;
                    var norm = (x * x) + (y * y);
                    if (norm < 0.8f || norm > 1.2f)
                        continue;
                    minimumX[index] = Math.Min(minimumX[index], x);
                    maximumX[index] = Math.Max(maximumX[index], x);
                    minimumY[index] = Math.Min(minimumY[index], y);
                    maximumY[index] = Math.Max(maximumY[index], y);
                    maxNormError[index] = Math.Max(maxNormError[index], Math.Abs(1f - norm));
                    validCounts[index]++;
                }
                await Task.Delay(100);
            }

            var rows = candidates.Select((address, index) => new
                {
                    Address = address,
                    MinX = minimumX[index], MaxX = maximumX[index],
                    MinY = minimumY[index], MaxY = maximumY[index],
                    RangeX = maximumX[index] - minimumX[index],
                    RangeY = maximumY[index] - minimumY[index],
                    NormError = maxNormError[index], Count = validCounts[index]
                })
                .Where(row => row.Count >= 120 && float.IsFinite(row.RangeX) &&
                              (row.RangeX + row.RangeY) > 0.05f && row.NormError < 0.08f)
                .OrderByDescending(row => row.RangeX + row.RangeY)
                .ToArray();
            var primaryHeapVector = rows.FirstOrDefault(row => row.Address < 0x0000010000000000L && row.NormError < 0.001f);
            IReadOnlyList<H3PointerScanMatch> pointerMatches = [];
            if (primaryHeapVector is not null)
            {
                BtnScanGlobalFacing.Content = "TRACING FACING POINTER...";
                TxtFooter.Text = $"Tracing pointers to live facing candidate 0x{primaryHeapVector.Address:X}.";
                var possibleObjectBases = Enumerable.Range(0, (0x1000 / 8) + 1)
                    .Select(index => primaryHeapVector.Address - (index * 8L))
                    .ToArray();
                pointerMatches = await Task.Run(() =>
                    _session.ScanReadablePointersToAny(possibleObjectBases, maxMatches: 512));
            }
            _session.TryResolveAddress(H3KnownAddresses.PlayerX, out var resolvedPlayerX);
            var report = new StringBuilder();
            report.AppendLine($"initial_candidates={candidates.Length}; scanned_bytes={scan.ScannedBytes}");
            report.AppendLine($"resolved_player_x=0x{resolvedPlayerX:X}");
            report.AppendLine("address,relative_to_player_x,min_x,max_x,min_y,max_y,range_x,range_y,max_norm_error,valid_samples");
            foreach (var row in rows.Take(2000))
                report.Append("0x").Append(row.Address.ToString("X"))
                    .Append(',').Append(row.Address - resolvedPlayerX)
                    .Append(',').Append(row.MinX.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.MaxX.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.MinY.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.MaxY.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.RangeX.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.RangeY.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.NormError.ToString("R", CultureInfo.InvariantCulture))
                    .Append(',').Append(row.Count).AppendLine();
            if (primaryHeapVector is not null)
            {
                report.AppendLine();
                report.AppendLine($"primary_heap_vector=0x{primaryHeapVector.Address:X}");
                report.AppendLine("pointer_address,target_address,vector_field_offset");
                foreach (var pointer in pointerMatches)
                    report.Append("0x").Append(pointer.Address.ToString("X"))
                        .Append(",0x").Append(pointer.TargetAddress.ToString("X"))
                        .Append(',').Append(primaryHeapVector.Address - pointer.TargetAddress)
                        .AppendLine();
            }
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HaloMCCToolbox");
            Directory.CreateDirectory(root);
            var path = Path.Combine(root, $"h3-global-facing-{DateTime.Now:yyyyMMdd-HHmmss}.csv");
            await File.WriteAllTextAsync(path, report.ToString());
            TxtFooter.Text = $"Global facing probe complete: {rows.Length} live candidates. {path}";
            LogSwivelDiagnostic($"global facing probe saved {path}; initial={candidates.Length}; live={rows.Length}");
        }
        catch (Exception ex)
        {
            TxtFooter.Text = $"Global facing probe failed: {ex.Message}";
            LogSwivelDiagnostic($"global facing probe failed: {ex}");
        }
        finally
        {
            _samplingPlayerFacing = false;
            BtnScanGlobalFacing.Content = "SCAN GLOBAL FACING VECTORS";
            BtnScanGlobalFacing.IsEnabled = _writesAllowed;
        }
    }

    private void BtnDollyFindLive_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForCoordinates())
            return;

        if (!EnableExperimentalDollyPlayback)
        {
            var report = BuildCameraArchitectureReport();
            var path = SaveDiscoveryReport("camera-architecture", "report", report);
            TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
            TxtFooter.Text = "Camera architecture report captured. Use this instead of scanning pan-cam coordinate copies.";
            return;
        }

        if (_cameraScanCandidates.Count > 0)
        {
            SelectNextCameraCandidate();
            return;
        }

        if (!TryParseCameraTextBoxes(out var targetX, out var targetY, out var targetZ) ||
            IsZeroVector(targetX, targetY, targetZ))
        {
            TxtFooter.Text = "Type the visible pan-cam X/Y/Z into the coordinate boxes, then click SCAN CAM.";
            return;
        }

        Mouse.OverrideCursor = Cursors.Wait;
        try
        {
            _activeCameraAddress = null;
            _cameraScanCandidateIndex = -1;
            TxtFooter.Text = "Camera scan running against typed pan-cam coordinates...";
            Dispatcher.Invoke(() => { }, DispatcherPriority.Background);

            var candidates = FindCameraCandidatesNear(targetX, targetY, targetZ);
            if (candidates.Count == 0)
            {
                TxtFooter.Text = "Camera scan found no nearby X/Y/Z values. Re-enter the exact pan-cam overlay numbers and try again.";
                RefreshDollyButtons();
                return;
            }

            _cameraScanCandidates = candidates;
            _cameraScanCandidateIndex = -1;
            SelectNextCameraCandidate();
        }
        finally
        {
            Mouse.OverrideCursor = null;
        }
    }

    private void BtnCameraArchBase_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        var preset = BuildCameraArchitecturePreset();
        _cameraArchitectureBaseline = CaptureDiscoverySnapshot(preset);
        var report = BuildCameraArchitectureReport() +
                     Environment.NewLine + Environment.NewLine +
                     FormatDiscoveryBaseline(_cameraArchitectureBaseline) +
                     Environment.NewLine + Environment.NewLine +
                     "Next: toggle exactly one camera-affecting control, then click FREE DIFF, 3P DIFF, or TIME DIFF.";
        var path = SaveDiscoveryReport("camera-architecture", "base", report);
        TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
        TxtFooter.Text = "Camera architecture baseline captured. Toggle one camera control, then click the matching diff.";
    }

    private void BtnCameraArchFreecamDiff_Click(object sender, RoutedEventArgs e)
        => CaptureCameraArchitectureDiff("FREECAM / COORD CONTROL");

    private void BtnCameraArchThirdPersonDiff_Click(object sender, RoutedEventArgs e)
        => CaptureCameraArchitectureDiff("THIRD PERSON");

    private void BtnCameraArchTimingDiff_Click(object sender, RoutedEventArgs e)
        => CaptureCameraArchitectureDiff("TIMING / FOV / TIMESCALE");

    private void BtnCameraMoveBase_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        var preset = BuildCameraMovementPreset();
        _cameraMovementBaseline = CaptureDiscoverySnapshot(preset);
        var report = FormatDiscoveryBaseline(_cameraMovementBaseline) +
                     Environment.NewLine + Environment.NewLine +
                     "Next: move only the freecam to a clearly different position, then click MOVE DIFF.";
        var path = SaveDiscoveryReport("camera-movement", "base", report);
        TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
        TxtFooter.Text = "Camera movement baseline captured. Move only the freecam, then click MOVE DIFF.";
    }

    private void BtnCameraMoveDiff_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        if (_cameraMovementBaseline is null)
        {
            TxtDiscoveryOutput.Text = "No camera movement baseline yet. Click MOVE BASE, move only the freecam, then click MOVE DIFF.";
            TxtFooter.Text = "Camera movement diff needs MOVE BASE first.";
            return;
        }

        var after = CaptureDiscoverySnapshot(_cameraMovementBaseline.Preset);
        var report = "CAMERA MOVEMENT DIFF" +
                     Environment.NewLine +
                     FormatCameraMovementDiff(_cameraMovementBaseline, after);
        var path = SaveDiscoveryReport("camera-movement", "diff", report);
        TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
        TxtFooter.Text = "Camera movement diff captured. Look for repeated XYZ triples that match the visible pan-cam movement.";
    }

    private void BtnCameraStructScan_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        if (!TryReadCameraPosition(out var x, out var y, out var z) || IsZeroVector(x, y, z))
        {
            TxtDiscoveryOutput.Text = "STRUCT SCAN needs a non-zero live camera readout first. Enable Freecam/Coord Control and wait for the coordinate boxes to update.";
            TxtFooter.Text = "Struct scan cancelled: live camera coordinates were not readable.";
            return;
        }

        Mouse.OverrideCursor = Cursors.Wait;
        try
        {
            _ownedCameraCandidate = null;
            var result = _session.ScanWritableFloatTriples(x, y, z, tolerance: 0.04f, maxBytesToScan: 2L * 1024 * 1024 * 1024, maxMatches: 512);
            _cameraStructCandidates = BuildCameraStructCandidates(result.Matches)
                .OrderByDescending(c => c.Score)
                .ThenBy(c => c.XAddress)
                .ToList();
            var report = FormatCameraStructScanReport(x, y, z, result, _cameraStructCandidates);
            var path = SaveDiscoveryReport("camera-struct-scan", "report", report);
            TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
            TxtFooter.Text = _cameraStructCandidates.Count == 0
                ? "Struct scan found no candidates beyond the known readout/cache copies."
                : $"Struct baseline saved with {_cameraStructCandidates.Count} candidate(s). Move the freecam, then click COMPARE.";
            RefreshCoordinateButtons();
        }
        finally
        {
            Mouse.OverrideCursor = null;
        }
    }

    private void BtnCameraStructTest_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForCoordinates())
            return;

        if (_cameraStructCandidates.Count == 0)
        {
            TxtFooter.Text = "No struct baseline is cached. Click STRUCT BASE first.";
            return;
        }

        if (!TryReadCameraPosition(out var currentX, out var currentY, out var currentZ))
        {
            TxtFooter.Text = "Could not read the current camera position for comparison.";
            return;
        }

        var report = FormatCameraStructComparisonReport(currentX, currentY, currentZ, _cameraStructCandidates);
        var path = SaveDiscoveryReport("camera-struct-compare", "report", report);
        TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
        TxtFooter.Text = "Read-only struct comparison complete. No game memory was changed.";
        RefreshCoordinateButtons();
    }

    private async void BtnCameraOwnedProbe_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForCoordinates())
            return;

        var candidate = _ownedCameraCandidate;
        if (candidate is null)
        {
            TxtFooter.Text = "No owned camera candidate is selected. Run STRUCT BASE, move, then COMPARE.";
            return;
        }

        if (!_session.TryReadFloatAbsolute(candidate.XAddress, out var x) ||
            !_session.TryReadFloatAbsolute(candidate.YAddress, out var y) ||
            !_session.TryReadFloatAbsolute(candidate.ZAddress, out var z))
        {
            TxtFooter.Text = "Owned camera candidate is no longer readable. Run the comparison again.";
            _ownedCameraCandidate = null;
            RefreshCoordinateButtons();
            return;
        }

        BtnCameraOwnedProbe.IsEnabled = false;
        var changed = _session.TryWriteFloatAbsolute(candidate.ZAddress, z + 0.25f);
        if (changed)
            await Task.Delay(150);

        var restored =
            _session.TryWriteFloatAbsolute(candidate.XAddress, x) &&
            _session.TryWriteFloatAbsolute(candidate.YAddress, y) &&
            _session.TryWriteFloatAbsolute(candidate.ZAddress, z);
        TxtFooter.Text = changed && restored
            ? $"Owned-vector probe completed at 0x{candidate.XAddress:X}: Z +0.25, then restored after 150 ms."
            : $"Owned-vector probe could not complete cleanly at 0x{candidate.XAddress:X}. Run COMPARE again before retrying.";
        RefreshCoordinateButtons();
    }

    private void CaptureCameraArchitectureDiff(string label)
    {
        if (!EnsureReadyForDiscovery())
            return;

        if (_cameraArchitectureBaseline is null)
        {
            TxtDiscoveryOutput.Text = "No camera architecture baseline yet. Click CAM BASE, perform one controlled camera action, then click the matching diff button.";
            TxtFooter.Text = "Camera architecture diff needs CAM BASE first.";
            return;
        }

        var after = CaptureDiscoverySnapshot(_cameraArchitectureBaseline.Preset);
        var report = BuildCameraArchitectureReport() +
                     Environment.NewLine + Environment.NewLine +
                     $"CAMERA ARCHITECTURE DIFF: {label}" +
                     Environment.NewLine +
                     FormatDiscoveryDiff(_cameraArchitectureBaseline, after);
        var path = SaveDiscoveryReport("camera-architecture", SlugifyReportKind(label), report);
        TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
        TxtFooter.Text = $"Camera architecture diff captured for {label}.";
    }

    private void BtnDollyPlay_Click(object sender, RoutedEventArgs e)
    {
        if (!EnableExperimentalDollyPlayback)
        {
            TxtFooter.Text = "Dolly playback is parked until we can write the true camera-control source.";
            return;
        }

        if (!EnsureReadyForCoordinates())
            return;

        StopDollyRecording("Dolly recording stopped for playback.");

        if (DollyMarkers.Count < 2)
        {
            TxtFooter.Text = "Dolly playback needs at least two markers.";
            return;
        }

        _dollyPlaying = true;
        var markers = DollyMarkers.ToList();
        if (!WriteCameraTransform(markers[0]))
        {
            _dollyPlaying = false;
            TxtFooter.Text = "Dolly playback could not start: the camera transform override write failed.";
            RefreshDollyButtons();
            return;
        }
        _dollyPlaybackCts?.Cancel();
        _dollyPlaybackCts?.Dispose();
        _dollyPlaybackCts = new CancellationTokenSource();
        var speedMultiplier = (float)SldDollySpeed.Value;
        StartDollyPlaybackLoop(markers, speedMultiplier, _dollyPlaybackCts.Token);
        TxtFooter.Text = $"Dolly playback started with {DollyMarkers.Count} marker(s) at {speedMultiplier:0.00}x.";
        RefreshDollyButtons();
    }

    private void BtnDollyStop_Click(object sender, RoutedEventArgs e)
    {
        StopDollyRecording("Dolly recording stopped.");
        StopDollyPlayback("Dolly playback stopped.");
    }

    private void BtnDollyEdit_Click(object sender, RoutedEventArgs e)
    {
        if (LstDollyMarkers.SelectedItem is not H3DollyMarker selected)
        {
            TxtFooter.Text = "Select a keyframe to edit.";
            return;
        }

        if (!_dollyEditing)
        {
            _dollyEditIndex = DollyMarkers.IndexOf(selected);
            if (_dollyEditIndex < 0)
                return;

            _dollyEditing = true;
            _session.DisableCameraTransformOverride();
            PopulateDollyEditor(selected);
            TxtFooter.Text = $"Editing keyframe {selected.Index}. Move the camera or type pose values, then click SAVE KEY.";
            RefreshDollyButtons();
            return;
        }

        if (!EnsureReadyForCoordinates())
            return;

        double timeSeconds;
        float x, y, z, yaw, pitch, roll;
        if (_dollyEditorDirty)
        {
            if (!TryReadDollyEditor(out timeSeconds, out x, out y, out z, out yaw, out pitch, out roll))
            {
                TxtFooter.Text = "Keyframe save failed: one or more typed values are invalid.";
                return;
            }
        }
        else if (!TryReadCameraPosition(out x, out y, out z) ||
                 !_session.TryReadCapturedCameraOrientation(out yaw, out pitch, out roll))
        {
            TxtFooter.Text = "Keyframe save failed: the live camera transform is unavailable.";
            return;
        }
        else
        {
            timeSeconds = selected.TimeSeconds;
        }

        if (_dollyEditIndex < 0 || _dollyEditIndex >= DollyMarkers.Count)
        {
            CancelDollyEdit();
            return;
        }

        var collectionIndex = _dollyEditIndex;
        if ((collectionIndex > 0 && timeSeconds <= DollyMarkers[collectionIndex - 1].TimeSeconds) ||
            (collectionIndex < DollyMarkers.Count - 1 && timeSeconds >= DollyMarkers[collectionIndex + 1].TimeSeconds))
        {
            TxtFooter.Text = "Keyframe save failed: its time must remain between the previous and next keyframes.";
            return;
        }

        var updated = new H3DollyMarker(selected.Index, Math.Max(0, timeSeconds), x, y, z, yaw, pitch, roll);
        _dollyEditing = false;
        _dollyEditIndex = -1;
        _suppressDollySelectionNavigation = true;
        try
        {
            DollyMarkers[collectionIndex] = updated;
            LstDollyMarkers.SelectedItem = updated;
        }
        finally
        {
            _suppressDollySelectionNavigation = false;
        }
        WriteCameraTransform(updated);
        DollyKeyEditor.Visibility = Visibility.Collapsed;
        SetCoordinateText(x, y, z);
        TxtFooter.Text = $"Keyframe {updated.Index} saved and held at its updated pose.";
        RefreshDollyButtons();
    }

    private void LstDollyMarkers_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_xamlReady || _suppressDollySelectionNavigation)
            return;

        if (_dollyEditing)
            CancelDollyEdit();

        if (LstDollyMarkers.SelectedItem is H3DollyMarker marker &&
            !_dollyPlaying && !_dollyRecording && _writesAllowed && _session.IsCameraCaptureHookInstalled)
        {
            if (WriteCameraTransform(marker))
            {
                SetCoordinateText(marker.X, marker.Y, marker.Z);
                TxtFooter.Text = $"Keyframe {marker.Index} selected and held. Click EDIT KEY to adjust it.";
            }
            else
            {
                TxtFooter.Text = $"Could not move the camera to keyframe {marker.Index}.";
            }
        }
        else if (LstDollyMarkers.SelectedItem is null && !_dollyPlaying && !_dollyRecording)
        {
            _session.DisableCameraTransformOverride();
        }

        RefreshDollyButtons();
    }

    private void SelectDollyMarkerWithoutNavigation(H3DollyMarker marker)
    {
        _suppressDollySelectionNavigation = true;
        try { LstDollyMarkers.SelectedItem = marker; }
        finally { _suppressDollySelectionNavigation = false; }
    }

    private void SelectDollyMarkerIndexWithoutNavigation(int index)
    {
        _suppressDollySelectionNavigation = true;
        try { LstDollyMarkers.SelectedIndex = index; }
        finally { _suppressDollySelectionNavigation = false; }
    }

    private void CancelDollyEdit()
    {
        _dollyEditing = false;
        _dollyEditIndex = -1;
        _dollyEditorDirty = false;
        if (_xamlReady)
            DollyKeyEditor.Visibility = Visibility.Collapsed;
    }

    private void PopulateDollyEditor(H3DollyMarker marker)
    {
        _loadingDollyEditor = true;
        try
        {
            TxtDollyEditTime.Text = marker.TimeSeconds.ToString("0.###", CultureInfo.InvariantCulture);
            TxtDollyEditX.Text = marker.X.ToString("0.###", CultureInfo.InvariantCulture);
            TxtDollyEditY.Text = marker.Y.ToString("0.###", CultureInfo.InvariantCulture);
            TxtDollyEditZ.Text = marker.Z.ToString("0.###", CultureInfo.InvariantCulture);
            TxtDollyEditYaw.Text = marker.A.ToString("0.###", CultureInfo.InvariantCulture);
            TxtDollyEditPitch.Text = marker.B.ToString("0.###", CultureInfo.InvariantCulture);
            TxtDollyEditRoll.Text = marker.C.ToString("0.###", CultureInfo.InvariantCulture);
            DollyKeyEditor.Visibility = Visibility.Visible;
            _dollyEditorDirty = false;
        }
        finally
        {
            _loadingDollyEditor = false;
        }
    }

    private void DollyEditor_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (!_loadingDollyEditor && _dollyEditing)
            _dollyEditorDirty = true;
    }

    private bool TryReadDollyEditor(
        out double timeSeconds,
        out float x, out float y, out float z,
        out float yaw, out float pitch, out float roll)
    {
        timeSeconds = 0;
        x = y = z = yaw = pitch = roll = 0;
        return double.TryParse(TxtDollyEditTime.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out timeSeconds) &&
               float.TryParse(TxtDollyEditX.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out x) &&
               float.TryParse(TxtDollyEditY.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out y) &&
               float.TryParse(TxtDollyEditZ.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out z) &&
               float.TryParse(TxtDollyEditYaw.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out yaw) &&
               float.TryParse(TxtDollyEditPitch.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out pitch) &&
               float.TryParse(TxtDollyEditRoll.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out roll) &&
               double.IsFinite(timeSeconds) && float.IsFinite(x) && float.IsFinite(y) && float.IsFinite(z) &&
               float.IsFinite(yaw) && float.IsFinite(pitch) && float.IsFinite(roll);
    }

    private void BtnDollyTrack_Click(object sender, RoutedEventArgs e)
    {
        if (_dollyTrackOverlay is not null)
        {
            _dollyTrackWanted = false;
            CloseDollyTrackOverlay();
            TxtFooter.Text = "Dolly track overlay hidden.";
            return;
        }

        _dollyTrackWanted = true;
        ShowDollyTrackOverlay();
        TxtFooter.Text = "Dolly track overlay shown.";
    }

    private void EnsureDollyTrackOverlayShown()
    {
        if (!_dollyTrackWanted || _dollyTrackOverlay is not null || !IsLoaded ||
            DollyMarkers.Count == 0 || !_writesAllowed || !_session.IsCameraCaptureHookInstalled)
        {
            return;
        }

        ShowDollyTrackOverlay();
    }

    private void ShowDollyTrackOverlay()
    {
        if (_dollyTrackOverlay is not null)
            return;

        _dollyTrackOverlay = new DollyTrackOverlayWindow(BuildDollyTrackSnapshot);
        _dollyTrackOverlay.Closed += DollyTrackOverlay_Closed;
        _dollyTrackOverlay.SetPreferredProcessId(_session.ProcessId);
        _dollyTrackOverlay.Show();
        BtnDollyTrack.Content = "HIDE TRACK";
    }

    private DollyTrackSnapshot? BuildDollyTrackSnapshot()
    {
        if (!_writesAllowed || DollyMarkers.Count == 0 ||
            !TryReadCameraPosition(out var x, out var y, out var z) ||
            !_session.TryReadCapturedCameraOrientation(out var yaw, out var pitch, out var roll))
        {
            return null;
        }

        return new DollyTrackSnapshot(
            x, y, z,
            yaw, pitch, roll,
            90.0,
            DollyMarkers.ToList(),
            LstDollyMarkers.SelectedIndex);
    }

    private void DollyTrackOverlay_Closed(object? sender, EventArgs e)
    {
        if (sender is DollyTrackOverlayWindow overlay)
            overlay.Closed -= DollyTrackOverlay_Closed;
        _dollyTrackOverlay = null;
        _dollyTrackWanted = false;
        if (_xamlReady)
            BtnDollyTrack.Content = "SHOW TRACK";
    }

    private void CloseDollyTrackOverlay()
    {
        if (_dollyTrackOverlay is null)
            return;

        var overlay = _dollyTrackOverlay;
        _dollyTrackOverlay = null;
        overlay.Closed -= DollyTrackOverlay_Closed;
        overlay.Close();
        if (_xamlReady)
            BtnDollyTrack.Content = "SHOW TRACK";
    }

    private void BtnDollyRecord_Click(object sender, RoutedEventArgs e)
    {
        if (!EnableExperimentalDollyPlayback)
        {
            TxtFooter.Text = "Dolly recording is parked until camera capture reads the true camera source.";
            return;
        }

        if (_dollyRecording)
        {
            StopDollyRecording("Dolly recording stopped.");
            return;
        }

        if (!EnsureReadyForCoordinates())
            return;

        if (!TryReadCameraPosition(out var x, out var y, out var z) || IsZeroVector(x, y, z))
        {
            TxtFooter.Text = "Record needs a readable coordinate buffer.";
            return;
        }

        StopDollyPlayback("Dolly playback stopped for recording.");
        _dollyRecording = true;
        _lastRecordedCameraTransform = null;
        _lastRecordedTimeSeconds = DollyMarkers.Count == 0 ? 0 : DollyMarkers[DollyMarkers.Count - 1].TimeSeconds;
        _dollyRecordBaseTimeSeconds = _lastRecordedTimeSeconds;
        if (!AddDollyMarkerFromPosition(x, y, z, _lastRecordedTimeSeconds))
            return;
        _dollyRecordClock.Restart();
        _dollyRecordTimer.Start();
        TxtFooter.Text = "Dolly recording started. Move the camera; markers are added after meaningful distance changes.";
        RefreshDollyButtons();
    }

    private void BtnDollyDelete_Click(object sender, RoutedEventArgs e)
    {
        if (LstDollyMarkers.SelectedItem is not H3DollyMarker marker)
        {
            TxtFooter.Text = "Select a dolly marker to delete.";
            return;
        }

        var index = DollyMarkers.IndexOf(marker);
        if (index >= 0)
            DollyMarkers.RemoveAt(index);

        RenumberDollyMarkers();
        if (DollyMarkers.Count > 0)
            SelectDollyMarkerIndexWithoutNavigation(Math.Clamp(index, 0, DollyMarkers.Count - 1));

        TxtFooter.Text = "Dolly marker deleted.";
        RefreshDollyButtons();
    }

    private void BtnDollyClear_Click(object sender, RoutedEventArgs e)
    {
        StopDollyPlayback("Dolly markers cleared.");
        StopDollyRecording("Dolly markers cleared.");
        CancelDollyEdit();
        DollyMarkers.Clear();
        RefreshDollyButtons();
    }

    private void DollyRecordTimer_Tick(object? sender, EventArgs e)
    {
        if (!_dollyRecording)
            return;

        if (!_writesAllowed)
        {
            StopDollyRecording("Dolly recording stopped: Halo 3 memory is no longer ready.");
            return;
        }

        if (!TryReadCameraPosition(out var x, out var y, out var z) || IsZeroVector(x, y, z))
            return;

        if (_lastRecordedCameraTransform is { } last &&
            Distance(last.X, last.Y, last.Z, x, y, z) < SldDollyRecordDistance.Value)
        {
            return;
        }

        _lastRecordedTimeSeconds = _dollyRecordBaseTimeSeconds + _dollyRecordClock.Elapsed.TotalSeconds;
        if (!AddDollyMarkerFromPosition(x, y, z, _lastRecordedTimeSeconds))
            return;
        TxtFooter.Text = $"Dolly recording: {DollyMarkers.Count} marker(s).";
    }

    private void DollyTimer_Tick(object? sender, EventArgs e)
    {
        if (!_dollyPlaying)
            return;

        if (!_writesAllowed || DollyMarkers.Count < 2)
        {
            StopDollyPlayback("Dolly playback stopped: Halo 3 memory is no longer ready.");
            return;
        }

        var marker = EvaluateDollyPath(DollyMarkers, _dollyClock.Elapsed.TotalSeconds);
        if (marker.TimeSeconds >= DollyMarkers[DollyMarkers.Count - 1].TimeSeconds)
        {
            WriteCameraTransform(marker);
            StopDollyPlayback("Dolly playback complete.");
            return;
        }

        WriteCameraTransform(marker);
    }

    private void StartDollyPlaybackLoop(IReadOnlyList<H3DollyMarker> markers, float speedMultiplier, CancellationToken token)
    {
        // Match the desktop DollyCam: evaluate once up front, then let a thread inside MCC
        // consume one pose per millisecond. External WriteProcessMemory scheduling was the
        // source of the visible hold/catch-up cadence in rendered frames.
        var duration = markers[^1].TimeSeconds;
        var sampleCount = Math.Max(2, checked((int)Math.Ceiling(duration * 1000.0)) + 1);
        var transforms = new List<float>(sampleCount * 6);
        for (var i = 0; i < sampleCount; i++)
        {
            var marker = EvaluateDollyPath(markers, Math.Min(duration, i / 1000.0));
            transforms.Add((float)marker.X);
            transforms.Add((float)marker.Y);
            transforms.Add((float)marker.Z);
            transforms.Add((float)marker.A);
            transforms.Add((float)marker.B);
            transforms.Add((float)marker.C);
        }

        if (!_session.TryStartCapturedCameraPlayback(transforms, speedMultiplier))
        {
            StopDollyPlayback("Dolly playback could not upload the in-process animation.");
            return;
        }

        var playbackThread = new Thread(() =>
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (!_writesAllowed || markers.Count < 2)
                    {
                        _ = Dispatcher.BeginInvoke(() => StopDollyPlayback("Dolly playback stopped: Halo 3 memory is no longer ready."));
                        return;
                    }

                    if (!_session.IsCapturedCameraPlaybackActive())
                    {
                        _ = Dispatcher.BeginInvoke(() => StopDollyPlayback("Dolly playback complete."));
                        return;
                    }

                    Thread.Sleep(10);
                }
            }
            catch (Exception ex)
            {
                _ = Dispatcher.BeginInvoke(() => StopDollyPlayback($"Dolly playback stopped: {ex.Message}"));
            }
        })
        {
            IsBackground = true,
            Name = "H3 Dolly Playback"
        };
        playbackThread.Start();
    }

    private static H3DollyMarker EvaluateDollyPath(IReadOnlyList<H3DollyMarker> markers, double elapsed)
    {
        var last = markers[markers.Count - 1];
        if (elapsed >= last.TimeSeconds)
            return last;

        var segmentIndex = 0;
        for (int i = 0; i < markers.Count - 1; i++)
        {
            if (elapsed >= markers[i].TimeSeconds && elapsed <= markers[i + 1].TimeSeconds)
            {
                segmentIndex = i;
                break;
            }
        }

        var start = markers[segmentIndex];
        var end = markers[segmentIndex + 1];
        var duration = Math.Max(0.001, end.TimeSeconds - start.TimeSeconds);
        var t = Math.Clamp((elapsed - start.TimeSeconds) / duration, 0, 1);
        var previous = markers[Math.Max(0, segmentIndex - 1)];
        var next = markers[Math.Min(markers.Count - 1, segmentIndex + 2)];
        var position = markers.Count > 2
            ? CentripetalCatmullRom(previous, start, end, next, t)
            : ((double X, double Y, double Z))(Lerp(start.X, end.X, t), Lerp(start.Y, end.Y, t), Lerp(start.Z, end.Z, t));
        // HaloDirector does not spline Euler yaw/pitch independently. It splines a
        // look direction and derives the angles from that vector, avoiding angular
        // velocity breaks and wrap discontinuities at every keyframe.
        var d0 = CameraDirection(previous.A, previous.B);
        var d1 = CameraDirection(start.A, start.B);
        var d2 = CameraDirection(end.A, end.B);
        var d3 = CameraDirection(next.A, next.B);
        var direction = NormalizeDirection((
            CatmullRom(d0.X, d1.X, d2.X, d3.X, t),
            CatmullRom(d0.Y, d1.Y, d2.Y, d3.Y, t),
            CatmullRom(d0.Z, d1.Z, d2.Z, d3.Z, t)));
        var a = Math.Atan2(direction.Y, direction.X);
        var b = Math.Atan2(direction.Z, Math.Sqrt((direction.X * direction.X) + (direction.Y * direction.Y)));
        var rollPeriod = Math.Max(Math.Max(Math.Abs(previous.C), Math.Abs(start.C)), Math.Max(Math.Abs(end.C), Math.Abs(next.C))) <= 7
            ? Math.PI * 2
            : 360.0;
        var c0 = UnwrapAngle(previous.C, start.C, rollPeriod);
        var c2 = UnwrapAngle(end.C, start.C, rollPeriod);
        var c3 = UnwrapAngle(next.C, c2, rollPeriod);
        var c = CatmullRom(c0, start.C, c2, c3, t);

        return new H3DollyMarker(start.Index, elapsed, (float)position.X, (float)position.Y, (float)position.Z, (float)a, (float)b, (float)c);
    }

    private static List<H3DollyMarker> BuildVelocityCleanPlaybackMarkers(IReadOnlyList<H3DollyMarker> markers)
    {
        if (markers.Count < 2)
            return markers.ToList();

        var distances = new double[markers.Count];
        for (var i = 1; i < markers.Count; i++)
        {
            distances[i] = distances[i - 1] + Distance(
                markers[i - 1].X,
                markers[i - 1].Y,
                markers[i - 1].Z,
                markers[i].X,
                markers[i].Y,
                markers[i].Z);
        }

        var totalDistance = distances[^1];
        var totalDuration = Math.Max(0.001, markers[^1].TimeSeconds);
        if (totalDistance < 0.001)
            return markers.ToList();

        var retimed = new List<H3DollyMarker>(markers.Count);
        for (var i = 0; i < markers.Count; i++)
        {
            var time = totalDuration * (distances[i] / totalDistance);
            retimed.Add(markers[i] with { TimeSeconds = time });
        }

        return retimed;
    }

    private static List<float> BuildArcLengthPlaybackTransforms(IReadOnlyList<H3DollyMarker> markers)
    {
        var duration = Math.Max(0.001, markers[^1].TimeSeconds);
        var outputCount = Math.Max(2, checked((int)Math.Ceiling(duration * 1000.0)) + 1);
        // Oversample the actual curve; chord lengths between this many closely spaced
        // evaluations accurately approximate its arc length without altering its shape.
        var denseCount = Math.Max(4096, outputCount * 4);
        var dense = new H3DollyMarker[denseCount];
        var cumulative = new double[denseCount];
        dense[0] = EvaluateDollyPath(markers, 0);
        for (var i = 1; i < denseCount; i++)
        {
            dense[i] = EvaluateDollyPath(markers, duration * i / (denseCount - 1.0));
            cumulative[i] = cumulative[i - 1] + Distance(
                dense[i - 1].X, dense[i - 1].Y, dense[i - 1].Z,
                dense[i].X, dense[i].Y, dense[i].Z);
        }

        var totalLength = cumulative[^1];
        var transforms = new List<float>(outputCount * 6);
        var upper = 1;
        for (var i = 0; i < outputCount; i++)
        {
            var desired = totalLength * i / (outputCount - 1.0);
            while (upper < denseCount - 1 && cumulative[upper] < desired)
                upper++;
            var lower = Math.Max(0, upper - 1);
            var span = cumulative[upper] - cumulative[lower];
            var amount = span > 0.0000001 ? (desired - cumulative[lower]) / span : 0;
            var a = dense[lower];
            var b = dense[upper];
            transforms.Add((float)Lerp(a.X, b.X, amount));
            transforms.Add((float)Lerp(a.Y, b.Y, amount));
            transforms.Add((float)Lerp(a.Z, b.Z, amount));
            transforms.Add((float)LerpAngle(a.A, b.A, amount));
            transforms.Add((float)LerpAngle(a.B, b.B, amount));
            transforms.Add((float)LerpAngle(a.C, b.C, amount));
        }

        return transforms;
    }

    private void StopDollyPlayback(string message)
    {
        if (_dollyPlaying)
        {
            _dollyPlaying = false;
            _dollyPlaybackCts?.Cancel();
            _dollyPlaybackCts?.Dispose();
            _dollyPlaybackCts = null;
            _dollyTimer.Stop();
            _dollyClock.Reset();
        }
        _session.DisableCameraTransformOverride();

        if (_xamlReady)
        {
            TxtFooter.Text = message;
            RefreshDollyButtons();
        }
    }

    private void StopDollyRecording(string message)
    {
        if (_dollyRecording)
        {
            _dollyRecording = false;
            _dollyRecordTimer.Stop();
            _dollyRecordClock.Reset();
            _lastRecordedCameraTransform = null;
        }

        if (_xamlReady)
        {
            TxtFooter.Text = message;
            RefreshDollyButtons();
        }
    }

    private void RefreshDollyButtons()
    {
        if (!_xamlReady)
            return;

        var enabled = _writesAllowed;
        var cameraHookReady = enabled && _session.IsCameraCaptureHookInstalled;
        BtnSwivelCam.IsEnabled = enabled && !_dollyPlaying && !_dollyRecording && !_dollyEditing;
        BtnSwivelCam.Content = _swivelCamActive ? "SWIVEL CAM: ON" : "SWIVEL CAM: OFF";
        BtnDollyRecord.IsEnabled = cameraHookReady && EnableExperimentalDollyPlayback && !_swivelCamActive && !_dollyPlaying && !_dollyEditing;
        BtnDollyRecord.Content = _dollyRecording ? "STOP REC" : "RECORD";
        BtnDollyCapture.IsEnabled = cameraHookReady && EnableExperimentalDollyPlayback && !_swivelCamActive && !_dollyPlaying && !_dollyRecording && !_dollyEditing;
        BtnDollyFindLive.IsEnabled = enabled && !_dollyPlaying && !_dollyRecording;
        BtnDollyFindLive.Content = EnableExperimentalDollyPlayback
            ? (_cameraScanCandidates.Count > 0 ? "NEXT CAM" : "SCAN CAM")
            : "ARCH REPORT";
        BtnDollyPlay.IsEnabled = cameraHookReady && EnableExperimentalDollyPlayback && !_swivelCamActive && !_dollyPlaying && !_dollyRecording && !_dollyEditing && DollyMarkers.Count >= 2;
        BtnDollyStop.IsEnabled = _dollyPlaying || _dollyRecording;
        BtnDollyDelete.IsEnabled = !_dollyPlaying && !_dollyRecording && !_dollyEditing && LstDollyMarkers.SelectedItem is H3DollyMarker;
        BtnDollyEdit.IsEnabled = cameraHookReady && !_dollyPlaying && !_dollyRecording && LstDollyMarkers.SelectedItem is H3DollyMarker;
        BtnDollyEdit.Content = _dollyEditing ? "SAVE KEY" : "EDIT KEY";
        BtnDollyClear.IsEnabled = !_dollyPlaying && !_dollyEditing && DollyMarkers.Count > 0;
        TxtDollySeconds.IsEnabled = EnableExperimentalDollyPlayback && !_dollyPlaying && !_dollyRecording && !_dollyEditing;
        SldDollySpeed.IsEnabled = EnableExperimentalDollyPlayback && !_dollyPlaying && !_dollyRecording && !_dollyEditing;
        TxtDollySpeed.IsEnabled = SldDollySpeed.IsEnabled;
        SldDollyRecordDistance.IsEnabled = EnableExperimentalDollyPlayback && !_dollyPlaying && !_dollyRecording && !_dollyEditing;
        BtnDollyTrack.IsEnabled = _dollyTrackOverlay is not null || (cameraHookReady && !_dollyEditing && DollyMarkers.Count > 0);
        EnsureDollyTrackOverlayShown();
    }

    private bool TryReadCameraPosition(out float x, out float y, out float z)
    {
        x = y = z = 0;
        if (_session.TryReadCapturedCameraPosition(out x, out y, out z))
            return true;

        if (_session.TryReadFloat(H3KnownAddresses.CameraLiveX, out x) &&
            _session.TryReadFloat(H3KnownAddresses.CameraLiveY, out y) &&
            _session.TryReadFloat(H3KnownAddresses.CameraLiveZ, out z))
        {
            return true;
        }

        if (_activeCameraAddress.HasValue)
        {
            var address = _activeCameraAddress.Value;
            if (_session.TryReadFloatAbsolute(address.XAddress, out x) &&
                _session.TryReadFloatAbsolute(address.YAddress, out y) &&
                _session.TryReadFloatAbsolute(address.ZAddress, out z))
            {
                return true;
            }

            _activeCameraAddress = null;
            _cameraScanCandidates.Clear();
            _cameraScanCandidateIndex = -1;
            x = y = z = 0;
        }

        return false;
    }

    private bool TryReadPlayerPosition(out float x, out float y, out float z)
    {
        x = y = z = 0;
        return _session.TryReadFloat(H3KnownAddresses.PlayerX, out x) &&
               _session.TryReadFloat(H3KnownAddresses.PlayerY, out y) &&
               _session.TryReadFloat(H3KnownAddresses.PlayerZ, out z);
    }

    private void SetPlayerCoordinateText(float x, float y, float z)
    {
        TxtPlayerX.Text = x.ToString("0.000", CultureInfo.InvariantCulture);
        TxtPlayerY.Text = y.ToString("0.000", CultureInfo.InvariantCulture);
        TxtPlayerZ.Text = z.ToString("0.000", CultureInfo.InvariantCulture);
    }

    private void SelectNextCameraCandidate()
    {
        if (_cameraScanCandidates.Count == 0)
        {
            TxtFooter.Text = "No camera candidates are cached. Click SCAN CAM first.";
            RefreshDollyButtons();
            return;
        }

        _cameraScanCandidateIndex = (_cameraScanCandidateIndex + 1) % _cameraScanCandidates.Count;
        var candidate = _cameraScanCandidates[_cameraScanCandidateIndex];
        _activeCameraAddress = candidate.Addresses;
        SetCoordinateText(candidate.CurrentX, candidate.CurrentY, candidate.CurrentZ);
        TxtFooter.Text = $"Camera candidate {_cameraScanCandidateIndex + 1}/{_cameraScanCandidates.Count}: X 0x{candidate.Addresses.XAddress:X}, Y 0x{candidate.Addresses.YAddress:X}, Z 0x{candidate.Addresses.ZAddress:X}. Click NEXT CAM if these numbers drift from the overlay.";
        RefreshDollyButtons();
    }

    private List<H3CameraScanCandidate> FindCameraCandidatesNear(float targetX, float targetY, float targetZ)
    {
        const float tolerance = 0.075f;
        const long maxScanBytes = 2L * 1024 * 1024 * 1024;
        var xSamples = _session.ScanReadableFloats(targetX - tolerance, targetX + tolerance, maxScanBytes, maxSamples: 5000).Samples;
        var ySamples = _session.ScanReadableFloats(targetY - tolerance, targetY + tolerance, maxScanBytes, maxSamples: 5000).Samples;
        var zSamples = _session.ScanReadableFloats(targetZ - tolerance, targetZ + tolerance, maxScanBytes, maxSamples: 5000).Samples;
        if (xSamples.Count == 0 || ySamples.Count == 0 || zSamples.Count == 0)
            return [];

        var candidates = new List<H3CameraScanCandidate>();
        foreach (var xSample in xSamples)
        {
            foreach (var ySample in ySamples.Where(s => Math.Abs(s.Address - xSample.Address) <= 0x400))
            {
                foreach (var zSample in zSamples.Where(s =>
                             Math.Abs(s.Address - xSample.Address) <= 0x400 &&
                             Math.Abs(s.Address - ySample.Address) <= 0x400))
                {
                    var minAddress = Math.Min(xSample.Address, Math.Min(ySample.Address, zSample.Address));
                    var maxAddress = Math.Max(xSample.Address, Math.Max(ySample.Address, zSample.Address));
                    var valueError =
                        Math.Abs(xSample.Value - targetX) +
                        Math.Abs(ySample.Value - targetY) +
                        Math.Abs(zSample.Value - targetZ);
                    var addressSpread = maxAddress - minAddress;
                    var score = addressSpread + (valueError * 1000);
                    candidates.Add(new H3CameraScanCandidate(
                        new H3CameraAddressSet(xSample.Address, ySample.Address, zSample.Address),
                        xSample.Value,
                        ySample.Value,
                        zSample.Value,
                        score));
                }
            }
        }

        return candidates
            .GroupBy(c => c.Addresses)
            .Select(g => g.OrderBy(c => c.Score).First())
            .OrderBy(c => c.Score)
            .Take(24)
            .ToList();
    }

    private bool TryParseCameraTextBoxes(out float x, out float y, out float z)
    {
        x = y = z = 0;
        return float.TryParse(TxtCamX.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out x) &&
               float.TryParse(TxtCamY.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out y) &&
               float.TryParse(TxtCamZ.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out z);
    }

    private bool TryGetTypedNonZeroCameraPosition(out float x, out float y, out float z)
    {
        if (TryParseCameraTextBoxes(out x, out y, out z) &&
            (!IsNearZero(x) || !IsNearZero(y) || !IsNearZero(z)))
        {
            return true;
        }

        x = y = z = 0;
        return false;
    }

    private static bool IsNearZero(float value)
        => Math.Abs(value) < 0.0001f;

    private static bool IsZeroVector(float x, float y, float z)
        => IsNearZero(x) && IsNearZero(y) && IsNearZero(z);

    private bool WriteCameraTransform(H3DollyMarker marker)
        => WriteCameraTransform(marker.X, marker.Y, marker.Z, marker.A, marker.B, marker.C);

    private bool WriteCameraTransform(double x, double y, double z, double a, double b, double c)
        => _session.TryWriteCapturedCameraTransform((float)x, (float)y, (float)z, (float)a, (float)b, (float)c);

    private static double LerpAngle(double start, double end, double amount)
    {
        var period = Math.Max(Math.Abs(start), Math.Abs(end)) <= 7 ? Math.PI * 2 : 360.0;
        var delta = (end - start) % period;
        if (delta > period / 2) delta -= period;
        if (delta < -period / 2) delta += period;
        return start + (delta * amount);
    }

    private static (double X, double Y, double Z) CameraDirection(double yaw, double pitch)
    {
        var horizontal = Math.Cos(pitch);
        return (Math.Cos(yaw) * horizontal, Math.Sin(yaw) * horizontal, Math.Sin(pitch));
    }

    private static (double X, double Y, double Z) NormalizeDirection((double X, double Y, double Z) direction)
    {
        var length = Math.Sqrt((direction.X * direction.X) + (direction.Y * direction.Y) + (direction.Z * direction.Z));
        return length > 0.0000001
            ? (direction.X / length, direction.Y / length, direction.Z / length)
            : (1, 0, 0);
    }

    private static double UnwrapAngle(double angle, double reference, double period)
    {
        var delta = (angle - reference) % period;
        if (delta > period / 2) delta -= period;
        if (delta < -period / 2) delta += period;
        return reference + delta;
    }

    private static double CatmullRom(double p0, double p1, double p2, double p3, double t)
    {
        var t2 = t * t;
        var t3 = t2 * t;
        return 0.5 * ((2 * p1) + ((-p0 + p2) * t) +
                      ((2 * p0 - 5 * p1 + 4 * p2 - p3) * t2) +
                      ((-p0 + 3 * p1 - 3 * p2 + p3) * t3));
    }

    private static (double X, double Y, double Z) CentripetalCatmullRom(
        H3DollyMarker p0,
        H3DollyMarker p1,
        H3DollyMarker p2,
        H3DollyMarker p3,
        double t)
    {
        const double alpha = 0.5;
        var t0 = 0.0;
        var t1 = NextCentripetalTime(t0, p0, p1, alpha);
        var t2 = NextCentripetalTime(t1, p1, p2, alpha);
        var t3 = NextCentripetalTime(t2, p2, p3, alpha);
        var target = Lerp(t1, t2, t);

        if (Math.Abs(t1 - t0) < 0.000001 ||
            Math.Abs(t2 - t1) < 0.000001 ||
            Math.Abs(t3 - t2) < 0.000001)
        {
            return (Lerp(p1.X, p2.X, t), Lerp(p1.Y, p2.Y, t), Lerp(p1.Z, p2.Z, t));
        }

        var a1 = InterpolatePoint(p0, p1, t0, t1, target);
        var a2 = InterpolatePoint(p1, p2, t1, t2, target);
        var a3 = InterpolatePoint(p2, p3, t2, t3, target);
        var b1 = InterpolatePoint(a1, a2, t0, t2, target);
        var b2 = InterpolatePoint(a2, a3, t1, t3, target);
        return InterpolatePoint(b1, b2, t1, t2, target);
    }

    private static double NextCentripetalTime(double current, H3DollyMarker a, H3DollyMarker b, double alpha)
    {
        var distance = Distance(a.X, a.Y, a.Z, b.X, b.Y, b.Z);
        return current + Math.Pow(Math.Max(distance, 0.000001), alpha);
    }

    private static (double X, double Y, double Z) InterpolatePoint(
        H3DollyMarker a,
        H3DollyMarker b,
        double ta,
        double tb,
        double t)
        => InterpolatePoint((a.X, a.Y, a.Z), (b.X, b.Y, b.Z), ta, tb, t);

    private static (double X, double Y, double Z) InterpolatePoint(
        (double X, double Y, double Z) a,
        (double X, double Y, double Z) b,
        double ta,
        double tb,
        double t)
    {
        var denominator = tb - ta;
        if (Math.Abs(denominator) < 0.000001)
            return b;

        var left = (tb - t) / denominator;
        var right = (t - ta) / denominator;
        return ((left * a.X) + (right * b.X), (left * a.Y) + (right * b.Y), (left * a.Z) + (right * b.Z));
    }

    private async Task<bool> BurstWriteAbsoluteTripleAsync(long xAddress, long yAddress, long zAddress, float x, float y, float z, int milliseconds)
    {
        _coordinateWriteBurstActive = true;
        BtnCameraStructTest.IsEnabled = false;
        try
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(milliseconds);
            var wroteAny = false;
            while (DateTime.UtcNow < deadline && _writesAllowed)
            {
                wroteAny |=
                    _session.TryWriteFloatAbsolute(xAddress, x) &&
                    _session.TryWriteFloatAbsolute(yAddress, y) &&
                    _session.TryWriteFloatAbsolute(zAddress, z);
                await Task.Delay(5);
            }

            return wroteAny;
        }
        finally
        {
            _coordinateWriteBurstActive = false;
            BtnCameraStructTest.IsEnabled = _writesAllowed && _cameraStructCandidates.Count > 0;
        }
    }

    private double ParseDollySegmentSeconds()
    {
        if (double.TryParse(TxtDollySeconds.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var seconds))
            return Math.Clamp(seconds, 0.1, 120);

        TxtDollySeconds.Text = "2.0";
        return 2.0;
    }

    private bool AddDollyMarkerFromPosition(float x, float y, float z, double timeSeconds)
    {
        if (!_session.TryReadCapturedCameraOrientation(out var a, out var b, out var c))
        {
            StopDollyRecording("Dolly recording stopped: the live camera rotation is unavailable.");
            return false;
        }
        var marker = new H3DollyMarker(DollyMarkers.Count + 1, timeSeconds, x, y, z, a, b, c);
        DollyMarkers.Add(marker);
        SelectDollyMarkerWithoutNavigation(marker);
        SetCoordinateText(x, y, z);
        _lastRecordedCameraTransform = (x, y, z, a, b, c);
        RefreshDollyButtons();
        return true;
    }

    private void SetCoordinateText(float x, float y, float z)
    {
        TxtCamX.Text = x.ToString("0.###", CultureInfo.InvariantCulture);
        TxtCamY.Text = y.ToString("0.###", CultureInfo.InvariantCulture);
        TxtCamZ.Text = z.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private string BuildCameraArchitectureReport()
    {
        var lines = new List<string>
        {
            $"CAMERA ARCHITECTURE REPORT @ {DateTime.Now:HH:mm:ss}",
            "",
            "Status",
            $"  Experimental dolly playback: {(EnableExperimentalDollyPlayback ? "enabled" : "parked")}",
            "  Current finding: pan-cam XYZ float scans find readable/writable copies, not the render camera source.",
            "",
            "Required real targets",
            "  CameraHookAddress: unknown",
            "  FOV: unknown",
            "  serverSeconds: unknown",
            "  serverTime: unknown",
            "  timescale: unknown",
            "",
            "Known camera-affecting patches"
        };

        AppendPatchBytes(lines, "Freecam state", H3KnownAddresses.FreecamState, 6);
        AppendPatchBytes(lines, "Disable camera control", H3KnownAddresses.DisableCameraControl, 6);
        AppendPatchBytes(lines, "Third-person branch", H3KnownAddresses.ThirdPersonBranch, 2);
        AppendPatchBytes(lines, "Coordinate gate A", H3KnownAddresses.CoordinatesA, 6);
        AppendPatchBytes(lines, "Coordinate gate B", H3KnownAddresses.CoordinatesB, 6);
        AppendPatchBytes(lines, "Coordinate gate C", H3KnownAddresses.CoordinatesC, 2);
        AppendPatchBytes(lines, "Coordinate gate D", H3KnownAddresses.CoordinatesD, 2);

        lines.Add("");
        lines.Add("Confirmed camera coordinate copies");
        AppendFloatValue(lines, "Live Camera X", H3KnownAddresses.CameraLiveX);
        AppendFloatValue(lines, "Live Camera Y", H3KnownAddresses.CameraLiveY);
        AppendFloatValue(lines, "Live Camera Z", H3KnownAddresses.CameraLiveZ);
        AppendFloatValue(lines, "Mirror Camera X", H3KnownAddresses.CameraMirrorX);
        AppendFloatValue(lines, "Mirror Camera Y", H3KnownAddresses.CameraMirrorY);
        AppendFloatValue(lines, "Mirror Camera Z", H3KnownAddresses.CameraMirrorZ);

        lines.Add("");
        lines.Add("Stale legacy coordinate chain");
        AppendFloatValue(lines, "Legacy Camera X", H3KnownAddresses.CameraX);
        AppendFloatValue(lines, "Legacy Camera Y", H3KnownAddresses.CameraY);
        AppendFloatValue(lines, "Legacy Camera Z", H3KnownAddresses.CameraZ);
        lines.Add("  Note: this chain reads 0/0/0 in sessions where pan-cam overlay is non-zero.");

        lines.Add("");
        lines.Add("Next implementation rule");
        lines.Add("  Do not promote float-scan candidates to dolly playback unless a controlled write visibly moves the camera.");
        lines.Add("  Treat overlay XYZ matches as readout/discovery only.");
        lines.Add("  Promote only stable signatures or pointer chains for camera hook, FOV, server time, and timescale.");

        return string.Join(Environment.NewLine, lines);
    }

    private static H3DiscoveryPreset BuildCameraArchitecturePreset()
        => new(
            "camera-architecture",
            "CAMERA ARCHITECTURE",
            [
                new("third-person-code", "Third-person branch neighborhood", new H3Address("halo3.dll", 0x132850), 0x80),
                new("freecam-code", "Freecam / camera patch neighborhood", new H3Address("halo3.dll", 0x131080), 0x380),
                new("coord-code-a", "Coordinate gate code neighborhood", new H3Address("halo3.dll", 0xF7600), 0x140),
                new("coord-code-b", "Coordinate/freecam code neighborhood", new H3Address("halo3.dll", 0x211E80), 0x120),
                new("state-camera", "Legacy camera state window", new H3Address("halo3.dll", 0x2030288, 0x2BAF00), 0x300),
                new("state-timing", "Timing / pause / tick state window", new H3Address("halo3.dll", 0x2030288, 0x10000), 0x1600),
                new("sim-functions", "Simulation function variable IDs", new H3Address("halo3.dll", 0x781B00), 0x400),
            ]);

    private static H3DiscoveryPreset BuildCameraMovementPreset()
        => new(
            "camera-movement",
            "CAMERA MOVEMENT",
            [
                new("state-core-a", "Game-state 0x00000-0x0FFFF", new H3Address("halo3.dll", 0x2030288, 0x0), 0x10000),
                new("state-core-b", "Game-state 0x10000-0x1FFFF", new H3Address("halo3.dll", 0x2030288, 0x10000), 0x10000),
                new("state-core-c", "Game-state 0x20000-0x2FFFF", new H3Address("halo3.dll", 0x2030288, 0x20000), 0x10000),
                new("state-core-d", "Game-state 0x30000-0x3FFFF", new H3Address("halo3.dll", 0x2030288, 0x30000), 0x10000),
                new("camera-near-a", "Legacy camera region 0x2BA000-0x2BBFFF", new H3Address("halo3.dll", 0x2030288, 0x2BA000), 0x2000),
                new("camera-near-b", "Legacy camera region 0x2BC000-0x2BDFFF", new H3Address("halo3.dll", 0x2030288, 0x2BC000), 0x2000),
                new("render-state-a", "Render-ish state 0x40000-0x4FFFF", new H3Address("halo3.dll", 0x2030288, 0x40000), 0x10000),
                new("render-state-b", "Render-ish state 0x50000-0x5FFFF", new H3Address("halo3.dll", 0x2030288, 0x50000), 0x10000),
            ]);

    private static string SlugifyReportKind(string value)
    {
        var chars = value
            .ToLowerInvariant()
            .Select(ch => char.IsLetterOrDigit(ch) ? ch : '-')
            .ToArray();
        return string.Join('-', new string(chars).Split('-', StringSplitOptions.RemoveEmptyEntries));
    }

    private void AppendPatchBytes(List<string> lines, string label, H3Address address, int length)
    {
        if (_session.TryReadBytes(address, length, out var bytes))
        {
            lines.Add($"  {label}: {address} = {FormatBytes(bytes)}");
            return;
        }

        lines.Add($"  {label}: {address} = unreadable");
    }

    private void AppendFloatValue(List<string> lines, string label, H3Address address)
    {
        if (_session.TryReadFloat(address, out var value))
        {
            lines.Add($"  {label}: {address} = {value:R}");
            return;
        }

        lines.Add($"  {label}: {address} = unreadable");
    }

    private static string FormatBytes(byte[] bytes)
        => string.Join(' ', bytes.Select(b => b.ToString("X2", CultureInfo.InvariantCulture)));


    private void AppendPointerScan(List<string> lines, string label, H3Address target)
    {
        if (!_session.TryResolveAddress(target, out var absoluteAddress))
        {
            lines.Add($"{label}: {target} unresolved");
            lines.Add("");
            return;
        }

        var matches = _session.ScanReadablePointersTo(absoluteAddress, maxBytesToScan: 768L * 1024 * 1024, maxMatches: 128);
        var first = matches.FirstOrDefault();
        lines.Add($"{label}: {target} abs 0x{absoluteAddress:X} -> {matches.Count} pointer match(es)");
        if (first is not null)
            lines.Add($"  scanned {first.ScannedRegions} region(s), {first.ReadableRegions} readable, 0x{first.ScannedBytes:X} byte(s)");

        foreach (var match in matches.Take(32))
        {
            lines.Add($"  ptr at 0x{match.Address:X} -> 0x{match.TargetAddress:X}");
            AppendPointerNeighborhood(lines, match.Address, absoluteAddress);

            var parentMatches = _session.ScanReadablePointersTo(match.Address, maxBytesToScan: 256L * 1024 * 1024, maxMatches: 16);
            if (parentMatches.Count > 0)
            {
                lines.Add($"    parent pointers to 0x{match.Address:X}: {parentMatches.Count}");
                foreach (var parent in parentMatches.Take(8))
                    lines.Add($"      ptr at 0x{parent.Address:X} -> 0x{parent.TargetAddress:X}");
            }
        }

        if (matches.Count > 32)
            lines.Add($"  ... {matches.Count - 32} more pointer match(es) hidden");

        lines.Add("");
    }

    private void AppendCodeWindow(List<string> lines, string label, H3Address center, int before, int length)
    {
        var startOffset = Math.Max(0, center.BaseOffset - before);
        var relativeMark = (int)(center.BaseOffset - startOffset);
        var address = new H3Address(center.ModuleName, startOffset);
        if (!_session.TryReadBytes(address, length, out var bytes))
        {
            lines.Add($"{label}: {address} len 0x{length:X} unreadable");
            lines.Add("");
            return;
        }

        lines.Add($"{label}: {address} len 0x{length:X}, marker +0x{relativeMark:X}");
        for (int offset = 0; offset < bytes.Length; offset += 16)
        {
            var count = Math.Min(16, bytes.Length - offset);
            var chunk = bytes.Skip(offset).Take(count).ToArray();
            var marker = relativeMark >= offset && relativeMark < offset + count ? " <PATCH>" : "";
            lines.Add($"  +0x{offset:X4}: {FormatBytes(chunk)}{marker}");
        }

        lines.Add("");
    }

    private void AppendPointerNeighborhood(List<string> lines, long pointerAddress, long targetAddress)
    {
        const int radius = 0x40;
        var start = pointerAddress - radius;
        if (!_session.TryReadBytesAbsolute(start, radius * 2, out var bytes))
        {
            lines.Add("    neighborhood unreadable");
            return;
        }

        lines.Add("    nearby qwords:");
        for (int offset = 0; offset + 8 <= bytes.Length; offset += 8)
        {
            var address = start + offset;
            var qword = BitConverter.ToInt64(bytes, offset);
            var marker = address == pointerAddress
                ? " <MATCH>"
                : Math.Abs(qword - targetAddress) <= 0x2000
                    ? " <CAMERA-NEAR>"
                    : "";

            if (marker.Length == 0 && (qword < 0x10000 || qword > 0x0000800000000000))
                continue;

            lines.Add($"      0x{address:X}: 0x{qword:X}{marker}");
        }

        lines.Add("    nearby f32:");
        for (int offset = 0; offset + 4 <= bytes.Length; offset += 4)
        {
            var value = BitConverter.ToSingle(bytes, offset);
            if (!IsLikelyCameraComponent(value) || Math.Abs(value) < 0.0001f)
                continue;

            lines.Add($"      0x{start + offset:X}: {value:0.###}");
        }
    }

    private static string FormatPatternContext(H3ModulePatternMatch match)
    {
        var parts = new List<string>();
        for (int i = 0; i < match.ContextBytes.Length; i++)
        {
            var text = match.ContextBytes[i].ToString("X2", CultureInfo.InvariantCulture);
            parts.Add(i >= match.PatternOffsetInContext && i < match.PatternOffsetInContext + 4 ? $"[{text}]" : text);
        }

        return string.Join(' ', parts);
    }

    private static string FormatRipContext(H3RipRelativeReference match)
    {
        var parts = new List<string>();
        for (int i = 0; i < match.ContextBytes.Length; i++)
        {
            var text = match.ContextBytes[i].ToString("X2", CultureInfo.InvariantCulture);
            parts.Add(i >= match.InstructionOffsetInContext && i < match.InstructionOffsetInContext + match.InstructionLength
                ? $"[{text}]"
                : text);
        }

        return string.Join(' ', parts);
    }

    private void RenumberDollyMarkers()
    {
        for (int i = 0; i < DollyMarkers.Count; i++)
            DollyMarkers[i] = DollyMarkers[i] with { Index = i + 1 };
    }

    private static double Lerp(double start, double end, double t)
        => start + ((end - start) * t);

    private static double Distance(float ax, float ay, float az, float bx, float by, float bz)
    {
        var dx = ax - bx;
        var dy = ay - by;
        var dz = az - bz;
        return Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz));
    }

    private void BtnDiscoveryBase_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        var preset = GetSelectedDiscoveryPreset();
        var snapshot = CaptureDiscoverySnapshot(preset);
        _discoveryBaselines[preset.Id] = snapshot;
        var report = FormatDiscoveryBaseline(snapshot);
        var path = SaveDiscoveryReport(snapshot.Preset.Id, "base", report);
        TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
        TxtFooter.Text = $"Discovery baseline captured for {preset.Name}. Perform one in-game action, then click DIFF.";
    }

    private void BtnDiscoveryDiff_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        var preset = GetSelectedDiscoveryPreset();
        if (!_discoveryBaselines.TryGetValue(preset.Id, out var baseline))
        {
            TxtDiscoveryOutput.Text = "No baseline for this preset. Click BASE first, perform one controlled action, then click DIFF.";
            return;
        }

        var after = CaptureDiscoverySnapshot(preset);
        var report = FormatDiscoveryDiff(baseline, after);
        var path = SaveDiscoveryReport(after.Preset.Id, "diff", report);
        TxtDiscoveryOutput.Text = $"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}";
        TxtFooter.Text = $"Discovery diff complete for {preset.Name}. Review changed addresses before promoting anything to a toggle.";
    }

    private void BtnDiscoveryCopy_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(TxtDiscoveryOutput.Text))
        {
            Clipboard.SetText(TxtDiscoveryOutput.Text);
            TxtFooter.Text = "Discovery output copied to clipboard.";
        }
    }

    private void BtnDiscoveryClear_Click(object sender, RoutedEventArgs e)
    {
        _discoveryBaselines.Clear();
        TxtDiscoveryOutput.Text = "Discovery baselines cleared.";
        TxtFooter.Text = "Discovery baselines cleared.";
    }

    private void BtnSkullByteRead_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        var address = GetSelectedSkullLabAddress();
        if (_session.TryReadByte(address, out var value))
        {
            var message = $"{address}: 0x{value:X2}  bits {Convert.ToString(value, 2).PadLeft(8, '0')}";
            TxtSkullByteStatus.Text = message;
            TxtDiscoveryOutput.Text = $"SKULL BIT LAB READ{Environment.NewLine}{message}";
            TxtFooter.Text = "Skull bit lab byte read.";
        }
        else
        {
            SetSkullLabStatus($"{address}: READ FAILED");
            TxtFooter.Text = "Skull bit lab could not read that byte.";
        }
    }

    private void SkullBitButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        if (sender is not Button button || !TryParseHex(button.Tag?.ToString() ?? "", out var parsedMask))
            return;

        ToggleSkullLabBit(GetSelectedSkullLabAddress(), (byte)parsedMask, button.Content?.ToString() ?? "BIT");
    }

    private void KnownSkullBitButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        if (sender is not Button button)
            return;

        var parts = (button.Tag?.ToString() ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 2 || !TryParseHex(parts[0], out var offset) || !TryParseHex(parts[1], out var mask))
            return;

        ToggleSkullLabBit(new H3Address("halo3.dll", 0x2030288, offset), (byte)mask, button.Content?.ToString() ?? "KNOWN");
    }

    private void ToggleSkullLabBit(H3Address address, byte mask, string label)
    {
        if (!_session.TryReadByte(address, out var current))
        {
            SetSkullLabStatus($"{address}: READ FAILED");
            TxtFooter.Text = "Skull bit lab could not read before toggling.";
            return;
        }

        RememberOriginalBytes(_session, address, 1);
        var next = (byte)(current ^ mask);
        if (!_session.TryWriteByte(address, next))
        {
            SetSkullLabStatus($"{address}: WRITE FAILED");
            TxtFooter.Text = "Skull bit lab write failed.";
            return;
        }

        var verified = _session.TryReadByte(address, out var actual);
        var line = verified
            ? $"{address}: 0x{current:X2} -> requested 0x{next:X2}, actual 0x{actual:X2}  toggled {label}"
            : $"{address}: 0x{current:X2} -> requested 0x{next:X2}, readback failed  toggled {label}";
        SetSkullLabStatus(line);
        TxtFooter.Text = $"Skull bit lab toggled {label}. Use RESTORE ALL to revert lab writes.";
        RefreshRows();
    }

    private void SetSkullLabStatus(string message)
    {
        TxtSkullByteStatus.Text = message;
        TxtDiscoveryOutput.Text = $"SKULL BIT LAB{Environment.NewLine}{message}";
    }

    private void BtnLiveProbe_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        var lines = new List<string>
        {
            $"LIVE PROBE @ {DateTime.Now:HH:mm:ss}",
            "Candidate values are guesses until they repeat across controlled tests.",
            ""
        };

        foreach (var target in BuildLiveProbeTargets())
        {
            if (!_session.TryReadFloat(target.Address, out var value))
            {
                lines.Add($"{target.Name}: READ FAILED");
                continue;
            }

            var delta = _liveProbeLastValues.TryGetValue(target.Id, out var previous)
                ? $"  delta {value - previous:+0.###;-0.###;0}"
                : "";
            _liveProbeLastValues[target.Id] = value;
            lines.Add($"{target.Name}: {value:0.###}{delta}");
            lines.Add($"  {target.Address}");
        }

        SetProbeOutput(string.Join(Environment.NewLine, lines));
        TxtFooter.Text = "Live probe sampled candidate values. Repeat after one controlled action.";
    }

    private void BtnLiveProbeReset_Click(object sender, RoutedEventArgs e)
    {
        _liveProbeLastValues.Clear();
        SetProbeOutput("Live probe baseline cleared.");
        TxtFooter.Text = "Live probe baseline cleared.";
    }

    private void BtnVitalScanBase_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        var previousCursor = Mouse.OverrideCursor;
        Mouse.OverrideCursor = Cursors.Wait;
        try
        {
            SetProbeOutput("VITAL MEMORY SCAN baseline is running. The UI may pause briefly while MCC memory is read.");
            Dispatcher.Invoke(() => { }, DispatcherPriority.Background);

            var scan = _session.ScanWritableFloats(0.001f, 1.25f);
            _vitalScanBaseline = scan.Samples
                .GroupBy(s => s.Address)
                .ToDictionary(g => g.Key, g => g.Last().Value);
            _vitalScanBaselineAt = DateTime.Now;

            var report = FormatVitalScanBaseline(scan);
            var path = SaveDiscoveryReport("vital-memory-scan", "base", report);
            SetProbeOutput($"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}");
            TxtFooter.Text = $"Vital scan baseline captured: {_vitalScanBaseline.Count:N0} health-like floats. Take damage, then click VDIFF.";
        }
        catch (Exception ex)
        {
            SetProbeOutput($"VITAL MEMORY SCAN baseline failed: {ex.Message}");
            TxtFooter.Text = "Vital scan baseline failed. See output for details.";
        }
        finally
        {
            Mouse.OverrideCursor = previousCursor;
        }
    }

    private void BtnVitalScanDiff_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureReadyForDiscovery())
            return;

        if (_vitalScanBaseline.Count == 0 || _vitalScanBaselineAt is null)
        {
            SetProbeOutput("No vitality scan baseline yet. Click VBASE while healthy, take one controlled damage event, then click VDIFF.");
            return;
        }

        var previousCursor = Mouse.OverrideCursor;
        Mouse.OverrideCursor = Cursors.Wait;
        try
        {
            SetProbeOutput("VITAL MEMORY SCAN diff is running. The UI may pause briefly while MCC memory is read.");
            Dispatcher.Invoke(() => { }, DispatcherPriority.Background);

            var scan = _session.ScanWritableFloats(0.0f, 1.25f);
            var report = FormatVitalScanDiff(_vitalScanBaseline, _vitalScanBaselineAt.Value, scan);
            var path = SaveDiscoveryReport("vital-memory-scan", "diff", report);
            SetProbeOutput($"{report}{Environment.NewLine}{Environment.NewLine}Saved: {path}");
            TxtFooter.Text = "Vital scan diff complete. Strong candidates are addresses that drop only when damage happens.";
        }
        catch (Exception ex)
        {
            SetProbeOutput($"VITAL MEMORY SCAN diff failed: {ex.Message}");
            TxtFooter.Text = "Vital scan diff failed. See output for details.";
        }
        finally
        {
            Mouse.OverrideCursor = previousCursor;
        }
    }

    private void SetProbeOutput(string text)
    {
        TxtDiscoveryOutput.Text = text;
    }

    private H3Address GetSelectedSkullLabAddress()
    {
        var tag = (CmbSkullByte.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "0xFDE6";
        return TryParseHex(tag, out var offset)
            ? new H3Address("halo3.dll", 0x2030288, offset)
            : H3KnownAddresses.Bandana;
    }

    private bool EnsureReadyForDiscovery()
    {
        if (_writesAllowed)
            return true;

        TxtDiscoveryOutput.Text = "Discovery needs an attached Halo 3 session with EAC not detected. It only reads memory, but it still needs halo3.dll visible.";
        TxtFooter.Text = "Discovery is gated until MCC is attached with EAC not detected and halo3.dll loaded.";
        return false;
    }

    private H3DiscoveryPreset GetSelectedDiscoveryPreset()
    {
        var id = (CmbDiscoveryPreset.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "trainer-pack";
        return BuildDiscoveryPresets().FirstOrDefault(p => p.Id.Equals(id, StringComparison.OrdinalIgnoreCase))
            ?? BuildDiscoveryPresets()[0];
    }

    private H3DiscoverySnapshot CaptureDiscoverySnapshot(H3DiscoveryPreset preset)
    {
        var captures = new List<H3DiscoveryCapture>();
        foreach (var probe in preset.Probes)
        {
            var readable = _session.TryReadBytes(probe.Address, probe.Length, out var bytes);
            captures.Add(new H3DiscoveryCapture(probe, readable, readable ? bytes : []));
        }

        return new H3DiscoverySnapshot(preset, DateTime.Now, captures);
    }

    private static string FormatDiscoveryBaseline(H3DiscoverySnapshot snapshot)
    {
        var lines = new List<string>
        {
            $"{snapshot.Preset.Name} baseline @ {snapshot.CapturedAt:HH:mm:ss}",
            "Perform one controlled action, then click DIFF.",
            ""
        };

        foreach (var capture in snapshot.Captures)
        {
            var state = capture.Readable ? $"{capture.Bytes.Length} bytes" : "READ FAILED";
            lines.Add($"{capture.Probe.Name}: {state}");
            lines.Add($"  {capture.Probe.Address} len 0x{capture.Probe.Length:X}");
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static string FormatDiscoveryDiff(H3DiscoverySnapshot before, H3DiscoverySnapshot after)
    {
        var lines = new List<string>
        {
            $"{after.Preset.Name} diff",
            $"BASE {before.CapturedAt:HH:mm:ss} -> DIFF {after.CapturedAt:HH:mm:ss}",
            ""
        };

        int totalChanges = 0;
        foreach (var afterCapture in after.Captures)
        {
            var beforeCapture = before.Captures.FirstOrDefault(c => c.Probe.Key == afterCapture.Probe.Key);
            if (beforeCapture is null || !beforeCapture.Readable || !afterCapture.Readable)
            {
                lines.Add($"{afterCapture.Probe.Name}: skipped (read failed)");
                continue;
            }

            var allChanges = FindChangedBytes(beforeCapture.Bytes, afterCapture.Bytes).ToList();
            var changes = allChanges.Take(96).ToList();
            var changedCount = allChanges.Count;
            totalChanges += changedCount;

            lines.Add($"{afterCapture.Probe.Name}: {changedCount} changed byte(s)");
            foreach (var change in changes)
            {
                var address = afterCapture.Probe.FormatAddress(change.Offset);
                var detail = FormatDiscoveryChangeDetail(beforeCapture.Bytes, afterCapture.Bytes, change.Offset);
                lines.Add($"  {address}: {change.Before:X2} -> {change.After:X2}{detail}");
            }

            if (changedCount > changes.Count)
                lines.Add($"  ... {changedCount - changes.Count} more change(s) hidden");

            var summaries = SummarizeDiscoveryCandidates(afterCapture.Probe, beforeCapture.Bytes, afterCapture.Bytes, allChanges);
            if (summaries.Count > 0)
            {
                lines.Add("  Candidate values:");
                foreach (var summary in summaries.Take(16))
                    lines.Add($"    {summary}");
            }

            lines.Add("");
        }

        if (totalChanges == 0)
            lines.Add("No byte changes found in these probes. Try a narrower controlled action, a different preset, or a longer interval between BASE and DIFF.");
        else
            lines.Add("Promote only repeatable changes that move with the same action across multiple BASE/DIFF passes.");

        return string.Join(Environment.NewLine, lines);
    }

    private static string FormatCameraMovementDiff(H3DiscoverySnapshot before, H3DiscoverySnapshot after)
    {
        var lines = new List<string>
        {
            $"{after.Preset.Name} diff",
            $"BASE {before.CapturedAt:HH:mm:ss} -> DIFF {after.CapturedAt:HH:mm:ss}",
            "",
            "Looking for adjacent f32 XYZ triples that changed when the freecam moved.",
            "Strong candidates should resemble the visible pan-cam numbers and repeat across multiple MOVE BASE / MOVE DIFF passes.",
            ""
        };

        int totalChanges = 0;
        int totalTriples = 0;
        foreach (var afterCapture in after.Captures)
        {
            var beforeCapture = before.Captures.FirstOrDefault(c => c.Probe.Key == afterCapture.Probe.Key);
            if (beforeCapture is null || !beforeCapture.Readable || !afterCapture.Readable)
            {
                lines.Add($"{afterCapture.Probe.Name}: skipped (read failed)");
                continue;
            }

            var allChanges = FindChangedBytes(beforeCapture.Bytes, afterCapture.Bytes).ToList();
            totalChanges += allChanges.Count;
            var triples = SummarizeCameraMovementTriples(afterCapture.Probe, beforeCapture.Bytes, afterCapture.Bytes, allChanges).ToList();
            totalTriples += triples.Count;

            lines.Add($"{afterCapture.Probe.Name}: {allChanges.Count} changed byte(s), {triples.Count} XYZ-like candidate(s)");
            foreach (var triple in triples.Take(24))
                lines.Add($"  {triple}");

            if (triples.Count > 24)
                lines.Add($"  ... {triples.Count - 24} more XYZ-like candidate(s) hidden");

            if (triples.Count == 0 && allChanges.Count > 0)
            {
                foreach (var change in allChanges.Take(12))
                {
                    var address = afterCapture.Probe.FormatAddress(change.Offset);
                    var detail = FormatDiscoveryChangeDetail(beforeCapture.Bytes, afterCapture.Bytes, change.Offset);
                    lines.Add($"  raw {address}: {change.Before:X2} -> {change.After:X2}{detail}");
                }

                if (allChanges.Count > 12)
                    lines.Add($"  ... {allChanges.Count - 12} raw change(s) hidden");
            }

            lines.Add("");
        }

        if (totalChanges == 0)
            lines.Add("No movement changes found. Verify Freecam is enabled, move a visible distance, then repeat MOVE BASE / MOVE DIFF.");
        else if (totalTriples == 0)
            lines.Add("Changes were found, but no useful XYZ-like float triples appeared. We may need a different base pointer or a larger region.");
        else
            lines.Add("Next: repeat the test from another position. Promote only triples whose before/after values track the visible pan-cam coordinates.");

        return string.Join(Environment.NewLine, lines);
    }

    private IEnumerable<H3CameraStructCandidate> BuildCameraStructCandidates(IEnumerable<H3FloatTripleScanMatch> matches)
    {
        var excluded = new[]
        {
            H3KnownAddresses.CameraLiveX,
            H3KnownAddresses.CameraMirrorX,
            H3KnownAddresses.CameraX,
        }
        .Select(address => _session.TryResolveAddress(address, out var absolute) ? absolute : 0)
        .Where(address => address != 0)
        .ToArray();

        foreach (var match in matches)
        {
            if (excluded.Any(address => Math.Abs(match.Address - address) < 0x400))
                continue;

            var baseAddress = match.Address - 0x28;
            var score = ScoreCameraStructCandidate(baseAddress, match.Address);
            yield return new H3CameraStructCandidate(baseAddress, match.Address, match.Address + 4, match.Address + 8, match.X, match.Y, match.Z, score);
        }
    }

    private double ScoreCameraStructCandidate(long baseAddress, long xAddress)
    {
        var score = 0.0;
        if (_session.TryReadBytesAbsolute(baseAddress, 0x80, out var bytes))
        {
            var usefulFloats = 0;
            for (int offset = 0; offset + 4 <= bytes.Length; offset += 4)
            {
                var value = BitConverter.ToSingle(bytes, offset);
                if (IsLikelyCameraComponent(value) && Math.Abs(value) > 0.0001f)
                    usefulFloats++;
            }

            score += Math.Min(usefulFloats, 32);
        }

        if ((xAddress & 0xF) == 0)
            score += 4;
        if ((baseAddress & 0xF) == 0)
            score += 2;

        return score;
    }

    private static string FormatCameraStructScanReport(
        float x,
        float y,
        float z,
        H3FloatTripleScanResult result,
        IReadOnlyList<H3CameraStructCandidate> candidates)
    {
        var lines = new List<string>
        {
            $"CAMERA STRUCT SCAN @ {DateTime.Now:HH:mm:ss}",
            "",
            $"Target XYZ: {x:0.###}, {y:0.###}, {z:0.###}",
            $"Scanned {result.ScannedRegions} region(s), {result.ReadableRegions} writable/readable, 0x{result.ScannedBytes:X} byte(s)",
            $"Raw XYZ matches: {result.Matches.Count}",
            $"Candidates after excluding known readout/cache copies: {candidates.Count}",
            "",
            "Assumption",
            "  The coord-control code window uses values shaped like [struct+0x28], [struct+0x2C], [struct+0x30].",
            "  Candidate Base below is therefore XAddress - 0x28.",
            ""
        };

        if (candidates.Count == 0)
        {
            lines.Add("No candidates survived filtering. Try moving the freecam, wait for the coordinate boxes to update, then scan again.");
            return string.Join(Environment.NewLine, lines);
        }

        foreach (var candidate in candidates
                     .OrderByDescending(c => c.Score)
                     .ThenBy(c => c.XAddress)
                     .Take(48)
                     .Select((candidate, index) => (candidate, index)))
        {
            lines.Add($"{candidate.index + 1:00}. base 0x{candidate.candidate.BaseAddress:X}  XYZ @ 0x{candidate.candidate.XAddress:X}/0x{candidate.candidate.YAddress:X}/0x{candidate.candidate.ZAddress:X}  score {candidate.candidate.Score:0.##}");
            lines.Add($"    {candidate.candidate.X:0.###}, {candidate.candidate.Y:0.###}, {candidate.candidate.Z:0.###}");
        }

        lines.Add("");
        lines.Add("Move the freecam a visible distance, wait for the live coordinates to update, then click COMPARE.");
        return string.Join(Environment.NewLine, lines);
    }

    private string FormatCameraStructComparisonReport(
        float currentX,
        float currentY,
        float currentZ,
        IReadOnlyList<H3CameraStructCandidate> candidates)
    {
        var results = new List<(H3CameraStructCandidate Candidate, float X, float Y, float Z, double Error, double MovementError)>();
        foreach (var candidate in candidates)
        {
            if (!_session.TryReadFloatAbsolute(candidate.XAddress, out var x) ||
                !_session.TryReadFloatAbsolute(candidate.YAddress, out var y) ||
                !_session.TryReadFloatAbsolute(candidate.ZAddress, out var z) ||
                !IsLikelyCameraComponent(x) || !IsLikelyCameraComponent(y) || !IsLikelyCameraComponent(z))
                continue;

            var error = Math.Sqrt(
                Math.Pow(x - currentX, 2) +
                Math.Pow(y - currentY, 2) +
                Math.Pow(z - currentZ, 2));
            var movementError = Math.Sqrt(
                Math.Pow((x - candidate.X) - (currentX - candidates[0].X), 2) +
                Math.Pow((y - candidate.Y) - (currentY - candidates[0].Y), 2) +
                Math.Pow((z - candidate.Z) - (currentZ - candidates[0].Z), 2));
            results.Add((candidate, x, y, z, error, movementError));
        }

        var lines = new List<string>
        {
            $"CAMERA STRUCT COMPARISON @ {DateTime.Now:HH:mm:ss}",
            "",
            $"Current live XYZ: {currentX:0.###}, {currentY:0.###}, {currentZ:0.###}",
            $"Baseline candidates checked: {candidates.Count}",
            $"Readable candidates: {results.Count}",
            "",
            "Ranked by current-coordinate error (lower is better)"
        };

        foreach (var result in results.OrderBy(r => r.Error).ThenBy(r => r.MovementError).Take(24).Select((result, index) => (result, index)))
        {
            lines.Add($"{result.index + 1:00}. XYZ @ 0x{result.result.Candidate.XAddress:X}/0x{result.result.Candidate.YAddress:X}/0x{result.result.Candidate.ZAddress:X}  error {result.result.Error:0.####}  movement error {result.result.MovementError:0.####}");
            lines.Add($"    baseline {result.result.Candidate.X:0.###}, {result.result.Candidate.Y:0.###}, {result.result.Candidate.Z:0.###} -> now {result.result.X:0.###}, {result.result.Y:0.###}, {result.result.Z:0.###}");
        }

        var exact = results
            .Where(result => result.Error < 0.01)
            .OrderBy(result => result.MovementError)
            .ToList();
        var ownershipTargets = exact
            .SelectMany(result => new[] { result.Candidate.BaseAddress, result.Candidate.XAddress })
            .Distinct()
            .ToArray();
        var owners = ownershipTargets.Length == 0
            ? []
            : _session.ScanReadablePointersToAny(ownershipTargets, maxBytesToScan: 2L * 1024 * 1024 * 1024, maxMatches: 512);

        _ownedCameraCandidate = exact
            .Select(result => new
            {
                result.Candidate,
                OwnerCount = owners.Count(owner =>
                    owner.TargetAddress == result.Candidate.BaseAddress ||
                    owner.TargetAddress == result.Candidate.XAddress)
            })
            .Where(result => result.OwnerCount > 0)
            .OrderByDescending(result => result.OwnerCount)
            .Select(result => result.Candidate)
            .FirstOrDefault();

        lines.Add("");
        lines.Add($"Pointer ownership scan: {exact.Count} exact tracker(s), {owners.Count} pointer reference(s)");
        lines.Add(_ownedCameraCandidate is null
            ? "Selected owned probe: none"
            : $"Selected owned probe: XYZ @ 0x{_ownedCameraCandidate.XAddress:X}/0x{_ownedCameraCandidate.YAddress:X}/0x{_ownedCameraCandidate.ZAddress:X}");
        foreach (var result in exact.Select((result, index) => (result, index)))
        {
            var baseOwners = owners.Where(owner => owner.TargetAddress == result.result.Candidate.BaseAddress).ToList();
            var xyzOwners = owners.Where(owner => owner.TargetAddress == result.result.Candidate.XAddress).ToList();
            lines.Add($"{result.index + 1:00}. base 0x{result.result.Candidate.BaseAddress:X}: {baseOwners.Count} owner(s); XYZ 0x{result.result.Candidate.XAddress:X}: {xyzOwners.Count} owner(s)");
            foreach (var owner in baseOwners.Concat(xyzOwners).Take(8))
                lines.Add($"    pointer @ 0x{owner.Address:X} -> 0x{owner.TargetAddress:X}");
        }

        lines.Add("");
        lines.Add("This comparison and ownership scan are read-only. No candidate was written.");
        return string.Join(Environment.NewLine, lines);
    }

    private static IEnumerable<string> SummarizeCameraMovementTriples(
        H3DiscoveryProbe probe,
        byte[] before,
        byte[] after,
        IReadOnlyCollection<H3DiscoveryByteChange> changes)
    {
        var changedOffsets = changes.Select(c => c.Offset).ToHashSet();
        var length = Math.Min(before.Length, after.Length);
        var seen = new HashSet<int>();

        for (int offset = 0; offset + 12 <= length; offset += 4)
        {
            if (!Enumerable.Range(offset, 12).Any(changedOffsets.Contains))
                continue;

            var bx = BitConverter.ToSingle(before, offset);
            var by = BitConverter.ToSingle(before, offset + 4);
            var bz = BitConverter.ToSingle(before, offset + 8);
            var ax = BitConverter.ToSingle(after, offset);
            var ay = BitConverter.ToSingle(after, offset + 4);
            var az = BitConverter.ToSingle(after, offset + 8);
            if (!IsLikelyCameraComponent(bx) || !IsLikelyCameraComponent(by) || !IsLikelyCameraComponent(bz) ||
                !IsLikelyCameraComponent(ax) || !IsLikelyCameraComponent(ay) || !IsLikelyCameraComponent(az))
            {
                continue;
            }

            var distance = Math.Sqrt(
                ((ax - bx) * (ax - bx)) +
                ((ay - by) * (ay - by)) +
                ((az - bz) * (az - bz)));
            if (distance < 0.05 || distance > 10000)
                continue;

            if (!seen.Add(offset))
                continue;

            yield return $"{probe.FormatAddress(offset)} XYZ {bx:0.###},{by:0.###},{bz:0.###} -> {ax:0.###},{ay:0.###},{az:0.###}  dist {distance:0.###}";
        }
    }

    private static bool IsLikelyCameraComponent(float value)
        => !float.IsNaN(value) &&
           !float.IsInfinity(value) &&
           Math.Abs(value) < 10000f;

    private static List<string> SummarizeDiscoveryCandidates(
        H3DiscoveryProbe probe,
        byte[] before,
        byte[] after,
        IReadOnlyCollection<H3DiscoveryByteChange> changes)
    {
        var changedOffsets = changes.Select(c => c.Offset).ToHashSet();
        var summaries = new List<string>();

        for (int offset = 0; offset + 4 <= Math.Min(before.Length, after.Length); offset += 4)
        {
            if (!Enumerable.Range(offset, 4).Any(changedOffsets.Contains))
                continue;

            var beforeFloat = BitConverter.ToSingle(before, offset);
            var afterFloat = BitConverter.ToSingle(after, offset);
            if (IsUsefulDiscoveryFloat(beforeFloat) &&
                IsUsefulDiscoveryFloat(afterFloat) &&
                Math.Abs(beforeFloat - afterFloat) > 0.01f)
            {
                summaries.Add($"{probe.FormatAddress(offset)} f32 {beforeFloat:0.###}->{afterFloat:0.###}");
                continue;
            }

            var beforeInt = BitConverter.ToInt32(before, offset);
            var afterInt = BitConverter.ToInt32(after, offset);
            if (beforeInt != afterInt &&
                Math.Abs((long)beforeInt) < 1_000_000 &&
                Math.Abs((long)afterInt) < 1_000_000)
            {
                summaries.Add($"{probe.FormatAddress(offset)} i32 {beforeInt}->{afterInt}");
            }
        }

        return summaries;
    }

    private static string FormatDiscoveryChangeDetail(byte[] before, byte[] after, int offset)
    {
        var parts = new List<string>();

        if (offset + 2 <= before.Length && offset + 2 <= after.Length)
        {
            var before16 = BitConverter.ToUInt16(before, offset);
            var after16 = BitConverter.ToUInt16(after, offset);
            if (before16 != after16)
                parts.Add($"u16 {before16}->{after16}");
        }

        if (offset + 4 <= before.Length && offset + 4 <= after.Length)
        {
            var before32 = BitConverter.ToInt32(before, offset);
            var after32 = BitConverter.ToInt32(after, offset);
            if (before32 != after32 && Math.Abs((long)before32) < 1_000_000 && Math.Abs((long)after32) < 1_000_000)
                parts.Add($"i32 {before32}->{after32}");

            var beforeFloat = BitConverter.ToSingle(before, offset);
            var afterFloat = BitConverter.ToSingle(after, offset);
            if (IsUsefulDiscoveryFloat(beforeFloat) && IsUsefulDiscoveryFloat(afterFloat) && Math.Abs(beforeFloat - afterFloat) > 0.0001f)
                parts.Add($"f32 {beforeFloat:0.###}->{afterFloat:0.###}");
        }

        return parts.Count == 0 ? "" : $"  [{string.Join(", ", parts)}]";
    }

    private static bool IsUsefulDiscoveryFloat(float value)
        => !float.IsNaN(value) &&
           !float.IsInfinity(value) &&
           Math.Abs(value) is > 0.00001f and < 100000f;

    private static string SaveDiscoveryReport(string presetId, string kind, string text)
    {
        var root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HaloMCCToolbox",
            "H3Discovery");
        Directory.CreateDirectory(root);

        var safePreset = string.Concat(presetId.Select(ch => char.IsLetterOrDigit(ch) || ch == '-' ? ch : '_'));
        var path = Path.Combine(root, $"{DateTime.Now:yyyyMMdd-HHmmss}-{safePreset}-{kind}.txt");
        File.WriteAllText(path, text);
        return path;
    }

    private static string FormatVitalScanBaseline(H3FloatScanResult scan)
    {
        var lines = new List<string>
        {
            $"VITAL MEMORY SCAN baseline @ {DateTime.Now:HH:mm:ss}",
            $"Scanned: {scan.ScannedBytes / 1024 / 1024:N0} MB across {scan.ReadableRegions:N0}/{scan.ScannedRegions:N0} writable committed region(s)",
            $"Health-like float samples: {scan.Samples.Count:N0}",
            "Take one controlled damage event, then click VDIFF.",
        };

        return string.Join(Environment.NewLine, lines);
    }

    private static string FormatVitalScanDiff(Dictionary<long, float> baseline, DateTime baselineAt, H3FloatScanResult scan)
    {
        var after = scan.Samples
            .GroupBy(s => s.Address)
            .ToDictionary(g => g.Key, g => g.Last().Value);

        var drops = baseline
            .Select(pair => after.TryGetValue(pair.Key, out var current)
                ? new H3VitalScanDrop(pair.Key, pair.Value, current, pair.Value - current)
                : null)
            .Where(drop => drop is not null)
            .Select(drop => drop!)
            .Where(drop =>
                drop.Before is >= 0.05f and <= 1.25f &&
                drop.After is >= 0.0f and <= 1.25f &&
                drop.Drop >= 0.05f)
            .OrderByDescending(drop => drop.Drop)
            .ThenBy(drop => drop.After)
            .Take(120)
            .ToList();

        var lines = new List<string>
        {
            "VITAL MEMORY SCAN diff",
            $"BASE {baselineAt:HH:mm:ss} -> DIFF {DateTime.Now:HH:mm:ss}",
            $"Scanned: {scan.ScannedBytes / 1024 / 1024:N0} MB across {scan.ReadableRegions:N0}/{scan.ScannedRegions:N0} writable committed region(s)",
            $"Baseline samples: {baseline.Count:N0}; current samples: {after.Count:N0}",
            $"Drop candidates: {drops.Count:N0}",
            ""
        };

        foreach (var drop in drops)
            lines.Add($"0x{drop.Address:X16}: {drop.Before:0.###} -> {drop.After:0.###}  drop {drop.Drop:0.###}");

        if (drops.Count == 0)
        {
            lines.Add("No health-like drops found. Try VBASE while fully healthy, immediately take damage, then VDIFF before regen settles.");
        }
        else
        {
            lines.Add("");
            lines.Add("Best next test: repeat the same BASE/VDIFF sequence. Addresses that drop both times are worth promoting to candidate writes.");
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static IEnumerable<H3DiscoveryByteChange> FindChangedBytes(byte[] before, byte[] after)
    {
        var length = Math.Min(before.Length, after.Length);
        for (int i = 0; i < length; i++)
        {
            if (before[i] != after[i])
                yield return new H3DiscoveryByteChange(i, before[i], after[i]);
        }
    }

    private bool EnsureReadyForCoordinates()
    {
        if (_writesAllowed)
            return true;

        TxtFooter.Text = "Camera coordinates are gated until MCC is attached with EAC not detected and halo3.dll loaded.";
        return false;
    }

    private void RememberOriginalBytes(H3MemorySession session, H3Address address, int length)
    {
        var key = $"{address}|{length}";
        if (_originalBytes.ContainsKey(key))
            return;

        if (session.TryReadBytes(address, length, out var original))
            _originalBytes[key] = original;
    }

    private void RestoreAll()
    {
        StopSwivelCam("Swivel Cam stopped for restore.");
        var hookRestored = _session.UninstallCameraCaptureHook();
        if (!_session.IsAttached || _originalBytes.Count == 0)
        {
            TxtPatchDetail.Text = hookRestored ? "Camera hook restored. No other patches applied by this Toolbox session." : "Camera hook restore failed.";
            RefreshCoordinateButtons();
            return;
        }

        int restored = 0;
        foreach (var (key, bytes) in _originalBytes.ToArray())
        {
            var separator = key.LastIndexOf('|');
            if (separator <= 0)
                continue;

            if (!TryParseAddressKey(key[..separator], out var address))
                continue;

            if (_session.TryWriteBytes(address, bytes))
                restored++;
        }

        foreach (var mod in _mods.OfType<H3FloatHoldMod>())
            mod.DisableWithoutWriting();

        _originalBytes.Clear();
        TxtFooter.Text = $"Restore complete: {restored} location(s) restored.";
        RefreshRows();
    }

    private static bool TryParseAddressKey(string value, out H3Address address)
    {
        address = new H3Address("halo3.dll", 0);
        var parts = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var moduleParts = parts[0].Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (moduleParts.Length != 2 || !TryParseHex(moduleParts[1], out var baseOffset))
            return false;

        var offsets = new List<long>();
        for (int i = 1; i < parts.Length; i++)
        {
            if (!TryParseHex(parts[i], out var offset))
                return false;
            offsets.Add(offset);
        }

        address = new H3Address(moduleParts[0], baseOffset, [.. offsets]);
        return true;
    }

    private static bool TryParseHex(string value, out long result)
    {
        result = 0;
        var trimmed = value.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? value[2..] : value;
        return long.TryParse(trimmed, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result);
    }

    private static List<H3RuntimeMod> BuildModCatalog()
    {
        return
        [
            new H3BitToggleMod("acrophobia", "ACROPHOBIA", H3ModCategory.Skull,
                "Toggles the known Acrophobia skull flag in the live Halo 3 runtime state.",
                H3KnownAddresses.Acrophobia, 0x20),
            new H3BitToggleMod("bandana", "BANDANA", H3ModCategory.Skull,
                "Toggles the known Bandana skull flag in the live Halo 3 runtime state.",
                H3KnownAddresses.Bandana, 0x80),

            new H3BytePatchMod("freecam", "FREECAM", H3ModCategory.Camera,
                "Enables the known Halo 3 freecam patch set. Incompatible with Third Person.",
                H3KnownAddresses.FreecamState, 0xE9,
                [
                    new(H3KnownAddresses.DisableCameraControl, Bytes("90 90 90 90 90 90"), Bytes("0F 85 14 02 00 00")),
                    new(H3KnownAddresses.ThirdPersonBranch, Bytes("90 90"), Bytes("74 0E")),
                    new(H3KnownAddresses.FreecamState, Bytes("E9 3D 03 00 00 90"), Bytes("0F 84 A4 03 00 00")),
                    new(H3KnownAddresses.HudBlindPatch, Bytes("41 8B C9 90"), Bytes("41 0F 45 C9")),
                ],
                guard: session => !IsByte(session, H3KnownAddresses.ThirdPersonBranch, 0x90)),
            new H3BytePatchMod("freeze-camera", "FREEZE CAM", H3ModCategory.Camera,
                "Freezes the camera while Freecam is active.",
                new H3Address("halo3.dll", 0x131B9B), 0xFD,
                [
                    new(H3KnownAddresses.FreecamState, Bytes("E9 FD 01 00 00 90"), Bytes("E9 3D 03 00 00 90")),
                ],
                guard: session => IsByte(session, H3KnownAddresses.FreecamState, 0xE9)),
            new H3BytePatchMod("freeze-player", "FREEZE PLAYER", H3ModCategory.Camera,
                "Freezes player movement with the known runtime branch patch.",
                H3KnownAddresses.FreezePlayer, 0xE9,
                [
                    new(H3KnownAddresses.FreezePlayer, Bytes("E9 C7 01 00 00 90"), Bytes("90 90 90 90 90 90")),
                ]),
            new H3BytePatchMod("coordinates", "COORD CONTROL", H3ModCategory.Camera,
                "Allows direct camera coordinate read/write in the runtime camera structure.",
                H3KnownAddresses.CoordinatesA, 0x90,
                [
                    new(H3KnownAddresses.CoordinatesA, Bytes("90 90 90 90 90 90"), Bytes("0F 84 9F 01 00 00")),
                    new(H3KnownAddresses.CoordinatesB, Bytes("90 90 90 90 90 90"), Bytes("0F 84 92 01 00 00")),
                    new(H3KnownAddresses.CoordinatesC, Bytes("90 90"), Bytes("74 0A")),
                    new(H3KnownAddresses.CoordinatesD, Bytes("90 90"), Bytes("74 02")),
                ]),

            new H3BytePatchMod("third-person", "THIRD PERSON", H3ModCategory.Gameplay,
                "Toggles the known third-person branch patch. Incompatible with Freecam.",
                H3KnownAddresses.ThirdPersonBranch, 0x90,
                [
                    new(H3KnownAddresses.ThirdPersonBranch, Bytes("90 90"), Bytes("74 0E")),
                ],
                guard: session => !IsByte(session, H3KnownAddresses.FreecamState, 0xE9)),
            new H3BytePatchMod("barriers", "BARRIERS", H3ModCategory.Gameplay,
                "Disables known Halo 3 soft ceiling, kill trigger, and safe-zone barriers.",
                H3KnownAddresses.BarrierSoftCeiling, 0xEB,
                [
                    new(H3KnownAddresses.BarrierSoftCeiling, Bytes("EB 72"), Bytes("7E 72")),
                    new(H3KnownAddresses.BarrierKillTrigger, Bytes("EB 65"), Bytes("7E 65")),
                    new(H3KnownAddresses.BarrierSafeZone, Bytes("EB 6D"), Bytes("7E 6D")),
                ]),
            new H3BytePatchMod("pause-game", "PAUSE GAME", H3ModCategory.Gameplay,
                "Toggles the known runtime pause flag.",
                H3KnownAddresses.PauseGame, 0x18,
                [
                    new(H3KnownAddresses.PauseGame, Bytes("18"), Bytes("10")),
                ]),
            new H3BytePatchMod("thirty-tick", "30 TICK", H3ModCategory.Gameplay,
                "Toggles the known tick-rate value pair.",
                H3KnownAddresses.ThirtyTick, 0x1E,
                [
                    new(H3KnownAddresses.ThirtyTick, Bytes("1E 00 00 00 89 88 08 3D"), Bytes("3C 00 00 00 89 88 88 3C")),
                ]),
            new H3BytePatchMod("team-colors", "TEAM COLORS", H3ModCategory.Gameplay,
                "Disables team colors with the known branch-byte patch.",
                H3KnownAddresses.TeamColors, 0xEB,
                [
                    new(H3KnownAddresses.TeamColors, Bytes("EB"), Bytes("74")),
                ]),
        ];
    }

    private static List<H3LiveProbeTarget> BuildLiveProbeTargets()
        =>
        [
            new("candidate-shield-vitality", "Candidate shield/vitality", H3KnownAddresses.CandidateShieldVitality),
            new("movement-a", "Movement/state A", H3KnownAddresses.CandidateMovementA),
            new("movement-b", "Movement/state B", H3KnownAddresses.CandidateMovementB),
            new("repeat-a", "Repeated state A", H3KnownAddresses.CandidateRepeatedA),
            new("repeat-b", "Repeated state B", H3KnownAddresses.CandidateRepeatedB),
            new("repeat-c", "Repeated state C", H3KnownAddresses.CandidateRepeatedC),
            new("camera-x", "Camera X", H3KnownAddresses.CameraX),
            new("camera-y", "Camera Y", H3KnownAddresses.CameraY),
            new("camera-z", "Camera Z", H3KnownAddresses.CameraZ),
        ];

    private static bool IsByte(H3MemorySession session, H3Address address, byte expected)
        => session.TryReadByte(address, out var value) && value == expected;

    private static byte[] Bytes(string value)
        => value.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => byte.Parse(part.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? part[2..] : part, NumberStyles.HexNumber, CultureInfo.InvariantCulture))
            .ToArray();

    private static List<H3DiscoveryPreset> BuildDiscoveryPresets()
    {
        var trainerPack = new H3DiscoveryPreset(
            "trainer-pack",
            "TRAINER TARGET PACK",
            [
                new("state-core", "Game-state core window", new H3Address("halo3.dll", 0x2030288, 0x0), 0x8000),
                new("timing-flags", "Timing / pause / tick window", new H3Address("halo3.dll", 0x2030288, 0x10000), 0x1400),
                new("skull-flags", "Skull / session flag window", new H3Address("halo3.dll", 0x2030288, 0xFDC0), 0x90),
                new("camera-state", "Camera coordinate window", new H3Address("halo3.dll", 0x2030288, 0x2BAF00), 0x260),
                new("sim-functions", "Simulation function variable IDs", new H3Address("halo3.dll", 0x781B00), 0x400),
            ]);

        return
        [
            trainerPack,
            new H3DiscoveryPreset(
                "vitality-damage",
                "VITALITY / DAMAGE",
                [
                    new("state-core", "Game-state core window", new H3Address("halo3.dll", 0x2030288, 0x0), 0x10000),
                    new("damage-window", "Known damage / vitality function IDs", new H3Address("halo3.dll", 0x781B70), 0x90),
                    new("damage-code", "Damage-adjacent code/data window", new H3Address("halo3.dll", 0x676300), 0x160),
                ]),
            new H3DiscoveryPreset(
                "ammo-weapons",
                "AMMO / WEAPONS",
                [
                    new("state-core", "Game-state core window", new H3Address("halo3.dll", 0x2030288, 0x0), 0x10000),
                    new("weapon-functions", "Weapon function IDs", new H3Address("halo3.dll", 0x781CB0), 0x180),
                    new("weapon-code", "Weapon/camera patch neighborhood", new H3Address("halo3.dll", 0x131A80), 0x520),
                ]),
            new H3DiscoveryPreset(
                "movement-jump",
                "MOVEMENT / JUMP",
                [
                    new("state-core", "Game-state core window", new H3Address("halo3.dll", 0x2030288, 0x0), 0x10000),
                    new("movement-functions", "Movement function IDs", new H3Address("halo3.dll", 0x781E40), 0xC0),
                    new("freeze-player-code", "Movement patch neighborhood", new H3Address("halo3.dll", 0xE4B40), 0x260),
                ]),
            new H3DiscoveryPreset(
                "wide-state",
                "WIDE STATE SCAN",
                [
                    new("state-wide-a", "Game-state 0x00000-0x1FFFF", new H3Address("halo3.dll", 0x2030288, 0x0), 0x20000),
                    new("state-wide-b", "Game-state 0x20000-0x3FFFF", new H3Address("halo3.dll", 0x2030288, 0x20000), 0x20000),
                ]),
            new H3DiscoveryPreset(
                "skull-flags",
                "SKULL / SESSION FLAGS",
                [
                    new("skull-flags", "Skull / session flag window", new H3Address("halo3.dll", 0x2030288, 0xFDC0), 0x90),
                ]),
            new H3DiscoveryPreset(
                "camera-state",
                "CAMERA STATE",
                [
                    new("camera-state", "Camera coordinate window", new H3Address("halo3.dll", 0x2030288, 0x2BAF00), 0x260),
                    new("camera-patches", "Camera patch code window", new H3Address("halo3.dll", 0x131080), 0x360),
                ]),
            new H3DiscoveryPreset(
                "sim-functions",
                "SIM FUNCTION TABLE",
                [
                    new("sim-functions", "Simulation function variable IDs", new H3Address("halo3.dll", 0x781B00), 0x400),
                ]),
        ];
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        StopSwivelCam("Swivel Cam stopped because the H3 Mods tab was disposed.");
        StopDollyPlayback("Dolly playback stopped because the H3 Mods tab was disposed.");
        _attachTimer.Stop();
        _coordDisplayTimer.Stop();
        _dollyRecordTimer.Stop();
        CloseDollyTrackOverlay();
        RestoreAll();
        _session.Dispose();
    }
}

internal enum H3ModCategory
{
    Skull,
    Camera,
    Gameplay
}

public sealed class H3ModRow : INotifyPropertyChanged
{
    private bool _isActive;
    private bool _isReadable;
    private bool _canToggle;
    private string _stateText = "WAIT";
    private Brush _stateBrush = Brushes.Gray;

    internal H3ModRow(H3RuntimeMod definition)
    {
        Definition = definition;
        Name = definition.Name;
        Detail = definition.Detail;
    }

    internal H3RuntimeMod Definition { get; }
    public string Name { get; }
    public string Detail { get; }

    public bool IsActive
    {
        get => _isActive;
        set => SetField(ref _isActive, value);
    }

    public bool IsReadable
    {
        get => _isReadable;
        set => SetField(ref _isReadable, value);
    }

    public bool CanToggle
    {
        get => _canToggle;
        set => SetField(ref _canToggle, value);
    }

    public string StateText
    {
        get => _stateText;
        set => SetField(ref _stateText, value);
    }

    public Brush StateBrush
    {
        get => _stateBrush;
        set => SetField(ref _stateBrush, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return;

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

internal abstract class H3RuntimeMod
{
    protected H3RuntimeMod(string id, string name, H3ModCategory category, string detail, bool isMapped)
    {
        Id = id;
        Name = name;
        Category = category;
        Detail = detail;
        IsMapped = isMapped;
    }

    public string Id { get; }
    public string Name { get; }
    public H3ModCategory Category { get; }
    public string Detail { get; }
    public bool IsMapped { get; }
    public H3ModRow? Row { get; set; }

    public abstract bool TryReadActive(H3MemorySession session, out bool active);
    public abstract bool Toggle(H3MemorySession session, Action<H3MemorySession, H3Address, int> rememberOriginal, out string message);

    public static H3RuntimeMod Locked(string id, string name, H3ModCategory category, string detail)
        => new H3LockedMod(id, name, category, detail);
}

internal sealed class H3LockedMod : H3RuntimeMod
{
    public H3LockedMod(string id, string name, H3ModCategory category, string detail)
        : base(id, name, category, detail, isMapped: false)
    {
    }

    public override bool TryReadActive(H3MemorySession session, out bool active)
    {
        active = false;
        return false;
    }

    public override bool Toggle(H3MemorySession session, Action<H3MemorySession, H3Address, int> rememberOriginal, out string message)
    {
        message = $"{Name} is locked until its runtime location is verified.";
        return false;
    }
}

internal sealed class H3BitToggleMod : H3RuntimeMod
{
    private readonly H3Address _address;
    private readonly byte _mask;

    public H3BitToggleMod(string id, string name, H3ModCategory category, string detail, H3Address address, byte mask)
        : base(id, name, category, detail, isMapped: true)
    {
        _address = address;
        _mask = mask;
    }

    public override bool TryReadActive(H3MemorySession session, out bool active)
    {
        active = false;
        if (!session.TryReadByte(_address, out var value))
            return false;

        active = (value & _mask) == _mask;
        return true;
    }

    public override bool Toggle(H3MemorySession session, Action<H3MemorySession, H3Address, int> rememberOriginal, out string message)
    {
        if (!session.TryReadByte(_address, out var value))
        {
            message = $"{Name} failed: could not read current skull byte.";
            return false;
        }

        rememberOriginal(session, _address, 1);
        var next = (byte)(value ^ _mask);
        if (!session.TryWriteByte(_address, next))
        {
            message = $"{Name} failed: memory write did not complete.";
            return false;
        }

        message = $"{Name}: {(((next & _mask) == _mask) ? "ON" : "OFF")}.";
        return true;
    }
}

internal sealed class H3BytePatchMod : H3RuntimeMod
{
    private readonly H3Address _stateAddress;
    private readonly byte _activeMarker;
    private readonly H3PatchWrite[] _writes;
    private readonly Func<H3MemorySession, bool>? _guard;

    public H3BytePatchMod(
        string id,
        string name,
        H3ModCategory category,
        string detail,
        H3Address stateAddress,
        byte activeMarker,
        H3PatchWrite[] writes,
        Func<H3MemorySession, bool>? guard = null)
        : base(id, name, category, detail, isMapped: true)
    {
        _stateAddress = stateAddress;
        _activeMarker = activeMarker;
        _writes = writes;
        _guard = guard;
    }

    public override bool TryReadActive(H3MemorySession session, out bool active)
    {
        active = false;
        if (!session.TryReadByte(_stateAddress, out var value))
            return false;

        active = value == _activeMarker;
        return true;
    }

    public override bool Toggle(H3MemorySession session, Action<H3MemorySession, H3Address, int> rememberOriginal, out string message)
    {
        if (!TryReadActive(session, out var active))
        {
            message = $"{Name} failed: could not read patch state.";
            return false;
        }

        if (!active && _guard is not null && !_guard(session))
        {
            message = $"{Name} is blocked by another active patch or missing prerequisite.";
            return false;
        }

        var targetBytes = active
            ? _writes.Select(w => (w.Address, Bytes: w.DisabledBytes)).ToArray()
            : _writes.Select(w => (w.Address, Bytes: w.EnabledBytes)).ToArray();

        foreach (var (address, bytes) in targetBytes)
            rememberOriginal(session, address, bytes.Length);

        foreach (var (address, bytes) in targetBytes)
        {
            if (!session.TryWriteBytes(address, bytes))
            {
                message = $"{Name} failed: write did not complete.";
                return false;
            }
        }

        message = $"{Name}: {(active ? "OFF" : "ON")}.";
        return true;
    }
}

internal sealed class H3FloatHoldMod : H3RuntimeMod
{
    private readonly H3Address _address;
    private readonly float _targetValue;
    private readonly bool _useCameraXOnEnable;
    private byte[]? _restoreBytes;
    private float _activeValue;
    private bool _enabled;

    public H3FloatHoldMod(
        string id,
        string name,
        H3ModCategory category,
        string detail,
        H3Address address,
        float targetValue,
        bool useCameraXOnEnable = false)
        : base(id, name, category, detail, isMapped: true)
    {
        _address = address;
        _targetValue = targetValue;
        _activeValue = targetValue;
        _useCameraXOnEnable = useCameraXOnEnable;
    }

    public override bool TryReadActive(H3MemorySession session, out bool active)
    {
        active = false;
        if (!session.TryReadFloat(_address, out _))
            return false;

        active = _enabled;
        return true;
    }

    public void ApplyHold(H3MemorySession session, Action<H3MemorySession, H3Address, int> rememberOriginal)
    {
        if (!_enabled)
            return;

        rememberOriginal(session, _address, 4);
        session.TryWriteFloat(_address, _activeValue);
    }

    public void DisableWithoutWriting()
        => _enabled = false;

    public override bool Toggle(H3MemorySession session, Action<H3MemorySession, H3Address, int> rememberOriginal, out string message)
    {
        if (!session.TryReadBytes(_address, 4, out var currentBytes))
        {
            message = $"{Name} failed: could not read candidate value.";
            return false;
        }

        if (_enabled)
        {
            _enabled = false;
            var restored = _restoreBytes is not null && session.TryWriteBytes(_address, _restoreBytes);
            message = restored
                ? $"{Name}: OFF. Candidate value restored."
                : $"{Name}: OFF. Candidate restore failed or had no cached value.";
            return true;
        }

        _restoreBytes = currentBytes;
        rememberOriginal(session, _address, 4);

        _activeValue = _targetValue;
        if (_useCameraXOnEnable && session.TryReadFloat(H3KnownAddresses.CameraX, out var cameraX))
            _activeValue = cameraX;

        if (!session.TryWriteFloat(_address, _activeValue))
        {
            message = $"{Name} failed: candidate write did not complete.";
            return false;
        }

        _enabled = true;
        message = $"{Name}: ON. Holding {_activeValue:0.###} at {_address}.";
        return true;
    }
}

internal sealed record H3PatchWrite(H3Address Address, byte[] EnabledBytes, byte[] DisabledBytes);

public sealed record H3DollyMarker(int Index, double TimeSeconds, float X, float Y, float Z, float A, float B, float C)
{
    public string TimeText => $"{Index:00}  {TimeSeconds:0.00}s";
    public string CoordinateText => $"X {X:0.###}   Y {Y:0.###}   Z {Z:0.###}   R {A:0.##}/{B:0.##}/{C:0.##}";
}

internal readonly record struct H3CameraAddressSet(long XAddress, long YAddress, long ZAddress);

internal sealed record H3CameraScanCandidate(
    H3CameraAddressSet Addresses,
    float CurrentX,
    float CurrentY,
    float CurrentZ,
    double Score);

internal sealed record H3CameraStructCandidate(
    long BaseAddress,
    long XAddress,
    long YAddress,
    long ZAddress,
    float X,
    float Y,
    float Z,
    double Score);

internal sealed record H3DiscoveryPreset(string Id, string Name, H3DiscoveryProbe[] Probes);

internal sealed record H3DiscoveryProbe(string Key, string Name, H3Address Address, int Length)
{
    public string FormatAddress(int byteOffset)
    {
        if (Address.PointerOffsets.Length == 0)
            return new H3Address(Address.ModuleName, Address.BaseOffset + byteOffset).ToString();

        var offsets = Address.PointerOffsets.ToArray();
        offsets[^1] += byteOffset;
        return new H3Address(Address.ModuleName, Address.BaseOffset, offsets).ToString();
    }
}

internal sealed record H3DiscoverySnapshot(H3DiscoveryPreset Preset, DateTime CapturedAt, List<H3DiscoveryCapture> Captures);

internal sealed record H3DiscoveryCapture(H3DiscoveryProbe Probe, bool Readable, byte[] Bytes);

internal sealed record H3DiscoveryByteChange(int Offset, byte Before, byte After);

internal sealed record H3LiveProbeTarget(string Id, string Name, H3Address Address);

internal sealed record H3VitalScanDrop(long Address, float Before, float After, float Drop);

internal static class H3KnownAddresses
{
    public static readonly H3Address Acrophobia = new("halo3.dll", 0x2030288, 0xFDE9);
    public static readonly H3Address Bandana = new("halo3.dll", 0x2030288, 0xFDE6);
    public static readonly H3Address PauseGame = new("halo3.dll", 0x2030288, 0x10D76);
    public static readonly H3Address ThirtyTick = new("halo3.dll", 0x2030288, 0x10408);
    public static readonly H3Address CameraMirrorX = new("halo3.dll", 0x2030288, 0x2BA3A8);
    public static readonly H3Address CameraMirrorY = new("halo3.dll", 0x2030288, 0x2BA3AC);
    public static readonly H3Address CameraMirrorZ = new("halo3.dll", 0x2030288, 0x2BA3B0);
    public static readonly H3Address CameraLiveX = new("halo3.dll", 0x2030288, 0x2BA4A8);
    public static readonly H3Address CameraLiveY = new("halo3.dll", 0x2030288, 0x2BA4AC);
    public static readonly H3Address CameraLiveZ = new("halo3.dll", 0x2030288, 0x2BA4B0);
    public static readonly H3Address CameraX = new("halo3.dll", 0x2030288, 0x2BB088);
    public static readonly H3Address CameraY = new("halo3.dll", 0x2030288, 0x2BB08C);
    public static readonly H3Address CameraZ = new("halo3.dll", 0x2030288, 0x2BB090);
    public static readonly H3Address PlayerX = new("halo3.dll", 0x2030288, 0x10720);
    public static readonly H3Address PlayerY = new("halo3.dll", 0x2030288, 0x10724);
    public static readonly H3Address PlayerZ = new("halo3.dll", 0x2030288, 0x10728);
    // Identified by the global normalized-vector probe. This vector lives in the
    // same respawn-safe state allocation as PlayerX/Y/Z and traverses a unit circle
    // exactly with the Spartan's in-place rotation.
    public static readonly H3Address PlayerFacingX = new("halo3.dll", 0x2030288, -0x16C20);
    public static readonly H3Address PlayerFacingY = new("halo3.dll", 0x2030288, -0x16C1C);
    public static readonly H3Address PlayerFacingZ = new("halo3.dll", 0x2030288, -0x16C18);

    public static readonly H3Address FreecamState = new("halo3.dll", 0x131B9A);
    public static readonly H3Address DisableCameraControl = new("halo3.dll", 0x211EFC);
    public static readonly H3Address ThirdPersonBranch = new("halo3.dll", 0x132872);
    public static readonly H3Address HudBlindPatch = new("halo3.dll", 0x2D249C);
    public static readonly H3Address FreezePlayer = new("halo3.dll", 0xE4B75);
    public static readonly H3Address BarrierSoftCeiling = new("halo3.dll", 0x1BB869);
    public static readonly H3Address BarrierKillTrigger = new("halo3.dll", 0x1B8E49);
    public static readonly H3Address BarrierSafeZone = new("halo3.dll", 0x1B8EC2);
    public static readonly H3Address CoordinatesA = new("halo3.dll", 0x131113);
    public static readonly H3Address CoordinatesB = new("halo3.dll", 0x131120);
    public static readonly H3Address CoordinatesC = new("halo3.dll", 0xF767E);
    public static readonly H3Address CoordinatesD = new("halo3.dll", 0xF7686);
    public static readonly H3Address TeamColors = new("halo3.dll", 0xD1C0);
    public static readonly H3Address CandidateShieldVitality = new("halo3.dll", 0x2030288, 0x10418);
    public static readonly H3Address CandidateMovementA = new("halo3.dll", 0x2030288, 0x10720);
    public static readonly H3Address CandidateMovementB = new("halo3.dll", 0x2030288, 0x10724);
    public static readonly H3Address CandidateRepeatedA = new("halo3.dll", 0x2030288, 0x10980);
    public static readonly H3Address CandidateRepeatedB = new("halo3.dll", 0x2030288, 0x10984);
    public static readonly H3Address CandidateRepeatedC = new("halo3.dll", 0x2030288, 0x10988);
}
