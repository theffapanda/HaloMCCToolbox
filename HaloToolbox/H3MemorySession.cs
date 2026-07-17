using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace HaloToolbox;

internal sealed class H3MemorySession : IDisposable
{
    private const int CameraHookAllocationSize = 0x1000;
    private const int CameraHookCodeOffset = 0x400;
    private const int CameraManualTarget0Offset = 0x10;
    private const int CameraManualTarget1Offset = 0x140;
    private const int CameraManualTargetSize = 24;
    private const int CameraRawTransformOffset = 0x170;
    private const int CameraSharedDataEnd = CameraRawTransformOffset + CameraManualTargetSize;
    private const int ProcessVmOperation = 0x0008;
    private const int ProcessVmRead = 0x0010;
    private const int ProcessVmWrite = 0x0020;
    private const int ProcessQueryInformation = 0x0400;
    private const uint Th32csSnapModule = 0x00000008;
    private const uint Th32csSnapModule32 = 0x00000010;
    private const uint MemCommit = 0x1000;
    private const uint MemReserve = 0x2000;
    private const uint MemRelease = 0x8000;
    private const uint PageNoAccess = 0x01;
    private const uint PageReadOnly = 0x02;
    private const uint PageReadWrite = 0x04;
    private const uint PageWriteCopy = 0x08;
    private const uint PageExecuteRead = 0x20;
    private const uint PageExecuteReadWrite = 0x40;
    private const uint PageExecuteWriteCopy = 0x80;
    private const uint PageGuard = 0x100;
    private static readonly IntPtr InvalidHandleValue = new(-1);

    private readonly Dictionary<string, H3ModuleInfo> _modules = new(StringComparer.OrdinalIgnoreCase);
    private Process? _process;
    private IntPtr _handle = IntPtr.Zero;
    private long _cameraHookSite;
    private long _cameraHookAllocation;
    private byte[]? _cameraHookOriginal;
    private IntPtr _cameraPlaybackThread = IntPtr.Zero;
    private readonly List<long> _cameraPlaybackBuffers = [];
    private long _cameraGameClockAddress;

    public int? ProcessId => _process?.HasExited == false ? _process.Id : null;
    public string ProcessName => _process?.HasExited == false ? _process.ProcessName : "";
    public bool IsAttached => _process?.HasExited == false && _handle != IntPtr.Zero;
    public bool HasHalo3Module => _modules.ContainsKey("halo3.dll");
    public bool IsCameraCaptureHookInstalled => _cameraHookSite != 0 && _cameraHookAllocation != 0;

    public static Process? FindMccProcess()
    {
        var names = new[] { "MCC-Win64-Shipping", "MCCWinStore-Win64-Shipping" };
        return names
            .SelectMany(name =>
            {
                try { return Process.GetProcessesByName(name); }
                catch { return []; }
            })
            .OrderByDescending(p =>
            {
                try { return p.StartTime; }
                catch { return DateTime.MinValue; }
            })
            .FirstOrDefault();
    }

    public static bool IsEasyAntiCheatLikelyLoaded(Process process)
    {
        try
        {
            foreach (ProcessModule module in process.Modules)
            {
                var moduleName = module.ModuleName;
                if (moduleName.Contains("EasyAntiCheat", StringComparison.OrdinalIgnoreCase) ||
                    moduleName.Contains("easyanticheat", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        catch
        {
            return false;
        }

        if (EnumerateLikelyRelatedEacProcesses(process).Any())
            return true;

        return false;
    }

    public bool Attach(Process process)
    {
        if (IsAttached && _process?.Id == process.Id)
        {
            RefreshModules();
            return true;
        }

        Detach();
        _process = process;
        _handle = OpenProcess(
            ProcessQueryInformation | ProcessVmOperation | ProcessVmRead | ProcessVmWrite,
            false,
            process.Id);

        if (_handle == IntPtr.Zero)
        {
            _process = null;
            return false;
        }

        RefreshModules();
        return true;
    }

    public void Detach()
    {
        UninstallCameraCaptureHook();
        _modules.Clear();

        if (_handle != IntPtr.Zero)
        {
            CloseHandle(_handle);
            _handle = IntPtr.Zero;
        }

        _process = null;
    }

    public bool InstallCameraCaptureHook(out string detail)
    {
        detail = "Camera hook installation failed.";
        if (!IsAttached || !HasHalo3Module)
        {
            detail = "Halo 3 is not attached.";
            return false;
        }
        if (IsCameraCaptureHookInstalled)
        {
            detail = "Camera capture hook is already installed.";
            return true;
        }

        byte[] pattern = [0xF3, 0x0F, 0, 0, 0xF3, 0x0F, 0x11, 0, 0, 0xF3, 0x0F, 0x59, 0, 0xF3, 0x0F, 0x59, 0, 0xF3, 0x0F, 0x58];
        var matches = FindModulePattern("halo3.dll", pattern, "xx??xxx??xxx?xxx?xxx", contextBytes: 0, maxMatches: 3);
        if (matches.Count != 1)
        {
            detail = $"Camera hook signature expected 1 match, found {matches.Count}.";
            return false;
        }

        // This verified boundary is active for Freecam and snapshots camera+18h before
        // Halo builds the view basis. Swivel Cam intentionally uses that proven path.
        var site = matches[0].AbsoluteAddress + 0x186;
        byte[] expected = [
            0x4C, 0x8D, 0x6E, 0x18,
            0x45, 0x0F, 0x28, 0xFD,
            0xF2, 0x41, 0x0F, 0x10, 0x45, 0x00
        ];
        if (!TryReadBytesAbsolute(site, expected.Length, out var original) || !original.SequenceEqual(expected))
        {
            // Recover a hook left behind only when a previous Toolbox process was
            // terminated before Dispose could restore this exact 14-byte detour.
            if (original.Length != expected.Length ||
                !original.Take(6).SequenceEqual(new byte[] { 0xFF, 0x25, 0, 0, 0, 0 }) ||
                !WriteProtectedCode(site, expected))
            {
                detail = $"Camera hook site bytes do not match the verified instruction window ({Convert.ToHexString(original)}).";
                return false;
            }

            original = expected.ToArray();
        }

        var allocation = VirtualAllocEx(_handle, IntPtr.Zero, (UIntPtr)CameraHookAllocationSize, MemCommit | MemReserve, PageExecuteReadWrite).ToInt64();
        if (allocation == 0)
        {
            detail = $"VirtualAllocEx failed ({Marshal.GetLastWin32Error()}).";
            return false;
        }

        byte[] clockPattern = [0x0F, 0x8F, 0, 0, 0, 0, 0xE8, 0, 0, 0, 0, 0x0F, 0x28, 0xC8];
        var clockMatches = FindModulePattern("halo3.dll", clockPattern, "xx????x????xxx", contextBytes: 0, maxMatches: 3);
        if (clockMatches.Count != 1 ||
            !TryReadBytesAbsolute(clockMatches[0].AbsoluteAddress + 0x1D, sizeof(int), out var clockDisplacementBytes))
        {
            VirtualFreeEx(_handle, new IntPtr(allocation), UIntPtr.Zero, MemRelease);
            detail = $"Game clock signature expected 1 match, found {clockMatches.Count}.";
            return false;
        }
        var gameClockAddress = clockMatches[0].AbsoluteAddress + BitConverter.ToInt32(clockDisplacementBytes, 0) + 0x21;

        var stubAddress = allocation + CameraHookCodeOffset;
        var stub = BuildCameraCaptureStub(stubAddress, allocation, site + expected.Length, gameClockAddress, expected);
        if (CameraSharedDataEnd > CameraHookCodeOffset ||
            CameraHookCodeOffset + stub.Length > CameraHookAllocationSize)
        {
            VirtualFreeEx(_handle, new IntPtr(allocation), UIntPtr.Zero, MemRelease);
            detail = $"Camera hook layout is unsafe (data end 0x{CameraSharedDataEnd:X}, code 0x{CameraHookCodeOffset:X}-0x{CameraHookCodeOffset + stub.Length - 1:X}).";
            return false;
        }

        var block = new byte[CameraHookCodeOffset + stub.Length];
        stub.CopyTo(block, CameraHookCodeOffset);
        if (!WriteBytes(new IntPtr(allocation), block))
        {
            VirtualFreeEx(_handle, new IntPtr(allocation), UIntPtr.Zero, MemRelease);
            detail = "Failed to write camera capture trampoline.";
            return false;
        }

        var patch = new List<byte> { 0xFF, 0x25, 0, 0, 0, 0 };
        patch.AddRange(BitConverter.GetBytes(stubAddress));
        while (patch.Count < expected.Length)
            patch.Add(0x90);
        if (!WriteProtectedCode(site, patch.ToArray()))
        {
            VirtualFreeEx(_handle, new IntPtr(allocation), UIntPtr.Zero, MemRelease);
            detail = "Failed to patch the verified camera hook site.";
            return false;
        }

        _cameraHookSite = site;
        _cameraHookAllocation = allocation;
        _cameraHookOriginal = original;
        _cameraGameClockAddress = gameClockAddress;
        detail = $"Camera capture hook installed at 0x{site:X}; capture slot 0x{allocation:X}.";
        return true;
    }

    public bool UninstallCameraCaptureHook()
    {
        if (!IsCameraCaptureHookInstalled)
            return true;

        WriteBytes(new IntPtr(_cameraHookAllocation + 0x09), [0]);
        WriteBytes(new IntPtr(_cameraHookAllocation + 0x08), [0]);
        if (_cameraPlaybackThread != IntPtr.Zero)
        {
            WaitForSingleObject(_cameraPlaybackThread, 1000);
            CloseHandle(_cameraPlaybackThread);
            _cameraPlaybackThread = IntPtr.Zero;
        }

        var restored = _cameraHookOriginal is not null && WriteProtectedCode(_cameraHookSite, _cameraHookOriginal);
        if (restored)
        {
            foreach (var buffer in _cameraPlaybackBuffers)
                VirtualFreeEx(_handle, new IntPtr(buffer), UIntPtr.Zero, MemRelease);
            _cameraPlaybackBuffers.Clear();
            VirtualFreeEx(_handle, new IntPtr(_cameraHookAllocation), UIntPtr.Zero, MemRelease);
        }
        _cameraHookSite = 0;
        _cameraHookAllocation = 0;
        _cameraHookOriginal = null;
        _cameraGameClockAddress = 0;
        return restored;
    }

    public bool TryReadCapturedCameraPosition(out float x, out float y, out float z)
    {
        x = y = z = 0;
        if (!IsCameraCaptureHookInstalled || !TryReadBytesAbsolute(_cameraHookAllocation, 8, out var pointerBytes))
            return false;
        var camera = BitConverter.ToInt64(pointerBytes, 0);
        return camera > 0x10000 &&
               TryReadFloatAbsolute(camera + 0x18, out x) &&
               TryReadFloatAbsolute(camera + 0x1C, out y) &&
               TryReadFloatAbsolute(camera + 0x20, out z);
    }

    public bool TryWriteCapturedCameraPosition(float x, float y, float z)
    {
        if (!TryReadCapturedCameraOrientation(out var a, out var b, out var c))
            return false;
        return TryWriteCapturedCameraTransform(x, y, z, a, b, c);
    }

    public bool TryReadCapturedCameraOrientation(out float a, out float b, out float c)
    {
        a = b = c = 0;
        if (!IsCameraCaptureHookInstalled || !TryReadBytesAbsolute(_cameraHookAllocation, 8, out var pointerBytes))
            return false;
        var camera = BitConverter.ToInt64(pointerBytes, 0);
        return camera > 0x10000 &&
               TryReadFloatAbsolute(camera + 0x24, out a) &&
               TryReadFloatAbsolute(camera + 0x28, out b) &&
               TryReadFloatAbsolute(camera + 0x2C, out c);
    }

    public bool TryReadRawCapturedCameraTransform(out float x, out float y, out float z, out float a, out float b, out float c)
    {
        x = y = z = a = b = c = 0;
        if (!IsCameraCaptureHookInstalled)
            return false;
        var address = _cameraHookAllocation + CameraRawTransformOffset;
        return TryReadFloatAbsolute(address, out x) &&
               TryReadFloatAbsolute(address + 4, out y) &&
               TryReadFloatAbsolute(address + 8, out z) &&
               TryReadFloatAbsolute(address + 12, out a) &&
               TryReadFloatAbsolute(address + 16, out b) &&
               TryReadFloatAbsolute(address + 20, out c);
    }

    public bool TryWriteCapturedCameraTransform(float x, float y, float z, float a, float b, float c)
    {
        if (!IsCameraCaptureHookInstalled)
            return false;
        var values = new[] { x, y, z, a, b, c };
        var bytes = new byte[24];
        Buffer.BlockCopy(values, 0, bytes, 0, bytes.Length);
        if (!TryReadBytesAbsolute(_cameraHookAllocation + 0x0A, 1, out var activeBytes))
            return false;
        var nextBuffer = activeBytes[0] == 0 ? (byte)1 : (byte)0;
        var targetOffset = nextBuffer == 0 ? CameraManualTarget0Offset : CameraManualTarget1Offset;
        return WriteBytes(new IntPtr(_cameraHookAllocation + targetOffset), bytes) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x0C), [0]) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x0A), [nextBuffer]) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x09), [1]);
    }

    public bool TryWriteCapturedCameraPositionOnly(float x, float y, float z)
    {
        if (!IsCameraCaptureHookInstalled)
            return false;
        var values = new[] { x, y, z };
        var bytes = new byte[12];
        Buffer.BlockCopy(values, 0, bytes, 0, bytes.Length);
        if (!TryReadBytesAbsolute(_cameraHookAllocation + 0x0A, 1, out var activeBytes))
            return false;
        var nextBuffer = activeBytes[0] == 0 ? (byte)1 : (byte)0;
        var targetOffset = nextBuffer == 0 ? CameraManualTarget0Offset : CameraManualTarget1Offset;
        return WriteBytes(new IntPtr(_cameraHookAllocation + targetOffset), bytes) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x0C), [1]) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x0A), [nextBuffer]) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x09), [1]);
    }

    public void DisableCameraTransformOverride()
    {
        if (IsCameraCaptureHookInstalled)
        {
            WriteBytes(new IntPtr(_cameraHookAllocation + 0x0B), [0]);
            WriteBytes(new IntPtr(_cameraHookAllocation + 0x09), [0]);
            WriteBytes(new IntPtr(_cameraHookAllocation + 0x0C), [0]);
        }
    }

    public bool TryStartCapturedCameraPlayback(IReadOnlyList<float> transforms, float speedMultiplier)
    {
        if (!IsCameraCaptureHookInstalled || transforms.Count < 6 || transforms.Count % 6 != 0 ||
            !float.IsFinite(speedMultiplier) || speedMultiplier < 0.1f || speedMultiplier > 2.0f)
            return false;

        var bytes = new byte[transforms.Count * sizeof(float)];
        for (var i = 0; i < transforms.Count; i++)
            BitConverter.GetBytes(transforms[i]).CopyTo(bytes, i * sizeof(float));
        var buffer = VirtualAllocEx(_handle, IntPtr.Zero, (UIntPtr)bytes.Length, MemCommit | MemReserve, PageReadWrite).ToInt64();
        if (buffer == 0 || !WriteBytes(new IntPtr(buffer), bytes))
        {
            if (buffer != 0) VirtualFreeEx(_handle, new IntPtr(buffer), UIntPtr.Zero, MemRelease);
            return false;
        }

        if (!TryReadFloatAbsolute(_cameraGameClockAddress, out var startGameTime))
        {
            VirtualFreeEx(_handle, new IntPtr(buffer), UIntPtr.Zero, MemRelease);
            return false;
        }
        _cameraPlaybackBuffers.Add(buffer);
        WriteBytes(new IntPtr(_cameraHookAllocation + 0x0B), [0]);
        WriteBytes(new IntPtr(_cameraHookAllocation + 0x0C), [0]);
        return WriteBytes(new IntPtr(_cameraHookAllocation + 0x110), BitConverter.GetBytes(buffer)) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x118), BitConverter.GetBytes(transforms.Count / 6)) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x11C), BitConverter.GetBytes(0)) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x128), BitConverter.GetBytes(startGameTime)) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x12C), BitConverter.GetBytes(1000.0f * speedMultiplier)) &&
               WriteBytes(new IntPtr(_cameraHookAllocation + 0x0B), [1]);
    }

    public bool IsCapturedCameraPlaybackActive()
        => IsCameraCaptureHookInstalled && TryReadBytesAbsolute(_cameraHookAllocation + 0x0B, 1, out var value) && value[0] != 0;

    private bool TryGetRemoteProcAddress(string moduleName, string exportName, out long address)
    {
        address = 0;
        if (!_modules.TryGetValue(moduleName, out var remoteModule))
            return false;
        var localModule = GetModuleHandle(moduleName);
        if (localModule == IntPtr.Zero)
            return false;
        var localExport = GetProcAddress(localModule, exportName);
        if (localExport == IntPtr.Zero)
            return false;
        address = remoteModule.BaseAddress.ToInt64() + (localExport.ToInt64() - localModule.ToInt64());
        return true;
    }

    private static byte[] BuildCameraCaptureStub(
        long stubAddress,
        long sharedAddress,
        long returnAddress,
        long gameClockAddress,
        byte[] originalInstructions)
    {
        var code = new List<byte>(256);
        code.AddRange([0x48, 0x89, 0x35]); // mov [captured camera],rsi
        AddRipDisplacement(code, stubAddress, sharedAddress);

        // The transform is committed on the camera update itself. Preserve every register,
        // flag, and SIMD value used by the displaced game instructions.
        code.Add(0x9C); // pushfq
        code.Add(0x50); // push rax
        code.AddRange([0x41, 0x50]); // push r8
        code.AddRange([0x48, 0x83, 0xEC, 0x20]);
        code.AddRange([0xF3, 0x0F, 0x7F, 0x04, 0x24]); // movdqu [rsp],xmm0
        code.AddRange([0xF3, 0x0F, 0x7F, 0x4C, 0x24, 0x10]); // movdqu [rsp+10h],xmm1

        // Publish Freecam's requested transform before any manual constraint is applied.
        code.AddRange([0x4C, 0x8D, 0x05]);
        AddRipDisplacement(code, stubAddress, sharedAddress + CameraRawTransformOffset);
        for (var i = 0; i < 6; i++)
        {
            code.AddRange([0x8B, 0x46, (byte)(0x18 + (i * 4))]);
            code.AddRange([0x41, 0x89, 0x40, (byte)(i * 4)]);
        }

        code.AddRange([0x80, 0x3D]); // cmp playback flag,0
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x0B, trailingBytes: 1);
        code.Add(0);
        code.AddRange([0x0F, 0x84]);
        var manualJump = code.Count;
        code.AddRange([0, 0, 0, 0]);

        code.AddRange([0x48, 0xB8]); // mov rax,game clock
        code.AddRange(BitConverter.GetBytes(gameClockAddress));
        code.AddRange([0xF3, 0x0F, 0x10, 0x00]); // movss xmm0,[rax]
        code.AddRange([0xF3, 0x0F, 0x5C, 0x05]); // subss xmm0,start time
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x128);
        code.AddRange([0xF3, 0x0F, 0x59, 0x05]); // mulss xmm0,1000
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x12C);
        code.AddRange([0x0F, 0x57, 0xC9]); // xorps xmm1,xmm1
        code.AddRange([0xF3, 0x0F, 0x5F, 0xC1]); // maxss xmm0,xmm1
        code.AddRange([0xF3, 0x0F, 0x2C, 0xC0]); // cvttss2si eax,xmm0
        code.AddRange([0x89, 0x05]); // publish current sample index
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x11C);
        code.AddRange([0x3B, 0x05]); // cmp eax,playback count
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x118);
        code.AddRange([0x0F, 0x83]);
        var finishedJump = code.Count;
        code.AddRange([0, 0, 0, 0]);
        code.AddRange([0x4C, 0x8B, 0x05]); // mov r8,playback samples
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x110);
        code.AddRange([0x6B, 0xC0, 0x18]); // imul eax,eax,24
        code.AddRange([0x49, 0x01, 0xC0]); // add r8,rax
        code.Add(0xE9);
        var playbackCopyJump = code.Count;
        code.AddRange([0, 0, 0, 0]);

        var manualOffset = code.Count;
        code.AddRange([0x80, 0x3D]); // cmp manual target ready,0
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x09, trailingBytes: 1);
        code.Add(0);
        code.AddRange([0x0F, 0x84]);
        var restoreFromManualJump = code.Count;
        code.AddRange([0, 0, 0, 0]);
        code.AddRange([0x0F, 0xB6, 0x05]); // movzx eax,active target
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x0A);
        code.AddRange([0x85, 0xC0, 0x74, 0x09]);
        code.AddRange([0x4C, 0x8D, 0x05]); // lea r8,target 1
        AddRipDisplacement(code, stubAddress, sharedAddress + CameraManualTarget1Offset);
        code.AddRange([0xEB, 0x07]);
        code.AddRange([0x4C, 0x8D, 0x05]); // lea r8,target 0
        AddRipDisplacement(code, stubAddress, sharedAddress + CameraManualTarget0Offset);

        code.AddRange([0x80, 0x3D]); // cmp position-only flag,0
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x0C, trailingBytes: 1);
        code.Add(0);
        code.AddRange([0x0F, 0x84]);
        var manualFullCopyJump = code.Count;
        code.AddRange([0, 0, 0, 0]);
        for (var i = 0; i < 3; i++)
        {
            code.AddRange([0x41, 0x8B, 0x40, (byte)(i * 4)]);
            code.AddRange([0x89, 0x46, (byte)(0x18 + (i * 4))]);
        }
        code.Add(0xE9);
        var restoreFromPositionCopyJump = code.Count;
        code.AddRange([0, 0, 0, 0]);

        var copyOffset = code.Count;
        for (var i = 0; i < 6; i++)
        {
            code.AddRange([0x41, 0x8B, 0x40, (byte)(i * 4)]);
            code.AddRange([0x89, 0x46, (byte)(0x18 + (i * 4))]);
        }
        code.Add(0xE9);
        var restoreFromCopyJump = code.Count;
        code.AddRange([0, 0, 0, 0]);

        var finishedOffset = code.Count;
        code.AddRange([0xC6, 0x05]); // clear playback and stale manual hold
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x0B, trailingBytes: 1);
        code.Add(0);
        code.AddRange([0xC6, 0x05]);
        AddRipDisplacement(code, stubAddress, sharedAddress + 0x09, trailingBytes: 1);
        code.Add(0);

        var restoreOffset = code.Count;
        code.AddRange([0xF3, 0x0F, 0x6F, 0x04, 0x24]);
        code.AddRange([0xF3, 0x0F, 0x6F, 0x4C, 0x24, 0x10]);
        code.AddRange([0x48, 0x83, 0xC4, 0x20]);
        code.AddRange([0x41, 0x58, 0x58, 0x9D]);
        // MCC's displaced instructions snapshot camera+18h for all downstream view
        // calculations. Execute them only after the override has been committed so
        // the entire render update observes one coherent transform.
        // The verified normal-camera window ends in E8 rel32. A remote allocation
        // is not guaranteed to be within rel32 range, so replay that call absolutely.
        if (originalInstructions.Length == 14 && originalInstructions[9] == 0xE8)
        {
            code.AddRange(originalInstructions.Take(9));
            var originalCallTarget = returnAddress + BitConverter.ToInt32(originalInstructions, 10);
            code.AddRange([0x48, 0xB8]); // mov rax,absolute call target
            code.AddRange(BitConverter.GetBytes(originalCallTarget));
            code.AddRange([0xFF, 0xD0]); // call rax
        }
        else
        {
            code.AddRange(originalInstructions);
        }
        code.AddRange([0xFF, 0x25, 0, 0, 0, 0]);
        code.AddRange(BitConverter.GetBytes(returnAddress));

        PatchRelativeJump(code, manualJump, manualOffset);
        PatchRelativeJump(code, finishedJump, finishedOffset);
        PatchRelativeJump(code, playbackCopyJump, copyOffset);
        PatchRelativeJump(code, manualFullCopyJump, copyOffset);
        PatchRelativeJump(code, restoreFromPositionCopyJump, restoreOffset);
        PatchRelativeJump(code, restoreFromManualJump, restoreOffset);
        PatchRelativeJump(code, restoreFromCopyJump, restoreOffset);
        return code.ToArray();
    }

    private static void AddRipDisplacement(List<byte> code, long codeAddress, long targetAddress, int trailingBytes = 0)
    {
        var nextInstruction = codeAddress + code.Count + sizeof(int) + trailingBytes;
        code.AddRange(BitConverter.GetBytes(checked((int)(targetAddress - nextInstruction))));
    }

    private static byte[] BuildCameraPlaybackWorker(
        long workerAddress,
        long sharedAddress,
        long sleepAddress,
        long gameClockAddress,
        long timerResolutionAddress)
    {
        var code = new List<byte>(128);
        code.AddRange([0x48, 0x83, 0xEC, 0x28]); // sub rsp, 28h
        code.AddRange([0xB9, 0x10, 0x27, 0, 0]); // ecx = 10,000 (one millisecond in 100ns units)
        code.AddRange([0xBA, 0x01, 0, 0, 0]); // edx = TRUE
        code.AddRange([0x4C, 0x8D, 0x05]); // r8 = storage for actual resolution
        var timerStorageNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x120) - timerStorageNext))));
        code.AddRange([0x48, 0xB8]);
        code.AddRange(BitConverter.GetBytes(timerResolutionAddress));
        code.AddRange([0xFF, 0xD0]); // NtSetTimerResolution(1ms, TRUE, ...)
        var loopOffset = code.Count;
        code.AddRange([0x80, 0x3D]); // cmp byte ptr [rip+disp32], 0 (lifetime)
        var lifetimeNext = workerAddress + code.Count + 5;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x08) - lifetimeNext))));
        code.Add(0x00);
        code.AddRange([0x0F, 0x84]);
        var exitJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        code.AddRange([0xB9, 0x01, 0, 0, 0]); // mov ecx, 1
        code.AddRange([0x48, 0xB8]); // mov rax, Sleep
        code.AddRange(BitConverter.GetBytes(sleepAddress));
        code.AddRange([0xFF, 0xD0]); // call rax
        code.AddRange([0x48, 0x8B, 0x15]); // mov rdx, captured camera
        var pointerNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)(sharedAddress - pointerNext))));
        code.AddRange([0x48, 0x85, 0xD2]); // test rdx, rdx
        code.AddRange([0x0F, 0x84]);
        var noPointerJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        code.AddRange([0x80, 0x3D]); // cmp playback flag, 0
        var playbackFlagNext = workerAddress + code.Count + 5;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x0B) - playbackFlagNext))));
        code.Add(0);
        code.AddRange([0x0F, 0x84]);
        var manualJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        code.AddRange([0x48, 0xB8]); // exact live MCC game-time float resolved from HaloDirector's signature
        code.AddRange(BitConverter.GetBytes(gameClockAddress));
        code.AddRange([0xF3, 0x0F, 0x10, 0x00]); // movss xmm0,[rax]
        code.AddRange([0xF3, 0x0F, 0x5C, 0x05]); // subss xmm0,start game time
        var startTimeNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x128) - startTimeNext))));
        code.AddRange([0xF3, 0x0F, 0x59, 0x05]); // mulss xmm0,1000.0
        var millisecondsNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x12C) - millisecondsNext))));
        code.AddRange([0x0F, 0x57, 0xC9]); // xorps xmm1,xmm1
        code.AddRange([0xF3, 0x0F, 0x5F, 0xC1]); // maxss xmm0,0
        code.AddRange([0xF3, 0x0F, 0x2C, 0xC0]); // cvttss2si eax,xmm0
        code.AddRange([0x89, 0x05]); // publish current millisecond index for status
        var indexNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x11C) - indexNext))));
        code.AddRange([0x3B, 0x05]); // cmp eax, playback count
        var countNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x118) - countNext))));
        code.AddRange([0x0F, 0x83]);
        var finishedJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        code.AddRange([0x4C, 0x8B, 0x05]); // mov r8, playback samples
        var samplesNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x110) - samplesNext))));
        code.AddRange([0x6B, 0xC0, 0x18]); // imul eax,eax,24
        code.AddRange([0x49, 0x01, 0xC0]); // add r8,rax
        code.AddRange([0x48, 0x8B, 0x15]); // calls above clobber rdx; reload captured camera
        var playbackPointerNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)(sharedAddress - playbackPointerNext))));
        code.Add(0xE9);
        var playbackCopyJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        var manualOffset = code.Count;
        code.AddRange([0x80, 0x3D]); // cmp byte ptr [rip+disp32], 0 (target ready)
        var readyNext = workerAddress + code.Count + 5;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x09) - readyNext))));
        code.Add(0x00);
        code.AddRange([0x0F, 0x84]);
        var notReadyJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        code.AddRange([0x0F, 0xB6, 0x05]); // movzx eax, active target byte
        var activeTargetNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x0A) - activeTargetNext))));
        code.AddRange([0x85, 0xC0]); // test eax, eax
        code.AddRange([0x74, 0x09]); // je use target 0
        code.AddRange([0x4C, 0x8D, 0x05]); // lea r8, target 1
        var targetOneNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x100) - targetOneNext))));
        code.AddRange([0xEB, 0x07]); // jmp target selected
        code.AddRange([0x4C, 0x8D, 0x05]); // lea r8, target 0
        var targetZeroNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x10) - targetZeroNext))));
        var copyOffset = code.Count;
        for (var i = 0; i < 6; i++)
        {
            code.AddRange([0x41, 0x8B, 0x40, (byte)(i * 4)]); // mov eax, [r8+offset]
            code.AddRange([0x89, 0x42, (byte)(0x18 + (i * 4))]); // mov [rdx+offset], eax
        }
        code.Add(0xE9);
        var loopJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        var finishedOffset = code.Count;
        code.AddRange([0xC6, 0x05]); // mov playback flag,0
        var stopFlagNext = workerAddress + code.Count + 5;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x0B) - stopFlagNext))));
        code.Add(0);
        code.Add(0xE9);
        var finishedLoopJumpOffset = code.Count;
        code.AddRange([0, 0, 0, 0]);
        var exitOffset = code.Count;
        code.AddRange([0xB9, 0x10, 0x27, 0, 0]);
        code.AddRange([0x31, 0xD2]); // edx = FALSE
        code.AddRange([0x4C, 0x8D, 0x05]);
        var releaseStorageNext = workerAddress + code.Count + 4;
        code.AddRange(BitConverter.GetBytes(checked((int)((sharedAddress + 0x120) - releaseStorageNext))));
        code.AddRange([0x48, 0xB8]);
        code.AddRange(BitConverter.GetBytes(timerResolutionAddress));
        code.AddRange([0xFF, 0xD0]);
        code.AddRange([0x48, 0x83, 0xC4, 0x28, 0x31, 0xC0, 0xC3]);

        PatchRelativeJump(code, exitJumpOffset, exitOffset);
        PatchRelativeJump(code, noPointerJumpOffset, loopOffset);
        PatchRelativeJump(code, manualJumpOffset, manualOffset);
        PatchRelativeJump(code, finishedJumpOffset, finishedOffset);
        PatchRelativeJump(code, playbackCopyJumpOffset, copyOffset);
        PatchRelativeJump(code, notReadyJumpOffset, loopOffset);
        PatchRelativeJump(code, loopJumpOffset, loopOffset);
        PatchRelativeJump(code, finishedLoopJumpOffset, loopOffset);
        return code.ToArray();
    }

    private static void PatchRelativeJump(List<byte> code, int displacementOffset, int targetOffset)
    {
        var displacement = targetOffset - (displacementOffset + sizeof(int));
        var bytes = BitConverter.GetBytes(displacement);
        for (var i = 0; i < bytes.Length; i++)
            code[displacementOffset + i] = bytes[i];
    }

    private bool WriteProtectedCode(long address, byte[] bytes)
    {
        if (!VirtualProtectEx(_handle, new IntPtr(address), (UIntPtr)bytes.Length, PageExecuteReadWrite, out var oldProtect))
            return false;
        var wrote = WriteBytes(new IntPtr(address), bytes);
        FlushInstructionCache(_handle, new IntPtr(address), (UIntPtr)bytes.Length);
        VirtualProtectEx(_handle, new IntPtr(address), (UIntPtr)bytes.Length, oldProtect, out _);
        return wrote;
    }

    public void RefreshModules()
    {
        _modules.Clear();
        if (_process is null || _process.HasExited)
            return;

        try
        {
            _process.Refresh();
            foreach (ProcessModule module in _process.Modules)
                _modules[module.ModuleName] = new H3ModuleInfo(module.BaseAddress, module.ModuleMemorySize);
        }
        catch
        {
            _modules.Clear();
        }

        if (!_modules.ContainsKey("halo3.dll"))
            RefreshModulesWithToolhelp(_process.Id);
    }

    private void RefreshModulesWithToolhelp(int processId)
    {
        var snapshot = CreateToolhelp32Snapshot(Th32csSnapModule | Th32csSnapModule32, (uint)processId);
        if (snapshot == IntPtr.Zero || snapshot == InvalidHandleValue)
            return;

        try
        {
            var entry = new ModuleEntry32
            {
                dwSize = (uint)Marshal.SizeOf<ModuleEntry32>()
            };

            if (!Module32First(snapshot, ref entry))
                return;

            do
            {
                if (!string.IsNullOrWhiteSpace(entry.szModule))
                    _modules[entry.szModule] = new H3ModuleInfo(entry.modBaseAddr, (int)entry.modBaseSize);
            }
            while (Module32Next(snapshot, ref entry));
        }
        finally
        {
            CloseHandle(snapshot);
        }
    }

    public bool TryReadByte(H3Address address, out byte value)
    {
        value = 0;
        try
        {
            var bytes = ReadBytes(ResolveAddress(address), 1);
            value = bytes[0];
            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool TryReadFloat(H3Address address, out float value)
    {
        value = 0f;
        try
        {
            value = BitConverter.ToSingle(ReadBytes(ResolveAddress(address), 4), 0);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool TryReadFloatAbsolute(long address, out float value)
    {
        value = 0f;
        try
        {
            value = BitConverter.ToSingle(ReadBytes(new IntPtr(address), 4), 0);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool TryReadBytesAbsolute(long address, int length, out byte[] bytes)
    {
        bytes = [];
        try
        {
            bytes = ReadBytes(new IntPtr(address), length);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool TryWriteFloatAbsolute(long address, float value)
    {
        try
        {
            return WriteBytes(new IntPtr(address), BitConverter.GetBytes(value));
        }
        catch
        {
            return false;
        }
    }

    public bool TryReadBytes(H3Address address, int length, out byte[] bytes)
    {
        bytes = [];
        try
        {
            bytes = ReadBytes(ResolveAddress(address), length);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool TryWriteByte(H3Address address, byte value)
        => TryWriteBytes(address, [value]);

    public bool TryWriteFloat(H3Address address, float value)
        => TryWriteBytes(address, BitConverter.GetBytes(value));

    public bool TryWriteBytes(H3Address address, byte[] bytes)
    {
        try
        {
            return WriteBytes(ResolveAddress(address), bytes);
        }
        catch
        {
            return false;
        }
    }

    public bool TryResolveAddress(H3Address address, out long absoluteAddress)
    {
        absoluteAddress = 0;
        try
        {
            absoluteAddress = ResolveAddress(address).ToInt64();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public IReadOnlyList<H3PointerScanMatch> ScanReadablePointersTo(long targetAddress, long maxBytesToScan = 2L * 1024 * 1024 * 1024, int maxMatches = 256)
        => ScanReadablePointersToAny([targetAddress], maxBytesToScan, maxMatches);

    public IReadOnlyList<H3PointerScanMatch> ScanReadablePointersToAny(IEnumerable<long> targetAddresses, long maxBytesToScan = 2L * 1024 * 1024 * 1024, int maxMatches = 512)
    {
        var targets = targetAddresses.Where(address => address != 0).ToHashSet();
        if (!IsAttached || targets.Count == 0)
            return [];

        var matches = new List<H3PointerScanMatch>();
        long scannedBytes = 0;
        int scannedRegions = 0;
        int readableRegions = 0;
        var address = IntPtr.Zero;
        var infoSize = (UIntPtr)Marshal.SizeOf<MemoryBasicInformation>();

        while (VirtualQueryEx(_handle, address, out var info, infoSize) != UIntPtr.Zero)
        {
            var baseAddress = info.BaseAddress.ToInt64();
            var regionSize = info.RegionSize.ToInt64();
            if (regionSize <= 0)
                break;

            scannedRegions++;
            if (IsReadableCommittedRegion(info))
            {
                readableRegions++;
                ScanPointerRegion(info.BaseAddress, regionSize, targets, matches, maxMatches, ref scannedBytes, maxBytesToScan);
                if (matches.Count >= maxMatches || scannedBytes >= maxBytesToScan)
                    break;
            }

            var next = baseAddress + regionSize;
            if (next <= baseAddress || next >= 0x0000800000000000)
                break;

            address = new IntPtr(next);
        }

        return matches
            .Select(m => m with { ScannedRegions = scannedRegions, ReadableRegions = readableRegions, ScannedBytes = scannedBytes })
            .ToList();
    }

    private void ScanPointerRegion(
        IntPtr baseAddress,
        long regionSize,
        IReadOnlySet<long> targetAddresses,
        List<H3PointerScanMatch> matches,
        int maxMatches,
        ref long scannedBytes,
        long maxBytesToScan)
    {
        const int chunkSize = 64 * 1024;
        var start = baseAddress.ToInt64();
        for (long offset = 0; offset < regionSize && scannedBytes < maxBytesToScan && matches.Count < maxMatches; offset += chunkSize)
        {
            var length = (int)Math.Min(chunkSize, regionSize - offset);
            byte[] bytes;
            try
            {
                bytes = ReadBytes(new IntPtr(start + offset), length);
            }
            catch
            {
                continue;
            }

            scannedBytes += bytes.Length;
            for (int i = 0; i + 8 <= bytes.Length && matches.Count < maxMatches; i += 8)
            {
                var value = BitConverter.ToInt64(bytes, i);
                if (targetAddresses.Contains(value))
                    matches.Add(new H3PointerScanMatch(start + offset + i, value, 0, 0, 0));
            }
        }
    }

    public IReadOnlyList<H3ModulePatternMatch> FindModulePattern(string moduleName, byte[] pattern, int contextBytes = 16, int maxMatches = 64)
        => FindModulePattern(moduleName, pattern, new string('x', pattern.Length), contextBytes, maxMatches);

    public IReadOnlyList<H3ModulePatternMatch> FindModulePattern(string moduleName, byte[] pattern, string mask, int contextBytes = 16, int maxMatches = 64)
    {
        if (!IsAttached || pattern.Length == 0 || mask.Length != pattern.Length)
            return [];

        if (!_modules.TryGetValue(moduleName, out var module))
        {
            RefreshModules();
            if (!_modules.TryGetValue(moduleName, out module))
                return [];
        }

        byte[] bytes;
        try
        {
            bytes = ReadBytes(module.BaseAddress, module.Size);
        }
        catch
        {
            return [];
        }

        var matches = new List<H3ModulePatternMatch>();
        for (int i = 0; i <= bytes.Length - pattern.Length && matches.Count < maxMatches; i++)
        {
            var matched = true;
            for (int p = 0; p < pattern.Length; p++)
            {
                if (mask[p] == '?' || bytes[i + p] == pattern[p])
                    continue;

                matched = false;
                break;
            }

            if (!matched)
                continue;

            var contextStart = Math.Max(0, i - contextBytes);
            var contextLength = Math.Min(bytes.Length - contextStart, pattern.Length + (contextBytes * 2));
            matches.Add(new H3ModulePatternMatch(
                moduleName,
                i,
                module.BaseAddress.ToInt64() + i,
                bytes.Skip(contextStart).Take(contextLength).ToArray(),
                i - contextStart));
        }

        return matches;
    }

    public IReadOnlyList<H3RipRelativeReference> FindRipRelativeReferences(
        string moduleName,
        long targetModuleOffset,
        int contextBytes = 16,
        int maxMatches = 128,
        int toleranceBytes = 0x20)
    {
        if (!IsAttached)
            return [];

        if (!_modules.TryGetValue(moduleName, out var module))
        {
            RefreshModules();
            if (!_modules.TryGetValue(moduleName, out module))
                return [];
        }

        byte[] bytes;
        try
        {
            bytes = ReadBytes(module.BaseAddress, module.Size);
        }
        catch
        {
            return [];
        }

        var targetAbsolute = module.BaseAddress.ToInt64() + targetModuleOffset;
        var matches = new List<H3RipRelativeReference>();
        for (int i = 0; i + 7 <= bytes.Length && matches.Count < maxMatches; i++)
        {
            foreach (var displacementOffset in CandidateRipDisplacementOffsets(bytes, i))
            {
                if (i + displacementOffset + 4 > bytes.Length)
                    continue;

                var displacement = BitConverter.ToInt32(bytes, i + displacementOffset);
                var instructionLength = displacementOffset + 4;
                var resolved = module.BaseAddress.ToInt64() + i + instructionLength + displacement;
                if (Math.Abs(resolved - targetAbsolute) > toleranceBytes)
                    continue;

                var contextStart = Math.Max(0, i - contextBytes);
                var contextLength = Math.Min(bytes.Length - contextStart, instructionLength + (contextBytes * 2));
                matches.Add(new H3RipRelativeReference(
                    moduleName,
                    i,
                    module.BaseAddress.ToInt64() + i,
                    displacementOffset,
                    instructionLength,
                    resolved,
                    resolved - module.BaseAddress.ToInt64(),
                    bytes.Skip(contextStart).Take(contextLength).ToArray(),
                    i - contextStart));
                break;
            }
        }

        return matches;
    }

    private static IEnumerable<int> CandidateRipDisplacementOffsets(byte[] bytes, int offset)
    {
        if (offset + 7 > bytes.Length)
            yield break;

        // Common x64 RIP-relative forms:
        // 48/4C 8B/89/8D xx disp32, F3 0F 10/11 xx disp32, 0F 28/29 xx disp32.
        if ((bytes[offset] == 0x48 || bytes[offset] == 0x4C) &&
            (bytes[offset + 1] == 0x8B || bytes[offset + 1] == 0x89 || bytes[offset + 1] == 0x8D) &&
            IsRipModRm(bytes[offset + 2]))
        {
            yield return 3;
        }

        if (offset + 8 <= bytes.Length &&
            bytes[offset] == 0xF3 &&
            bytes[offset + 1] == 0x0F &&
            (bytes[offset + 2] == 0x10 || bytes[offset + 2] == 0x11) &&
            IsRipModRm(bytes[offset + 3]))
        {
            yield return 4;
        }

        if (offset + 7 <= bytes.Length &&
            bytes[offset] == 0x0F &&
            (bytes[offset + 1] == 0x28 || bytes[offset + 1] == 0x29) &&
            IsRipModRm(bytes[offset + 2]))
        {
            yield return 3;
        }
    }

    private static bool IsRipModRm(byte value)
        => (value & 0xC7) == 0x05;

    public H3FloatScanResult ScanReadableFloats(float minValue, float maxValue, long maxBytesToScan = 768L * 1024 * 1024, int maxSamples = 250_000)
        => ScanFloats(minValue, maxValue, maxBytesToScan, maxSamples, writableOnly: false);

    public H3FloatScanResult ScanWritableFloats(float minValue, float maxValue, long maxBytesToScan = 768L * 1024 * 1024, int maxSamples = 250_000)
        => ScanFloats(minValue, maxValue, maxBytesToScan, maxSamples, writableOnly: true);

    private H3FloatScanResult ScanFloats(float minValue, float maxValue, long maxBytesToScan, int maxSamples, bool writableOnly)
    {
        if (!IsAttached)
            return new H3FloatScanResult([], 0, 0, 0);

        var samples = new List<H3FloatScanSample>();
        long scannedBytes = 0;
        int scannedRegions = 0;
        int readableRegions = 0;
        var address = IntPtr.Zero;
        var infoSize = (UIntPtr)Marshal.SizeOf<MemoryBasicInformation>();

        while (VirtualQueryEx(_handle, address, out var info, infoSize) != UIntPtr.Zero)
        {
            var baseAddress = info.BaseAddress.ToInt64();
            var regionSize = info.RegionSize.ToInt64();
            if (regionSize <= 0)
                break;

            scannedRegions++;
            var canScan = writableOnly
                ? IsWritableCommittedRegion(info)
                : IsReadableCommittedRegion(info);
            if (canScan)
            {
                readableRegions++;
                ScanFloatRegion(info.BaseAddress, regionSize, minValue, maxValue, samples, maxSamples, ref scannedBytes, maxBytesToScan);
                if (samples.Count >= maxSamples || scannedBytes >= maxBytesToScan)
                    break;
            }

            var next = baseAddress + regionSize;
            if (next <= baseAddress || next >= 0x0000800000000000)
                break;

            address = new IntPtr(next);
        }

        return new H3FloatScanResult(samples, scannedRegions, readableRegions, scannedBytes);
    }

    public H3FloatTripleScanResult ScanReadableFloatTriples(float targetX, float targetY, float targetZ, float tolerance, long maxBytesToScan = 768L * 1024 * 1024, int maxMatches = 64)
        => ScanFloatTriples(targetX, targetY, targetZ, tolerance, maxBytesToScan, maxMatches, writableOnly: false);

    public H3FloatTripleScanResult ScanWritableFloatTriples(float targetX, float targetY, float targetZ, float tolerance, long maxBytesToScan = 2L * 1024 * 1024 * 1024, int maxMatches = 64)
        => ScanFloatTriples(targetX, targetY, targetZ, tolerance, maxBytesToScan, maxMatches, writableOnly: true);

    public H3FloatTripleScanResult ScanWritableFloatTriplesInRange(float minValue, float maxValue, long maxBytesToScan = 2L * 1024 * 1024 * 1024, int maxMatches = 200_000)
    {
        if (!IsAttached)
            return new H3FloatTripleScanResult([], 0, 0, 0);

        var matches = new List<H3FloatTripleScanMatch>();
        long scannedBytes = 0;
        int scannedRegions = 0;
        int readableRegions = 0;
        var address = IntPtr.Zero;
        var infoSize = (UIntPtr)Marshal.SizeOf<MemoryBasicInformation>();

        while (VirtualQueryEx(_handle, address, out var info, infoSize) != UIntPtr.Zero)
        {
            var baseAddress = info.BaseAddress.ToInt64();
            var regionSize = info.RegionSize.ToInt64();
            if (regionSize <= 0)
                break;

            scannedRegions++;
            if (IsWritableCommittedRegion(info))
            {
                readableRegions++;
                ScanFloatTripleRangeRegion(info.BaseAddress, regionSize, minValue, maxValue, matches, maxMatches, ref scannedBytes, maxBytesToScan);
                if (matches.Count >= maxMatches || scannedBytes >= maxBytesToScan)
                    break;
            }

            var next = baseAddress + regionSize;
            if (next <= baseAddress || next >= 0x0000800000000000)
                break;

            address = new IntPtr(next);
        }

        return new H3FloatTripleScanResult(matches, scannedRegions, readableRegions, scannedBytes);
    }

    public H3FloatTripleScanResult ScanWritableUnitVectorPairs(long maxBytesToScan = 2L * 1024 * 1024 * 1024, int maxMatches = 150_000)
    {
        if (!IsAttached)
            return new H3FloatTripleScanResult([], 0, 0, 0);

        var matches = new List<H3FloatTripleScanMatch>();
        long scannedBytes = 0;
        int scannedRegions = 0;
        int readableRegions = 0;
        var address = IntPtr.Zero;
        var infoSize = (UIntPtr)Marshal.SizeOf<MemoryBasicInformation>();
        while (VirtualQueryEx(_handle, address, out var info, infoSize) != UIntPtr.Zero)
        {
            var baseAddress = info.BaseAddress.ToInt64();
            var regionSize = info.RegionSize.ToInt64();
            if (regionSize <= 0)
                break;
            scannedRegions++;
            if (IsWritableCommittedRegion(info))
            {
                readableRegions++;
                ScanUnitVectorPairRegion(info.BaseAddress, regionSize, matches, maxMatches, ref scannedBytes, maxBytesToScan);
                if (matches.Count >= maxMatches || scannedBytes >= maxBytesToScan)
                    break;
            }
            var next = baseAddress + regionSize;
            if (next <= baseAddress || next >= 0x0000800000000000)
                break;
            address = new IntPtr(next);
        }
        return new H3FloatTripleScanResult(matches, scannedRegions, readableRegions, scannedBytes);
    }

    private void ScanUnitVectorPairRegion(IntPtr baseAddress, long regionSize,
        List<H3FloatTripleScanMatch> matches, int maxMatches, ref long scannedBytes, long maxBytesToScan)
    {
        const int chunkSize = 64 * 1024;
        var start = baseAddress.ToInt64();
        for (long offset = 0; offset < regionSize && scannedBytes < maxBytesToScan && matches.Count < maxMatches; offset += chunkSize)
        {
            var length = (int)Math.Min(chunkSize, regionSize - offset);
            byte[] bytes;
            try { bytes = ReadBytes(new IntPtr(start + offset), length); }
            catch { continue; }
            scannedBytes += bytes.Length;
            for (int i = 0; i + 8 <= bytes.Length && matches.Count < maxMatches; i += 4)
            {
                var x = BitConverter.ToSingle(bytes, i);
                var y = BitConverter.ToSingle(bytes, i + 4);
                if (!float.IsFinite(x) || !float.IsFinite(y) || Math.Abs(x) > 1.01f || Math.Abs(y) > 1.01f)
                    continue;
                var lengthSquared = (x * x) + (y * y);
                if (lengthSquared >= 0.94f && lengthSquared <= 1.06f)
                    matches.Add(new H3FloatTripleScanMatch(start + offset + i, x, y, 0));
            }
        }
    }

    private H3FloatTripleScanResult ScanFloatTriples(float targetX, float targetY, float targetZ, float tolerance, long maxBytesToScan, int maxMatches, bool writableOnly)
    {
        if (!IsAttached)
            return new H3FloatTripleScanResult([], 0, 0, 0);

        var matches = new List<H3FloatTripleScanMatch>();
        long scannedBytes = 0;
        int scannedRegions = 0;
        int readableRegions = 0;
        var address = IntPtr.Zero;
        var infoSize = (UIntPtr)Marshal.SizeOf<MemoryBasicInformation>();

        while (VirtualQueryEx(_handle, address, out var info, infoSize) != UIntPtr.Zero)
        {
            var baseAddress = info.BaseAddress.ToInt64();
            var regionSize = info.RegionSize.ToInt64();
            if (regionSize <= 0)
                break;

            scannedRegions++;
            var canScan = writableOnly
                ? IsWritableCommittedRegion(info)
                : IsReadableCommittedRegion(info);
            if (canScan)
            {
                readableRegions++;
                ScanFloatTripleRegion(info.BaseAddress, regionSize, targetX, targetY, targetZ, tolerance, matches, maxMatches, ref scannedBytes, maxBytesToScan);
                if (matches.Count >= maxMatches || scannedBytes >= maxBytesToScan)
                    break;
            }

            var next = baseAddress + regionSize;
            if (next <= baseAddress || next >= 0x0000800000000000)
                break;

            address = new IntPtr(next);
        }

        return new H3FloatTripleScanResult(matches, scannedRegions, readableRegions, scannedBytes);
    }

    private void ScanFloatTripleRegion(
        IntPtr baseAddress,
        long regionSize,
        float targetX,
        float targetY,
        float targetZ,
        float tolerance,
        List<H3FloatTripleScanMatch> matches,
        int maxMatches,
        ref long scannedBytes,
        long maxBytesToScan)
    {
        const int chunkSize = 64 * 1024;
        var start = baseAddress.ToInt64();
        for (long offset = 0; offset < regionSize && scannedBytes < maxBytesToScan && matches.Count < maxMatches; offset += chunkSize)
        {
            var length = (int)Math.Min(chunkSize, regionSize - offset);
            byte[] bytes;
            try
            {
                bytes = ReadBytes(new IntPtr(start + offset), length);
            }
            catch
            {
                continue;
            }

            scannedBytes += bytes.Length;
            for (int i = 0; i + 12 <= bytes.Length && matches.Count < maxMatches; i += 4)
            {
                var x = BitConverter.ToSingle(bytes, i);
                var y = BitConverter.ToSingle(bytes, i + 4);
                var z = BitConverter.ToSingle(bytes, i + 8);
                if (Math.Abs(x - targetX) <= tolerance &&
                    Math.Abs(y - targetY) <= tolerance &&
                    Math.Abs(z - targetZ) <= tolerance)
                {
                    matches.Add(new H3FloatTripleScanMatch(start + offset + i, x, y, z));
                }
            }
        }
    }

    private void ScanFloatTripleRangeRegion(
        IntPtr baseAddress,
        long regionSize,
        float minValue,
        float maxValue,
        List<H3FloatTripleScanMatch> matches,
        int maxMatches,
        ref long scannedBytes,
        long maxBytesToScan)
    {
        const int chunkSize = 64 * 1024;
        var start = baseAddress.ToInt64();
        for (long offset = 0; offset < regionSize && scannedBytes < maxBytesToScan && matches.Count < maxMatches; offset += chunkSize)
        {
            var length = (int)Math.Min(chunkSize, regionSize - offset);
            byte[] bytes;
            try
            {
                bytes = ReadBytes(new IntPtr(start + offset), length);
            }
            catch
            {
                continue;
            }

            scannedBytes += bytes.Length;
            for (int i = 0; i + 12 <= bytes.Length && matches.Count < maxMatches; i += 4)
            {
                var x = BitConverter.ToSingle(bytes, i);
                var y = BitConverter.ToSingle(bytes, i + 4);
                var z = BitConverter.ToSingle(bytes, i + 8);
                if (IsCameraLikeFloat(x, minValue, maxValue) &&
                    IsCameraLikeFloat(y, minValue, maxValue) &&
                    IsCameraLikeFloat(z, minValue, maxValue) &&
                    (Math.Abs(x) > 0.0001f || Math.Abs(y) > 0.0001f || Math.Abs(z) > 0.0001f))
                {
                    matches.Add(new H3FloatTripleScanMatch(start + offset + i, x, y, z));
                }
            }
        }
    }

    private static bool IsCameraLikeFloat(float value, float minValue, float maxValue)
        => !float.IsNaN(value) && !float.IsInfinity(value) && value >= minValue && value <= maxValue;

    private void ScanFloatRegion(
        IntPtr baseAddress,
        long regionSize,
        float minValue,
        float maxValue,
        List<H3FloatScanSample> samples,
        int maxSamples,
        ref long scannedBytes,
        long maxBytesToScan)
    {
        const int chunkSize = 64 * 1024;
        var start = baseAddress.ToInt64();
        for (long offset = 0; offset < regionSize && scannedBytes < maxBytesToScan && samples.Count < maxSamples; offset += chunkSize)
        {
            var length = (int)Math.Min(chunkSize, regionSize - offset);
            byte[] bytes;
            try
            {
                bytes = ReadBytes(new IntPtr(start + offset), length);
            }
            catch
            {
                continue;
            }

            scannedBytes += bytes.Length;
            for (int i = 0; i + 4 <= bytes.Length && samples.Count < maxSamples; i += 4)
            {
                var value = BitConverter.ToSingle(bytes, i);
                if (float.IsNaN(value) || float.IsInfinity(value) || value < minValue || value > maxValue)
                    continue;

                samples.Add(new H3FloatScanSample(start + offset + i, value));
            }
        }
    }

    private static bool IsWritableCommittedRegion(MemoryBasicInformation info)
    {
        if (info.State != MemCommit)
            return false;

        var protection = info.Protect;
        if ((protection & PageGuard) != 0 || (protection & PageNoAccess) != 0)
            return false;

        return (protection & (PageReadWrite | PageWriteCopy | PageExecuteReadWrite | PageExecuteWriteCopy)) != 0;
    }

    private static bool IsReadableCommittedRegion(MemoryBasicInformation info)
    {
        if (info.State != MemCommit)
            return false;

        var protection = info.Protect;
        if ((protection & PageGuard) != 0 || (protection & PageNoAccess) != 0)
            return false;

        return (protection & (PageReadOnly | PageReadWrite | PageWriteCopy | PageExecuteRead | PageExecuteReadWrite | PageExecuteWriteCopy)) != 0;
    }

    private IntPtr ResolveAddress(H3Address address)
    {
        if (!IsAttached)
            throw new InvalidOperationException("Not attached to MCC.");

        if (!_modules.TryGetValue(address.ModuleName, out var module))
        {
            RefreshModules();
            if (!_modules.TryGetValue(address.ModuleName, out module))
                throw new InvalidOperationException($"Module not loaded: {address.ModuleName}");
        }

        var current = IntPtr.Add(module.BaseAddress, checked((int)address.BaseOffset));
        if (address.PointerOffsets.Length == 0)
            return current;

        current = ReadPointer(current);
        for (int i = 0; i < address.PointerOffsets.Length; i++)
        {
            var offset = checked((int)address.PointerOffsets[i]);
            current = i < address.PointerOffsets.Length - 1
                ? ReadPointer(IntPtr.Add(current, offset))
                : IntPtr.Add(current, offset);
        }

        return current;
    }

    private IntPtr ReadPointer(IntPtr address)
    {
        var bytes = ReadBytes(address, IntPtr.Size);
        var pointer = IntPtr.Size == 8
            ? BitConverter.ToInt64(bytes, 0)
            : BitConverter.ToInt32(bytes, 0);
        return new IntPtr(pointer);
    }

    private byte[] ReadBytes(IntPtr address, int length)
    {
        var buffer = new byte[length];
        if (!ReadProcessMemory(_handle, address, buffer, buffer.Length, out var read) ||
            read.ToInt64() != buffer.Length)
        {
            throw new InvalidOperationException("Failed to read process memory.");
        }

        return buffer;
    }

    private bool WriteBytes(IntPtr address, byte[] bytes)
        => WriteProcessMemory(_handle, address, bytes, bytes.Length, out var written) &&
           written.ToInt64() == bytes.Length;

    private static IEnumerable<Process> EnumerateLikelyRelatedEacProcesses(Process mccProcess)
    {
        DateTime mccStart;
        try { mccStart = mccProcess.StartTime; }
        catch { yield break; }

        Process[] processes;
        try { processes = Process.GetProcesses(); }
        catch { yield break; }

        foreach (var process in processes)
        {
            var isMatch = false;
            try
            {
                var name = process.ProcessName;
                if (name.Contains("EasyAntiCheat", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("EasyAntiCheat_EOS", StringComparison.OrdinalIgnoreCase))
                {
                    var start = process.StartTime;
                    isMatch = start >= mccStart.AddMinutes(-2);
                }
            }
            catch { }

            if (isMatch)
                yield return process;
            else
                process.Dispose();
        }
    }

    public void Dispose()
        => Detach();

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(int processAccess, bool inheritHandle, int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool ReadProcessMemory(IntPtr processHandle, IntPtr baseAddress, byte[] buffer, int size, out IntPtr bytesRead);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool WriteProcessMemory(IntPtr processHandle, IntPtr baseAddress, byte[] buffer, int size, out IntPtr bytesWritten);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr handle);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string moduleName);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true)]
    private static extern IntPtr GetProcAddress(IntPtr module, string procedureName);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateRemoteThread(
        IntPtr processHandle,
        IntPtr threadAttributes,
        UIntPtr stackSize,
        IntPtr startAddress,
        IntPtr parameter,
        uint creationFlags,
        out uint threadId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint processId);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool Module32First(IntPtr snapshot, ref ModuleEntry32 moduleEntry);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool Module32Next(IntPtr snapshot, ref ModuleEntry32 moduleEntry);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr VirtualQueryEx(IntPtr processHandle, IntPtr address, out MemoryBasicInformation buffer, UIntPtr length);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr VirtualAllocEx(IntPtr processHandle, IntPtr address, UIntPtr size, uint allocationType, uint protect);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool VirtualFreeEx(IntPtr processHandle, IntPtr address, UIntPtr size, uint freeType);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool VirtualProtectEx(IntPtr processHandle, IntPtr address, UIntPtr size, uint newProtect, out uint oldProtect);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FlushInstructionCache(IntPtr processHandle, IntPtr baseAddress, UIntPtr size);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ModuleEntry32
    {
        public uint dwSize;
        public uint th32ModuleID;
        public uint th32ProcessID;
        public uint GlblcntUsage;
        public uint ProccntUsage;
        public IntPtr modBaseAddr;
        public uint modBaseSize;
        public IntPtr hModule;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string szModule;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szExePath;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MemoryBasicInformation
    {
        public IntPtr BaseAddress;
        public IntPtr AllocationBase;
        public uint AllocationProtect;
        public ushort PartitionId;
        public IntPtr RegionSize;
        public uint State;
        public uint Protect;
        public uint Type;
    }
}

internal sealed record H3FloatScanResult(List<H3FloatScanSample> Samples, int ScannedRegions, int ReadableRegions, long ScannedBytes);

internal sealed record H3FloatScanSample(long Address, float Value);

internal sealed record H3FloatTripleScanResult(List<H3FloatTripleScanMatch> Matches, int ScannedRegions, int ReadableRegions, long ScannedBytes);

internal sealed record H3FloatTripleScanMatch(long Address, float X, float Y, float Z);

internal sealed record H3PointerScanMatch(long Address, long TargetAddress, int ScannedRegions, int ReadableRegions, long ScannedBytes);

internal readonly record struct H3ModuleInfo(IntPtr BaseAddress, int Size);

internal sealed record H3ModulePatternMatch(
    string ModuleName,
    int ModuleOffset,
    long AbsoluteAddress,
    byte[] ContextBytes,
    int PatternOffsetInContext);

internal sealed record H3RipRelativeReference(
    string ModuleName,
    int ModuleOffset,
    long AbsoluteAddress,
    int DisplacementOffset,
    int InstructionLength,
    long ResolvedAbsoluteAddress,
    long ResolvedModuleOffset,
    byte[] ContextBytes,
    int InstructionOffsetInContext);

internal sealed record H3Address(string ModuleName, long BaseOffset, params long[] PointerOffsets)
{
    public override string ToString()
    {
        var baseText = $"{ModuleName}+0x{BaseOffset.ToString("X", CultureInfo.InvariantCulture)}";
        return PointerOffsets.Length == 0
            ? baseText
            : $"{baseText},{string.Join(",", PointerOffsets.Select(o => "0x" + o.ToString("X", CultureInfo.InvariantCulture)))}";
    }
}
