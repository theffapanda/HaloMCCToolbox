using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace HaloToolbox;

internal sealed class ReachBloomSession : IDisposable
{
    private const uint ProcessCreateThread = 0x0002;
    private const uint ProcessQueryInformation = 0x0400;
    private const uint ProcessVmOperation = 0x0008;
    private const uint ProcessVmWrite = 0x0020;
    private const uint MemCommit = 0x1000;
    private const uint MemReserve = 0x2000;
    private const uint MemRelease = 0x8000;
    private const uint PageReadWrite = 0x04;
    private const uint FileMapRead = 0x0004;
    private const uint EventModifyState = 0x0002;
    private const uint WaitObject0 = 0;
    private const uint StateMagic = 0x52424C4D;

    private Process? _process;
    private IntPtr _stateMap;
    private IntPtr _stateView;

    public int? ProcessId => _process is { HasExited: false } ? _process.Id : null;
    public bool IsConnected => _stateView != IntPtr.Zero && _process is { HasExited: false };

    public ReachBloomState ReadState()
    {
        if (!IsConnected || unchecked((uint)Marshal.ReadInt32(_stateView, 0)) != StateMagic)
            return default;

        return new ReachBloomState(
            Marshal.ReadInt32(_stateView, 8) > 0,
            Marshal.ReadInt32(_stateView, 12) != 0,
            Marshal.ReadInt32(_stateView, 16) != 0,
            Marshal.ReadInt64(_stateView, 24));
    }

    public void ConnectOrInject(Process process, string payloadPath)
    {
        ArgumentNullException.ThrowIfNull(process);
        if (H3MemorySession.IsEasyAntiCheatLikelyLoaded(process))
            throw new InvalidOperationException("Easy Anti-Cheat appears to be active. Restart MCC with anti-cheat disabled.");
        if (!File.Exists(payloadPath))
            throw new FileNotFoundException("The Reach lighting hook is missing from the Toolbox installation.", payloadPath);

        DisposeHandles();
        _process = process;
        if (TryOpenState(process.Id))
            return;

        InjectLibrary(process, Path.GetFullPath(payloadPath));
        var deadline = Stopwatch.StartNew();
        while (deadline.Elapsed < TimeSpan.FromSeconds(6))
        {
            if (process.HasExited)
                throw new InvalidOperationException("MCC closed while the Reach lighting hook was being armed.");
            if (TryOpenState(process.Id))
                return;
            Thread.Sleep(100);
        }

        throw new InvalidOperationException("The Reach lighting hook loaded but did not report ready. Restart MCC and try again from the main menu.");
    }

    public bool SetBloomDisabled(bool disabled)
    {
        if (ProcessId is not int pid)
            return false;
        var name = $"Local\\HaloMCCToolbox.ReachBloom.{(disabled ? "Disable" : "Enable")}.{pid}";
        var handle = OpenEvent(EventModifyState, false, name);
        if (handle == IntPtr.Zero)
            return false;
        try { return SetEvent(handle); }
        finally { CloseHandle(handle); }
    }

    private bool TryOpenState(int pid)
    {
        DisposeState();
        _stateMap = OpenFileMapping(FileMapRead, false, $"Local\\HaloMCCToolbox.ReachBloom.State.{pid}");
        if (_stateMap == IntPtr.Zero)
            return false;
        _stateView = MapViewOfFile(_stateMap, FileMapRead, 0, 0, UIntPtr.Zero);
        if (_stateView != IntPtr.Zero)
            return true;
        CloseHandle(_stateMap);
        _stateMap = IntPtr.Zero;
        return false;
    }

    private static void InjectLibrary(Process process, string payloadPath)
    {
        var access = ProcessCreateThread | ProcessQueryInformation | ProcessVmOperation | ProcessVmWrite;
        var processHandle = OpenProcess(access, false, process.Id);
        if (processHandle == IntPtr.Zero)
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not open MCC. Run the Toolbox as Administrator.");

        IntPtr remotePath = IntPtr.Zero;
        IntPtr thread = IntPtr.Zero;
        try
        {
            var bytes = System.Text.Encoding.Unicode.GetBytes(payloadPath + '\0');
            remotePath = VirtualAllocEx(processHandle, IntPtr.Zero, (UIntPtr)bytes.Length, MemCommit | MemReserve, PageReadWrite);
            if (remotePath == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not allocate memory in MCC.");
            if (!WriteProcessMemory(processHandle, remotePath, bytes, (UIntPtr)bytes.Length, out _))
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not write the Reach hook path into MCC.");

            var kernel32 = GetModuleHandle("kernel32.dll");
            var loadLibrary = GetProcAddress(kernel32, "LoadLibraryW");
            if (loadLibrary == IntPtr.Zero)
                throw new InvalidOperationException("Could not resolve LoadLibraryW.");
            thread = CreateRemoteThread(processHandle, IntPtr.Zero, UIntPtr.Zero, loadLibrary, remotePath, 0, out _);
            if (thread == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error(), "MCC rejected the Reach hook.");
            if (WaitForSingleObject(thread, 5000) != WaitObject0)
                throw new TimeoutException("MCC did not finish loading the Reach hook.");
            if (!GetExitCodeThread(thread, out var result) || result == 0)
                throw new InvalidOperationException("Windows could not load the Reach lighting hook into MCC.");
        }
        finally
        {
            if (thread != IntPtr.Zero) CloseHandle(thread);
            if (remotePath != IntPtr.Zero) VirtualFreeEx(processHandle, remotePath, UIntPtr.Zero, MemRelease);
            CloseHandle(processHandle);
        }
    }

    public void Dispose()
    {
        if (IsConnected)
            SetBloomDisabled(false);
        DisposeHandles();
    }

    private void DisposeHandles()
    {
        DisposeState();
        _process = null;
    }

    private void DisposeState()
    {
        if (_stateView != IntPtr.Zero) UnmapViewOfFile(_stateView);
        if (_stateMap != IntPtr.Zero) CloseHandle(_stateMap);
        _stateView = IntPtr.Zero;
        _stateMap = IntPtr.Zero;
    }

    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, int processId);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr VirtualAllocEx(IntPtr process, IntPtr address, UIntPtr size, uint allocationType, uint protect);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool VirtualFreeEx(IntPtr process, IntPtr address, UIntPtr size, uint freeType);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool WriteProcessMemory(IntPtr process, IntPtr address, byte[] buffer, UIntPtr size, out UIntPtr written);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr CreateRemoteThread(IntPtr process, IntPtr attributes, UIntPtr stackSize, IntPtr start, IntPtr parameter, uint flags, out uint threadId);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetExitCodeThread(IntPtr thread, out uint exitCode);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string moduleName);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)] private static extern IntPtr GetProcAddress(IntPtr module, string procedureName);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] private static extern IntPtr OpenFileMapping(uint access, bool inherit, string name);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr MapViewOfFile(IntPtr mapping, uint access, uint offsetHigh, uint offsetLow, UIntPtr bytes);
    [DllImport("kernel32.dll")] private static extern bool UnmapViewOfFile(IntPtr address);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] private static extern IntPtr OpenEvent(uint access, bool inherit, string name);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetEvent(IntPtr handle);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
}

internal readonly record struct ReachBloomState(bool InstallOk, bool ShaderFound, bool BloomDisabled, long BlockedDraws);
