param(
    [Parameter(Mandatory = $true)]
    [string]$DllPath
)

$ErrorActionPreference = 'Stop'
$DllPath = (Resolve-Path -LiteralPath $DllPath).Path

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'This injector must be run as administrator.'
}

$source = @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

public static class ReachBloomInjector
{
    [Flags]
    enum ProcessAccess : uint
    {
        CreateThread = 0x0002,
        QueryInformation = 0x0400,
        VmOperation = 0x0008,
        VmWrite = 0x0020,
        VmRead = 0x0010
    }

    const uint MEM_COMMIT = 0x1000;
    const uint MEM_RESERVE = 0x2000;
    const uint MEM_RELEASE = 0x8000;
    const uint PAGE_READWRITE = 0x04;
    const uint WAIT_OBJECT_0 = 0;

    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr OpenProcess(ProcessAccess access, bool inherit, int pid);
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr VirtualAllocEx(IntPtr process, IntPtr address, UIntPtr size, uint allocationType, uint protect);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool VirtualFreeEx(IntPtr process, IntPtr address, UIntPtr size, uint freeType);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool WriteProcessMemory(IntPtr process, IntPtr address, byte[] buffer, UIntPtr size, out UIntPtr written);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr GetModuleHandle(string name);
    [DllImport("kernel32.dll", CharSet=CharSet.Ansi, SetLastError=true)] static extern IntPtr GetProcAddress(IntPtr module, string name);
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr CreateRemoteThread(IntPtr process, IntPtr attrs, UIntPtr stackSize, IntPtr start, IntPtr parameter, uint flags, out uint threadId);
    [DllImport("kernel32.dll", SetLastError=true)] static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool GetExitCodeThread(IntPtr thread, out uint exitCode);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);

    static void Check(bool ok, string operation)
    {
        if (!ok) throw new Win32Exception(Marshal.GetLastWin32Error(), operation);
    }

    public static void Inject(int pid, string dllPath)
    {
        byte[] path = Encoding.Unicode.GetBytes(dllPath + "\0");
        IntPtr process = OpenProcess(ProcessAccess.CreateThread | ProcessAccess.QueryInformation |
            ProcessAccess.VmOperation | ProcessAccess.VmWrite | ProcessAccess.VmRead, false, pid);
        Check(process != IntPtr.Zero, "OpenProcess");
        IntPtr remote = IntPtr.Zero;
        IntPtr thread = IntPtr.Zero;
        try
        {
            remote = VirtualAllocEx(process, IntPtr.Zero, (UIntPtr)path.Length,
                MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
            Check(remote != IntPtr.Zero, "VirtualAllocEx");
            UIntPtr written;
            Check(WriteProcessMemory(process, remote, path, (UIntPtr)path.Length, out written) &&
                written.ToUInt64() == (ulong)path.Length, "WriteProcessMemory");
            IntPtr kernel32 = GetModuleHandle("kernel32.dll");
            Check(kernel32 != IntPtr.Zero, "GetModuleHandle(kernel32)");
            IntPtr loadLibrary = GetProcAddress(kernel32, "LoadLibraryW");
            Check(loadLibrary != IntPtr.Zero, "GetProcAddress(LoadLibraryW)");
            uint threadId;
            thread = CreateRemoteThread(process, IntPtr.Zero, UIntPtr.Zero, loadLibrary, remote, 0, out threadId);
            Check(thread != IntPtr.Zero, "CreateRemoteThread");
            Check(WaitForSingleObject(thread, 15000) == WAIT_OBJECT_0, "WaitForSingleObject");
            uint result;
            Check(GetExitCodeThread(thread, out result) && result != 0, "LoadLibraryW in MCC");
        }
        finally
        {
            if (thread != IntPtr.Zero) CloseHandle(thread);
            if (remote != IntPtr.Zero) VirtualFreeEx(process, remote, UIntPtr.Zero, MEM_RELEASE);
            CloseHandle(process);
        }
    }
}
'@

Add-Type -TypeDefinition $source
$targets = @(Get-Process -Name 'mcc-win64-shipping' -ErrorAction SilentlyContinue)
if ($targets.Count -ne 1) {
    throw "Expected exactly one mcc-win64-shipping process; found $($targets.Count)."
}

$target = $targets[0]
Write-Host "Injecting bloom POC into MCC PID $($target.Id)..."
[ReachBloomInjector]::Inject($target.Id, $DllPath)
Start-Sleep -Milliseconds 1500
Write-Host 'Injection completed.'
if (Test-Path -LiteralPath 'C:\tmp\reach_bloom_hook.log') {
    Get-Content -LiteralPath 'C:\tmp\reach_bloom_hook.log' -Tail 20
}

