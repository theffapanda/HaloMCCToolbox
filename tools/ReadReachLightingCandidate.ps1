$ErrorActionPreference = 'Stop'
$output = 'C:\tmp\reach_lighting_candidate.txt'
try {
    if (-not ('ReachCandidate.Native' -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
namespace ReachCandidate {
    public static class Native {
        [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll", SetLastError=true)] public static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] data, int size, out IntPtr read);
        [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
    }
}
'@
    }
    $process = Get-Process MCC-Win64-Shipping,MCCWinStore-Win64-Shipping -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1
    if (-not $process) { throw 'MCC is not running.' }
    $module = $process.Modules | Where-Object ModuleName -IEq 'haloreach.dll' | Select-Object -First 1
    if (-not $module) { throw 'haloreach.dll is not loaded.' }
    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo($module.FileName).FileVersion
    if ($version -ne '1.3528.0.0') { throw "Unsupported Reach build $version." }
    $handle = [ReachCandidate.Native]::OpenProcess(0x410, $false, [uint32]$process.Id)
    if ($handle -eq [IntPtr]::Zero) { throw "OpenProcess failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
    try {
        $candidate = $module.BaseAddress.ToInt64() + 0x198F2C
        $start = $candidate - 0x2C
        $bytes = New-Object byte[] 0x80
        $read = [IntPtr]::Zero
        if (-not [ReachCandidate.Native]::ReadProcessMemory($handle, [IntPtr]$start, $bytes, $bytes.Length, [ref]$read)) {
            throw "ReadProcessMemory failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
        }
        $lines = @(
            "pid=$($process.Id)",
            "module_base=0x$('{0:X}' -f $module.BaseAddress.ToInt64())",
            "version=$version",
            "candidate=0x$('{0:X}' -f $candidate)",
            "candidate_bytes=$([BitConverter]::ToString($bytes,0x2C,16))",
            "window=$([BitConverter]::ToString($bytes))",
            'result=read-only; no process memory changed'
        )
        $lines | Out-File -LiteralPath $output -Encoding utf8
    }
    finally { [void][ReachCandidate.Native]::CloseHandle($handle) }
}
catch {
    $_ | Out-String | Out-File -LiteralPath $output -Encoding utf8
    exit 1
}
