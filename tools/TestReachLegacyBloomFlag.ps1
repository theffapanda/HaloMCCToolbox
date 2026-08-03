param([int]$Seconds = 20)

$ErrorActionPreference = 'Stop'
$log = 'C:\tmp\reach_legacy_bloom_test.txt'
$expectedCurrent = [byte[]](0x21, 0x00, 0x20, 0x00)
$legacyValue = [byte[]](0x21, 0x00, 0x00, 0x00)
$descriptorRva = 0xBB98C0L
$flagsOffset = 8L

if (-not ('ReachBloomTest.Native' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
namespace ReachBloomTest {
    public static class Native {
        [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll", SetLastError=true)] public static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] data, int size, out IntPtr read);
        [DllImport("kernel32.dll", SetLastError=true)] public static extern bool WriteProcessMemory(IntPtr process, IntPtr address, byte[] data, int size, out IntPtr written);
        [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
    }
}
'@
}

function Log([string]$message) {
    $line = "$(Get-Date -Format o) $message"
    Add-Content -LiteralPath $log -Value $line
    Write-Output $line
}

Remove-Item -LiteralPath $log -Force -ErrorAction SilentlyContinue
$handle = [IntPtr]::Zero
$original = $null
$patched = $false
try {
    $process = Get-Process MCC-Win64-Shipping,MCCWinStore-Win64-Shipping -ErrorAction SilentlyContinue |
        Sort-Object StartTime -Descending | Select-Object -First 1
    if (-not $process) { throw 'MCC is not running.' }
    $module = $process.Modules | Where-Object ModuleName -IEq 'haloreach.dll' | Select-Object -First 1
    if (-not $module) { throw 'haloreach.dll is not loaded.' }
    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo($module.FileName).FileVersion
    if ($version -ne '1.3528.0.0') { throw "Unsupported Reach build $version." }

    $handle = [ReachBloomTest.Native]::OpenProcess(0x438, $false, [uint32]$process.Id)
    if ($handle -eq [IntPtr]::Zero) { throw "OpenProcess failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }

    $descriptor = $module.BaseAddress.ToInt64() + $descriptorRva
    $window = New-Object byte[] 16
    $read = [IntPtr]::Zero
    if (-not [ReachBloomTest.Native]::ReadProcessMemory($handle, [IntPtr]($descriptor - 16), $window, $window.Length, [ref]$read)) {
        throw "Descriptor validation read failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
    }
    if ([BitConverter]::ToUInt64($window, 0) -ne 0 -or [BitConverter]::ToUInt64($window, 8) -ne 0x13) {
        throw "Bloom descriptor identity mismatch at RVA 0x$('{0:X}' -f $descriptorRva)."
    }

    $flagsAddress = $descriptor + $flagsOffset
    $original = New-Object byte[] 4
    if (-not [ReachBloomTest.Native]::ReadProcessMemory($handle, [IntPtr]$flagsAddress, $original, 4, [ref]$read)) {
        throw "Flag read failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
    }
    if ([BitConverter]::ToString($original) -ne [BitConverter]::ToString($expectedCurrent)) {
        throw "Expected bloom flags 21-00-20-00, found $([BitConverter]::ToString($original))."
    }

    $written = [IntPtr]::Zero
    if (-not [ReachBloomTest.Native]::WriteProcessMemory($handle, [IntPtr]$flagsAddress, $legacyValue, 4, [ref]$written) -or $written.ToInt64() -ne 4) {
        throw "Legacy flag write failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
    }
    $patched = $true
    Log "APPLIED pid=$($process.Id) module=0x$('{0:X}' -f $module.BaseAddress.ToInt64()) address=0x$('{0:X}' -f $flagsAddress) 00200021 -> 00000021"
    Log "Observe bloom and image softness for $Seconds seconds. Automatic restore is armed."
    for ($remaining = $Seconds; $remaining -gt 0; $remaining--) {
        Start-Sleep -Seconds 1
        if ($process.HasExited) { break }
    }
}
catch {
    Log "ERROR $($_.Exception.Message)"
    exit 1
}
finally {
    if ($patched -and $handle -ne [IntPtr]::Zero -and $original) {
        $written = [IntPtr]::Zero
        if ([ReachBloomTest.Native]::WriteProcessMemory($handle, [IntPtr]$flagsAddress, $original, 4, [ref]$written) -and $written.ToInt64() -eq 4) {
            Log "RESTORED original bloom flags $([BitConverter]::ToString($original))."
        } else {
            Log "RESTORE FAILED Win32=$([Runtime.InteropServices.Marshal]::GetLastWin32Error()). Exit Reach before continuing tests."
        }
    }
    if ($handle -ne [IntPtr]::Zero) { [void][ReachBloomTest.Native]::CloseHandle($handle) }
}
