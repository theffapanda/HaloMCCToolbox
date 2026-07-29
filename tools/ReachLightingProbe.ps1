param(
    [int]$WaitSeconds = 60,
    [string[]]$Names = @(
        'render_bloom',
        'render_bloom_source',
        'render_postprocess_exposure',
        'render_autoexposure_enable',
        'render_exposure',
        'render_postprocess',
        'render_downsample'
    )
)

$ErrorActionPreference = 'Stop'

if (-not ('ReachProbe.Native' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ReachProbe {
    public static class Native {
        public const uint TH32CS_SNAPMODULE = 0x00000008;
        public const uint TH32CS_SNAPMODULE32 = 0x00000010;
        public const uint PROCESS_QUERY_INFORMATION = 0x0400;
        public const uint PROCESS_VM_READ = 0x0010;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct MODULEENTRY32 {
            public uint dwSize;
            public uint th32ModuleID;
            public uint th32ProcessID;
            public uint GlblcntUsage;
            public uint ProccntUsage;
            public IntPtr modBaseAddr;
            public uint modBaseSize;
            public IntPtr hModule;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string szModule;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szExePath;
        }

        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint processId);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        public static extern bool Module32FirstW(IntPtr snapshot, ref MODULEENTRY32 entry);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        public static extern bool Module32NextW(IntPtr snapshot, ref MODULEENTRY32 entry);
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern IntPtr OpenProcess(uint access, bool inherit, uint processId);
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] buffer, int size, out IntPtr read);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr handle);
    }
}
'@
}

function Get-ReachModule {
    param([int]$ProcessId)
    $snapshot = [ReachProbe.Native]::CreateToolhelp32Snapshot(
        [ReachProbe.Native]::TH32CS_SNAPMODULE -bor [ReachProbe.Native]::TH32CS_SNAPMODULE32,
        [uint32]$ProcessId)
    if ($snapshot -eq [IntPtr]::Zero -or $snapshot.ToInt64() -eq -1) { return $null }
    try {
        $entry = New-Object ReachProbe.Native+MODULEENTRY32
        $entry.dwSize = [Runtime.InteropServices.Marshal]::SizeOf($entry)
        if (-not [ReachProbe.Native]::Module32FirstW($snapshot, [ref]$entry)) { return $null }
        do {
            if ($entry.szModule -ieq 'haloreach.dll') {
                return [pscustomobject]@{
                    Base = $entry.modBaseAddr.ToInt64()
                    Size = [int64]$entry.modBaseSize
                    Path = $entry.szExePath
                }
            }
            $entry.dwSize = [Runtime.InteropServices.Marshal]::SizeOf($entry)
        } while ([ReachProbe.Native]::Module32NextW($snapshot, [ref]$entry))
        return $null
    }
    finally { [void][ReachProbe.Native]::CloseHandle($snapshot) }
}

function Find-Pattern {
    param([byte[]]$Bytes, [byte[]]$Needle)
    $hits = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -le $Bytes.Length - $Needle.Length; $i++) {
        if ($Bytes[$i] -ne $Needle[0]) { continue }
        $match = $true
        for ($j = 1; $j -lt $Needle.Length; $j++) {
            if ($Bytes[$i + $j] -ne $Needle[$j]) { $match = $false; break }
        }
        if ($match) { $hits.Add($i) }
    }
    return $hits
}

$deadline = [DateTime]::UtcNow.AddSeconds($WaitSeconds)
$process = $null
$module = $null
do {
    $process = Get-Process MCC-Win64-Shipping,MCCWinStore-Win64-Shipping -ErrorAction SilentlyContinue |
        Sort-Object StartTime -Descending | Select-Object -First 1
    if ($process) { $module = Get-ReachModule -ProcessId $process.Id }
    if (-not $module) { Start-Sleep -Milliseconds 500 }
} while (-not $module -and [DateTime]::UtcNow -lt $deadline)

if (-not $process -or -not $module) {
    throw "haloreach.dll was not found in a running MCC process within $WaitSeconds seconds. Enter Reach gameplay first."
}

$access = [ReachProbe.Native]::PROCESS_QUERY_INFORMATION -bor [ReachProbe.Native]::PROCESS_VM_READ
$handle = [ReachProbe.Native]::OpenProcess($access, $false, [uint32]$process.Id)
if ($handle -eq [IntPtr]::Zero) {
    throw "OpenProcess(read-only) failed with Win32 error $([Runtime.InteropServices.Marshal]::GetLastWin32Error()). Run this probe at the same elevation as MCC."
}

try {
    $image = New-Object byte[] $module.Size
    $read = [IntPtr]::Zero
    if (-not [ReachProbe.Native]::ReadProcessMemory($handle, [IntPtr]$module.Base, $image, $image.Length, [ref]$read)) {
        throw "ReadProcessMemory failed with Win32 error $([Runtime.InteropServices.Marshal]::GetLastWin32Error())."
    }
    if ($read.ToInt64() -ne $image.Length) {
        throw "Short module read: $($read.ToInt64()) of $($image.Length) bytes."
    }

    $version = if (Test-Path -LiteralPath $module.Path) {
        [Diagnostics.FileVersionInfo]::GetVersionInfo($module.Path).FileVersion
    } else { '<path unavailable>' }
    "PROCESS pid=$($process.Id) name=$($process.ProcessName)"
    "MODULE  base=0x$('{0:X}' -f $module.Base) size=0x$('{0:X}' -f $module.Size) version=$version"
    "PATH    $($module.Path)"

    foreach ($name in $Names) {
        $needle = [Text.Encoding]::ASCII.GetBytes($name + [char]0)
        $stringHits = @(Find-Pattern -Bytes $image -Needle $needle)
        if ($stringHits.Count -ne 1) {
            "NAME    $name string_hits=$($stringHits.Count)"
            continue
        }

        $stringRva = [int64]$stringHits[0]
        $stringVa = $module.Base + $stringRva
        $pointerBytes = [BitConverter]::GetBytes([uint64]$stringVa)
        $pointerHits = @(Find-Pattern -Bytes $image -Needle $pointerBytes)
        "NAME    $name string_rva=0x$('{0:X}' -f $stringRva) pointer_hits=$($pointerHits.Count)"
        foreach ($recordRva in $pointerHits) {
            $type = if ($recordRva + 16 -le $image.Length) { [BitConverter]::ToUInt64($image, $recordRva + 8) } else { 0 }
            $slot = if ($recordRva + 24 -le $image.Length) { [BitConverter]::ToUInt64($image, $recordRva + 16) } else { 0 }
            "RECORD  rva=0x$('{0:X}' -f $recordRva) type=0x$('{0:X}' -f $type) slot=0x$('{0:X16}' -f $slot)"
        }
    }

    "RESULT  Read-only discovery complete. No process memory was changed."
}
finally {
    [void][ReachProbe.Native]::CloseHandle($handle)
}
