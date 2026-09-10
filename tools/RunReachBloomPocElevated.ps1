param([string]$DllPath)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$injector = Join-Path $PSScriptRoot 'InjectReachBloomPoc.ps1'
$dll = if ($DllPath) { (Resolve-Path -LiteralPath $DllPath).Path } else {
    Join-Path $root '.codex_tmp\ReachBloomPoc\x64\Release\ReachBloomPoc.dll'
}

if (-not (Test-Path -LiteralPath $dll)) {
    throw "POC DLL not found: $dll"
}

$argumentList = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', ('"' + $injector + '"'),
    '-DllPath', ('"' + $dll + '"')
)

Start-Process -FilePath 'powershell.exe' -Verb RunAs -Wait -ArgumentList $argumentList
