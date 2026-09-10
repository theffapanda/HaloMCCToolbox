$ErrorActionPreference = 'Continue'
$probe = Join-Path $PSScriptRoot 'ReachLightingProbe.ps1'
$output = 'C:\tmp\reach_lighting_probe_output.txt'
& $probe -WaitSeconds 15 *>&1 | Out-File -LiteralPath $output -Encoding utf8
exit $LASTEXITCODE
