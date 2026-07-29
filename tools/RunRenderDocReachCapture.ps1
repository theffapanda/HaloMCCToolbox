$renderDoc = 'C:\tmp\renderdoc144\RenderDoc_1.44_64\renderdoccmd.exe'
$game = 'C:\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64\mcc-win64-shipping.exe'
$working = Split-Path -Parent $game
$log = 'C:\tmp\renderdoc_reach_capture_launch.txt'
$existing = Get-Process MCC-Win64-Shipping,MCCWinStore-Win64-Shipping -ErrorAction SilentlyContinue
if ($existing) {
    $existing | Stop-Process -Force
    $existing | Wait-Process -Timeout 15 -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}
& $renderDoc capture --opt-hook-children --opt-capture-callstacks --opt-capture-callstacks-only-actions -d $working -c 'C:\tmp\reach_bloom_stack2' $game -no-eac *>&1 |
    Out-File -LiteralPath $log -Encoding utf8
exit $LASTEXITCODE
