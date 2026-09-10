param([int]$ProcessId)

$renderDoc = 'C:\tmp\renderdoc144\RenderDoc_1.44_64\renderdoccmd.exe'
$log = 'C:\tmp\renderdoc_reach_inject.txt'
& $renderDoc inject "--PID=$ProcessId" -c 'C:\tmp\reach_bloom' *>&1 |
    Out-File -LiteralPath $log -Encoding utf8
exit $LASTEXITCODE
