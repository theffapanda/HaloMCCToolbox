param([int]$ProcessId)
$output = 'C:\tmp\mcc_launch_info.txt'
$process = Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId"
if (-not $process) { 'PROCESS_NOT_FOUND' | Out-File $output; exit 1 }
@(
    "ExecutablePath=$($process.ExecutablePath)",
    "CommandLine=$($process.CommandLine)"
) | Out-File -LiteralPath $output -Encoding utf8
