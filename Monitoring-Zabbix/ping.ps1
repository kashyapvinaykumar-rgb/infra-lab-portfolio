$target = "8.8.8.8"
$logfile = "D:\Logs\ping_log.txt"

while ($true) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $result = Test-Connection -ComputerName $target -Count 1 -ErrorAction SilentlyContinue

    if ($result) {
        "$timestamp - Reply from $($result.Address) time=$($result.ResponseTime)ms" | Out-File -Append $logfile
    } else {
        "$timestamp - Request timed out" | Out-File -Append $logfile
    }

    Start-Sleep -Seconds 1
}