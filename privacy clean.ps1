# Clear Activity History
Write-Host "Clearing Activity History..."
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist" /f

# Clear Location History
Write-Host "Clearing Location History..."
Get-Process LocationNotificationWnd | Stop-Process
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs" /f

# Clear Diagnostic Data
Write-Host "Deleting Diagnostic Data..."
wevtutil cl Microsoft-Windows-Diagnostics-Performance/Operational

# Run Disk Cleanup for Temp Files
Write-Host "Cleaning Temporary Files..."
Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait