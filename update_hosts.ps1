# Download hostscompress-x64.exe
$exeUrl = 'https://github.com/Lateralus138/hosts-compress-windows/releases/latest/download/hostscompress-x64.exe'
$exePath = '.\hostscompress-x64.exe'
Invoke-WebRequest -Uri $exeUrl -OutFile $exePath

# Download hosts file
$url = 'https://raw.githubusercontent.com/StevenBlack/hosts/master/hosts'
$response = Invoke-WebRequest -Uri $url -OutFile '.\hosts'

# Wait for download to complete
while ($response.IsCompleted -eq $false) {
    Start-Sleep -Seconds 1
}

# Compress hosts file
$compressedFilePath = '.\hosts'  # Adjust the path if needed
$compressionProcess = Start-Process -FilePath '.\hostscompress-x64.exe' -ArgumentList '/c', '9', '/d', '/o', $compressedFilePath -PassThru -Wait

# Wait for compression process to complete
$compressionProcess.WaitForExit()

# Read compressed content
$text = Get-Content $compressedFilePath -Raw

# Replace system's hosts file
$hostsPath = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
Set-Content -Path $hostsPath -Value $text -Force

# Optionally, remove the uncompressed file after replacing the hosts file
Remove-Item $compressedFilePath -Force