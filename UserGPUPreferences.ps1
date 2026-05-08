# Define the GPU preference (2 = High Performance, 1 = Power Saving)
$pref = "GpuPreference=2;"
$registryPath = "HKCU:\Software\Microsoft\DirectX\UserGpuPreferences"

# Base directories to scan completely
$basePaths = @(
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)},
    "$env:LocalAppData\Programs"
)

# BLACKLIST ARRAY
$blacklistKeywords = @(
    "unins", "update", "setup", "crash", "dump", 
    "bug", "redist", "helper", "anticheat", "battleye", 
    "beservice", "vgtray", "dotnet", "antimalware", "defender"
)
$blacklistPattern = '(?i)(' + ($blacklistKeywords -join '|') + ')'

# Initialize counters
$addedCount = 0
$skippedConfiguredCount = 0
$skippedBlacklistCount = 0
$errorCount = 0

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "        GPU PREFERENCE MANAGER          " -ForegroundColor White
Write-Host "========================================" -ForegroundColor Magenta
Write-Host " [1] Scan and Add missing apps" -ForegroundColor Cyan
Write-Host " [2] WIPE existing rules, then Scan and Add" -ForegroundColor Yellow
Write-Host " [3] WIPE existing rules ONLY (Do not scan)" -ForegroundColor Red
Write-Host " [4] Exit" -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor Magenta

$choice = Read-Host -Prompt "Select an action (1-4)"

$runScan = $false
$wipeRules = $false

switch ($choice) {
    '1' { $runScan = $true }
    '2' { $wipeRules = $true; $runScan = $true }
    '3' { $wipeRules = $true }
    '4' { exit }
    default { 
        Write-Host "Invalid selection. Exiting." -ForegroundColor Red
        Start-Sleep -Seconds 2
        exit 
    }
}

# --- STEP 1: WIPE LOGIC ---
if ($wipeRules) {
    Write-Host "`nClearing existing GPU preferences..." -ForegroundColor Yellow
    if (Test-Path $registryPath) {
        Remove-Item -Path $registryPath -Recurse -Force
    }
    New-Item -Path $registryPath -Force | Out-Null
    Write-Host "[+] Registry rules successfully wiped." -ForegroundColor Green
}

# --- STEP 2: SCAN LOGIC ---
if ($runScan) {
    Write-Host "`nScanning directories deeply... This will take a moment." -ForegroundColor Cyan

    foreach ($baseDir in $basePaths) {
        if (Test-Path $baseDir) {
            try {
                $files = Get-ChildItem -Path $baseDir -Filter "*.exe" -Recurse -File -ErrorAction Ignore
                
                foreach ($file in $files) {
                    if ($file.Name -match $blacklistPattern) {
                        $skippedBlacklistCount++
                        continue
                    }

                    $exePath = $file.FullName
                    $existingEntry = Get-ItemProperty -Path $registryPath -Name $exePath -ErrorAction Ignore
                    
                    if ($null -eq $existingEntry) {
                        try {
                            New-ItemProperty -Path $registryPath -Name $exePath -Value $pref -PropertyType String -Force | Out-Null
                            $addedCount++
                            Write-Host "[+] Added: $($file.Name)" -ForegroundColor Green
                        }
                        catch {
                            $errorCount++
                            Write-Host "[!] Failed to add: $($file.Name)" -ForegroundColor Red
                        }
                    }
                    else {
                        $skippedConfiguredCount++
                    }
                }
            }
            catch {
                Write-Host "[!] Critical error scanning base directory: $baseDir" -ForegroundColor Red
            }
        }
    }

    # Print the final success message and summary ONLY if a scan was run
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "         GPU PREFERENCE SUMMARY         " -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Successfully Added : $addedCount" -ForegroundColor Green
    Write-Host "Already Configured : $skippedConfiguredCount" -ForegroundColor Yellow
    Write-Host "Blacklisted Apps   : $skippedBlacklistCount" -ForegroundColor DarkGray
    if ($errorCount -gt 0) {
        Write-Host "Errors Encountered : $errorCount" -ForegroundColor Red
    }
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Success! GPU preferences have been updated." -ForegroundColor Green
}
else {
    Write-Host "`nScan skipped as requested." -ForegroundColor DarkGray
}

Write-Host "Please restart your applications for the changes to take effect.`n" -ForegroundColor Cyan
Read-Host -Prompt "Press Enter to exit"