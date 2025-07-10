<#
.SYNOPSIS
    PowerShell Driver Store Explorer (RAPR-like functionality)
.DESCRIPTION
    Enhanced version with dependency checking and stronger force removal options
    Now includes support for removing stubborn drivers and non-OEM INFs
.NOTES
    Version: 2.4
    Author: Your Name
    Last Updated: 2023-11-18
#>

$Error.Clear()

function Show-Menu {
    Clear-Host
    Write-Host "`nDriver Store Explorer (PowerShell Version)`n" -ForegroundColor Cyan
    Write-Host "1. List all drivers (Microsoft/System in " -NoNewline
    Write-Host "RED" -ForegroundColor Red -NoNewline
    Write-Host ")"
    Write-Host "2. Add driver to store"
    Write-Host "3. Save drivers list to CSV"
    Write-Host "4. Advanced driver removal options" -ForegroundColor Yellow
    Write-Host "0/q. Exit`n"
}

function Get-ProperDriverDate {
    param($driverDate)
    
    if ([string]::IsNullOrEmpty($driverDate)) {
        return "Unknown"
    }

    if ($driverDate -match '^(\d{8})') {
        $datePart = $Matches[1]
        try {
            return [datetime]::ParseExact($datePart, 'yyyyMMdd', $null).ToString("yyyy-MM-dd")
        }
        catch {
            return $datePart
        }
    }
    
    return $driverDate
}

function Get-DriverStoreDrivers {
    try {
        $allDrivers = Get-WmiObject Win32_PnPSignedDriver -ErrorAction Stop | 
                     Where-Object { $_.DeviceName -ne $null }
        
        if (-not $allDrivers) {
            Write-Host "No drivers found in driver store" -ForegroundColor Yellow
            return @()
        }

        $processedDrivers = foreach ($driver in $allDrivers) {
            $isSystem = $driver.Manufacturer -match "Microsoft|Standard" -or 
                       $driver.InfName -match "ms_|oem[01]\.inf|netwtw|netvwif|netrtwlan"
            
            [PSCustomObject]@{
                DeviceName      = if ($driver.DeviceName) { $driver.DeviceName.Trim() } else { "Unknown" }
                DriverVersion   = if ($driver.DriverVersion) { $driver.DriverVersion.Trim() } else { "Unknown" }
                Manufacturer    = if ($driver.Manufacturer) { $driver.Manufacturer.Trim() } else { "Unknown" }
                DriverDate      = Get-ProperDriverDate $driver.DriverDate
                InfName        = if ($driver.InfName) { $driver.InfName.Trim() } else { "Unknown" }
                OriginalInfName = if ($driver.InfName) { $driver.InfName.Split('\')[-1] } else { "Unknown" }
                IsSystem       = $isSystem
                DeviceID       = $driver.DeviceID
                HardwareID     = if ($driver.HardwareID) { $driver.HardwareID -join ", " } else { "Unknown" }
            }
        }

        return $processedDrivers
    }
    catch {
        Write-Host "Failed to enumerate drivers: $_" -ForegroundColor Red
        return @()
    }
}

function Show-ColorCodedDrivers {
    param(
        [Parameter(Mandatory=$true)]
        $Items,
        [string]$Title
    )
    
    Clear-Host
    Write-Host "`n$Title`n" -ForegroundColor Yellow
    Write-Host "Microsoft/System drivers shown in " -NoNewline
    Write-Host "RED" -ForegroundColor Red -NoNewline
    Write-Host ", third-party drivers in default color`n"
    
    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "No items to display" -ForegroundColor Yellow
        Wait-ForEnter
        return $null
    }

    $formatString = "{0,-4} {1,-30} {2,-15} {3,-20} {4,-12} {5}"
    
    Write-Host ($formatString -f "NUM", "DEVICE", "VERSION", "MANUFACTURER", "DATE", "INF NAME")
    Write-Host ("-" * 98)

    $index = 1
    foreach ($item in $Items) {
        $line = $formatString -f "[$index]",
            ($item.DeviceName -replace "`n|`r", ""),
            ($item.DriverVersion -replace "`n|`r", ""),
            ($item.Manufacturer -replace "`n|`r", ""),
            ($item.DriverDate -replace "`n|`r", ""),
            ($item.InfName -replace "`n|`r", "")

        if ($item.IsSystem) {
            Write-Host $line -ForegroundColor Red
        } else {
            Write-Host $line
        }
        $index++
    }

    Write-Host "`nOptions:"
    Write-Host "R # - Remove driver by number" -ForegroundColor Yellow
    Write-Host "F # - Force remove driver (advanced)" -ForegroundColor Red
    Write-Host "S - Save this list to CSV on Desktop"
    Write-Host "0/Q - Return to main menu"
    
    do {
        $choice = Read-Host "`nEnter your choice"
        
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -match '[\x00-\x1F]') {
            continue
        }
        
        switch -regex ($choice.ToLower()) {
            '^r\s?(\d+)$' {
                $num = [int]$Matches[1]
                if ($num -ge 1 -and $num -le $Items.Count) {
                    $driverToRemove = $Items[$num-1]
                    Remove-Driver -InfName $driverToRemove.OriginalInfName -DeviceID $driverToRemove.DeviceID
                }
                else {
                    Write-Host "Invalid driver number!" -ForegroundColor Red
                    Wait-ForEnter
                }
                return $Items
            }
            '^f\s?(\d+)$' {
                $num = [int]$Matches[1]
                if ($num -ge 1 -and $num -le $Items.Count) {
                    $driverToRemove = $Items[$num-1]
                    Remove-Driver -InfName $driverToRemove.OriginalInfName -Force -DeviceID $driverToRemove.DeviceID
                }
                else {
                    Write-Host "Invalid driver number!" -ForegroundColor Red
                    Wait-ForEnter
                }
                return $Items
            }
            '^s$' {
                Save-DriversToCSV -Drivers $Items
                return $Items
            }
            '^(0|q)$' {
                return $null
            }
            default {
                Write-Host "Invalid option!" -ForegroundColor Red
                Start-Sleep -Seconds 1
                break
            }
        }
    } while ($true)
}

function Save-DriversToCSV {
    param(
        [Parameter(Mandatory=$true)]
        $Drivers
    )
    
    try {
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $csvPath = Join-Path $desktopPath "DriverStore_$timestamp.csv"
        
        if ($Drivers -and $Drivers.Count -gt 0) {
            $Drivers | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nDrivers list saved to: $csvPath" -ForegroundColor Green
        }
        else {
            Write-Host "No drivers found to save" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error saving drivers list: $_" -ForegroundColor Red
    }
    
    Wait-ForEnter
}

function Add-Driver {
    $path = Read-Host "`nEnter full path to .inf file"
    if (-not (Test-Path $path -PathType Leaf)) {
        Write-Host "File not found!" -ForegroundColor Red
        Wait-ForEnter
        return
    }
    
    try {
        Write-Host "`nAdding driver..." -ForegroundColor Yellow
        $result = pnputil /add-driver $path 2>&1
        $result | Out-Host
    }
    catch {
        Write-Host "Error adding driver: $_" -ForegroundColor Red
    }
    
    Wait-ForEnter
}

function Remove-Driver {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InfName,
        [switch]$Force,
        [string]$DeviceID
    )
    
    try {
        Write-Host "`nRemoving driver $InfName..." -ForegroundColor Yellow
        
        if ($InfName -match "ms_|oem[01]\.inf") {
            Write-Host "WARNING: This appears to be a system driver!" -ForegroundColor Red
            Write-Host "Removing system drivers can cause serious system instability." -ForegroundColor Red
            $confirm = Read-Host "Are you absolutely sure you want to continue? (y/n)"
            if ($confirm -ne 'y') {
                Write-Host "Aborted driver removal." -ForegroundColor Yellow
                Wait-ForEnter
                return
            }
        }

        # Initialize removal steps
        $removalSteps = @(
            "Standard pnputil removal",
            "Force pnputil removal",
            "Device removal (alternative methods)",
            "Filesystem cleanup"
        )
        $currentStep = 0

        # 1. Try standard removal first
        $currentStep++
        Write-Host "`n[$currentStep/$($removalSteps.Count)] $($removalSteps[$currentStep-1])..." -ForegroundColor Cyan
        $result = pnputil /delete-driver $InfName 2>&1
        $result | Out-Host
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Driver removed successfully" -ForegroundColor Green
            Wait-ForEnter
            return
        }

        # 2. Try force removal if requested or standard failed
        $currentStep++
        Write-Host "`n[$currentStep/$($removalSteps.Count)] $($removalSteps[$currentStep-1])..." -ForegroundColor Cyan
        $result = pnputil /delete-driver $InfName /force 2>&1
        $result | Out-Host
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Driver force-removed successfully" -ForegroundColor Green
            Wait-ForEnter
            return
        }

        # 3. Try device removal if DeviceID provided
        if ($DeviceID) {
            $currentStep++
            Write-Host "`n[$currentStep/$($removalSteps.Count)] $($removalSteps[$currentStep-1])..." -ForegroundColor Cyan
            try {
                # Try modern method first
                if (Get-Command Remove-PnpDevice -ErrorAction SilentlyContinue) {
                    $device = Get-PnpDevice | Where-Object { $_.InstanceId -eq $DeviceID }
                    if ($device) {
                        $device | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
                        $device | Remove-PnpDevice -Confirm:$false -ErrorAction Stop
                        Write-Host "Device removed successfully (modern method)" -ForegroundColor Green
                    }
                } 
                # Fallback to devcon.exe
                elseif (Test-Path "$env:SystemRoot\System32\devcon.exe") {
                    Write-Host "Using devcon.exe for removal..." -ForegroundColor Yellow
                    & "$env:SystemRoot\System32\devcon.exe" remove "@$DeviceID" 2>&1 | Out-Host
                }
                # Final fallback to WMI
                else {
                    Write-Host "Using WMI for removal..." -ForegroundColor Yellow
                    $device = Get-WmiObject Win32_PnPEntity | Where-Object { $_.DeviceID -eq $DeviceID }
                    if ($device) {
                        $device | ForEach-Object { $_.Disable() }
                        $device | ForEach-Object { $_.Uninstall() }
                        Write-Host "Device removal attempted via WMI" -ForegroundColor Yellow
                    }
                }
            } catch {
                Write-Host "Device removal failed: $_" -ForegroundColor Red
            }
        }

        # 4. Final filesystem cleanup
        $currentStep++
        Write-Host "`n[$currentStep/$($removalSteps.Count)] $($removalSteps[$currentStep-1])..." -ForegroundColor Cyan
        $driverStorePath = Join-Path $env:SystemRoot "System32\DriverStore\FileRepository"
        $infStubPath = Join-Path $env:SystemRoot "INF\$InfName"
        
        # Remove from DriverStore
        Get-ChildItem $driverStorePath -Recurse -Filter "*$InfName*" | ForEach-Object {
            try {
                Remove-Item $_.FullName -Force -Recurse -ErrorAction Stop
                Write-Host "Removed: $($_.FullName)" -ForegroundColor Yellow
            } catch {
                Write-Host "Failed to remove $($_.FullName): $_" -ForegroundColor Red
            }
        }
        
        # Remove INF stub
        if (Test-Path $infStubPath) {
            try {
                Remove-Item $infStubPath -Force -ErrorAction Stop
                Write-Host "Removed INF stub: $infStubPath" -ForegroundColor Yellow
            } catch {
                Write-Host "Failed to remove INF stub: $_" -ForegroundColor Red
            }
        }

        Write-Host "`nRemoval process completed. Some components may require reboot to fully remove." -ForegroundColor Yellow
        Write-Host "If driver persists, try these additional steps:" -ForegroundColor Cyan
        Write-Host "1. Reboot into Safe Mode and run this tool again"
        Write-Host "2. Use: pnputil /delete-driver $InfName /force /reboot"
        Write-Host "3. Manually delete from: C:\Windows\System32\DriverStore\FileRepository"
    }
    catch {
        Write-Host "Error removing driver: $_" -ForegroundColor Red
    }
    
    Wait-ForEnter
}

function Wait-ForEnter {
    Write-Host "`nPress Enter to continue..." -ForegroundColor Gray
    do {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        if ($key.VirtualKeyCode -eq 13) {  # 13 is Enter key
            return
        }
    } while ($true)
}

# Check for admin rights
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This tool requires Administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    Wait-ForEnter
    exit
}

# Main execution loop
while ($true) {
    try {
        Show-Menu
        $choice = Read-Host "`nSelect an option"
        
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -match '[\x00-\x1F]') {
            continue
        }
        
        switch ($choice.ToLower()) {
            '1' { 
                $drivers = Get-DriverStoreDrivers
                if ($drivers.Count -gt 0) {
                    do {
                        $shouldExit = $false
                        $drivers = Show-ColorCodedDrivers -Items $drivers -Title "All Drivers in Store (Microsoft/System in RED)"
                        if ($drivers -eq $null) { $shouldExit = $true }
                    } while (-not $shouldExit)
                } else {
                    Wait-ForEnter
                }
            }
            '2' { 
                Add-Driver 
            }
            '3' { 
                $drivers = Get-DriverStoreDrivers
                Save-DriversToCSV -Drivers $drivers
            }
            '4' {
                Write-Host "`nAdvanced Driver Removal Options" -ForegroundColor Yellow
                Write-Host "1. Remove driver by INF name"
                Write-Host "2. Remove all drivers by manufacturer"
                Write-Host "0. Back to main menu"
                $advChoice = Read-Host "`nSelect an option"
                
                switch ($advChoice) {
                    '1' {
                        $infName = Read-Host "Enter INF name to remove (e.g., printqueue.inf)"
                        if ($infName) {
                            Remove-Driver -InfName $infName -Force
                        }
                    }
                    '2' {
                        $manufacturer = Read-Host "Enter manufacturer name to remove (e.g., RustDesk)"
                        if ($manufacturer) {
                            $drivers = Get-DriverStoreDrivers | Where-Object { $_.Manufacturer -match $manufacturer }
                            if ($drivers) {
                                $drivers | ForEach-Object {
                                    Write-Host "Removing $($_.DeviceName) ($($_.OriginalInfName))..." -ForegroundColor Yellow
                                    Remove-Driver -InfName $_.OriginalInfName -Force -DeviceID $_.DeviceID
                                }
                            } else {
                                Write-Host "No drivers found for manufacturer: $manufacturer" -ForegroundColor Yellow
                                Wait-ForEnter
                            }
                        }
                    }
                }
            }
            { $_ -in '0','q' } { exit }
            default {
                Write-Host "Invalid option!" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    }
    catch {
        Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
        Wait-ForEnter
    }
}