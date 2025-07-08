<#
.SYNOPSIS
    PowerShell Driver Store Explorer (RAPR-like functionality)
.DESCRIPTION
    Enhanced version with dependency checking and stronger force removal options
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

    # Define column widths and format string
    $formatString = "{0,-4} {1,-30} {2,-15} {3,-20} {4,-12} {5}"
    
    # Display header
    Write-Host ($formatString -f "NUM", "DEVICE", "VERSION", "MANUFACTURER", "DATE", "INF NAME")
    Write-Host ("-" * 98)

    # Display numbered list with color coding
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
    Write-Host "R # - Remove driver by number (DANGEROUS!)" -ForegroundColor Red
    Write-Host "S - Save this list to CSV on Desktop"
    Write-Host "0/Q - Return to main menu"
    
    # Safe input handling that ignores Ctrl keys
    do {
        $choice = Read-Host "`nEnter your choice"
        
        # Ignore empty input or control characters
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -match '[\x00-\x1F]') {
            continue
        }
        
        switch -regex ($choice.ToLower()) {
            '^r\s?(\d+)$' {
                $num = [int]$Matches[1]
                if ($num -ge 1 -and $num -le $Items.Count) {
                    $driverToRemove = $Items[$num-1]
                    Write-Host "`nYou are about to remove:" -ForegroundColor Yellow
                    Write-Host "Device:      $($driverToRemove.DeviceName)"
                    Write-Host "Manufacturer: $($driverToRemove.Manufacturer)"
                    Write-Host "Driver:      $($driverToRemove.OriginalInfName)`n"
                    
                    Write-Host "WARNING: Removing drivers can cause system instability!" -ForegroundColor Red
                    Write-Host "Only remove drivers you are absolutely sure about.`n" -ForegroundColor Yellow
                    
                    $confirmation = Read-Host "Are you sure you want to remove this driver? (y/n)"
                    if ($confirmation -eq 'y') {
                        $force = Read-Host "Force removal (may be needed if driver is in use)? (y/n)"
                        Remove-Driver -InfName $driverToRemove.OriginalInfName -Force:($force -eq 'y') -DeviceID $driverToRemove.DeviceID
                    }
                    else {
                        Write-Host "Driver removal cancelled" -ForegroundColor Yellow
                        Wait-ForEnter
                    }
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
                break  # Continue the loop for new input
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
        if ($Force) {
            # First try normal force removal
            $result = pnputil /delete-driver $InfName /force 2>&1
            $result | Out-Host
            
            # If that fails, try disabling devices first
            if ($result -like "*devices are presently installed*") {
                Write-Host "`nAttempting to disable dependent devices first..." -ForegroundColor Yellow
                
                # Get devices using this driver
                $devices = @()
                if ($DeviceID) {
                    $devices += Get-WmiObject Win32_PnPSignedDriver | 
                              Where-Object { $_.InfName -like "*$InfName*" -or $_.DeviceID -eq $DeviceID }
                }
                else {
                    $devices += Get-WmiObject Win32_PnPSignedDriver | 
                              Where-Object { $_.InfName -like "*$InfName*" }
                }
                
                if ($devices) {
                    Write-Host "`nDevices using this driver:" -ForegroundColor Yellow
                    $devices | ForEach-Object {
                        $hardwareIDs = if ($_.HardwareID) { $_.HardwareID -join ", " } else { "Unknown" }
                        Write-Host "- $($_.DeviceName) (HardwareID: $hardwareIDs)"
                    }
                    
                    Write-Host "`nAttempting to disable these devices..." -ForegroundColor Yellow
                    $disabledDevices = @()
                    foreach ($device in $devices) {
                        try {
                            $dev = Get-PnpDevice | Where-Object { $_.InstanceID -eq $device.DeviceID }
                            if ($dev -and $dev.Status -eq "OK") {
                                $dev | Disable-PnpDevice -Confirm:$false -ErrorAction Stop
                                Write-Host "Disabled: $($device.DeviceName)" -ForegroundColor Yellow
                                $disabledDevices += $device
                            }
                        } catch {
                            Write-Host "Error disabling $($device.DeviceName): $_" -ForegroundColor Red
                        }
                    }
                    
                    # Try removal again after disabling
                    if ($disabledDevices) {
                        Write-Host "`nRetrying driver removal..." -ForegroundColor Yellow
                        $result = pnputil /delete-driver $InfName /force 2>&1
                        $result | Out-Host
                    }
                    
                    # Re-enable devices if removal still failed
                    if ($result -like "*devices are presently installed*" -or -not $disabledDevices) {
                        Write-Host "`nRe-enabling devices..." -ForegroundColor Yellow
                        foreach ($device in $disabledDevices) {
                            try {
                                $dev = Get-PnpDevice | Where-Object { $_.InstanceID -eq $device.DeviceID }
                                if ($dev) {
                                    $dev | Enable-PnpDevice -Confirm:$false
                                    Write-Host "Re-enabled: $($device.DeviceName)" -ForegroundColor Yellow
                                }
                            } catch {
                                Write-Host "Error re-enabling $($device.DeviceName): $_" -ForegroundColor Red
                            }
                        }
                        Write-Host "`nCould not remove driver. Devices must be uninstalled first." -ForegroundColor Red
                        Write-Host "Try these steps:" -ForegroundColor Yellow
                        Write-Host "1. Open Device Manager"
                        Write-Host "2. Find and uninstall all devices using this driver"
                        Write-Host "3. Check 'Delete the driver software for this device' when uninstalling"
                        Write-Host "4. Try removing the driver again"
                        Write-Host "Alternatively, reboot into Safe Mode and try removing the driver." -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "No dependent devices found, but driver still in use." -ForegroundColor Red
                    Write-Host "Try these steps:" -ForegroundColor Yellow
                    Write-Host "1. Reboot into Safe Mode"
                    Write-Host "2. Run this tool again to remove the driver"
                    Write-Host "3. Or use the command: pnputil /delete-driver $InfName /force" -ForegroundColor Cyan
                }
            }
        }
        else {
            $result = pnputil /delete-driver $InfName 2>&1
            $result | Out-Host
        }
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
        # Only accept Enter key (13 is Enter key code)
        if ($key.VirtualKeyCode -eq 13) {
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
        
        # Ignore control characters
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