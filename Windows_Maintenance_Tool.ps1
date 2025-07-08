# ===== ADMIN CHECK =====
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrator privileges."
    Write-Host "Requesting elevation now ..."
    Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Pause-Menu {
    Write-Host
    Read-Host "Press ENTER to return to menu"
}

function Show-Menu {
    Clear-Host
    Write-Host "====================================================="
    Write-Host "         WINDOWS MAINTENANCE TOOL V3.3.0 - By Lil_Batti & Chaython"
    Write-Host "====================================================="
    Write-Host
    Write-Host "     === WINDOWS UPDATES ==="
    Write-Host "  [1]  Update Windows Apps / Programs (Winget upgrade)"
    Write-Host
    Write-Host "     === SYSTEM HEALTH CHECKS ==="
    Write-Host "  [2]  Scan for corrupt files (SFC /scannow) [Admin]"
    Write-Host "  [3]  Windows CheckHealth (DISM) [Admin]"
    Write-Host "  [4]  Restore Windows Health (DISM /RestoreHealth) [Admin]"
    Write-Host
    Write-Host "     === NETWORK TOOLS ==="
    Write-Host "  [5]  DNS Options (Flush/Set/Reset, IPv4/IPv6, DoH)"
    Write-Host "  [6]  Show network information (ipconfig /all)"
    Write-Host "  [7]  Restart Wi-Fi Adapters"
    Write-Host "  [8]  Network Repair - Automatic Troubleshooter"
    Write-Host
    Write-Host "     === CLEANUP & OPTIMIZATION ==="
    Write-Host "  [9]  Disk Cleanup (cleanmgr)"
    Write-Host " [10]  Run Advanced Error Scan (CHKDSK) [Admin]"
    Write-Host " [11]  Perform System Optimization (Delete Temporary Files)"
    Write-Host " [12]  Advanced Registry Cleanup-Optimization"
    Write-Host " [13]  Optimize SSDs (ReTrim)"
    Write-Host " [14]  Task Management (Scheduled Tasks) [Admin]"
    Write-Host
    Write-Host "     === SUPPORT ==="
    Write-Host " [15]  Contact and Support information (Discord)"
    Write-Host
    Write-Host "     === UTILITIES & EXTRAS ==="
    Write-Host " [20]  Show installed drivers"
    Write-Host " [21]  Windows Update Repair Tool"
    Write-Host " [22]  Generate Full System Report"
    Write-Host " [23]  Windows Update Utility & Service Reset"
    Write-Host " [24]  View Network Routing Table [Advanced]"
    Write-Host
    Write-Host " [0]  EXIT"
    Write-Host "------------------------------------------------------"
}

function Choice-1 {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "    Windows Update (via Winget)"
    Write-Host "==============================================="

    # Check if Winget is installed
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "Winget is not installed. Attempting to install it automatically..."
        Write-Host
        
        try {
            # Method 1: Try installing via Microsoft Store (App Installer)
            Write-Host "Installing Winget via Microsoft Store..."
            $result = Start-Process "ms-windows-store://pdp/?productid=9NBLGGH4NNS1" -Wait -PassThru
            
            if ($result.ExitCode -eq 0) {
                Write-Host "Microsoft Store opened successfully. Please complete the installation."
                Write-Host "After installation, restart this tool to use Winget features."
                Pause-Menu
                return
            } else {
                # Method 2: Alternative direct download if Store method fails
                Write-Host "Microsoft Store method failed, trying direct download..."
                $wingetUrl = "https://aka.ms/getwinget"
                $installerPath = "$env:TEMP\winget-cli.msixbundle"
                
                # Download the installer
                try {
                    Invoke-WebRequest -Uri $wingetUrl -OutFile $installerPath -ErrorAction Stop
                } catch {
                    Write-Host "Failed to download Winget installer: $_"
                    Pause-Menu
                    return
                }
                
                # Install Winget
                try {
                    Add-AppxPackage -Path $installerPath -ErrorAction Stop
                } catch {
                    Write-Host "Failed to install Winget: $_"
                    Pause-Menu
                    return
                }
                
                # Verify installation
                if (Get-Command winget -ErrorAction SilentlyContinue) {
                    Write-Host "Winget installed successfully!"
                    Start-Sleep -Seconds 2
                } else {
                    Write-Host "Installation failed. Please install manually from Microsoft Store."
                    Pause-Menu
                    return
                }
            }
        } catch {
            Write-Host "Failed to install Winget automatically. Error: $_"
            Write-Host "Please install 'App Installer' from Microsoft Store manually."
            Pause-Menu
            return
        }
    }

    # Main Winget functionality
    Write-Host "Listing available upgrades..."
    Write-Host
    winget upgrade --include-unknown
    Write-Host
    
    while ($true) {
        Write-Host "==============================================="
        Write-Host "Options:"
        Write-Host "[1] Upgrade all packages"
        Write-Host "[2] Upgrade selected packages"
        Write-Host "[0] Cancel"
        Write-Host
        $upopt = Read-Host "Choose an option"
        $upopt = $upopt.Trim()
        switch ($upopt) {
            "0" {
                Write-Host "Cancelled. Returning to menu..."
                Start-Sleep -Seconds 1
                return
            }
            "1" {
                Write-Host "Running full upgrade..."
                try {
                    $upgradeOutput = winget upgrade --all --include-unknown 2>&1 | Out-String
                    Write-Host $upgradeOutput
                    
                    if ($upgradeOutput -match "Installer failed with exit code" -or $upgradeOutput -match "Files modified by the installer are currently in use") {
                        Write-Host "`nSome upgrades failed. Checking for common issues..."
                        Check-InstallationBlockers
                    }
                } catch {
                    Write-Host "Error during upgrade: $_"
                    Check-InstallationBlockers
                }
                Pause-Menu
                return
            }
            "2" {
                Clear-Host
                Write-Host "==============================================="
                Write-Host "  Available Packages [Copy ID to upgrade]"
                Write-Host "==============================================="
                winget upgrade --include-unknown
                Write-Host
                Write-Host "Enter one or more package IDs to upgrade (comma-separated, no spaces)"
                $packlist = Read-Host "IDs"
                $packlist = $packlist -replace ' ', ''
                if ([string]::IsNullOrWhiteSpace($packlist)) {
                    Write-Host "No package IDs entered."
                    Pause-Menu
                    return
                }
                $ids = $packlist.Split(",")
                foreach ($id in $ids) {
                    Write-Host "Upgrading ${id}..."
                    try {
                        $upgradeOutput = winget upgrade --id $id --include-unknown 2>&1 | Out-String
                        Write-Host $upgradeOutput
                        
                        if ($upgradeOutput -match "Installer failed with exit code" -or $upgradeOutput -match "Files modified by the installer are currently in use") {
                            Write-Host "`nUpgrade failed for ${id}. Checking for common issues..."
                            Check-InstallationBlockers $id
                        }
                    } catch {
                        Write-Host "Error upgrading ${id}: $_"
                        Check-InstallationBlockers $id
                    }
                    Write-Host
                }
                Pause-Menu
                return
            }
            default {
                Write-Host "Invalid option. Please choose 1, 2, or 0."
                continue
            }
        }
    }
}

function Check-InstallationBlockers {
    param (
        [string]$packageId = ""
    )

    Write-Host "`n[INSTALLATION TROUBLESHOOTING]"
    Write-Host "Checking for common installation blockers..."
    
    # 1. Check for pending reboots
    if (Test-PendingReboot) {
        Write-Host "`nWARNING: System has pending reboots that may block installations."
        Write-Host "Recommendation: Restart your computer and try again."
    }
    
    # 2. Check for running installers
    Write-Host "`nChecking for running installer processes..."
    $installerProcesses = Get-Process | Where-Object {
        $_.ProcessName -match "msiexec|installer|setup|update"
    }
    
    if ($installerProcesses) {
        Write-Host "The following installer processes are running:"
        $installerProcesses | Format-Table Id, ProcessName, MainWindowTitle -AutoSize
        
        $answer = Read-Host "`nWould you like to try closing these processes? (Y/N)"
        if ($answer -eq "Y" -or $answer -eq "y") {
            $installerProcesses | ForEach-Object {
                try {
                    Stop-Process -Id $_.Id -Force
                    Write-Host "Closed process: $($_.ProcessName) (ID: $($_.Id))"
                } catch {
                    Write-Host "Failed to close process $($_.ProcessName): $_"
                }
            }
            Write-Host "`nProcesses closed. Try installing again."
        }
    } else {
        Write-Host "No conflicting installer processes detected."
    }
    
    # 3. Check for file locks (download handle.exe if needed)
    $handlePath = "$env:TEMP\handle.exe"
    if (-not (Test-Path $handlePath)) {
        try {
            Write-Host "`nDownloading handle.exe from Sysinternals..."
            $handleUrl = "https://download.sysinternals.com/files/Handle.zip"
            $zipPath = "$env:TEMP\Handle.zip"
            
            Invoke-WebRequest -Uri $handleUrl -OutFile $zipPath
            Expand-Archive -Path $zipPath -DestinationPath $env:TEMP -Force
            Remove-Item $zipPath -Force
        } catch {
            Write-Host "Failed to download handle.exe: $_"
        }
    }
    
    if (Test-Path $handlePath)) {
        Write-Host "`nChecking for locked files in temp and program folders..."
        try {
            $lockedFiles = & $handlePath -a -nobanner $env:TEMP,"C:\Program Files","C:\Program Files (x86)" 2>&1 | 
                Where-Object { $_ -match "pid:" } |
                Select-Object -Unique
            
            if ($lockedFiles) {
                Write-Host "The following files/directories are locked by other processes:"
                $lockedFiles | ForEach-Object { Write-Host " - $_" }
                Write-Host "Recommendation: Close the applications using these files and try again."
            } else {
                Write-Host "No locked files detected in critical locations."
            }
        } catch {
            Write-Host "Error while checking for locked files: $_"
        }
    } else {
        Write-Host "`nNote: Could not access handle.exe for detailed file lock detection"
    }
    
    # 4. Generic advice
    Write-Host "`nGeneral recommendations:"
    Write-Host "1. Close all running applications"
    Write-Host "2. Ensure you have enough disk space"
    Write-Host "3. Try running this tool as Administrator"
    Write-Host "4. Check Windows Update for any pending updates"
    
    if ($packageId) {
        Write-Host "`nPackage-specific advice for ${packageId}:"
        Write-Host "- Check if the application is currently running"
        Write-Host "- Visit the vendor's website for troubleshooting tips"
        Write-Host "- Try manual installation from the vendor's website"
    }
    
    Pause-Menu
}

function Test-PendingReboot {
    $pendingReboot = $false
    
    # Check Component Based Servicing
    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue) {
        $pendingReboot = $true
    }
    
    # Check Windows Update Auto Update
    if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) {
        $pendingReboot = $true
    }
    
    # Check PendingFileRenameOperations
    if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue) {
        $pendingReboot = $true
    }
    
    return $pendingReboot
}

function Choice-2 {
    Clear-Host
    Write-Host "Scanning for corrupt files (SFC /scannow)..."
    sfc /scannow
    Pause-Menu
}

function Choice-3 {
    Clear-Host
    Write-Host "Checking Windows health status (DISM /CheckHealth)..."
    dism /online /cleanup-image /checkhealth
    Pause-Menu
}

function Choice-4 {
    Clear-Host
    Write-Host "Restoring Windows health status (DISM /RestoreHealth)..."
    dism /online /cleanup-image /restorehealth
    Pause-Menu
}

function Choice-5 {
    function Get-ActiveAdapters {
        # Exclude virtual adapters like vEthernet
        Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -notlike '*Virtual*' -and $_.Name -notlike '*vEthernet*' } | Select-Object -ExpandProperty Name
    }

    # Check if DoH is supported (Windows 11 or recent Windows 10)
    function Test-DoHSupport {
        $osVersion = [System.Environment]::OSVersion.Version
        return ($osVersion.Major -eq 10 -and $osVersion.Build -ge 19041) -or ($osVersion.Major -gt 10)
    }

    # Check if running as Administrator
    function Test-Admin {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    # Function to enable DoH for all known DNS servers using netsh
    function Enable-DoHAllServers {
        $dnsServers = @(
            # Cloudflare DNS
            @{ Server = "1.1.1.1"; Template = "https://cloudflare-dns.com/dns-query" },
            @{ Server = "1.0.0.1"; Template = "https://cloudflare-dns.com/dns-query" },
            @{ Server = "2606:4700:4700::1111"; Template = "https://cloudflare-dns.com/dns-query" },
            @{ Server = "2606:4700:4700::1001"; Template = "https://cloudflare-dns.com/dns-query" },
            # Google DNS
            @{ Server = "8.8.8.8"; Template = "https://dns.google/dns-query" },
            @{ Server = "8.8.4.4"; Template = "https://dns.google/dns-query" },
            @{ Server = "2001:4860:4860::8888"; Template = "https://dns.google/dns-query" },
            @{ Server = "2001:4860:4860::8844"; Template = "https://dns.google/dns-query" },
            # Quad9 DNS
            @{ Server = "9.9.9.9"; Template = "https://dns.quad9.net/dns-query" },
            @{ Server = "149.112.112.112"; Template = "https://dns.quad9.net/dns-query" },
            @{ Server = "2620:fe::fe"; Template = "https://dns.quad9.net/dns-query" },
            @{ Server = "2620:fe::fe:9"; Template = "https://dns.quad9.net/dns-query" },
            # AdGuard DNS
            @{ Server = "94.140.14.14"; Template = "https://dns.adguard.com/dns-query" },
            @{ Server = "94.140.15.15"; Template = "https://dns.adguard.com/dns-query" },
            @{ Server = "2a10:50c0::ad1:ff"; Template = "https://dns.adguard.com/dns-query" },
            @{ Server = "2a10:50c0::ad2:ff"; Template = "https://dns.adguard.com/dns-query" }
        )
        Write-Host "Enabling DoH for all known DNS servers..."
        $successCount = 0
        foreach ($dns in $dnsServers) {
            try {
                $command = "netsh dns add encryption server=$($dns.Server) dohtemplate=$($dns.Template) autoupgrade=yes udpfallback=no"
                $result = Invoke-Expression $command 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "  - DoH enabled for $($dns.Server) with template $($dns.Template)" -ForegroundColor Green
                    $successCount++
                } else {
                    Write-Host "  - Failed to enable DoH for $($dns.Server): $result" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "  - Failed to enable DoH for $($dns.Server): $_" -ForegroundColor Yellow
            }
        }
        if ($successCount -eq 0) {
            Write-Host "  - No DoH settings were applied successfully. Check system permissions or Windows version." -ForegroundColor Red
            return $false
        }
        # Flush DNS cache to ensure changes are applied
        try {
            Invoke-Expression "ipconfig /flushdns" | Out-Null
            Write-Host "  - DNS cache flushed to apply changes" -ForegroundColor Green
        } catch {
            Write-Host "  - Failed to flush DNS cache: $_" -ForegroundColor Yellow
        }
        # Attempt to restart DNS client service if running as Administrator
        if (Test-Admin) {
            $service = Get-Service -Name Dnscache -ErrorAction SilentlyContinue
            if ($service.Status -eq "Running" -and $service.StartType -ne "Disabled") {
                try {
                    Restart-Service -Name Dnscache -Force -ErrorAction Stop
                    Write-Host "  - DNS client service restarted to apply DoH settings" -ForegroundColor Green
                } catch {
                    Write-Host "  - Failed to restart DNS client service: $_" -ForegroundColor Yellow
                    try {
                        $stopResult = Invoke-Expression "net stop dnscache" 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            Start-Sleep -Seconds 2
                            $startResult = Invoke-Expression "net start dnscache" 2>&1
                            if ($LASTEXITCODE -eq 0) {
                                Write-Host "  - DNS client service restarted using net stop/start" -ForegroundColor Green
                            } else {
                                Write-Host "  - Failed to start DNS client service: $startResult" -ForegroundColor Yellow
                            }
                        } else {
                            Write-Host "  - Failed to stop DNS client service: $stopResult" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "  - Failed to restart DNS client service via net commands: $_" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "  - DNS client service is not running or is disabled. Please enable and start it manually." -ForegroundColor Yellow
            }
            Write-Host "  - Please reboot your system to apply DoH settings or manually restart the 'DNS Client' service in services.msc." -ForegroundColor Yellow
        } else {
            Write-Host "  - Not running as Administrator. Cannot restart DNS client service. Please reboot to apply DoH settings." -ForegroundColor Yellow
        }
        return $true
    }

    # Function to check DoH status
    function Check-DoHStatus {
        try {
            $netshOutput = Invoke-Expression "netsh dns show encryption" | Out-String
            if ($netshOutput -match "cloudflare-dns\.com|dns\.google|dns\.quad9\.net|dns\.adguard\.com") {
                Write-Host "DoH Status:"
                Write-Host $netshOutput -ForegroundColor Green
                Write-Host "DoH is enabled for at least one known DNS server." -ForegroundColor Green
            } else {
                Write-Host "DoH Status:"
                Write-Host $netshOutput -ForegroundColor Yellow
                Write-Host "No DoH settings detected. Ensure DNS servers are set and DoH was applied successfully." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Failed to check DoH status: $_" -ForegroundColor Red
        }
        Pause-Menu
    }

    # Function to update hosts file with ad-blocking entries
    function Update-HostsFile {
        Clear-Host
        Write-Host "==============================================="
        Write-Host "   Updating Windows Hosts File with Ad-Blocking"
        Write-Host "==============================================="
        
        # Check for admin privileges
        if (-not (Test-Admin)) {
            Write-Host "Error: This operation requires administrator privileges." -ForegroundColor Red
            Write-Host "Please run the script as Administrator and try again."
            Pause-Menu
            return
        }
        
        $hostsPath = "$env:windir\System32\drivers\etc\hosts"
        $backupPath = "$env:windir\System32\drivers\etc\hosts.bak"
        
        try {
            # Create backup of current hosts file
            if (Test-Path $hostsPath) {
                Copy-Item $hostsPath $backupPath -Force
                Write-Host "Created backup of hosts file at: $backupPath" -ForegroundColor Green
            }
            
            # Download ad-blocking hosts file
            $adBlockUrl = "https://o0.pages.dev/Lite/hosts.win"
            Write-Host "Downloading ad-blocking hosts file from: $adBlockUrl"
            $webClient = New-Object System.Net.WebClient
            $adBlockContent = $webClient.DownloadString($adBlockUrl)
            
            # Get existing custom entries from current hosts file (if any)
            $existingContent = ""
            if (Test-Path $hostsPath) {
                $existingContent = Get-Content $hostsPath | Where-Object {
                    $_ -notmatch "^# Ad-blocking entries" -and 
                    $_ -notmatch "^0\.0\.0\.0" -and 
                    $_ -notmatch "^127\.0\.0\.1" -and
                    $_ -notmatch "^::1" -and
                    $_ -notmatch "^$"
                }
            }
            
            # Combine existing custom entries with new ad-blocking entries
            $newContent = @"
# Ad-blocking entries - Updated $(Get-Date)
# Original hosts file backed up to: $backupPath

$existingContent

$adBlockContent
"@
            
            # Write new content to hosts file
            Set-Content -Path $hostsPath -Value $newContent -Encoding UTF8 -Force
            Write-Host "Successfully updated hosts file with ad-blocking entries." -ForegroundColor Green
            Write-Host "Total entries added: $($adBlockContent.Split("`n").Count)"
            
            # Flush DNS to apply changes
            try {
                ipconfig /flushdns | Out-Null
                Write-Host "DNS cache flushed to apply changes." -ForegroundColor Green
            } catch {
                Write-Host "Warning: Could not flush DNS cache. Changes may require a reboot." -ForegroundColor Yellow
            }
            
        } catch {
            Write-Host "Error updating hosts file: $_" -ForegroundColor Red
        }
        
        Pause-Menu
    }

    $dohSupported = Test-DoHSupport
    if (-not $dohSupported) {
        Write-Host "Warning: DNS over HTTPS (DoH) is not supported on this system. Option 5 will not be available." -ForegroundColor Yellow
    }

    while ($true) {
        Clear-Host
        Write-Host "======================================================"
        Write-Host "DNS / Network Tool"
        Write-Host "======================================================"
        Write-Host "[1] Set DNS to Google (8.8.8.8 / 8.8.4.4, IPv6)"
        Write-Host "[2] Set DNS to Cloudflare (1.1.1.1 / 1.0.0.1, IPv6)"
        Write-Host "[3] Restore automatic DNS (DHCP)"
        Write-Host "[4] Use your own DNS (IPv4/IPv6)"
        if ($dohSupported) {
            Write-Host "[5] Encrypt DNS: Enable DoH using netsh on all known DNS servers"
        }
        Write-Host "[6] Update Windows Hosts File with Ad-Blocking"
        Write-Host "[0] Return to menu"
        Write-Host "======================================================"
        $dns_choice = Read-Host "Enter your choice"
        switch ($dns_choice) {
            "1" {
                $adapters = Get-ActiveAdapters
                if (!$adapters) { Write-Host "No active network adapters found!" -ForegroundColor Red; Pause-Menu; return }
                Write-Host "Applying Google DNS (IPv4: 8.8.8.8/8.8.4.4, IPv6: 2001:4860:4860::8888/2001:4860:4860::8844) to:"
                foreach ($adapter in $adapters) {
                    Write-Host "  - $adapter"
                    $dnsAddresses = @("8.8.8.8", "8.8.4.4", "2001:4860:4860::8888", "2001:4860:4860::8844")
                    try {
                        Set-DnsClientServerAddress -InterfaceAlias $adapter -ServerAddresses $dnsAddresses -ErrorAction Stop
                        Write-Host "  - Google DNS applied successfully on $adapter" -ForegroundColor Green
                    } catch {
                        Write-Host "  - Failed to configure Google DNS on $adapter : $_" -ForegroundColor Yellow
                    }
                }
                Write-Host "Done. Google DNS set with IPv4 and IPv6."
                Write-Host "To enable DoH, use option [5] or configure manually in Settings."
                Pause-Menu
                return
            }
            "2" {
                $adapters = Get-ActiveAdapters
                if (!$adapters) { Write-Host "No active network adapters found!" -ForegroundColor Red; Pause-Menu; return }
                Write-Host "Applying Cloudflare DNS (IPv4: 1.1.1.1/1.0.0.1, IPv6: 2606:4700:4700::1111/2606:4700:4700::1001) to:"
                foreach ($adapter in $adapters) {
                    Write-Host "  - $adapter"
                    $dnsAddresses = @("1.1.1.1", "1.0.0.1", "2606:4700:4700::1111", "2606:4700:4700::1001")
                    try {
                        Set-DnsClientServerAddress -InterfaceAlias $adapter -ServerAddresses $dnsAddresses -ErrorAction Stop
                        Write-Host "  - Cloudflare DNS applied successfully on $adapter" -ForegroundColor Green
                    } catch {
                        Write-Host "  - Failed to configure Cloudflare DNS on $adapter : $_" -ForegroundColor Yellow
                    }
                }
                Write-Host "Done. Cloudflare DNS set with IPv4 and IPv6."
                Write-Host "To enable DoH, use option [5] or configure manually in Settings."
                Pause-Menu
                return
            }
            "3" {
                $adapters = Get-ActiveAdapters
                if (!$adapters) { Write-Host "No active network adapters found!" -ForegroundColor Red; Pause-Menu; return }
                Write-Host "Restoring automatic DNS (DHCP) on:"
                foreach ($adapter in $adapters) {
                    Write-Host "  - $adapter"
                    try {
                        Set-DnsClientServerAddress -InterfaceAlias $adapter -ResetServerAddresses -ErrorAction Stop
                        Write-Host "  - DNS set to automatic on $adapter" -ForegroundColor Green
                    } catch {
                        Write-Host "  - Failed to reset DNS on $adapter : $_" -ForegroundColor Yellow
                    }
                }
                Write-Host "Done. DNS set to automatic."
                Pause-Menu
                return
            }
            "4" {
                $adapters = Get-ActiveAdapters
                if (!$adapters) { Write-Host "No active network adapters found!" -ForegroundColor Red; Pause-Menu; return }
                while ($true) {
                    Clear-Host
                    Write-Host "==============================================="
                    Write-Host "          Enter your custom DNS"
                    Write-Host "==============================================="
                    Write-Host "Enter at least one DNS server (IPv4 or IPv6). Multiple addresses can be comma-separated."
                    $customDNS = Read-Host "Enter DNS addresses (e.g., 8.8.8.8,2001:4860:4860::8888)"
                    Clear-Host
                    Write-Host "==============================================="
                    Write-Host "         Validating DNS addresses..."
                    Write-Host "==============================================="
                    $dnsAddresses = $customDNS.Split(",", [StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
                    if ($dnsAddresses.Count -eq 0) {
                        Write-Host "[!] ERROR: No DNS addresses entered." -ForegroundColor Red
                        Pause-Menu
                        continue
                    }
                    $validDnsAddresses = @()
                    $allValid = $true
                    foreach ($dns in $dnsAddresses) {
                        $isIPv6 = $dns -match ":"
                        $reachable = Test-Connection -ComputerName $dns -Count 1 -Quiet -ErrorAction SilentlyContinue
                        if ($reachable) {
                            $validDnsAddresses += $dns
                            Write-Host "Validated: $dns" -ForegroundColor Green
                        } else {
                            Write-Host "[!] ERROR: The DNS address `"$dns`" is not reachable and will be skipped." -ForegroundColor Yellow
                            $allValid = $false
                        }
                    }
                    if ($validDnsAddresses.Count -eq 0) {
                        Write-Host "[!] ERROR: No valid DNS addresses provided." -ForegroundColor Red
                        Pause-Menu
                        continue
                    }
                    break
                }
                Clear-Host
                Write-Host "==============================================="
                Write-Host "    Setting DNS for all active adapters..."
                Write-Host "==============================================="
                foreach ($adapter in $adapters) {
                    Write-Host "  - $adapter"
                    try {
                        Set-DnsClientServerAddress -InterfaceAlias $adapter -ServerAddresses $validDnsAddresses -ErrorAction Stop
                        Write-Host "  - Custom DNS applied successfully on $adapter" -ForegroundColor Green
                    } catch {
                        Write-Host "  - Failed to configure custom DNS on $adapter : $_" -ForegroundColor Yellow
                    }
                }
                Write-Host
                Write-Host "==============================================="
                Write-Host "    DNS has been successfully updated:"
                foreach ($dns in $validDnsAddresses) {
                    Write-Host "      - $dns"
                }
                Write-Host "To enable DoH, use option [5] or configure manually in Settings."
                Write-Host "==============================================="
                Pause-Menu
                return
            }
            "5" {
                if (-not $dohSupported) {
                    Write-Host "Error: DoH is not supported on this system. Option 5 is unavailable." -ForegroundColor Red
                    Pause-Menu
                    return
                }
                $dohApplied = Enable-DoHAllServers
                while ($true) {
                    Clear-Host
                    Write-Host "======================================================"
                    Write-Host "DoH Configuration Menu"
                    Write-Host "======================================================"
                    if ($dohApplied) {
                        Write-Host "DoH was applied for $successCount DNS servers."
                    } else {
                        Write-Host "DoH application failed. Check system permissions or Windows version."
                    }
                    Write-Host "[1] Check DoH status"
                    Write-Host "[2] Return to menu"
                    Write-Host "======================================================"
                    $doh_choice = Read-Host "Enter your choice"
                    switch ($doh_choice) {
                        "1" { Check-DoHStatus }
                        "2" { return }
                        default { Write-Host "Invalid choice, please try again." -ForegroundColor Red; Pause-Menu }
                    }
                }
            }
            "6" { Update-HostsFile }
            "0" { return }
            default { Write-Host "Invalid choice, please try again." -ForegroundColor Red; Pause-Menu }
        }
    }
}
function Choice-6 { Clear-Host; Write-Host "Displaying Network Information..."; ipconfig /all; Pause-Menu }

function Choice-7 {
    Clear-Host
    Write-Host "=========================================="
    Write-Host "    Restarting all Wi-Fi adapters..."
    Write-Host "=========================================="

    $wifiAdapters = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match "Wi-Fi|Wireless" -and $_.Status -eq "Up" -or $_.Status -eq "Disabled" }

    if (-not $wifiAdapters) {
        Write-Host "No Wi-Fi adapters found!"
        Pause-Menu
        return
    }

    foreach ($adapter in $wifiAdapters) {
        Write-Host "Restarting '$($adapter.Name)'..."

        Disable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue

        Start-Sleep -Seconds 5

        # Check connection
        $status = Get-NetAdapter -Name $adapter.Name
        if ($status.Status -eq "Up") {
            Write-Host "SUCCESS: '$($adapter.Name)' is back online!" -ForegroundColor Green
        } else {
            Write-Host "WARNING: '$($adapter.Name)' is still offline!" -ForegroundColor Yellow
        }
    }

    Pause-Menu
}

function Choice-8 {
    $Host.UI.RawUI.WindowTitle = "Network Repair - Automatic Troubleshooter"
    Clear-Host
    Write-Host
    Write-Host "==============================="
    Write-Host "    Automatic Network Repair"
    Write-Host "==============================="
    Write-Host
    Write-Host "Step 1: Renewing your IP address..."
    ipconfig /release | Out-Null
    ipconfig /renew  | Out-Null
    Write-Host
    Write-Host "Step 2: Refreshing DNS settings..."
    ipconfig /flushdns | Out-Null
    Write-Host
    Write-Host "Step 3: Resetting network components..."
    netsh winsock reset | Out-Null
    netsh int ip reset  | Out-Null
    Write-Host
    Write-Host "Your network settings have been refreshed."
    Write-Host "A system restart is recommended for full effect."
    Write-Host
    while ($true) {
        $restart = Read-Host "Would you like to restart now? (Y/N)"
        switch ($restart.ToUpper()) {
            "Y" { shutdown /r /t 5; return }
            "N" { return }
            default { Write-Host "Invalid input. Please enter Y or N." }
        }
    }
}

function Choice-9 { Clear-Host; Write-Host "Running Disk Cleanup..."; Start-Process "cleanmgr.exe"; Pause-Menu }

function Choice-10 {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "Running advanced error scan on all drives..."
    Write-Host "==============================================="
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -ne $null } | Select-Object -ExpandProperty Name
    foreach ($drive in $drives) {
        Write-Host
        Write-Host "Scanning drive $drive`:" ...
        chkdsk "${drive}:" /f /r /x
    }
    Write-Host
    Write-Host "All drives scanned."
    Pause-Menu
}

function Choice-11 {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "   Delete Temporary Files and System Cache"
    Write-Host "==============================================="
    Write-Host
    Write-Host "This will permanently delete temporary files for your user and Windows."
    Write-Host "Warning: Close all applications to avoid file conflicts."
    Write-Host

    $deleteOption = ""
    while ($true) {
        Write-Host "==============================================="
        Write-Host "   Choose Cleanup Option"
        Write-Host "==============================================="
        Write-Host "[1] Permanently delete temporary files"
        Write-Host "[2] Permanently delete temporary files and empty Recycle Bin"
        Write-Host "[3] Advanced Privacy Cleanup (includes temp files + privacy data)"
        Write-Host "[0] Cancel"
        Write-Host
        $optionChoice = Read-Host "Select an option"
        switch ($optionChoice) {
            "1" { $deleteOption = "DeleteOnly"; break }
            "2" { $deleteOption = "DeleteAndEmpty"; break }
            "3" { $deleteOption = "PrivacyCleanup"; break }
            "0" {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
                Pause-Menu
                return
            }
            default { Write-Host "Invalid input. Please enter 1, 2, 3, or 0." -ForegroundColor Red }
        }
        if ($deleteOption) { break }
    }

    # Define paths to clean (remove redundant paths)
    $paths = @(
        $env:TEMP,              # User temp folder
        "C:\Windows\Temp"       # System temp folder
    )

    # Remove duplicates
    $paths = $paths | Select-Object -Unique

    # Load assembly for Recycle Bin if needed (only for DeleteAndEmpty option)
    if ($deleteOption -eq "DeleteAndEmpty" -or $deleteOption -eq "PrivacyCleanup") {
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
        } catch {
            Write-Host "[ERROR] Failed to load Microsoft.VisualBasic assembly for Recycle Bin operations." -ForegroundColor Red
            Write-Host "Proceeding with deletion only (Recycle Bin will not be emptied)." -ForegroundColor Yellow
            $deleteOption = "DeleteOnly"
        }
    }

    $deletedCount = 0
    $skippedCount = 0

    # Perform permanent deletion
    foreach ($path in $paths) {
        # Validate path
        if (-not (Test-Path $path)) {
            Write-Host "[ERROR] Path does not exist: $path" -ForegroundColor Red
            continue
        }

        # Additional safety check for user temp path
        if ($path -eq $env:TEMP -and -not ($path.ToLower() -like "*$($env:USERNAME.ToLower())*")) {
            Write-Host "[ERROR] TEMP path unsafe or invalid: $path" -ForegroundColor Red
            Write-Host "Skipping to prevent system damage." -ForegroundColor Red
            continue
        }

        Write-Host "Cleaning path: $path"
        try {
            Get-ChildItem -Path $path -Recurse -Force -ErrorAction Stop | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                    if ($_.PSIsContainer) {
                        Write-Host "Permanently deleted directory: $($_.FullName)" -ForegroundColor Green
                    } else {
                        Write-Host "Permanently deleted file: $($_.FullName)" -ForegroundColor Green
                    }
                    $deletedCount++
                } catch {
                    $skippedCount++
                    Write-Host "Skipped: $($_.FullName) ($($_.Exception.Message))" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "Error processing path $path : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Empty Recycle Bin if selected
    if ($deleteOption -eq "DeleteAndEmpty" -or $deleteOption -eq "PrivacyCleanup") {
        try {
            Write-Host "Emptying Recycle Bin..." -ForegroundColor Green
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
                "C:\`$Recycle.Bin",
                'OnlyErrorDialogs',
                'DeletePermanently'
            )
            Write-Host "Recycle Bin emptied successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error emptying Recycle Bin: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Perform privacy cleanup if selected
    if ($deleteOption -eq "PrivacyCleanup") {
        Write-Host
        Write-Host "==============================================="
        Write-Host "   Performing Advanced Privacy Cleanup"
        Write-Host "==============================================="
        
        # Clear Activity History
        try {
            Write-Host "Clearing Activity History..."
            reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist" /f 2>&1 | Out-Null
            Write-Host "Activity History cleared." -ForegroundColor Green
        } catch {
            Write-Host "Failed to clear Activity History: $_" -ForegroundColor Yellow
        }

        # Clear Location History
        try {
            Write-Host "Clearing Location History..."
            Get-Process LocationNotificationWindows -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs" /f 2>&1 | Out-Null
            Write-Host "Location History cleared." -ForegroundColor Green
        } catch {
            Write-Host "Failed to clear Location History: $_" -ForegroundColor Yellow
        }

        # Clear Diagnostic Data
        try {
            Write-Host "Clearing Diagnostic Data..."
            wevtutil cl Microsoft-Windows-Diagnostics-Performance/Operational 2>&1 | Out-Null
            Write-Host "Diagnostic Data cleared." -ForegroundColor Green
        } catch {
            Write-Host "Failed to clear Diagnostic Data: $_" -ForegroundColor Yellow
        }

        # Additional privacy cleanup commands
        try {
            Write-Host "Clearing Recent Items..."
            Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Recent\*" -Force -Recurse -ErrorAction SilentlyContinue
            Write-Host "Recent Items cleared." -ForegroundColor Green
        } catch {
            Write-Host "Failed to clear Recent Items: $_" -ForegroundColor Yellow
        }

        try {
            Write-Host "Clearing Thumbnail Cache..."
            Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache_*.db" -Force -ErrorAction SilentlyContinue
            Write-Host "Thumbnail Cache cleared." -ForegroundColor Green
        } catch {
            Write-Host "Failed to clear Thumbnail Cache: $_" -ForegroundColor Yellow
        }
    }

    Write-Host
    Write-Host "Cleanup complete. Processed $deletedCount files/directories, skipped $skippedCount files/directories." -ForegroundColor Green
    if ($deleteOption -eq "PrivacyCleanup") {
        Write-Host "Privacy-related data was also cleared."
    } else {
        Write-Host "Files and directories were permanently deleted."
    }

    Pause-Menu
}

function Choice-12 {
    while ($true) {
        Clear-Host
        Write-Host "======================================================"
        Write-Host " Advanced Registry Cleanup & Optimization"
        Write-Host "======================================================"
        Write-Host "[1] List 'safe to delete' registry keys under Uninstall"
        Write-Host "[2] Delete all 'safe to delete' registry keys (with backup)"
        Write-Host "[3] Create Registry Backup"
        Write-Host "[4] Restore Registry Backup"
        Write-Host "[5] Scan for corrupt registry entries"
        Write-Host "[0] Return to main menu"
        Write-Host
        $rchoice = Read-Host "Enter your choice"
        switch ($rchoice) {
            "1" {
                Write-Host
                Write-Host "Listing registry keys matching: IE40, IE4Data, DirectDrawEx, DXM_Runtime, SchedulingAgent"
                Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
                  Where-Object { $_.PSChildName -match 'IE40|IE4Data|DirectDrawEx|DXM_Runtime|SchedulingAgent' } |
                  ForEach-Object { Write-Host $_.PSChildName }
                Pause-Menu
            }
            "2" {
                Write-Host
                $backupFolder = "$env:SystemRoot\Temp\RegistryBackups"
                if (-not (Test-Path $backupFolder)) { New-Item -Path $backupFolder -ItemType Directory | Out-Null }

                $now = Get-Date
                $existingBackup = Get-ChildItem -Path $backupFolder -Filter "RegistryBackup_*.reg" |
                    Where-Object { ($now - $_.CreationTime).TotalMinutes -lt 10 } |  # backup within last 10 min
                    Sort-Object CreationTime -Descending | Select-Object -First 1

                $backupFile = $null
                if ($existingBackup) {
                    Write-Host "A recent backup already exists: $($existingBackup.Name)"
                    $useOld = Read-Host "Use this backup? (Y/n)"
                    if ($useOld -notin @("n", "N")) {
                        $backupFile = $existingBackup.FullName
                        Write-Host "Using existing backup: $backupFile"
                    } else {
                        $backupName = "RegistryBackup_{0}.reg" -f ($now.ToString("yyyy-MM-dd_HH-mm"))
                        $backupFile = Join-Path $backupFolder $backupName
                        reg export "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" $backupFile /y | Out-Null
                        Write-Host "New backup created: $backupFile"
                    }
                } else {
                    $backupName = "RegistryBackup_{0}.reg" -f ($now.ToString("yyyy-MM-dd_HH-mm"))
                    $backupFile = Join-Path $backupFolder $backupName
                    reg export "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" $backupFile /y | Out-Null
                    Write-Host "Backup created: $backupFile"
                }

                Write-Host "`nDeleting registry keys matching: IE40, IE4Data, DirectDrawEx, DXM_Runtime, SchedulingAgent"
                $keys = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
                    Where-Object { $_.PSChildName -match 'IE40|IE4Data|DirectDrawEx|DXM_Runtime|SchedulingAgent' }
                
                if ($keys) {
                    foreach ($key in $keys) {
                        try {
                            Remove-Item $key.PSPath -Recurse -Force -ErrorAction Stop
                            Write-Host "Deleted:" $key.PSChildName
                        } catch {
                            Write-Host "Failed to delete:" $key.PSChildName "($_.Exception.Message)"
                        }
                    }
                } else {
                    Write-Host "No matching registry keys found."
                }
                Pause-Menu
            }
            "3" {
                $backupFolder = "$env:SystemRoot\Temp\RegistryBackups"
                if (-not (Test-Path $backupFolder)) { New-Item -Path $backupFolder -ItemType Directory | Out-Null }
                $backupName = "RegistryBackup_{0}.reg" -f (Get-Date -Format "yyyy-MM-dd_HH-mm")
                $backupFile = Join-Path $backupFolder $backupName
                reg export HKLM $backupFile /y
                Write-Host "Full HKLM backup created: $backupFile"
                Pause-Menu
            }
            "4" {
                $backupFolder = "$env:SystemRoot\Temp\RegistryBackups"
                Write-Host "Available backups:"
                Get-ChildItem "$backupFolder\*.reg" | ForEach-Object { Write-Host $_.Name }
                $backupFile = Read-Host "Enter the filename to restore"
                $fullBackup = Join-Path $backupFolder $backupFile
                if (Test-Path $fullBackup) {
                    reg import $fullBackup
                    Write-Host "Backup successfully restored."
                } else {
                    Write-Host "File not found."
                }
                Pause-Menu
            }
            "5" {
                Clear-Host
                Write-Host "Scanning for corrupt registry entries..."
                Start-Process "cmd.exe" "/c sfc /scannow" -Wait
                Start-Process "cmd.exe" "/c dism /online /cleanup-image /checkhealth" -Wait
                Write-Host "Registry scan complete. If errors were found, please restart your PC."
                Pause-Menu
            }
            "0" { return }
            default { Write-Host "Invalid input. Try again."; Pause-Menu }
        }
    }
}

function Choice-13 {
    Clear-Host
    Write-Host "=========================================="
    Write-Host "     Optimize SSDs (ReTrim/TRIM)"
    Write-Host "=========================================="
    Write-Host "This will automatically optimize (TRIM) all detected SSDs."
    Write-Host
    Write-Host "Listing all detected SSD drives..."

    $ssds = Get-PhysicalDisk | Where-Object MediaType -eq 'SSD'
    if (-not $ssds) {
        Write-Host "No SSDs detected."
        Pause-Menu
        return
    }

    $log = "$env:USERPROFILE\Desktop\SSD_OPTIMIZE_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HHmmss")
    $logContent = @()
    $logContent += "SSD Optimize Log - $(Get-Date)"

    foreach ($ssd in $ssds) {
        $disk = Get-Disk | Where-Object { $_.FriendlyName -eq $ssd.FriendlyName }
        if ($disk) {
            $volumes = $disk | Get-Partition | Get-Volume | Where-Object DriveLetter -ne $null
            foreach ($vol in $volumes) {
                Write-Host "Optimizing SSD: $($vol.DriveLetter):"
                $logContent += "Optimizing SSD: $($vol.DriveLetter):"
                $result = Optimize-Volume -DriveLetter $($vol.DriveLetter) -ReTrim -Verbose 4>&1
                $logContent += $result
            }
        } else {
            $logContent += "Could not find Disk for SSD: $($ssd.FriendlyName)"
        }
    }
    Write-Host
    Write-Host "SSD optimization completed. Log file saved on Desktop: $log"
    $logContent | Out-File -FilePath $log -Encoding UTF8
    Pause-Menu
}

function Choice-14 {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "     Scheduled Task Management [Admin]"
    Write-Host "==============================================="
    Write-Host "Listing all scheduled tasks..."
    Write-Host "Microsoft tasks are shown in Green, third-party tasks in Yellow."
    Write-Host

    # Check for admin privileges
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Error: This function requires administrator privileges." -ForegroundColor Red
        Write-Host "Please run the script as Administrator and try again."
        Pause-Menu
        return
    }

    # Helper function to display task list with dynamic alignment and modified author/taskname
    function Show-TaskList {
        # Retrieve scheduled tasks
        try {
            $tasks = schtasks /query /fo CSV /v | ConvertFrom-Csv | Where-Object {
                $_."TaskName" -ne "" -and                        # Exclude empty TaskName
                $_."TaskName" -ne "TaskName" -and               # Exclude placeholder "TaskName"
                $_."Author" -ne "Author" -and                   # Exclude placeholder "Author"
                $_."Status" -ne "Status" -and                   # Exclude placeholder "Status"
                $_."Author" -notlike "*Scheduling data is not available in this format.*" -and  # Exclude invalid scheduling data
                $_."TaskName" -notlike "*Enabled*" -and         # Exclude rows starting with "Enabled"
                $_."TaskName" -notlike "*Disabled*"             # Exclude rows starting with "Disabled"
            }
            if (-not $tasks) {
                Write-Host "No valid scheduled tasks found." -ForegroundColor Yellow
                return $null
            }
        } catch {
            Write-Host "Error retrieving scheduled tasks: $_" -ForegroundColor Red
            return $null
        }

        # Remove duplicates based on TaskName, Author, and Status
        $uniqueTasks = $tasks | Sort-Object "TaskName", "Author", "Status" -Unique

        # Calculate maximum lengths for dynamic alignment
        $maxIdLength = ($uniqueTasks.Count.ToString()).Length  # Length of largest ID
        $maxTaskNameLength = 50  # Default max length for TaskName, adjustable
        $maxAuthorLength = 30    # Default max length for Author, adjustable
        $maxStatusLength = 10    # Default max length for Status (e.g., "Running", "Ready", "Disabled")

        # Process tasks to adjust Author and TaskName, and calculate max lengths
        $processedTasks = @()
        foreach ($task in $uniqueTasks) {
            $taskName = if ($task."TaskName") { $task."TaskName" } else { "N/A" }
            $author = if ($task."Author") { $task."Author" } else { "N/A" }
            $status = if ($task."Status") { $task."Status" } else { "Unknown" }

            # Fix Author field for Microsoft tasks with resource strings (e.g., $(@%SystemRoot%\...))
            if ($author -like '$(@%SystemRoot%\*' -or $taskName -like '\Microsoft\*') {
                $author = "Microsoft Corporation"
            }

            # Extract first folder from TaskName for Author if still N/A
            if ($author -eq "N/A" -and $taskName -match '^\\([^\\]+)\\') {
                $author = $matches[1]  # Get first folder (e.g., "LGTV Companion")
            }

            # Remove first folder from TaskName
            $displayTaskName = $taskName -replace '^\\[^\\]+\\', ''  # Remove "\Folder\"
            if ($displayTaskName -eq $taskName) { $displayTaskName = $taskName.TrimStart('\') }  # Fallback for tasks without folder

            # Truncate long fields for alignment
            if ($displayTaskName.Length -gt $maxTaskNameLength) { $displayTaskName = $displayTaskName.Substring(0, $maxTaskNameLength - 3) + "..." }
            if ($author.Length -gt $maxAuthorLength) { $author = $author.Substring(0, $maxAuthorLength - 3) + "..." }

            # Update max lengths based on processed data
            $maxTaskNameLength = [Math]::Max($maxTaskNameLength, [Math]::Min($displayTaskName.Length, 50))
            $maxAuthorLength = [Math]::Max($maxAuthorLength, [Math]::Min($author.Length, 30))
            $maxStatusLength = [Math]::Max($maxStatusLength, $status.Length)

            $processedTasks += [PSCustomObject]@{
                OriginalTaskName = $task."TaskName"
                DisplayTaskName  = $displayTaskName
                Author           = $author
                Status           = $status
            }
        }

        # Print header with dynamic widths
        $headerFormat = "{0,-$maxIdLength} | {1,-$maxTaskNameLength} | {2,-$maxAuthorLength} | {3}"
        Write-Host ($headerFormat -f "ID", "Task Name", "Author", "Status")
        Write-Host ("-" * $maxIdLength + "-+-" + "-" * $maxTaskNameLength + "-+-" + "-" * $maxAuthorLength + "-+-" + "-" * $maxStatusLength)

        # Display tasks with index and color coding
        $taskList = @()
        $index = 1
        foreach ($task in $processedTasks) {
            $isMicrosoft = $task.OriginalTaskName -like "\Microsoft\*" -or $task.Author -like "*Microsoft*"
            $taskList += [PSCustomObject]@{
                Index      = $index
                TaskName   = $task.OriginalTaskName  # Store original for schtasks commands
                Author     = $task.Author
                Status     = $task.Status
                IsMicrosoft = $isMicrosoft
            }
            $color = if ($isMicrosoft) { "Green" } else { "Yellow" }
            Write-Host ($headerFormat -f $index, $task.DisplayTaskName, $task.Author, $task.Status) -ForegroundColor $color
            $index++
        }
        Write-Host
        return $taskList
    }

    # Display task list initially
    $taskList = Show-TaskList
    if (-not $taskList) {
        Pause-Menu
        return
    }

    # Main loop for task management options
    while ($true) {
        Write-Host "Options:"
        Write-Host "[1] Enable a task"
        Write-Host "[2] Disable a task"
        Write-Host "[3] Delete a task"
        Write-Host "[4] Refresh task list"
        Write-Host "[0] Return to main menu"
        Write-Host

        $action = Read-Host "Enter option (0-4) or task ID to manage"
        if ($action -eq "0") {
            return
        } elseif ($action -eq "1") {
            $id = Read-Host "Enter task ID to enable"
            if ($id -match '^\d+$' -and $id -ge 1 -and $id -le $taskList.Count) {
                $selectedTask = $taskList[$id - 1]
                Write-Host "Enabling task: $($selectedTask.TaskName)"
                try {
                    schtasks /change /tn "$($selectedTask.TaskName)" /enable | Out-Null
                    Write-Host "Task enabled successfully." -ForegroundColor Green
                } catch {
                    Write-Host "Error enabling task: $_" -ForegroundColor Red
                }
            } else {
                Write-Host "Invalid task ID." -ForegroundColor Red
            }
            Pause-Menu
            Clear-Host
            Write-Host "==============================================="
            Write-Host "     Scheduled Task Management [Admin]"
            Write-Host "==============================================="
            Write-Host "Refreshing task list..."
            Write-Host "Microsoft tasks are shown in Green, third-party tasks in Yellow."
            Write-Host
            $taskList = Show-TaskList
            if (-not $taskList) {
                Pause-Menu
                return
            }
        } elseif ($action -eq "2") {
            $id = Read-Host "Enter task ID to disable"
            if ($id -match '^\d+$' -and $id -ge 1 -and $id -le $taskList.Count) {
                $selectedTask = $taskList[$id - 1]
                Write-Host "Disabling task: $($selectedTask.TaskName)"
                try {
                    schtasks /change /tn "$($selectedTask.TaskName)" /disable | Out-Null
                    Write-Host "Task disabled successfully." -ForegroundColor Green
                } catch {
                    Write-Host "Error disabling task: $_" -ForegroundColor Red
                }
            } else {
                Write-Host "Invalid task ID." -ForegroundColor Red
            }
            Pause-Menu
            Clear-Host
            Write-Host "==============================================="
            Write-Host "     Scheduled Task Management [Admin]"
            Write-Host "==============================================="
            Write-Host "Refreshing task list..."
            Write-Host "Microsoft tasks are shown in Green, third-party tasks in Yellow."
            Write-Host
            $taskList = Show-TaskList
            if (-not $taskList) {
                Pause-Menu
                return
            }
        } elseif ($action -eq "3") {
            $id = Read-Host "Enter task ID to delete"
            if ($id -match '^\d+$' -and $id -ge 1 -and $id -le $taskList.Count) {
                $selectedTask = $taskList[$id - 1]
                Write-Host "WARNING: Deleting task: $($selectedTask.TaskName)" -ForegroundColor Yellow
                $confirm = Read-Host "Are you sure? (Y/N)"
                if ($confirm -eq "Y" -or $confirm -eq "y") {
                    try {
                        schtasks /delete /tn "$($selectedTask.TaskName)" /f | Out-Null
                        Write-Host "Task deleted successfully." -ForegroundColor Green
                    } catch {
                        Write-Host "Error deleting task: $_" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Action cancelled." -ForegroundColor Yellow
                }
            } else {
                Write-Host "Invalid task ID." -ForegroundColor Red
            }
            Pause-Menu
            Clear-Host
            Write-Host "==============================================="
            Write-Host "     Scheduled Task Management [Admin]"
            Write-Host "==============================================="
            Write-Host "Refreshing task list..."
            Write-Host "Microsoft tasks are shown in Green, third-party tasks in Yellow."
            Write-Host
            $taskList = Show-TaskList
            if (-not $taskList) {
                Pause-Menu
                return
            }
        } elseif ($action -eq "4") {
            Clear-Host
            Write-Host "==============================================="
            Write-Host "     Scheduled Task Management [Admin]"
            Write-Host "==============================================="
            Write-Host "Refreshing task list..."
            Write-Host "Microsoft tasks are shown in Green, third-party tasks in Yellow."
            Write-Host
            $taskList = Show-TaskList
            if (-not $taskList) {
                Pause-Menu
                return
            }
        } else {
            Write-Host "Invalid option. Please enter 0-4 or a valid task ID." -ForegroundColor Red
            Pause-Menu
        }
    }
}

function Choice-15 {
    Clear-Host
    Write-Host
    Write-Host "=================================================="
    Write-Host "               CONTACT AND SUPPORT"
    Write-Host "=================================================="
    Write-Host "Do you have any questions or need help?"
    Write-Host "You are always welcome to contact me."
    Write-Host
    Write-Host "Discord-Username: Lil_Batti"
    Write-Host "Support-server: https://discord.gg/bCQqKHGxja"
    Write-Host
    Read-Host "Press ENTER to return to the main menu"
}

function Choice-0 { Clear-Host; Write-Host "Exiting script..."; exit }

function Choice-20 {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "    Saving Installed Driver Report to Desktop"
    Write-Host "==============================================="
    $outfile = "$env:USERPROFILE\Desktop\Installed_Drivers.txt"
    driverquery /v > $outfile
    Write-Host
    Write-Host "Driver report has been saved to: $outfile"
    Pause-Menu
}

function Choice-21 {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "    Windows Update Repair Tool [Admin]"
    Write-Host "==============================================="
    Write-Host
    Write-Host "[1/4] Stopping update-related services..."
    $services = @('wuauserv','bits','cryptsvc','msiserver','usosvc','trustedinstaller')
    foreach ($service in $services) {
        $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -ne "Stopped") {
            Write-Host "Stopping $service"
            try { Stop-Service -Name $service -Force -ErrorAction Stop } catch {}
        }
    }
    Start-Sleep -Seconds 2
    Write-Host
    Write-Host "[2/4] Renaming update cache folders..."
    $SUFFIX = ".bak_{0}" -f (Get-Random -Maximum 99999)
    $SD = "$env:windir\SoftwareDistribution"
    $CR = "$env:windir\System32\catroot2"
    $renamedSD = "$env:windir\SoftwareDistribution$SUFFIX"
    $renamedCR = "$env:windir\System32\catroot2$SUFFIX"
    if (Test-Path $SD) {
        try {
            Rename-Item $SD -NewName ("SoftwareDistribution" + $SUFFIX) -ErrorAction Stop
            if (Test-Path $renamedSD) {
                Write-Host "Renamed: $renamedSD"
            } else {
                Write-Host "Warning: Could not rename SoftwareDistribution."
            }
        } catch { Write-Host "Warning: Could not rename SoftwareDistribution." }
    } else { Write-Host "Info: SoftwareDistribution not found." }
    if (Test-Path $CR) {
        try {
            Rename-Item $CR -NewName ("catroot2" + $SUFFIX) -ErrorAction Stop
            if (Test-Path $renamedCR) {
                Write-Host "Renamed: $renamedCR"
            } else {
                Write-Host "Warning: Could not rename catroot2."
            }
        } catch { Write-Host "Warning: Could not rename catroot2." }
    } else { Write-Host "Info: catroot2 not found." }
    Write-Host
    Write-Host "[3/4] Restarting services..."
    foreach ($service in $services) {
        $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -ne "Running") {
            Write-Host "Starting $service"
            try { Start-Service -Name $service -ErrorAction Stop } catch {}
        }
    }
    Write-Host
    Write-Host "[4/4] Windows Update components have been reset."
    Write-Host
    Write-Host "Renamed folders:"
    Write-Host "  - $renamedSD"
    Write-Host "  - $renamedCR"
    Write-Host "You may delete them manually after reboot if all is working."
    Write-Host
    Pause-Menu
}

function Choice-22 {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "    Generating Separated System Reports..."
    Write-Host "==============================================="
    Write-Host
    Write-Host "Choose output location:"
    Write-Host " [1] Desktop (recommended)"
    Write-Host " [2] Enter custom path"
    Write-Host " [3] Show guide for custom path setup"
    $opt = Read-Host ">"
    $outpath = ""
    if ($opt -eq "1") {
        $desktop = [Environment]::GetFolderPath('Desktop')
        $reportdir = "SystemReports_{0}" -f (Get-Date -Format "yyyy-MM-dd_HHmm")
        $outpath = Join-Path $desktop $reportdir
        if (-not (Test-Path $outpath)) { New-Item -Path $outpath -ItemType Directory | Out-Null }
    } elseif ($opt -eq "2") {
        $outpath = Read-Host "Enter full path (e.g. D:\Reports)"
        if (-not (Test-Path $outpath)) {
            Write-Host
            Write-Host "[ERROR] Folder not found: $outpath"
            Pause-Menu
            return
        }
    } elseif ($opt -eq "3") {
        Clear-Host
        Write-Host "==============================================="
        Write-Host "    How to Use a Custom Report Path"
        Write-Host "==============================================="
        Write-Host
        Write-Host "1. Open File Explorer and create a new folder, e.g.:"
        Write-Host "   C:\Users\YourName\Desktop\SystemReports"
        Write-Host "   or"
        Write-Host "   C:\Users\YourName\OneDrive\Documents\SystemReports"
        Write-Host
        Write-Host "2. Copy the folder's full path from the address bar."
        Write-Host "3. Re-run this and choose option [2], then paste it."
        Write-Host
        Pause-Menu
        return
    } else {
        Write-Host
        Write-Host "Invalid selection."
        Start-Sleep -Seconds 2
        return
    }
    $datestr = Get-Date -Format "yyyy-MM-dd"
    $sys   = Join-Path $outpath "System_Info_$datestr.txt"
    $net   = Join-Path $outpath "Network_Info_$datestr.txt"
    $drv   = Join-Path $outpath "Driver_List_$datestr.txt"
    Write-Host
    Write-Host "Writing system info to: $sys"
    systeminfo | Out-File -FilePath $sys -Encoding UTF8
    Write-Host "Writing network info to: $net"
    ipconfig /all | Out-File -FilePath $net -Encoding UTF8
    Write-Host "Writing driver list to: $drv"
    driverquery | Out-File -FilePath $drv -Encoding UTF8
    Write-Host
    Write-Host "Reports saved in:"
    Write-Host $outpath
    Write-Host
    Pause-Menu
}

function Choice-23 {
    while ($true) {
        Clear-Host
        Write-Host "======================================================"
        Write-Host "           Windows Update Utility & Service Reset"
        Write-Host "======================================================"
        Write-Host "This tool will restart core Windows Update services."
        Write-Host "Make sure no Windows Updates are installing right now."
        Pause-Menu
        Write-Host
        Write-Host "[1] Reset Update Services (wuauserv, cryptsvc, appidsvc, bits)"
        Write-Host "[2] Return to Main Menu"
        Write-Host
        $fixchoice = Read-Host "Select an option"
        switch ($fixchoice) {
            "1" {
                Clear-Host
                Write-Host "======================================================"
                Write-Host "    Resetting Windows Update & Related Services"
                Write-Host "======================================================"
                Write-Host "Stopping Windows Update service..."
                try { Stop-Service -Name wuauserv -Force -ErrorAction Stop } catch {}
                Write-Host "Stopping Cryptographic service..."
                try { Stop-Service -Name cryptsvc -Force -ErrorAction Stop } catch {}
                Write-Host "Starting Application Identity service..."
                try { Start-Service -Name appidsvc -ErrorAction Stop } catch {}
                Write-Host "Starting Windows Update service..."
                try { Start-Service -Name wuauserv -ErrorAction Stop } catch {}
                Write-Host "Starting Background Intelligent Transfer Service..."
                try { Start-Service -Name bits -ErrorAction Stop } catch {}
                Write-Host
                Write-Host "[OK] Update-related services have been restarted."
                Pause-Menu
                return
            }
            "2" { return }
            default { Write-Host "Invalid input. Try again."; Pause-Menu }
        }
    }
}

function Choice-24 {
    while ($true) {
        Clear-Host
        Write-Host "==============================================="
        Write-Host "     View Network Routing Table  [Advanced]"
        Write-Host "==============================================="
        Write-Host "This shows how your system handles network traffic."
        Write-Host
        Write-Host "[1] Display routing table in this window"
        Write-Host "[2] Save routing table as a text file on Desktop"
        Write-Host "[3] Return to Main Menu"
        Write-Host
        $routeopt = Read-Host "Choose an option"
        switch ($routeopt) {
            "1" {
                Clear-Host
                route print
                Write-Host
                Pause-Menu
                return
            }
            "2" {
                $desktop = "$env:USERPROFILE\Desktop"
                if (-not (Test-Path $desktop)) {
                    Write-Host "Desktop folder not found."
                    Pause-Menu
                    return
                }
                $dt = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
                if (-not $dt) { $dt = "manual_timestamp" }
                $file = Join-Path $desktop "routing_table_${dt}.txt"
                Clear-Host
                Write-Host "Saving routing table to: `"$file`""
                Write-Host
                route print | Out-File -FilePath $file -Encoding UTF8
                if (Test-Path $file) {
                    Write-Host "[OK] Routing table saved successfully."
                } else {
                    Write-Host "[ERROR] Failed to save routing table to file."
                }
                Write-Host
                Pause-Menu
                return
            }
            "3" { return }
            default {
                Write-Host "Invalid input. Please enter 1, 2 or 3."
                Pause-Menu
            }
        }
    }
}

# === MAIN MENU LOOP ===
while ($true) {
    Show-Menu
    $choice = Read-Host "Enter your choice"
    switch ($choice) {
        "1"  { Choice-1; continue }
        "2"  { Choice-2; continue }
        "3"  { Choice-3; continue }
        "4"  { Choice-4; continue }
        "5"  { Choice-5; continue }
        "6"  { Choice-6; continue }
        "7"  { Choice-7; continue }
        "8"  { Choice-8; continue }
        "9"  { Choice-9; continue }
        "10" { Choice-10; continue }
        "11" { Choice-11; continue }
        "12" { Choice-12; continue }
        "13" { Choice-13; continue }
        "14" { Choice-14; continue }
        "15" { Choice-15; continue }
        "0" { Choice-0 }
        "20" { Choice-20; continue }
        "21" { Choice-21; continue }
        "22" { Choice-22; continue }
        "23" { Choice-23; continue }
        "24" { Choice-24; continue }
        default { Write-Host "Invalid choice, please try again."; Pause-Menu }
    }
}
