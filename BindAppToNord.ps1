# 1. SELF-ELEVATION (Run as Admin)
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Restarting as Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# 2. DETECT NORDLYNX
Write-Host "Scanning for NordVPN Adapter..." -ForegroundColor Cyan
$vpnAdapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like "*NordLynx*" }

if (-not $vpnAdapter) {
    Write-Error "NordLynx adapter not found. Please connect NordVPN first."
    Read-Host "Press Enter to exit"
    exit
}

# 3. SAFETY CHECK (Ensure NordLynx isn't misidentified as a physical wire)
# If NordLynx claims to be "802.3" (Ethernet), blocking "Wired" would kill the VPN.
if ($vpnAdapter.PhysicalMediaType -eq "802.3") {
    $safeToBlockWired = $false
    Write-Warning "NordLynx is masquerading as Ethernet. Using fallback blocking mode."
} else {
    $safeToBlockWired = $true
}

# 4. LOAD FILE PICKER
Add-Type -AssemblyName System.Windows.Forms

function Select-Executable {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Title = "Select the Application (.exe) to Bind to NordVPN"
    $FileBrowser.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $FileBrowser.Filter = "Applications (*.exe)|*.exe|All Files (*.*)|*.*"
    
    if ($FileBrowser.ShowDialog() -eq "OK") {
        return $FileBrowser.FileName
    } else {
        return $null
    }
}

# 5. BLOCKING LOGIC
function Set-VPNRule {
    param ([string]$FilePath)

    if (-not $FilePath) { return }

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    Write-Host "`nConfiguring Firewall for: " -NoNewline; Write-Host $fileName -ForegroundColor Magenta
    Write-Host "Path: $FilePath" -ForegroundColor DarkGray

    # Clean old rules
    Get-NetFirewallRule -DisplayName "Block $fileName*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule -Confirm:$false -ErrorAction SilentlyContinue
    Get-NetFirewallRule -DisplayName "Allow $fileName*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule -Confirm:$false -ErrorAction SilentlyContinue

    # RULE A: Allow specific app on NordLynx ONLY
    New-NetFirewallRule -DisplayName "Allow $fileName via NordLynx" -Direction Outbound -Program $FilePath -Action Allow -InterfaceAlias $vpnAdapter.Name -Profile Any -ErrorAction SilentlyContinue | Out-Null
    Write-Host "  [+] Allowed on NordLynx" -ForegroundColor Green

    # RULE B: Block specific app on ALL WIRELESS (Covers Wi-Fi 6/7, MLO, Ghost Adapters)
    New-NetFirewallRule -DisplayName "Block $fileName - All Wireless" -Direction Outbound -Program $FilePath -Action Block -InterfaceType Wireless -Profile Any -ErrorAction SilentlyContinue | Out-Null
    Write-Host "  [-] Blocked on All Wireless Interfaces" -ForegroundColor Red

    # RULE C: Block specific app on WIRED (Ethernet)
    if ($safeToBlockWired) {
        New-NetFirewallRule -DisplayName "Block $fileName - All Wired" -Direction Outbound -Program $FilePath -Action Block -InterfaceType Wired -Profile Any -ErrorAction SilentlyContinue | Out-Null
        Write-Host "  [-] Blocked on All Wired Interfaces" -ForegroundColor Red
    } else {
        # Fallback: Block Ethernet by name if global wired block is unsafe
        $eth = Get-NetAdapter | Where-Object { $_.PhysicalMediaType -eq "802.3" -and $_.InterfaceDescription -notlike "*NordLynx*" }
        foreach ($e in $eth) {
            New-NetFirewallRule -DisplayName "Block $fileName on $($e.Name)" -Direction Outbound -Program $FilePath -Action Block -InterfaceAlias $e.Name -Profile Any -ErrorAction SilentlyContinue | Out-Null
        }
        Write-Host "  [-] Blocked on Wired Adapters (ByName)" -ForegroundColor Yellow
    }

    # Attempt to close the app so rules apply immediately
    $procName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    Stop-Process -Name $procName -ErrorAction SilentlyContinue
    Write-Host "  [!] App process restarted/killed." -ForegroundColor DarkGray
}

# 6. MAIN LOOP
do {
    $selectedFile = Select-Executable
    
    if ($selectedFile) {
        Set-VPNRule -FilePath $selectedFile
        Write-Host "`nSUCCESS! $selectedFile is now bound to NordVPN." -ForegroundColor Green
    } else {
        Write-Host "No file selected." -ForegroundColor Yellow
    }

    $response = Read-Host "`nDo you want to select another file? (y/n)"
} while ($response -eq "y")

Write-Host "Exiting..."
Start-Sleep -Seconds 1