<#
.SYNOPSIS
    Standalone script to completely remove leftover HP bloatware folders and registry keys.
.DESCRIPTION
    This script elevates to Administrator, scans for known HP residual directories and registry keys,
    and forcefully removes them to clean up the system after a standard uninstallation.
#>

# --- 1. ADMIN CHECK ---
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Administrator privileges required. Elevating..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "      HP Bloatware Leftover Cleaner   " -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Starting cleanup process..."

# --- 2. NUKE LEFTOVER REGISTRY KEYS ---
Write-Host "`n[1/2] Scanning Registry..." -ForegroundColor Yellow
$hpRegPaths = @(
    "HKCU:\SOFTWARE\Hewlett-Packard",
    "HKCU:\SOFTWARE\HP",
    "HKLM:\SOFTWARE\Hewlett-Packard",
    "HKLM:\SOFTWARE\HP",
    "HKLM:\SOFTWARE\WOW6432Node\Hewlett-Packard",
    "HKLM:\SOFTWARE\WOW6432Node\HP"
)

$regRemoved = 0
foreach ($regPath in $hpRegPaths) {
    if (Test-Path $regPath) {
        try {
            # Force remove the key and all subkeys
            Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
            Write-Host "  [DELETED] Key: $regPath" -ForegroundColor Green
            $regRemoved++
        }
        catch {
            # FIXED: Wrapped $regPath in {} to prevent the colon from breaking parsing
            Write-Host "  [ERROR] Failed to delete ${regPath}: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
if ($regRemoved -eq 0) { Write-Host "  No leftover registry keys found." -ForegroundColor DarkGray }

# --- 3. NUKE PERSISTENT FOLDERS ---
Write-Host "`n[2/2] Scanning File System..." -ForegroundColor Yellow
$hpFolderPaths = @(
    "$env:SystemDrive\swsetup",
    "$env:SystemDrive\hp",
    "$env:ProgramData\HP",
    "$env:ProgramData\Hewlett-Packard",
    "$env:ProgramData\HP Audio Switch",
    "$env:LOCALAPPDATA\HP",
    "$env:LOCALAPPDATA\Hewlett-Packard",
    "$env:LOCALAPPDATA\TouchSmartData",
    "$env:APPDATA\HP",
    "$env:APPDATA\hp active health",
    "$env:APPDATA\hpqlog"
)

$foldersRemoved = 0
foreach ($folder in $hpFolderPaths) {
    if (Test-Path $folder) {
        try {
            # Force remove the folder and all contents
            Remove-Item -Path $folder -Recurse -Force -ErrorAction Stop
            Write-Host "  [DELETED] Folder: $folder" -ForegroundColor Green
            $foldersRemoved++
        }
        catch {
            # FIXED: Wrapped $folder in {} to prevent the colon from breaking parsing
            Write-Host "  [ERROR] Failed to delete ${folder}: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
if ($foldersRemoved -eq 0) { Write-Host "  No leftover folders found." -ForegroundColor DarkGray }

Write-Host "`n======================================" -ForegroundColor Cyan
Write-Host "Cleanup Complete!" -ForegroundColor Green
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')