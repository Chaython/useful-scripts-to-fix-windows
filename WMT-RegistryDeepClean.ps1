<#
    Windows Maintenance Tool - Deep Registry Cleaner

    Standalone export of the WMT-GUI "Deep Clean (Invalid Paths)" registry logic.

    Examples:
      .\WMT-RegistryDeepClean.ps1
      .\WMT-RegistryDeepClean.ps1 -Targets Startup,Firewall,MuiCache -CsvPath .\registry-findings.csv
      .\WMT-RegistryDeepClean.ps1 -Clean -CreateRestorePoint
      .\WMT-RegistryDeepClean.ps1 -Clean -Force -Targets Startup,MuiCache
#>

[CmdletBinding()]
param(
    [ValidateSet(
        "SharedDLLs", "Ext", "FileExts", "ActiveX", "TypeLib", "AppPaths",
        "Apps", "Installer", "Uninstall", "Startup", "Icons", "ProgIDs",
        "Services", "MuiCache", "AppCompat", "Firewall"
    )]
    [string[]]$Targets = @(
        "SharedDLLs", "Ext", "FileExts", "ActiveX", "TypeLib", "AppPaths",
        "Apps", "Installer", "Uninstall", "Startup", "Icons", "ProgIDs",
        "Services", "MuiCache", "AppCompat", "Firewall"
    ),

    [switch]$Clean,
    [switch]$CreateRestorePoint,
    [switch]$Force,
    [switch]$ListTargets,
    [string]$BackupDirectory,
    [string]$CsvPath,
    [int]$PreviewLimit = 50
)

$ErrorActionPreference = "SilentlyContinue"

function Get-WmtRegistryDeepCleanCategories {
    [ordered]@{
        "Missing Shared DLLs"             = "SharedDLLs"
        "Unused File Extensions (System)" = "Ext"
        "Unused File Extensions (User)"   = "FileExts"
        "ActiveX & COM Issues"            = "ActiveX"
        "Type Libraries (TLB)"            = "TypeLib"
        "Application Paths"               = "AppPaths"
        "Applications (Registered)"       = "Apps"
        "Installer Folders"               = "Installer"
        "Obsolete Software (Uninstall)"   = "Uninstall"
        "Run At Startup"                  = "Startup"
        "Invalid Default Icons"           = "Icons"
        "File Associations"               = "ProgIDs"
        "Windows Services"                = "Services"
        "MUI Cache (MRU Lists)"           = "MuiCache"
        "Compatibility Store (Flags)"     = "AppCompat"
        "Firewall Rules"                  = "Firewall"
    }
}

function Get-WmtDefaultRegistryBackupDirectory {
    $root = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
    return (Join-Path $root "data\RegistryBackups")
}

function Write-WmtStatus {
    param([string]$Message)
    Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message)
}

function Test-WmtIsAdministrator {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Ensure-WmtRegistryDrives {
    if (-not (Get-PSDrive -Name HKCR -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
    }
}

function Enable-WmtRegistryPrivileges {
    if (-not ([System.Management.Automation.PSTypeName]'Win32.TokenManipulator').Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    public class TokenManipulator {
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall, ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern IntPtr GetCurrentProcess();
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);
        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        internal struct TokPriv1Luid { public int Count; public long Luid; public int Attr; }
        internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
        internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;
        internal const int TOKEN_QUERY = 0x00000008;
        public static bool EnablePrivilege(string privilege) {
            try {
                IntPtr htok = IntPtr.Zero;
                if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok)) return false;
                TokPriv1Luid tp; tp.Count = 1; tp.Attr = SE_PRIVILEGE_ENABLED; tp.Luid = 0;
                if (!LookupPrivilegeValue(null, privilege, ref tp.Luid)) return false;
                if (!AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero)) return false;
                return true;
            } catch { return false; }
        }
    }
}
"@
    }

    try {
        [Win32.TokenManipulator]::EnablePrivilege("SeTakeOwnershipPrivilege") | Out-Null
        [Win32.TokenManipulator]::EnablePrivilege("SeRestorePrivilege") | Out-Null
    }
    catch {}
}

function Test-WmtPathExists {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ($Path -match "%.*%" -or $Path -match "\$\(.*\)") { return $true }
    if (Test-Path -Path $Path) { return $true }
    if ($Path -match "(?i)System32") {
        $nativePath = $Path -replace "(?i)System32", "Sysnative"
        if (Test-Path -Path $nativePath) { return $true }
    }
    return $false
}

function Convert-WmtRegistryPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $normalized = $Path.Trim()
    $normalized = $normalized -replace "^Registry::", ""
    $normalized = $normalized -replace "^HKEY_LOCAL_MACHINE\\", "HKLM:\"
    $normalized = $normalized -replace "^HKEY_CURRENT_USER\\", "HKCU:\"
    $normalized = $normalized -replace "^HKEY_CLASSES_ROOT\\", "HKCR:\"

    if ($normalized -notmatch '^(?<Hive>HKLM|HKCU|HKCR):\\(?<SubPath>.*)$') { return $null }

    $hive = switch ($Matches.Hive) {
        "HKLM" { [Microsoft.Win32.RegistryHive]::LocalMachine }
        "HKCU" { [Microsoft.Win32.RegistryHive]::CurrentUser }
        "HKCR" { [Microsoft.Win32.RegistryHive]::ClassesRoot }
    }

    [PSCustomObject]@{
        Hive    = $hive
        Root    = $Matches.Hive
        SubPath = $Matches.SubPath
    }
}

function Convert-WmtRegistryPathToRegExePath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $normalized = $Path.Trim()
    $normalized = $normalized -replace "^Registry::", ""
    $normalized = $normalized -replace "^HKLM:?\\", "HKEY_LOCAL_MACHINE\"
    $normalized = $normalized -replace "^HKCU:?\\", "HKEY_CURRENT_USER\"
    $normalized = $normalized -replace "^HKCR:?\\", "HKEY_CLASSES_ROOT\"
    return $normalized
}

function Test-WmtRegistryValueExists {
    param(
        [string]$Path,
        [string]$ValueName
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or $null -eq $ValueName) { return $false }

    $parts = Convert-WmtRegistryPath -Path $Path
    if (-not $parts) { return $false }

    $views = @([Microsoft.Win32.RegistryView]::Default)
    foreach ($view in $views) {
        $baseKey = $null
        $key = $null
        try {
            $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($parts.Hive, $view)
            $key = $baseKey.OpenSubKey($parts.SubPath, $false)
            if ($key) {
                return ($key.GetValueNames() -contains $ValueName)
            }
        }
        catch {}
        finally {
            if ($key) { $key.Close() }
            if ($baseKey) { $baseKey.Close() }
        }
    }

    return $false
}

function Remove-WmtRegistryValueNative {
    param(
        [string]$Path,
        [string]$ValueName
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or $null -eq $ValueName) { return $false }

    $parts = Convert-WmtRegistryPath -Path $Path
    if (-not $parts) { return $false }

    $views = @([Microsoft.Win32.RegistryView]::Default)
    foreach ($view in $views) {
        $baseKey = $null
        $key = $null
        try {
            $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($parts.Hive, $view)
            $key = $baseKey.OpenSubKey($parts.SubPath, $true)
            if ($key) {
                $key.DeleteValue($ValueName, $false)
                $key.Close(); $key = $null
                if (-not (Test-WmtRegistryValueExists -Path $Path -ValueName $ValueName)) { return $true }
            }
        }
        catch {}
        finally {
            if ($key) { $key.Close() }
            if ($baseKey) { $baseKey.Close() }
        }
    }

    return $false
}

function Remove-WmtRegistryValueRegExe {
    param(
        [string]$Path,
        [string]$ValueName
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or $null -eq $ValueName) { return $false }

    $regPath = Convert-WmtRegistryPathToRegExePath -Path $Path
    if ([string]::IsNullOrWhiteSpace($regPath)) { return $false }

    try {
        & reg.exe delete $regPath /v $ValueName /f 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0 -or -not (Test-WmtRegistryValueExists -Path $Path -ValueName $ValueName)) { return $true }
    }
    catch {}

    return (-not (Test-WmtRegistryValueExists -Path $Path -ValueName $ValueName))
}

function Remove-WmtRegKeyForced {
    param($Path, $IsKey, $ValName)
    if ($Path -match "^HKEY_CLASSES_ROOT") { $Path = $Path -replace "^HKEY_CLASSES_ROOT", "HKCR:" }
    if ($Path -match "^HKEY_LOCAL_MACHINE") { $Path = $Path -replace "^HKEY_LOCAL_MACHINE", "HKLM:" }
    if ($Path -match "^HKEY_CURRENT_USER") { $Path = $Path -replace "^HKEY_CURRENT_USER", "HKCU:" }

    $realPaths = @()
    if ($Path -match "^HKLM" -or $Path -match "^HKCU") { $realPaths += $Path }
    elseif ($Path -match "^HKCR:\\(?<SubPath>.*)") {
        $sub = $Matches.SubPath
        if (Test-Path "HKLM:\SOFTWARE\Classes\$sub") { $realPaths += "HKLM:\SOFTWARE\Classes\$sub" }
        if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Classes\$sub") { $realPaths += "HKLM:\SOFTWARE\WOW6432Node\Classes\$sub" }
        if (Test-Path "HKCU:\Software\Classes\$sub") { $realPaths += "HKCU:\Software\Classes\$sub" }
    }
    if ($realPaths.Count -eq 0) { $realPaths += $Path }

    $globalSuccess = $true
    $literalValueName = if ($null -ne $ValName) { [System.Management.Automation.WildcardPattern]::Escape([string]$ValName) } else { $ValName }
    foreach ($targetPath in $realPaths) {
        try {
            if ($IsKey) {
                Remove-Item -LiteralPath $targetPath -Recurse -Force -ErrorAction Stop
            }
            else {
                [void](Remove-WmtRegistryValueNative -Path $targetPath -ValueName $ValName)
                if (Test-WmtRegistryValueExists -Path $targetPath -ValueName $ValName) {
                    try { Remove-ItemProperty -LiteralPath $targetPath -Name $literalValueName -ErrorAction Stop } catch {}
                }
                if (Test-WmtRegistryValueExists -Path $targetPath -ValueName $ValName) {
                    throw "Registry value still exists after native removal."
                }
            }
            continue
        }
        catch {}

        try {
            $sid = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)
            $adminUser = $sid.Translate([System.Security.Principal.NTAccount])
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule($adminUser, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
            $unlockItem = {
                param($p)
                try {
                    $acl = Get-Acl $p
                    $acl.SetOwner($adminUser)
                    Set-Acl $p $acl -ErrorAction SilentlyContinue
                    $acl = Get-Acl $p
                    $acl.SetAccessRule($rule)
                    Set-Acl $p $acl -ErrorAction SilentlyContinue
                }
                catch {}
            }
            if ($IsKey) {
                $children = Get-ChildItem -LiteralPath $targetPath -Recurse -ErrorAction SilentlyContinue
                foreach ($c in $children) { & $unlockItem -p $c.PSPath }
            }
            & $unlockItem -p $targetPath
            if ($IsKey) {
                Remove-Item -LiteralPath $targetPath -Recurse -Force -ErrorAction Stop
            }
            else {
                [void](Remove-WmtRegistryValueNative -Path $targetPath -ValueName $ValName)
                if (Test-WmtRegistryValueExists -Path $targetPath -ValueName $ValName) {
                    try { Remove-ItemProperty -LiteralPath $targetPath -Name $literalValueName -ErrorAction Stop } catch {}
                }
                if (Test-WmtRegistryValueExists -Path $targetPath -ValueName $ValName) {
                    [void](Remove-WmtRegistryValueRegExe -Path $targetPath -ValueName $ValName)
                }
                if (Test-WmtRegistryValueExists -Path $targetPath -ValueName $ValName) {
                    throw "Registry value still exists after forced removal."
                }
            }
        }
        catch { $globalSuccess = $false }
    }
    return $globalSuccess
}

function Backup-WmtRegKey {
    param($ItemObj, $FilePath)
    $path = $ItemObj.RegPath
    $targetValue = $ItemObj.ValueName
    $type = $ItemObj.Type
    if ([string]::IsNullOrWhiteSpace($path)) { return }

    $regKeyPath = $path -replace "^HKLM:?\\", "HKEY_LOCAL_MACHINE\" -replace "^HKCU:?\\", "HKEY_CURRENT_USER\" -replace "^HKCR:?\\", "HKEY_CLASSES_ROOT\"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("[$regKeyPath]")
    try {
        if ($type -eq "Key") {
            $tmpFile = Join-Path ([System.IO.Path]::GetTempPath()) ("WMT_RegBackup_{0}.reg" -f ([guid]::NewGuid().ToString("N")))
            try {
                & reg.exe export $regKeyPath $tmpFile /y | Out-Null
                if (Test-Path -LiteralPath $tmpFile) {
                    $content = Get-Content -LiteralPath $tmpFile -Raw -Encoding Unicode
                    $content = $content -replace "^\uFEFF?Windows Registry Editor Version 5\.00\r?\n\r?\n", ""
                    if (-not [string]::IsNullOrWhiteSpace($content)) {
                        Add-Content -Path $FilePath -Value $content -Encoding Unicode
                        return
                    }
                }
            }
            catch {}
            finally {
                Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue
            }
        }
        elseif ($type -eq "Value") {
            $val = Get-ItemProperty -Path $path -Name $targetValue -ErrorAction SilentlyContinue
            if ($val) {
                $vData = $val.$targetValue
                if ($vData -is [string]) { $vData = '"' + ($vData -replace '\\', '\\' -replace '"', '\"') + '"' }
                elseif ($vData -is [int]) { $vData = "dword:{0:x8}" -f $vData }
                [void]$sb.AppendLine("`"$targetValue`"=$vData")
            }
        }
        [void]$sb.AppendLine("")
        Add-Content -Path $FilePath -Value $sb.ToString() -Encoding Unicode
    }
    catch {}
}

function Test-WmtRegistryCleanerWhitelist {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $safeList = @(
        "TetheringSettingHandler", "CrossDevice", "Windows.Media.Protection", "psmachine", "WebView2",
        "System.Data.dll", "System.EnterpriseServices", "rundll32", "explorer.exe", "svchost",
        "dllhost", "wmiprvse", "mmgaserver", "pickerhost", "castsrv", "uihelper",
        "backgroundtaskhost", "smartscreen", "runtimebroker", "mousocoreworker",
        "spatialaudiolicensesrv", "speechruntime", "mstsc.exe", "searchprotocolhost",
        "AppX", "WindowsApps", "UIEOrchestrator", "control.exe", "sdclt.exe",
        "provtool.exe", "perfmon", "Diagnostic.Perfmon", "QuickActionsPS", "VailAudioProxy"
    )
    foreach ($safe in $safeList) {
        if ($Path -match "(?i)$safe") { return $true }
    }
    return $false
}

function Get-WmtRealExecutablePath {
    param([string]$RawString)
    if ([string]::IsNullOrWhiteSpace($RawString)) { return $null }
    $clean = $RawString.Trim()
    if ($clean -match "^(.*?),\s*-?\d+$") { $clean = $matches[1].Trim() }
    if ($clean.StartsWith('"')) {
        $endQuote = $clean.IndexOf('"', 1)
        if ($endQuote -gt 1) {
            $quotedPart = $clean.Substring(1, $endQuote - 1)
            if (Test-WmtPathExists $quotedPart) { return $quotedPart }
            $clean = $quotedPart
        }
    }
    if (Test-WmtPathExists $clean) { return $clean }
    if ($clean.Contains(" ")) {
        $parts = $clean -split " "
        $candidate = $parts[0]
        $check = {
            param($p)
            if (Test-WmtPathExists $p) { return $true }
            if (Test-WmtPathExists "$p.exe") { return $true }
            return $false
        }
        if (& $check $candidate) { return $candidate }
        for ($i = 1; $i -lt $parts.Count; $i++) {
            $candidate += " " + $parts[$i]
            if (& $check $candidate) { return $candidate }
        }
    }
    return $clean
}

function Test-WmtRegistryFlagEnabled {
    param($Value)
    if ($null -eq $Value) { return $false }
    try { return ([int]$Value -eq 1) }
    catch { return ([string]$Value -eq "1") }
}

function Test-WmtWindowsInstallerCommand {
    param([string]$RawString)
    if ([string]::IsNullOrWhiteSpace($RawString)) { return $false }
    $expanded = [Environment]::ExpandEnvironmentVariables($RawString.Trim())
    return ($expanded -match '(?i)(^|[\\/"\s])msiexec(\.exe)?(?=$|[\s/"])')
}

function Get-WmtUninstallExecutablePath {
    param([string]$RawString)
    if ([string]::IsNullOrWhiteSpace($RawString)) { return $null }
    $cmd = [Environment]::ExpandEnvironmentVariables($RawString.Trim())
    if ([string]::IsNullOrWhiteSpace($cmd) -or (Test-WmtWindowsInstallerCommand $cmd)) { return $null }

    $candidate = $null
    if ($cmd -match '^\s*"([^"]+)"') {
        $candidate = $matches[1].Trim()
    }
    elseif ($cmd -match '^\s*(?<Path>[a-zA-Z]:\\.+?\.(?:exe|msi|cmd|bat|com))(?=$|\s)') {
        $candidate = $matches.Path.Trim()
    }
    elseif ($cmd -match '^\s*(?<Path>[a-zA-Z]:\\[^\s]+)') {
        $candidate = $matches.Path.Trim()
    }

    if ([string]::IsNullOrWhiteSpace($candidate)) { return $null }
    $candidate = [Environment]::ExpandEnvironmentVariables($candidate)
    if ($candidate -notmatch '^[a-zA-Z]:\\') { return $null }
    if ($candidate -match '%.*%' -or $candidate -match '\$\(.*\)') { return $null }
    return $candidate
}

function Invoke-WmtRegistryDeepCleanScan {
    param([string[]]$SelectedTargets)

    Import-Module Microsoft.PowerShell.Management
    Import-Module Microsoft.PowerShell.Security

    $SelectedTargets = @($SelectedTargets | Select-Object -Unique)
    $findings = New-Object System.Collections.Generic.List[object]
    if ($SelectedTargets.Count -eq 0) { return $findings.ToArray() }

    $categoryWeight = 100.0 / $SelectedTargets.Count
    $currentBaseProgress = 0.0

    $tick = {
        param($DetailText, $CurrentIndex, $TotalEstimated)
        if ($CurrentIndex % 50 -eq 0) {
            $safeTotal = [math]::Max(1, $TotalEstimated)
            $fraction = [math]::Min(($CurrentIndex / $safeTotal), 1.0)
            $progress = [int]($currentBaseProgress + ($fraction * $categoryWeight))
            Write-Progress -Activity "Deep registry scan" -Status $DetailText -PercentComplete $progress
        }
    }

    $endCategory = {
        $currentBaseProgress += $categoryWeight
        Write-Progress -Activity "Deep registry scan" -Status "Completed category" -PercentComplete ([int]$currentBaseProgress)
    }

    if ($SelectedTargets -contains "ActiveX") {
        Write-WmtStatus "Scanning ActiveX/COM..."
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Default)
        $clsidKey = $root.OpenSubKey("CLSID", $false)
        if ($clsidKey) {
            $subKeys = $clsidKey.GetSubKeyNames(); $total = $subKeys.Count; $i = 0
            foreach ($id in $subKeys) {
                $i++; & $tick "Scanning CLSID: $id" $i $total
                try {
                    $sub = $clsidKey.OpenSubKey("$id\InProcServer32", $false)
                    if ($sub) {
                        $dll = $sub.GetValue($null)
                        if ($dll -and $dll -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $dll)) {
                            $cleanDll = Get-WmtRealExecutablePath $dll
                            if (-not (Test-WmtPathExists $cleanDll)) { [void]$findings.Add([PSCustomObject]@{ Problem = "ActiveX Issue"; Data = $cleanDll; DisplayKey = $id; RegPath = "HKCR:\CLSID\$id\InProcServer32"; ValueName = $null; Type = "Key" }) }
                        }
                        $sub.Close()
                    }
                    $sub2 = $clsidKey.OpenSubKey("$id\LocalServer32", $false)
                    if ($sub2) {
                        $exe = $sub2.GetValue($null)
                        if ($exe -and $exe -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $exe)) {
                            $cleanExe = Get-WmtRealExecutablePath $exe
                            if (-not (Test-WmtPathExists $cleanExe)) { [void]$findings.Add([PSCustomObject]@{ Problem = "ActiveX Issue"; Data = $cleanExe; DisplayKey = $id; RegPath = "HKCR:\CLSID\$id\LocalServer32"; ValueName = $null; Type = "Key" }) }
                        }
                        $sub2.Close()
                    }
                }
                catch {}
            }
            $clsidKey.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Ext") {
        Write-WmtStatus "Scanning file extensions..."
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Default)
        $names = $root.GetSubKeyNames(); $total = $names.Count; $i = 0
        foreach ($ext in $names) {
            $i++; & $tick "Scanning extension: $ext" $i $total
            if ($ext.StartsWith(".")) {
                try {
                    $sub = $root.OpenSubKey($ext)
                    if ($sub.SubKeyCount -eq 0 -and $null -eq $sub.GetValue($null)) {
                        [void]$findings.Add([PSCustomObject]@{ Problem = "Unused Extension"; Data = $ext; DisplayKey = $ext; RegPath = "HKCR:\$ext"; ValueName = $null; Type = "Key" })
                    }
                    $sub.Close()
                }
                catch {}
            }
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "FileExts") {
        Write-WmtStatus "Scanning user file associations..."
        $path = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
        $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($path)
        if ($root) {
            $names = $root.GetSubKeyNames(); $total = $names.Count; $i = 0
            foreach ($ext in $names) {
                $i++; & $tick "Scanning user extension: $ext" $i $total
                $owl = $root.OpenSubKey("$ext\OpenWithList")
                if ($owl) {
                    foreach ($valName in $owl.GetValueNames()) {
                        if ($valName -match "^[a-z]$") {
                            $val = $owl.GetValue($valName)
                            if ($val -match '^[a-zA-Z]:\\') {
                                $cleanPath = Get-WmtRealExecutablePath $val
                                if (-not (Test-WmtPathExists $cleanPath)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Invalid FileExt MRU"; Data = $cleanPath; DisplayKey = $ext; RegPath = "HKCU:\$path\$ext\OpenWithList"; ValueName = $valName; Type = "Value" }) }
                            }
                        }
                    }
                    $owl.Close()
                }
            }
            $root.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "AppPaths") {
        Write-WmtStatus "Scanning app paths..."
        $key = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
        $root = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($key)
        if ($root) {
            $names = $root.GetSubKeyNames(); $total = $names.Count; $i = 0
            foreach ($app in $names) {
                $i++; & $tick "Scanning app path: $app" $i $total
                try {
                    $sub = $root.OpenSubKey($app)
                    $path = $sub.GetValue($null)
                    if ($path -and $path -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $path)) {
                        $clean = Get-WmtRealExecutablePath $path
                        if (-not (Test-WmtPathExists $clean)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Missing App Path"; Data = $clean; DisplayKey = $app; RegPath = "HKLM:\$key\$app"; ValueName = $null; Type = "Key" }) }
                    }
                    $sub.Close()
                }
                catch {}
            }
            $root.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Apps" -or $SelectedTargets -contains "ProgIDs") {
        Write-WmtStatus "Scanning applications and ProgIDs..."
        $searchRoots = @()
        if ($SelectedTargets -contains "Apps") {
            $cr = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Default)
            $searchRoots += $cr.OpenSubKey("Applications")
        }
        if ($SelectedTargets -contains "ProgIDs") {
            $searchRoots += [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Default)
        }

        foreach ($root in $searchRoots) {
            if ($root) {
                $names = $root.GetSubKeyNames(); $total = $names.Count; $i = 0
                foreach ($app in $names) {
                    $i++; & $tick "Scanning app command: $app" $i $total
                    try {
                        $appKey = $root.OpenSubKey($app)
                        $shellKey = $appKey.OpenSubKey("shell")
                        if ($shellKey) {
                            foreach ($verb in $shellKey.GetSubKeyNames()) {
                                try {
                                    $cmdKey = $shellKey.OpenSubKey("$verb\command")
                                    if ($cmdKey) {
                                        $cmd = $cmdKey.GetValue($null)
                                        if ($cmd) {
                                            $clean = Get-WmtRealExecutablePath $cmd
                                            if ($clean -and $clean -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $clean)) {
                                                if (-not (Test-WmtPathExists $clean) -and $clean -notmatch "%1") {
                                                    [void]$findings.Add([PSCustomObject]@{ Problem = "Invalid App Command ($verb)"; Data = $clean; DisplayKey = $app; RegPath = "$($root.Name)\$app\shell\$verb\command"; ValueName = $null; Type = "Key" })
                                                }
                                            }
                                        }
                                        $cmdKey.Close()
                                    }
                                }
                                catch {}
                            }
                            $shellKey.Close()
                        }
                        $appKey.Close()
                    }
                    catch {}
                }
            }
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Uninstall") {
        Write-WmtStatus "Scanning uninstallers..."
        $paths = @("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
        $hives = @([Microsoft.Win32.Registry]::LocalMachine, [Microsoft.Win32.Registry]::CurrentUser)
        $total = 500; $i = 0
        foreach ($h in $hives) {
            foreach ($p in $paths) {
                try {
                    $rk = $h.OpenSubKey($p)
                    if ($rk) {
                        foreach ($subName in $rk.GetSubKeyNames()) {
                            $i++; & $tick "Scanning uninstaller: $subName" $i $total
                            $sub = $rk.OpenSubKey($subName)
                            $uString = $sub.GetValue("UninstallString")
                            $displayName = [string]$sub.GetValue("DisplayName")
                            $isWindowsInstaller = (Test-WmtRegistryFlagEnabled $sub.GetValue("WindowsInstaller")) -or (Test-WmtWindowsInstallerCommand $uString)
                            $isSystemComponent = Test-WmtRegistryFlagEnabled $sub.GetValue("SystemComponent")
                            $isNoRemove = Test-WmtRegistryFlagEnabled $sub.GetValue("NoRemove")
                            if ($uString -and -not ([string]::IsNullOrWhiteSpace($displayName)) -and -not $isWindowsInstaller -and -not $isSystemComponent -and -not $isNoRemove) {
                                $clean = Get-WmtUninstallExecutablePath $uString
                                if ($clean -and -not (Test-WmtRegistryCleanerWhitelist $clean) -and -not (Test-WmtPathExists $clean)) {
                                    $rootName = if ($h.Name -match "HKEY_LOCAL_MACHINE") { "HKLM" } else { "HKCU" }
                                    [void]$findings.Add([PSCustomObject]@{ Problem = "Missing Uninstaller"; Data = $clean; DisplayKey = $displayName; RegPath = "$rootName`:\$p\$subName"; ValueName = $null; Type = "Key" })
                                }
                            }
                            $sub.Close()
                        }
                        $rk.Close()
                    }
                }
                catch {}
            }
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "MuiCache") {
        Write-WmtStatus "Scanning MuiCache..."
        $key = "Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache"
        $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($key)
        if ($root) {
            $names = $root.GetValueNames(); $total = $names.Count; $i = 0
            foreach ($valName in $names) {
                $i++; & $tick "Scanning MuiCache: $valName" $i $total
                $cleanPath = $valName -replace '\.(FriendlyAppName|ApplicationCompany)$', ''
                if ($cleanPath -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $cleanPath)) {
                    if (-not (Test-WmtPathExists $cleanPath)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Obsolete MuiCache"; Data = $cleanPath; DisplayKey = "MuiCache"; RegPath = "HKCU:\$key"; ValueName = $valName; Type = "Value" }) }
                }
            }
            $root.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "AppCompat") {
        Write-WmtStatus "Scanning compatibility store..."
        $key = "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store"
        $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($key)
        if ($root) {
            $names = $root.GetValueNames(); $total = $names.Count; $i = 0
            foreach ($valName in $names) {
                $i++; & $tick "Scanning AppCompat: $valName" $i $total
                if ($valName -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $valName)) {
                    if (-not (Test-WmtPathExists $valName)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Obsolete Compatibility Ref"; Data = $valName; DisplayKey = "AppCompat"; RegPath = "HKCU:\$key"; ValueName = $valName; Type = "Value" }) }
                }
            }
            $root.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Firewall") {
        Write-WmtStatus "Scanning firewall rules..."
        $key = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules"
        $root = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($key)
        if ($root) {
            $names = $root.GetValueNames(); $total = $names.Count; $i = 0
            foreach ($valName in $names) {
                $i++; & $tick "Scanning firewall: $valName" $i $total
                $data = $root.GetValue($valName)
                if ($data -match "App=([^|]+)") {
                    $appPath = $matches[1]
                    $expanded = [Environment]::ExpandEnvironmentVariables($appPath)
                    if ($expanded -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $expanded)) {
                        if (-not (Test-WmtPathExists $expanded)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Invalid Firewall Rule"; Data = $expanded; DisplayKey = $valName; RegPath = "HKLM:\$key"; ValueName = $valName; Type = "Value" }) }
                    }
                }
            }
            $root.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Services") {
        Write-WmtStatus "Scanning services..."
        $key = "SYSTEM\CurrentControlSet\Services"
        $root = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($key)
        if ($root) {
            $names = $root.GetSubKeyNames(); $total = $names.Count; $i = 0
            foreach ($svc in $names) {
                $i++; & $tick "Scanning service: $svc" $i $total
                try {
                    $sub = $root.OpenSubKey($svc)
                    $img = $sub.GetValue("ImagePath")
                    if ($img -and $img -match '^[a-zA-Z]:\\' -and $img -notmatch "\\drivers\\") {
                        $clean = Get-WmtRealExecutablePath $img
                        if (-not (Test-WmtPathExists $clean)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Missing Service Binary"; Data = $clean; DisplayKey = $svc; RegPath = "HKLM:\$key\$svc"; ValueName = "ImagePath"; Type = "Value" }) }
                    }
                    $sub.Close()
                }
                catch {}
            }
            $root.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "TypeLib") {
        Write-WmtStatus "Scanning type libraries..."
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Default)
        $tlKey = $root.OpenSubKey("TypeLib", $false)
        if ($tlKey) {
            $names = $tlKey.GetSubKeyNames(); $total = $names.Count; $i = 0
            foreach ($guid in $names) {
                $i++; & $tick "Scanning TypeLib: $guid" $i $total
                try {
                    $verKey = $tlKey.OpenSubKey($guid)
                    if ($verKey) {
                        foreach ($ver in $verKey.GetSubKeyNames()) {
                            $numKey = $verKey.OpenSubKey($ver)
                            $helpDir = $numKey.GetValue("HELPDIR")
                            if ($helpDir -and $helpDir -match '^[a-zA-Z]:\\' -and -not (Test-WmtPathExists $helpDir)) {
                                [void]$findings.Add([PSCustomObject]@{ Problem = "Missing HelpDir"; Data = $helpDir; DisplayKey = "$guid"; RegPath = "HKCR:\TypeLib\$guid\$ver"; ValueName = "HELPDIR"; Type = "Value" })
                            }
                            $numKey.Close()
                        }
                        $verKey.Close()
                    }
                }
                catch {}
            }
            $tlKey.Close()
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Icons") {
        Write-WmtStatus "Scanning default icons..."
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Default)
        $names = $root.GetSubKeyNames(); $total = $names.Count; $i = 0
        foreach ($ext in $names) {
            $i++; & $tick "Scanning icon: $ext" $i $total
            try {
                $iconKey = $root.OpenSubKey("$ext\DefaultIcon")
                if ($iconKey) {
                    $val = $iconKey.GetValue($null)
                    if ($val) {
                        $cleanPath = Get-WmtRealExecutablePath $val
                        if ($cleanPath -match '^[a-zA-Z]:\\' -and -not (Test-WmtRegistryCleanerWhitelist $cleanPath) -and -not (Test-WmtPathExists $cleanPath) -and $cleanPath -notmatch "%1") {
                            [void]$findings.Add([PSCustomObject]@{ Problem = "Invalid Default Icon"; Data = $cleanPath; DisplayKey = $ext; RegPath = "HKCR:\$ext\DefaultIcon"; ValueName = $null; Type = "Key" })
                        }
                    }
                    $iconKey.Close()
                }
            }
            catch {}
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "SharedDLLs") {
        Write-WmtStatus "Scanning shared DLLs..."
        $keys = @("SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDlls", "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\SharedDlls")
        foreach ($k in $keys) {
            $rk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($k)
            if ($rk) {
                $names = $rk.GetValueNames(); $total = $names.Count; $i = 0
                foreach ($val in $names) {
                    $i++; & $tick "Scanning DLL: $val" $i $total
                    if ($val -match '^[a-zA-Z]:\\' -and -not(Test-WmtPathExists $val)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Missing Shared Ref"; Data = $val; DisplayKey = "SharedDlls"; RegPath = "HKLM:\$k"; ValueName = $val; Type = "Value" }) }
                }
                $rk.Close()
            }
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Startup") {
        Write-WmtStatus "Scanning startup items..."
        $paths = @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Run", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run")
        foreach ($p in $paths) {
            if (Test-Path $p) {
                $props = Get-ItemProperty $p
                $names = $props.PSObject.Properties.Name; $total = $names.Count; $i = 0
                foreach ($n in $names) {
                    $i++; & $tick "Scanning startup: $n" $i $total
                    $v = $props.$n
                    if ($v -is [string] -and $v -match '^[a-zA-Z]:\\') {
                        $cleanExe = Get-WmtRealExecutablePath $v
                        if (-not (Test-WmtPathExists $cleanExe)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Broken Startup"; Data = $cleanExe; DisplayKey = $n; RegPath = $p; ValueName = $n; Type = "Value" }) }
                    }
                }
            }
        }
        & $endCategory
    }

    if ($SelectedTargets -contains "Installer") {
        Write-WmtStatus "Scanning installer folders..."
        $k = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders"
        $rk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($k)
        if ($rk) {
            $names = $rk.GetValueNames(); $total = $names.Count; $i = 0
            foreach ($val in $names) {
                $i++; & $tick "Scanning installer: $val" $i $total
                if ($val -match '^[a-zA-Z]:\\' -and -not(Test-WmtPathExists $val)) { [void]$findings.Add([PSCustomObject]@{ Problem = "Missing Installer Folder"; Data = $val; DisplayKey = "Installer"; RegPath = "HKLM:\$k"; ValueName = $val; Type = "Value" }) }
            }
            $rk.Close()
        }
        & $endCategory
    }

    Write-Progress -Activity "Deep registry scan" -Completed
    return $findings.ToArray()
}

function Invoke-WmtRegistryDeepCleanRemoval {
    param(
        [object[]]$Items,
        [string]$BackupDirectory
    )

    if (-not (Test-Path -LiteralPath $BackupDirectory)) {
        New-Item -Path $BackupDirectory -ItemType Directory -Force | Out-Null
    }

    $bkFile = Join-Path $BackupDirectory ("DeepClean_Backup_{0}.reg" -f (Get-Date -Format "yyyyMMdd_HHmm"))
    Set-Content -Path $bkFile -Value "Windows Registry Editor Version 5.00`r`n`r`n" -Encoding Unicode

    $fixed = 0
    $skipped = 0
    $backedUpKeys = @()
    $count = @($Items).Count
    $index = 0

    foreach ($item in $Items) {
        $index++
        Write-Progress -Activity "Deep registry clean" -Status "Removing $index/$($count): $($item.DisplayKey)" -PercentComplete ([int](($index / [math]::Max(1, $count)) * 100))
        $uid = "$($item.RegPath):$($item.ValueName)"
        if ($uid -notin $backedUpKeys) {
            Backup-WmtRegKey -ItemObj $item -FilePath $bkFile
            $backedUpKeys += $uid
        }

        Write-WmtStatus "Removing: $($item.DisplayKey)"
        $isKey = ($item.Type -eq "Key")
        $success = Remove-WmtRegKeyForced -Path $item.RegPath -IsKey $isKey -ValName $item.ValueName

        if ($success) {
            $fixed++
            Write-WmtStatus "Removed: $($item.RegPath)"
        }
        else {
            $skipped++
            Write-WmtStatus "Failed to remove: $($item.RegPath)"
        }
    }

    Write-Progress -Activity "Deep registry clean" -Completed
    [PSCustomObject]@{
        Fixed   = $fixed
        Skipped = $skipped
        Backup  = $bkFile
    }
}

if ($ListTargets) {
    Get-WmtRegistryDeepCleanCategories | ForEach-Object {
        $_.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Key; Target = $_.Value }
        }
    }
    return
}

$isAdmin = Test-WmtIsAdministrator
if (-not $isAdmin) {
    if ($Clean) {
        throw "Clean mode requires an elevated PowerShell session. Re-run PowerShell as Administrator."
    }
    Write-Warning "Not running as Administrator. Scan-only mode will continue, but HKLM/HKCR results may be incomplete."
}

Ensure-WmtRegistryDrives
if ($isAdmin) { Enable-WmtRegistryPrivileges }

if ([string]::IsNullOrWhiteSpace($BackupDirectory)) {
    $BackupDirectory = Get-WmtDefaultRegistryBackupDirectory
}

$Targets = @($Targets | Select-Object -Unique)
Write-WmtStatus "Deep registry scan targets: $($Targets -join ', ')"
$findings = @(Invoke-WmtRegistryDeepCleanScan -SelectedTargets $Targets)

Write-Host ""
Write-WmtStatus "Registry Cleaner: found $($findings.Count) issue(s)."

if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
    $csvFullPath = [System.IO.Path]::GetFullPath($CsvPath)
    $csvDir = Split-Path -Parent $csvFullPath
    if (-not [string]::IsNullOrWhiteSpace($csvDir) -and -not (Test-Path -LiteralPath $csvDir)) {
        New-Item -Path $csvDir -ItemType Directory -Force | Out-Null
    }
    $findings | Export-Csv -Path $csvFullPath -NoTypeInformation -Encoding UTF8
    Write-WmtStatus "Findings exported to: $csvFullPath"
}

if ($findings.Count -gt 0 -and $PreviewLimit -gt 0) {
    $findings |
        Select-Object -First $PreviewLimit Problem, Data, DisplayKey, RegPath, ValueName, Type |
        Format-Table -AutoSize |
        Out-Host

    if ($findings.Count -gt $PreviewLimit) {
        Write-WmtStatus "Showing first $PreviewLimit of $($findings.Count) issue(s). Use -CsvPath for the full list."
    }
}

if (-not $Clean) {
    Write-WmtStatus "Scan only. Re-run with -Clean to remove the detected registry entries."
    return
}

if ($findings.Count -eq 0) {
    Write-WmtStatus "Nothing to clean."
    return
}

if ($CreateRestorePoint) {
    Write-WmtStatus "Creating restore point..."
    try {
        Checkpoint-Computer -Description "WMT DeepClean" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        Write-WmtStatus "Restore point created."
    }
    catch {
        Write-WmtStatus "Restore point failed or is disabled. Continuing with registry export backup."
    }
}

if (-not $Force) {
    Write-Host ""
    Write-Host "This will remove $($findings.Count) registry item(s). A .reg backup will be saved first." -ForegroundColor Yellow
    $answer = Read-Host "Type DELETE to continue"
    if ($answer -ne "DELETE") {
        Write-WmtStatus "Clean cancelled."
        return
    }
}

$summary = Invoke-WmtRegistryDeepCleanRemoval -Items $findings -BackupDirectory $BackupDirectory

Write-Host ""
Write-WmtStatus "Cleanup complete. Fixed: $($summary.Fixed); Skipped: $($summary.Skipped)"
Write-WmtStatus "Backup saved to: $($summary.Backup)"
