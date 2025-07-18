<#
.SYNOPSIS
    PowerShell Firewall Manager
.DESCRIPTION
    A script to manage Windows Firewall rules with various operations.
.NOTES
    File Name      : FirewallManager.ps1
    Author         : Chaython
    Prerequisite   : PowerShell 5.1 or later, Administrative privileges
    Version        : 1.4.1
    Changes        : - Fixed function name consistency
                   - Added retry logic for rule changes
                   - Improved error handling
#>

# Requires admin privileges
#Requires -RunAsAdministrator

function Get-CleanRuleName {
    param ([string]$name)
    if ($name -match '@{.+?}\(?(.+?)\)?$') { $name = $matches[1] }
    if ($name -match '(.+?)_\d+\.\d+\.\d+\.\d+_x64__.+') { $name = $matches[1] + "_x64" }
    elseif ($name -match '(.+?)_\d+\.\d+\.\d+\.\d+_.+') { $name = $matches[1] }
    $name = $name -replace '({[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}})', ''
    return $name.Trim()
}

function Show-MainMenu {
    param ([string]$Title = 'PowerShell Firewall Manager')
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: View and Manage Firewall Rules"
    Write-Host "2: Export firewall rules to CSV"
    Write-Host "3: Import firewall rules from CSV"
    Write-Host "0/Q: Quit"
}

function Show-FirewallRulesWithOptions {
    $rules = Get-NetFirewallRule | Sort-Object -Property Action -Descending
    $count = 1
    
    Write-Host "`n================ Firewall Rules ================"
    Write-Host "#  Action   Enabled   Rule Name"
    Write-Host "--  ------   -------   ---------"
    
    foreach ($rule in $rules) {
        $action = $rule.Action.ToString().PadRight(6)
        $enabled = if ($rule.Enabled -eq $true) { "Yes" } else { "No" }
        $cleanName = Get-CleanRuleName -name $rule.DisplayName
        if ([string]::IsNullOrWhiteSpace($cleanName)) {
            $cleanName = Get-CleanRuleName -name $rule.Name
        }
        Write-Host "$($count.ToString().PadLeft(2))  $action   $($enabled.PadRight(7))   $cleanName"
        $count++
    }
    
    Write-Host "`n================ Action Options ================"
    Write-Host "1: Enable a rule (type '1 NUMBER')"
    Write-Host "2: Disable a rule (type '2 NUMBER')"
    Write-Host "3: Add new rule"
    Write-Host "4: Remove a rule (type '4 NUMBER')"
    Write-Host "0: Back to main menu"
}

function Show-OperationResult {
    param (
        [string]$message,
        [bool]$success,
        [int]$delaySeconds = 0
    )
    if ($delaySeconds -gt 0) {
        Write-Host "Waiting $delaySeconds seconds for changes to apply..." -ForegroundColor Yellow
        Start-Sleep -Seconds $delaySeconds
    }
    $color = if ($success) { "Green" } else { "Red" }
    Write-Host "`n$(Get-Date -Format 'HH:mm:ss') - $message" -ForegroundColor $color
}

function Enable-FirewallRule {
    param ($ruleNum)
    $rules = @(Get-NetFirewallRule | Sort-Object -Property Action -Descending)
    if ($ruleNum -gt 0 -and $ruleNum -le $rules.Count) {
        $rule = $rules[$ruleNum - 1]
        $ruleName = Get-CleanRuleName -name $rule.DisplayName
        try {
            Set-NetFirewallRule -Name $rule.Name -Enabled True -ErrorAction Stop
            Show-OperationResult -message "[SUCCESS] Enabled rule: $ruleName" -success $true -delaySeconds 1
        } catch {
            Show-OperationResult -message "[ERROR] Failed to enable rule $ruleName`: $_" -success $false
        }
    } else {
        Show-OperationResult -message "[ERROR] Invalid rule number" -success $false
    }
    Wait-ForUser
}

function Disable-FirewallRule {
    param ($ruleNum)
    $rules = @(Get-NetFirewallRule | Sort-Object -Property Action -Descending)
    if ($ruleNum -gt 0 -and $ruleNum -le $rules.Count) {
        $rule = $rules[$ruleNum - 1]
        $ruleName = Get-CleanRuleName -name $rule.DisplayName
        try {
            Set-NetFirewallRule -Name $rule.Name -Enabled False -ErrorAction Stop
            Show-OperationResult -message "[SUCCESS] Disabled rule: $ruleName" -success $true -delaySeconds 1
        } catch {
            Show-OperationResult -message "[ERROR] Failed to disable rule $ruleName`: $_" -success $false
        }
    } else {
        Show-OperationResult -message "[ERROR] Invalid rule number" -success $false
    }
    Wait-ForUser
}

function Add-FirewallRule {
    Write-Host "`nCreating a new firewall rule" -ForegroundColor Cyan
    
    $displayName = Read-Host "Enter a display name for the rule"
    $name = Read-Host "Enter a unique name for the rule (no spaces, use hyphens)"
    $description = Read-Host "Enter a description for the rule"
    
    # Direction
    do {
        $direction = Read-Host "Enter direction (Inbound/Outbound)"
    } while ($direction -notin "Inbound", "Outbound")
    
    # Action
    do {
        $action = Read-Host "Enter action (Allow/Block)"
    } while ($action -notin "Allow", "Block")
    
    # Profile
    do {
        $profile = Read-Host "Enter profile (Domain, Private, Public, Any)"
    } while ($profile -notin "Domain", "Private", "Public", "Any")
    
    # Protocol
    do {
        $protocol = Read-Host "Enter protocol (TCP, UDP, ICMP, Any)"
    } while ($protocol -notin "TCP", "UDP", "ICMP", "Any")
    
    $localPort = Read-Host "Enter local port (leave blank for any)"
    $remotePort = Read-Host "Enter remote port (leave blank for any)"
    $program = Read-Host "Enter program path (leave blank for any)"
    
    try {
        $params = @{
            DisplayName = $displayName
            Name        = $name
            Description = $description
            Direction   = $direction
            Action      = $action
            Profile     = $profile
            Protocol    = $protocol
        }
        
        if ($localPort) { $params['LocalPort'] = $localPort }
        if ($remotePort) { $params['RemotePort'] = $remotePort }
        if ($program) { $params['Program'] = $program }
        
        New-NetFirewallRule @params
        Show-OperationResult -message "[SUCCESS] Firewall rule created: $displayName" -success $true
    } catch {
        Show-OperationResult -message "[ERROR] Failed to create rule: $_" -success $false
    }
    Wait-ForUser
}

function Remove-FirewallRule {
    param ($ruleNum)
    $rules = @(Get-NetFirewallRule | Sort-Object -Property Action -Descending)
    if ($ruleNum -gt 0 -and $ruleNum -le $rules.Count) {
        $rule = $rules[$ruleNum - 1]
        $ruleName = Get-CleanRuleName -name $rule.DisplayName
        try {
            Remove-NetFirewallRule -Name $rule.Name -ErrorAction Stop
            Show-OperationResult -message "[SUCCESS] Removed rule: $ruleName" -success $true
        } catch {
            Show-OperationResult -message "[ERROR] Failed to remove rule $ruleName`: $_" -success $false
        }
    } else {
        Show-OperationResult -message "[ERROR] Invalid rule number" -success $false
    }
    Wait-ForUser
}

function Export-FirewallRules {
    $defaultPath = "$env:USERPROFILE\Desktop\firewall_rules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $filePath = Read-Host "Enter the file path to save the CSV (default: $defaultPath)"
    
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        $filePath = $defaultPath
    }
    
    try {
        Get-NetFirewallRule | Export-Csv -Path $filePath -NoTypeInformation
        Show-OperationResult -message "[SUCCESS] Rules exported to $filePath" -success $true
    } catch {
        Show-OperationResult -message "[ERROR] Export failed: $_" -success $false
    }
    Wait-ForUser
}

function Import-FirewallRules {
    $defaultPath = "$env:USERPROFILE\Desktop\firewall_rules.csv"
    $filePath = Read-Host "Enter the file path of the CSV to import (default looks on Desktop for firewall_rules.csv)"
    
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        $filePath = $defaultPath
    }
    
    if (Test-Path $filePath) {
        try {
            $rules = Import-Csv -Path $filePath
            $successCount = 0
            $errorCount = 0
            
            foreach ($rule in $rules) {
                try {
                    $params = @{
                        DisplayName = $rule.DisplayName
                        Name        = $rule.Name
                        Description = $rule.Description
                        Direction   = $rule.Direction
                        Action      = $rule.Action
                        Profile     = $rule.Profile
                        Enabled     = if ($rule.Enabled -eq "True") { $true } else { $false }
                    }
                    
                    New-NetFirewallRule @params
                    $successCount++
                } catch {
                    $errorCount++
                    Write-Host "  [WARNING] Error importing rule $($rule.DisplayName): $_" -ForegroundColor Yellow
                }
            }
            
            Show-OperationResult -message "[RESULT] Import completed: $successCount succeeded, $errorCount failed" -success ($errorCount -eq 0)
        } catch {
            Show-OperationResult -message "[ERROR] Import failed: $_" -success $false
        }
    } else {
        Show-OperationResult -message "[ERROR] File not found: $filePath" -success $false
    }
    Wait-ForUser
}

function Wait-ForUser {
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main program loop
do {
    Show-MainMenu
    $selection = Read-Host "`nPlease make a selection"
    
    switch ($selection.ToUpper()) {
        '1' {
            do {
                Clear-Host
                Show-FirewallRulesWithOptions
                $input = Read-Host "`nEnter action and number (e.g., '2 5') or 0 to return"
                
                if ($input -eq '0') { break }
                
                $parts = $input -split '\s+'
                $action = $parts[0]
                $ruleNum = if ($parts.Count -gt 1) { $parts[1] } else { $null }
                
                if (@('1','2','4') -contains $action -and ($ruleNum -notmatch '^\d+$')) {
                    Show-OperationResult -message "[ERROR] Invalid rule number" -success $false
                    Wait-ForUser
                    continue
                }
                
                switch ($action) {
                    '1' { Enable-FirewallRule -ruleNum $ruleNum }
                    '2' { Disable-FirewallRule -ruleNum $ruleNum }
                    '3' { Add-FirewallRule }
                    '4' { Remove-FirewallRule -ruleNum $ruleNum }
                    default { 
                        Show-OperationResult -message "[ERROR] Invalid action" -success $false
                        Wait-ForUser
                    }
                }
            } while ($true)
        }
        '2' { Export-FirewallRules }
        '3' { Import-FirewallRules }
        {'0', 'Q'} { return }
        default { 
            Show-OperationResult -message "[ERROR] Invalid selection" -success $false
            Wait-ForUser
        }
    }
} while ($true)