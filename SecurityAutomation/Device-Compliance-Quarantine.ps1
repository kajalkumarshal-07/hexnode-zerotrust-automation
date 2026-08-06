<#
.SYNOPSIS
    Zero Trust Device Compliance Checker & Quarantine Automation
.DESCRIPTION
    Checks device compliance status and automatically quarantines non-compliant devices
    by moving them to a restricted policy group via Hexnode API.
.NOTES
    Run this script via Hexnode Custom Scripts or as a scheduled task.
#>

param(
    [string]$ConfigPath = "C:\Hexnode-ZeroTrust\Config\hexnode_config.json"
)

# Load configuration
$Config = Get-Content $ConfigPath | ConvertFrom-Json
$ApiUrl = $Config.hexnode.api_url
$ApiKey = $Config.hexnode.api_key
$QuarantineGroupId = $Config.quarantine.quarantine_group_id

# Compliance checks
$ComplianceResults = @{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    DeviceName = $env:COMPUTERNAME
    Checks = @()
    IsCompliant = $true
    Violations = @()
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $LogEntry = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    Write-Host $LogEntry
    return $LogEntry
}

function Test-FirewallEnabled {
    try {
        $FirewallProfiles = Get-NetFirewallProfile -ErrorAction Stop
        $AllEnabled = ($FirewallProfiles | Where-Object { $_.Enabled -eq $true }).Count -eq $FirewallProfiles.Count
        return @{
            Check = "Firewall Enabled"
            Status = if ($AllEnabled) { "PASS" } else { "FAIL" }
            Details = ($FirewallProfiles | ForEach-Object { "$($_.Name): $($_.Enabled)" }) -join ", "
        }
    } catch {
        return @{ Check = "Firewall Enabled"; Status = "ERROR"; Details = $_.Exception.Message }
    }
}

function Test-BitLockerEncryption {
    try {
        $EncryptionStatus = Get-BitLockerVolume -MountPoint "C:" -ErrorAction Stop
        $IsEncrypted = $EncryptionStatus.ProteStatus -eq "On"
        return @{
            Check = "Disk Encryption (BitLocker)"
            Status = if ($IsEncrypted) { "PASS" } else { "FAIL" }
            Details = "Protection Status: $($EncryptionStatus.ProteStatus)"
        }
    } catch {
        return @{ Check = "Disk Encryption (BitLocker)"; Status = "WARNING"; Details = "BitLocker not available or not configured" }
    }
}

function Test-WindowsDefender {
    try {
        $DefenderStatus = Get-MpComputerStatus -ErrorAction Stop
        $IsRunning = $DefenderStatus.RealTimeProtectionEnabled -eq $true
        $IsUpdated = ($DefenderStatus.AntivirusSignatureLastUpdated -gt (Get-Date).AddDays(-1))
        return @{
            Check = "Windows Defender"
            Status = if ($IsRunning -and $IsUpdated) { "PASS" } elseif ($IsRunning) { "WARNING" } else { "FAIL" }
            Details = "RealTime: $($DefenderStatus.RealTimeProtectionEnabled), LastUpdated: $($DefenderStatus.AntivirusSignatureLastUpdated)"
        }
    } catch {
        return @{ Check = "Windows Defender"; Status = "ERROR"; Details = $_.Exception.Message }
    }
}

function Test-WindowsUpdate {
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        $PendingUpdates = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
        $CriticalUpdates = ($PendingUpdates.Updates | Where-Object { $_.MsrcSeverity -eq "Critical" }).Count
        return @{
            Check = "Windows Updates"
            Status = if ($CriticalUpdates -eq 0) { "PASS" } else { "FAIL" }
            Details = "Pending: $($PendingUpdates.Updates.Count), Critical: $CriticalUpdates"
        }
    } catch {
        return @{ Check = "Windows Updates"; Status = "ERROR"; Details = $_.Exception.Message }
    }
}

function Test-JailbreakRoot {
    $SuspiciousPaths = @(
        "C:\Users\*\AppData\Local\Temp\*superuser*",
        "C:\Users\*\AppData\Local\Temp\*root*",
        "C:\Windows\System32\*rootkit*"
    )
    $FoundSuspicious = $false
    foreach ($Path in $SuspiciousPaths) {
        if (Test-Path $Path) { $FoundSuspicious = $true }
    }
    return @{
        Check = "Root/Jailbreak Detection"
        Status = if (-not $FoundSuspicious) { "PASS" } else { "FAIL" }
        Details = if ($FoundSuspicious) { "Suspicious files detected" } else { "Clean" }
    }
}

function Test-AntiTampering {
    try {
        $Services = @("WinDefend", "MpsSvc", "SecurityHealthService")
        $DisabledServices = @()
        foreach ($Service in $Services) {
            $Svc = Get-Service -Name $Service -ErrorAction SilentlyContinue
            if ($Svc -and $Svc.Status -ne "Running") {
                $DisabledServices += $Service
            }
        }
        return @{
            Check = "Security Services Running"
            Status = if ($DisabledServices.Count -eq 0) { "PASS" } else { "FAIL" }
            Details = if ($DisabledServices.Count -gt 0) { "Disabled: $($DisabledServices -join ', ')" } else { "All running" }
        }
    } catch {
        return @{ Check = "Security Services Running"; Status = "ERROR"; Details = $_.Exception.Message }
    }
}

# Execute all compliance checks
Write-Log "Starting Zero Trust compliance scan on $env:COMPUTERNAME"

$Checks = @(
    (Test-FirewallEnabled),
    (Test-BitLockerEncryption),
    (Test-WindowsDefender),
    (Test-WindowsUpdate),
    (Test-JailbreakRoot),
    (Test-AntiTampering)
)

foreach ($Check in $Checks) {
    $ComplianceResults.Checks += $Check
    Write-Log "$($Check.Check): $($Check.Status) - $($Check.Details)"
    
    if ($Check.Status -eq "FAIL") {
        $ComplianceResults.IsCompliant = $false
        $ComplianceResults.Violations += $Check.Check
    }
}

# Output results
$ComplianceResults | ConvertTo-Json -Depth 3 | Write-Host

# If non-compliant, trigger quarantine via Hexnode API
if (-not $ComplianceResults.IsCompliant) {
    Write-Log "DEVICE NON-COMPLIANT - Triggering quarantine" "WARNING"
    
    # Get device ID from Hexnode (you need to set this or query it)
    $DeviceId = $env:HEXNODE_DEVICE_ID  # Set via Hexnode environment variable
    
    if ($DeviceId) {
        $Body = @{
            devices = @([int]$DeviceId)
            devicegroups = @([int]$QuarantineGroupId)
        } | ConvertTo-Json

        try {
            $Response = Invoke-RestMethod -Uri "$ApiUrl/actions/add_to_device_group/" `
                -Method Post `
                -Headers @{
                    "Authorization" = $ApiKey
                    "Content-Type" = "application/json"
                } -Body $Body
            
            Write-Log "QUARANTINE SUCCESS: Device moved to quarantine group" "CRITICAL"
        } catch {
            Write-Log "QUARANTINE FAILED: $($_.Exception.Message)" "ERROR"
        }
    } else {
        Write-Log "Device ID not found - manual quarantine required" "WARNING"
    }
    
    # Return non-compliant exit code for Hexnode automation triggers
    exit 1
} else {
    Write-Log "DEVICE COMPLIANT - All checks passed" "INFO"
    exit 0
}
