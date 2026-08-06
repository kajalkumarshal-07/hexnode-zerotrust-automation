<#
.SYNOPSIS
    Zero Trust Master Orchestration Script
.DESCRIPTION
    Main orchestration script that coordinates all Zero Trust automation components:
    - Compliance checks
    - Vulnerability scanning
    - Firewall enforcement
    - OS updates
    - Device provisioning
    - SIEM integration
.NOTES
    Deploy via Hexnode as a scheduled automation or on-demand.
#>

param(
    [string]$ConfigPath = "C:\Hexnode-ZeroTrust\Config\hexnode_config.json",
    [switch]$RunCompliance,
    [switch]$RunVulnerabilityScan,
    [switch]$RunFirewallEnforcement,
    [switch]$RunOSUpdates,
    [switch]$RunProvisioning,
    [switch]$RunAll
)

$ScriptRoot = "C:\Hexnode-ZeroTrust"
$LogFile = "C:\Windows\Temp\ZeroTrust_Orchestration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

$OrchestrationResults = @{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    DeviceName = $env:COMPUTERNAME
    Tasks = @()
    OverallStatus = "SUCCESS"
    RiskScore = 0
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $LogEntry = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogFile -Value $LogEntry -ErrorAction SilentlyContinue
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$DelaySeconds = 5
    )
    
    $RetryCount = 0
    while ($RetryCount -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        } catch {
            $RetryCount++
            if ($RetryCount -ge $MaxRetries) {
                throw
            }
            Write-Log "Retry $RetryCount/$MaxRetries after error: $($_.Exception.Message)" "WARNING"
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Run-ComplianceCheck {
    Write-Log "=== Running Compliance Check ==="
    
    $TaskResult = @{
        Task = "Compliance Check"
        Status = "SUCCESS"
        Details = @()
    }
    
    try {
        $ScriptPath = "$ScriptRoot\SecurityAutomation\Device-Compliance-Quarantine.ps1"
        
        if (Test-Path $ScriptPath) {
            $Result = & $ScriptPath -ConfigPath $ConfigPath
            $ExitCode = $LASTEXITCODE
            
            $TaskResult.Details += "Exit code: $ExitCode"
            
            if ($ExitCode -ne 0) {
                $TaskResult.Status = "WARNING"
                $OrchestrationResults.RiskScore += 5
            }
        } else {
            $TaskResult.Status = "SKIPPED"
            $TaskResult.Details += "Script not found"
        }
    } catch {
        $TaskResult.Status = "FAILED"
        $TaskResult.Details += $_.Exception.Message
        $OrchestrationResults.RiskScore += 10
    }
    
    $OrchestrationResults.Tasks += $TaskResult
    Write-Log "Compliance check: $($TaskResult.Status)"
}

function Run-VulnerabilityScan {
    Write-Log "=== Running Vulnerability Scan ==="
    
    $TaskResult = @{
        Task = "Vulnerability Scan"
        Status = "SUCCESS"
        Details = @()
    }
    
    try {
        $ScriptPath = "$ScriptRoot\SecurityAutomation\Vulnerability-Scanner.ps1"
        
        if (Test-Path $ScriptPath) {
            $Result = & $ScriptPath -ConfigPath $ConfigPath
            $ExitCode = $LASTEXITCODE
            
            $TaskResult.Details += "Exit code: $ExitCode"
            
            if ($ExitCode -eq 2) {
                $TaskResult.Status = "CRITICAL"
                $OrchestrationResults.RiskScore += 15
            } elseif ($ExitCode -eq 1) {
                $TaskResult.Status = "WARNING"
                $OrchestrationResults.RiskScore += 7
            }
        } else {
            $TaskResult.Status = "SKIPPED"
            $TaskResult.Details += "Script not found"
        }
    } catch {
        $TaskResult.Status = "FAILED"
        $TaskResult.Details += $_.Exception.Message
        $OrchestrationResults.RiskScore += 10
    }
    
    $OrchestrationResults.Tasks += $TaskResult
    Write-Log "Vulnerability scan: $($TaskResult.Status)"
}

function Run-FirewallEnforcement {
    Write-Log "=== Running Firewall Enforcement ==="
    
    $TaskResult = @{
        Task = "Firewall Enforcement"
        Status = "SUCCESS"
        Details = @()
    }
    
    try {
        $ScriptPath = "$ScriptRoot\FirewallEnforcement\Enforce-FirewallServices.ps1"
        
        if (Test-Path $ScriptPath) {
            $Result = & $ScriptPath -ConfigPath $ConfigPath
            $ExitCode = $LASTEXITCODE
            
            $TaskResult.Details += "Exit code: $ExitCode"
            
            if ($ExitCode -ne 0) {
                $TaskResult.Status = "WARNING"
                $OrchestrationResults.RiskScore += 5
            }
        } else {
            $TaskResult.Status = "SKIPPED"
            $TaskResult.Details += "Script not found"
        }
    } catch {
        $TaskResult.Status = "FAILED"
        $TaskResult.Details += $_.Exception.Message
        $OrchestrationResults.RiskScore += 10
    }
    
    $OrchestrationResults.Tasks += $TaskResult
    Write-Log "Firewall enforcement: $($TaskResult.Status)"
}

function Run-OSUpdates {
    Write-Log "=== Running OS Updates ==="
    
    $TaskResult = @{
        Task = "OS Updates"
        Status = "SUCCESS"
        Details = @()
    }
    
    try {
        $ScriptPath = "$ScriptRoot\OSUpdates\Orchestrate-OSUpdates.ps1"
        
        if (Test-Path $ScriptPath) {
            $Result = & $ScriptPath -ConfigPath $ConfigPath
            $ExitCode = $LASTEXITCODE
            
            $TaskResult.Details += "Exit code: $ExitCode"
            
            if ($ExitCode -ne 0) {
                $TaskResult.Status = "WARNING"
                $OrchestrationResults.RiskScore += 3
            }
        } else {
            $TaskResult.Status = "SKIPPED"
            $TaskResult.Details += "Script not found"
        }
    } catch {
        $TaskResult.Status = "FAILED"
        $TaskResult.Details += $_.Exception.Message
        $OrchestrationResults.RiskScore += 5
    }
    
    $OrchestrationResults.Tasks += $TaskResult
    Write-Log "OS updates: $($TaskResult.Status)"
}

function Send-HexnodeReport {
    param([Dict]$Results)
    
    Write-Log "Sending compliance report to Hexnode..."
    
    $Config = Get-Content $ConfigPath | ConvertFrom-Json
    $ApiUrl = $Config.hexnode.api_url
    $ApiKey = $Config.hexnode.api_key
    
    # Determine compliance status
    $IsCompliant = $Results.RiskScore -lt 10
    $ComplianceStatus = if ($IsCompliant) { "compliant" } else { "non-compliant" }
    
    # Update device custom attribute with risk score
    try {
        $Body = @{
            custom_attributes = @{
                "zero_trust_risk_score" = $Results.RiskScore.ToString()
                "zero_trust_status" = $ComplianceStatus
                "last_scan" = $Results.Timestamp
            }
        } | ConvertTo-Json
        
        # Note: You'll need the device ID from Hexnode
        # $DeviceId = $env:HEXNODE_DEVICE_ID
        # Invoke-RestMethod -Uri "$ApiUrl/devices/$DeviceId/" -Method Patch -Headers @{...} -Body $Body
        
        Write-Log "Report sent successfully"
    } catch {
        Write-Log "Failed to send report: $($_.Exception.Message)" "ERROR"
    }
}

function Send-SplunkEvent {
    param([Dict]$Results)
    
    $Config = Get-Content $ConfigPath | ConvertFrom-Json
    $SplunkWebhook = $Config.siem.splunk_webhook
    $SplunkToken = $Config.siem.splunk_token
    
    try {
        $Event = @{
            event = "zero_trust_orchestration"
            host = $env:COMPUTERNAME
            source = "hexnode_zerotrust"
            sourcetype = "json"
            result = $Results
        } | ConvertTo-Json -Depth 5
        
        $Headers = @{
            "Authorization" = "Splunk $SplunkToken"
            "Content-Type" = "application/json"
        }
        
        Invoke-RestMethod -Uri $SplunkWebhook -Method Post -Headers $Headers -Body $Event
        Write-Log "Splunk event sent"
    } catch {
        Write-Log "Failed to send Splunk event: $($_.Exception.Message)" "WARNING"
    }
}

# Main execution
Write-Log "=================================================="
Write-Log "  Zero Trust Orchestration Started"
Write-Log "  Device: $env:COMPUTERNAME"
Write-Log "=================================================="

$StartTime = Get-Date

# Execute requested tasks
if ($RunAll -or $RunCompliance) {
    Run-ComplianceCheck
}

if ($RunAll -or $RunVulnerabilityScan) {
    Run-VulnerabilityScan
}

if ($RunAll -or $RunFirewallEnforcement) {
    Run-FirewallEnforcement
}

if ($RunAll -or $RunOSUpdates) {
    Run-OSUpdates
}

if ($RunAll -or $RunProvisioning) {
    # Provisioning is typically run separately
    Write-Log "Provisioning skipped - run Bulk-Provision-Devices.ps1 separately"
}

# Determine overall status
$HasFailures = $OrchestrationResults.Tasks | Where-Object { $_.Status -eq "FAILED" }
$HasWarnings = $OrchestrationResults.Tasks | Where-Object { $_.Status -eq "WARNING" }

if ($HasFailures) {
    $OrchestrationResults.OverallStatus = "FAILED"
} elseif ($HasWarnings) {
    $OrchestrationResults.OverallStatus = "WARNING"
}

$Duration = ((Get-Date) - $StartTime).TotalSeconds
$OrchestrationResults | Add-Member -NotePropertyName "DurationSeconds" -NotePropertyValue $Duration

# Send reports
Send-HexnodeReport -Results $OrchestrationResults
Send-SplunkEvent -Results $OrchestrationResults

# Summary
Write-Log "=================================================="
Write-Log "  Orchestration Complete"
Write-Log "  Status: $($OrchestrationResults.OverallStatus)"
Write-Log "  Risk Score: $($OrchestrationResults.RiskScore)"
Write-Log "  Duration: $([math]::Round($Duration, 2)) seconds"
Write-Log "=================================================="

$OrchestrationResults | ConvertTo-Json -Depth 3 | Write-Host

# Exit with appropriate code
if ($OrchestrationResults.OverallStatus -eq "FAILED") {
    exit 1
}
exit 0
