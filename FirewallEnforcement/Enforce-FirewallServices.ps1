<#
.SYNOPSIS
    Zero Trust Firewall & Service Enforcement
.DESCRIPTION
    Enforces Windows Firewall policies, blocks dangerous ports,
    and ensures critical security services are running.
.NOTES
    Deploy via Hexnode Custom Scripts for continuous enforcement.
#>

param(
    [string]$ConfigPath = "C:\Hexnode-ZeroTrust\Config\hexnode_config.json"
)

$Config = Get-Content $ConfigPath | ConvertFrom-Json
$BlockedPorts = $Config.firewall.blocked_ports
$RequiredServices = $Config.firewall.required_services

$EnforcementResults = @{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    DeviceName = $env:COMPUTERNAME
    FirewallStatus = @()
    ServicesStatus = @()
    Violations = @()
    IsCompliant = $true
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
}

function Enable-WindowsFirewall {
    Write-Log "Enforcing Windows Firewall..."
    
    $Profiles = @("Domain", "Private", "Public")
    $Results = @()
    
    foreach ($Profile in $Profiles) {
        try {
            $CurrentState = Get-NetFirewallProfile -Name $Profile -ErrorAction Stop
            
            if (-not $CurrentState.Enabled) {
                Set-NetFirewallProfile -Name $Profile -Enabled True -ErrorAction Stop
                $Results += @{
                    Profile = $Profile
                    Status = "ENABLED"
                    Message = "Firewall was disabled, now enabled"
                }
                Write-Log "Enabled $Profile firewall profile" "WARNING"
            } else {
                $Results += @{
                    Profile = $Profile
                    Status = "OK"
                    Message = "Already enabled"
                }
            }
        } catch {
            $Results += @{
                Profile = $Profile
                Status = "ERROR"
                Message = $_.Exception.Message
            }
            Write-Log "Failed to configure $Profile firewall: $($_.Exception.Message)" "ERROR"
        }
    }
    
    return $Results
}

function Block-DangerousPorts {
    Write-Log "Blocking dangerous ports..."
    
    $Results = @()
    
    foreach ($Port in $BlockedPorts) {
        try {
            # Check if rule already exists
            $ExistingRule = Get-NetFirewallRule -DisplayName "ZeroTrust_Block_Port_$Port" -ErrorAction SilentlyContinue
            
            if (-not $ExistingRule) {
                # Create inbound block rule
                New-NetFirewallRule -DisplayName "ZeroTrust_Block_Port_$Port" `
                    -Direction Inbound `
                    -Protocol TCP `
                    -LocalPort $Port `
                    -Action Block `
                    -Profile Any `
                    -Description "Zero Trust: Blocked port $Port" `
                    -ErrorAction Stop | Out-Null
                
                # Create outbound block rule
                New-NetFirewallRule -DisplayName "ZeroTrust_Block_Outbound_$Port" `
                    -Direction Outbound `
                    -Protocol TCP `
                    -RemotePort $Port `
                    -Action Block `
                    -Profile Any `
                    -Description "Zero Trust: Blocked outbound port $Port" `
                    -ErrorAction Stop | Out-Null
                
                $Results += @{
                    Port = $Port
                    Status = "BLOCKED"
                    Message = "Port $Port blocked (inbound/outbound)"
                }
                Write-Log "Blocked port $Port"
            } else {
                $Results += @{
                    Port = $Port
                    Status = "EXISTS"
                    Message = "Rule already exists"
                }
            }
        } catch {
            $Results += @{
                Port = $Port
                Status = "ERROR"
                Message = $_.Exception.Message
            }
            Write-Log "Failed to block port $Port: $($_.Exception.Message)" "ERROR"
        }
    }
    
    return $Results
}

function Block-MaliciousIPs {
    Write-Log "Blocking known malicious IPs..."
    
    # Known malicious IP ranges (update as needed)
    $MaliciousIPs = @(
        @{ IP = "103.224.182.251"; Description = "Known C2 Server" },
        @{ IP = "185.220.101.0/24"; Description = "Tor Exit Nodes" },
        @{ IP = "45.33.32.0/24"; Description = "Shodan Scanners" }
    )
    
    foreach ($MaliciousIP in $MaliciousIPs) {
        try {
            $RuleName = "ZeroTrust_Block_IP_$($MaliciousIP.IP -replace '[\.\/]', '_')"
            $ExistingRule = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
            
            if (-not $ExistingRule) {
                New-NetFirewallRule -DisplayName $RuleName `
                    -Direction Inbound `
                    -RemoteAddress $MaliciousIP.IP `
                    -Action Block `
                    -Description "Zero Trust: $($MaliciousIP.Description)" `
                    -ErrorAction Stop | Out-Null
                
                Write-Log "Blocked IP range: $($MaliciousIP.IP) ($($MaliciousIP.Description))"
            }
        } catch {
            Write-Log "Failed to block IP $($MaliciousIP.IP): $($_.Exception.Message)" "ERROR"
        }
    }
}

function Test-SecurityServices {
    Write-Log "Checking required security services..."
    
    $Results = @()
    
    foreach ($ServiceName in $RequiredServices) {
        try {
            $Service = Get-Service -Name $ServiceName -ErrorAction Stop
            
            $Status = @{
                Service = $ServiceName
                CurrentStatus = $Service.Status.ToString()
                StartType = $Service.StartType.ToString()
                Required = "Running"
            }
            
            if ($Service.Status -ne "Running") {
                # Try to start the service
                Start-Service -Name $ServiceName -ErrorAction Stop
                $Status.Status = "STARTED"
                $Status.Message = "Service was not running, now started"
                Write-Log "Started service: $ServiceName" "WARNING"
            } else {
                $Status.Status = "OK"
                $Status.Message = "Service is running"
            }
            
            # Ensure service is set to automatic start
            if ($Service.StartType -ne "Automatic") {
                Set-Service -Name $ServiceName -StartupType Automatic -ErrorAction Stop
                $Status.Message += " (Startup type set to Automatic)"
            }
            
            $Results += $Status
        } catch {
            $Results += @{
                Service = $ServiceName
                Status = "ERROR"
                Message = $_.Exception.Message
            }
            Write-Log "Failed to verify service $ServiceName: $($_.Exception.Message)" "ERROR"
        }
    }
    
    return $Results
}

function Disable-UnnecessaryServices {
    Write-Log "Disabling unnecessary services..."
    
    $UnnecessaryServices = @(
        "RemoteRegistry",      # Remote Registry
        "SharedAccess",        # ICS/Internet Connection Sharing
        "TapiSrv",             # Telephony
        "W32Time",             # Windows Time (use NTP instead)
        "WMSvc",               # Web Management Service
        "icssvc"               # Windows Mobile Hotspot
    )
    
    foreach ($ServiceName in $UnnecessaryServices) {
        try {
            $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
            
            if ($Service -and $Service.Status -eq "Running") {
                Stop-Service -Name $ServiceName -Force -ErrorAction Stop
                Set-Service -Name $ServiceName -StartupType Disabled -ErrorAction Stop
                Write-Log "Disabled service: $ServiceName" "WARNING"
            }
        } catch {
            Write-Log "Could not disable service $ServiceName: $($_.Exception.Message)" "WARNING"
        }
    }
}

function Set-SecureFirewallPolicies {
    Write-Log "Applying secure firewall policies..."
    
    # Block all inbound by default
    try {
        Set-NetFirewallProfile -Profile Domain,Private,Public `
            -DefaultInboundAction Block `
            -DefaultOutboundAction Allow `
            -ErrorAction Stop
        Write-Log "Set default inbound action to Block"
    } catch {
        Write-Log "Failed to set default inbound action: $($_.Exception.Message)" "ERROR"
    }
    
    # Enable logging
    try {
        Set-NetFirewallProfile -Profile Domain,Private,Public `
            -LogFileName "C:\Windows\System32\LogFiles\Firewall\pfirewall.log" `
            -LogMaxSizeKilobytes 4096 `
            -LogBlocked True `
            -LogAllowed True `
            -ErrorAction Stop
        Write-Log "Enabled firewall logging"
    } catch {
        Write-Log "Failed to enable logging: $($_.Exception.Message)" "ERROR"
    }
}

function Test-ExistingConnections {
    Write-Log "Checking existing connections for violations..."
    
    $SuspiciousConnections = @()
    $Connections = Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue
    
    foreach ($Conn in $Connections) {
        if ($BlockedPorts -contains $Conn.RemotePort) {
            $SuspiciousConnections += @{
                LocalPort = $Conn.LocalPort
                RemotePort = $Conn.RemotePort
                RemoteAddress = $Conn.RemoteAddress
                Process = (Get-Process -Id $Conn.OwningProcess -ErrorAction SilentlyContinue).ProcessName
            }
        }
    }
    
    return $SuspiciousConnections
}

# Execute enforcement
Write-Log "=== Zero Trust Firewall Enforcement Started ==="

# Enable and configure firewall
$EnforcementResults.FirewallStatus = Enable-WindowsFirewall
Set-SecureFirewallPolicies
Block-DangerousPorts
Block-MaliciousIPs

# Check and enforce services
$EnforcementResults.ServicesStatus = Test-SecurityServices
Disable-UnnecessaryServices

# Check for existing violations
$Violations = Test-ExistingConnections
if ($Violations.Count -gt 0) {
    $EnforcementResults.Violations = $Violations
    $EnforcementResults.IsCompliant = $false
    Write-Log "Found $($Violations.Count) suspicious connections" "WARNING"
}

# Summary
Write-Log "=== Enforcement Complete ==="
Write-Log "Compliant: $($EnforcementResults.IsCompliant)"

$EnforcementResults | ConvertTo-Json -Depth 3 | Write-Host

if (-not $EnforcementResults.IsCompliant) {
    exit 1  # Non-compliant - trigger Hexnode automation
}
exit 0
