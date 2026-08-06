<#
.SYNOPSIS
    Zero Trust Bulk Device Provisioning
.DESCRIPTION
    Reads device inventory from CSV and automatically assigns devices to appropriate
    policy groups based on location, department, or other attributes.
.NOTES
    CSV Format: DeviceName,SerialNumber,Location,Department,GroupID
#>

param(
    [string]$ConfigPath = "C:\Hexnode-ZeroTrust\Config\hexnode_config.json",
    [string]$CSVPath = "C:\Hexnode-ZeroTrust\BulkProvisioning\devices.csv"
)

$Config = Get-Content $ConfigPath | ConvertFrom-Json
$ApiUrl = $Config.hexnode.api_url
$ApiKey = $Config.hexnode.api_key

$ProvisionResults = @{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    TotalDevices = 0
    Provisioned = 0
    Failed = 0
    Errors = @()
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
}

function Get-HexnodeDevice {
    param([string]$SerialNumber)
    
    try {
        $Response = Invoke-RestMethod -Uri "$ApiUrl/devices/?serial_number=$SerialNumber" `
            -Headers @{ "Authorization" = $ApiKey }
        
        return $Response.devices | Select-Object -First 1
    } catch {
        return $null
    }
}

function Add-DeviceToGroup {
    param([int]$DeviceId, [int]$GroupId)
    
    try {
        $Body = @{
            devices = @($DeviceId)
            devicegroups = @($GroupId)
        } | ConvertTo-Json
        
        $Response = Invoke-RestMethod -Uri "$ApiUrl/actions/add_to_device_group/" `
            -Method Post `
            -Headers @{
                "Authorization" = $ApiKey
                "Content-Type" = "application/json"
            } -Body $Body
        
        return $true
    } catch {
        return $false
    }
}

function Set-DevicePolicy {
    param([int]$DeviceId, [int]$PolicyId)
    
    try {
        $Body = @{
            devices = @($DeviceId)
            policies = @($PolicyId)
        } | ConvertTo-Json
        
        $Response = Invoke-RestMethod -Uri "$ApiUrl/actions/apply_policy/" `
            -Method Post `
            -Headers @{
                "Authorization" = $ApiKey
                "Content-Type" = "application/json"
            } -Body $Body
        
        return $true
    } catch {
        return $false
    }
}

function New-HexnodeDevice {
    param(
        [string]$DeviceName,
        [string]$SerialNumber,
        [string]$Platform = "windows"
    )
    
    try {
        $Body = @{
            device_name = $DeviceName
            serial_number = $SerialNumber
            platform = $Platform
        } | ConvertTo-Json
        
        $Response = Invoke-RestMethod -Uri "$ApiUrl/devices/" `
            -Method Post `
            -Headers @{
                "Authorization" = $ApiKey
                "Content-Type" = "application/json"
            } -Body $Body
        
        return $Response
    } catch {
        return $null
    }
}

function Invoke-BulkProvision {
    Write-Log "=== Zero Trust Bulk Provisioning Started ==="
    
    # Import CSV
    if (-not (Test-Path $CSVPath)) {
        Write-Log "CSV file not found: $CSVPath" "ERROR"
        return
    }
    
    $Devices = Import-Csv -Path $CSVPath
    $ProvisionResults.TotalDevices = $Devices.Count
    
    Write-Log "Loaded $($Devices.Count) devices from CSV"
    
    foreach ($Device in $Devices) {
        $DeviceName = $Device.DeviceName
        $SerialNumber = $Device.SerialNumber
        $Location = $Device.Location
        $Department = $Device.Department
        $TargetGroupId = [int]$Device.GroupID
        
        Write-Log "Processing: $DeviceName ($SerialNumber)"
        
        # Find device in Hexnode
        $HexnodeDevice = Get-HexnodeDevice -SerialNumber $SerialNumber
        
        if ($HexnodeDevice) {
            $DeviceId = $HexnodeDevice.id
            
            # Add to target group
            $GroupSuccess = Add-DeviceToGroup -DeviceId $DeviceId -GroupId $TargetGroupId
            
            if ($GroupSuccess) {
                Write-Log "  Added to group $TargetGroupId"
                $ProvisionResults.Provisioned++
            } else {
                Write-Log "  Failed to add to group" "ERROR"
                $ProvisionResults.Failed++
                $ProvisionResults.Errors += "Failed to add $DeviceName to group $TargetGroupId"
            }
        } else {
            # Device not found - create enrollment invitation
            Write-Log "  Device not found in Hexnode - creating enrollment" "WARNING"
            
            $NewDevice = New-HexnodeDevice -DeviceName $DeviceName -SerialNumber $SerialNumber
            if ($NewDevice) {
                Write-Log "  Device created - enrollment pending"
            } else {
                Write-Log "  Failed to create device" "ERROR"
                $ProvisionResults.Failed++
            }
        }
    }
    
    Write-Log "=== Provisioning Complete ==="
    Write-Log "Total: $($ProvisionResults.TotalDevices)"
    Write-Log "Provisioned: $($ProvisionResults.Provisioned)"
    Write-Log "Failed: $($ProvisionResults.Failed)"
}

# Create sample CSV if it doesn't exist
if (-not (Test-Path $CSVPath)) {
    $SampleCSV = @"
DeviceName,SerialNumber,Location,Department,GroupID
NYC-PC-001,SN12345678,NYC,Corporate,201
NYC-PC-002,SN23456789,NYC,Corporate,201
LA-PC-001,SN34567890,LA,Retail,202
LA-PC-002,SN45678901,LA,Retail,202
CHI-PC-001,SN56789012,CHI,Corporate,203
"@
    
    $SampleCSV | Out-File -FilePath $CSVPath -Encoding UTF8
    Write-Log "Created sample CSV at $CSVPath"
    Write-Log "Please update the CSV with your device inventory" "WARNING"
}

# Execute provisioning
Invoke-BulkProvision
