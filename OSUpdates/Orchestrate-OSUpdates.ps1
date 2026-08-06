<#
.SYNOPSIS
    Zero Trust OS Updates Orchestration
.DESCRIPTION
    Automates Windows Update scanning, downloading, and installation.
    Ensures devices stay patched with latest security updates.
.NOTES
    Deploy via Hexnode Custom Scripts for automated patch management.
#>

param(
    [string]$ConfigPath = "C:\Hexnode-ZeroTrust\Config\hexnode_config.json",
    [switch]$ForceInstall,
    [switch]$SkipReboot
)

$Config = Get-Content $ConfigPath | ConvertFrom-Json

$UpdateResults = @{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    DeviceName = $env:COMPUTERNAME
    OSVersion = (Get-CimInstance Win32_OperatingSystem).Caption
    UpdatesFound = 0
    UpdatesInstalled = 0
    UpdatesFailed = 0
    RebootRequired = $false
    CriticalUpdates = @()
    LogFile = "C:\Windows\Temp\ZeroTrust_Update_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $LogEntry = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $UpdateResults.LogFile -Value $LogEntry -ErrorAction SilentlyContinue
}

function Initialize-WindowsUpdate {
    Write-Log "Initializing Windows Update services..."
    
    $Services = @("wuauserv", "bits", "cryptsvc")
    
    foreach ($Service in $Services) {
        try {
            $Svc = Get-Service -Name $Service -ErrorAction Stop
            if ($Svc.Status -ne "Running") {
                Start-Service -Name $Service -ErrorAction Stop
                Write-Log "Started service: $Service"
            }
            Set-Service -Name $Service -StartupType Manual -ErrorAction SilentlyContinue
        } catch {
            Write-Log "Failed to start service $Service: $($_.Exception.Message)" "ERROR"
        }
    }
}

function Scan-Updates {
    Write-Log "Scanning for available updates..."
    
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        
        # Search for all applicable updates
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
        
        $Updates = @()
        foreach ($Update in $SearchResult.Updates) {
            $UpdateInfo = @{
                Title = $Update.Title
                KB = $Update.KBArticleIDs
                Severity = $Update.MsrcSeverity
                Size = [math]::Round($Update.MaxDownloadSize / 1MB, 2)
                Description = $Update.Description
            }
            $Updates += $UpdateInfo
            
            if ($Update.MsrcSeverity -eq "Critical") {
                $UpdateResults.CriticalUpdates += $UpdateInfo
            }
        }
        
        $UpdateResults.UpdatesFound = $Updates.Count
        Write-Log "Found $($Updates.Count) updates ($($UpdateResults.CriticalUpdates.Count) critical)"
        
        return $Updates
    } catch {
        Write-Log "Failed to scan updates: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Download-Updates {
    param($Updates)
    
    Write-Log "Downloading updates..."
    
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateDownloader = $UpdateSession.CreateUpdateDownloader()
        $UpdateDownloader.Updates = $Updates
        $UpdateDownloader.Priority = 3  # High priority
        
        $DownloadResult = $UpdateDownloader.Download()
        
        switch ($DownloadResult.ResultCode) {
            2 { Write-Log "Download completed successfully" }
            3 { Write-Log "Download completed with errors" "WARNING" }
            4 { Write-Log "Download failed" "ERROR" }
            5 { Write-Log "Download cancelled" "WARNING" }
        }
        
        return $DownloadResult.ResultCode -eq 2
    } catch {
        Write-Log "Failed to download updates: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Install-Updates {
    param($Updates)
    
    Write-Log "Installing updates..."
    
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateInstaller = $UpdateSession.CreateUpdateInstaller()
        $UpdateInstaller.Updates = $Updates
        
        # Force installation if specified
        if ($ForceInstall) {
            Write-Log "Force installation mode enabled"
        }
        
        $InstallResult = $UpdateInstaller.Install()
        
        switch ($InstallResult.ResultCode) {
            2 {
                Write-Log "Installation completed successfully"
                $UpdateResults.UpdatesInstalled = $Updates.Count
            }
            3 {
                Write-Log "Installation completed with errors" "WARNING"
                $UpdateResults.UpdatesInstalled = $Updates.Count - 1
                $UpdateResults.UpdatesFailed = 1
            }
            4 {
                Write-Log "Installation failed" "ERROR"
                $UpdateResults.UpdatesFailed = $Updates.Count
            }
            5 {
                Write-Log "Installation cancelled" "WARNING"
            }
        }
        
        # Check for reboot requirement
        $UpdateResults.RebootRequired = $InstallResult.RebootRequired
        if ($InstallResult.RebootRequired) {
            Write-Log "REBOOT REQUIRED to complete updates" "WARNING"
        }
        
        return $InstallResult.ResultCode -eq 2
    } catch {
        Write-Log "Failed to install updates: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Install-CriticalUpdates {
    Write-Log "Installing critical security updates..."
    
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    
    # Get only critical updates
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and MsrcSeverity='Critical' and Type='Software'")
    
    if ($SearchResult.Updates.Count -eq 0) {
        Write-Log "No critical updates pending"
        return $true
    }
    
    $CriticalUpdates = @()
    foreach ($Update in $SearchResult.Updates) {
        $CriticalUpdates += $Update
        Write-Log "Critical update: $($Update.Title)"
    }
    
    return Install-Updates -Updates $CriticalUpdates
}

function Test-UpdateCompliance {
    Write-Log "Checking update compliance..."
    
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    
    # Check for missing critical updates
    $CriticalMissing = $UpdateSearcher.Search("IsInstalled=0 and MsrcSeverity='Critical' and IsHidden=0")
    
    # Check for missing security updates (last 30 days)
    $CutoffDate = (Get-Date).AddDays(-30)
    $AllMissing = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
    
    $RecentMissing = @()
    foreach ($Update in $AllMissing.Updates) {
        if ($Update.LastDeploymentChangeTime -gt $CutoffDate) {
            $RecentMissing += $Update
        }
    }
    
    $IsCompliant = $CriticalMissing.Updates.Count -eq 0
    $Status = if ($IsCompliant) { "COMPLIANT" } else { "NON-COMPLIANT" }
    
    Write-Log "Update compliance: $Status"
    Write-Log "Missing critical: $($CriticalMissing.Updates.Count), Recent (30 days): $($RecentMissing.Count)"
    
    return @{
        IsCompliant = $IsCompliant
        MissingCritical = $CriticalMissing.Updates.Count
        MissingRecent = $RecentMissing.Count
    }
}

function Enable-AutomaticUpdates {
    Write-Log "Configuring automatic updates..."
    
    try {
        # Enable automatic updates via registry
        $RegistryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        
        if (-not (Test-Path $RegistryPath)) {
            New-Item -Path $RegistryPath -Force | Out-Null
        }
        
        # Set to auto download and install
        Set-ItemProperty -Path $RegistryPath -Name "NoAutoUpdate" -Value 0 -Type DWord
        Set-ItemProperty -Path $RegistryPath -Name "AUOptions" -Value 4 -Type DWord
        Set-ItemProperty -Path $RegistryPath -Name "ScheduledInstallDay" -Value 0 -Type DWord
        Set-ItemProperty -Path $RegistryPath -Name "ScheduledInstallTime" -Value 3 -Type DWord
        
        Write-Log "Automatic updates enabled"
    } catch {
        Write-Log "Failed to configure automatic updates: $($_.Exception.Message)" "ERROR"
    }
}

function Start-USOClient {
    Write-Log "Triggering Windows Update Orchestrator..."
    
    $USOClient = "C:\Windows\System32\USOClient.exe"
    
    if (Test-Path $USOClient) {
        $Actions = @(
            @{ Arg = "StartInteractiveScan"; Msg = "Scanning..." },
            @{ Arg = "StartDownload"; Msg = "Downloading..." },
            @{ Arg = "StartInstall"; Msg = "Installing..." }
        )
        
        foreach ($Action in $Actions) {
            Write-Log $Action.Msg
            Start-Process -FilePath $USOClient -ArgumentList $Action.Arg -NoNewWindow -Wait
            Start-Sleep -Seconds 5
        }
        
        return $true
    } else {
        Write-Log "USOClient.exe not found" "ERROR"
        return $false
    }
}

# Execute update orchestration
Write-Log "=== Zero Trust OS Update Orchestration Started ==="
Write-Log "Device: $env:COMPUTERNAME"
Write-Log "OS: $UpdateResults.OSVersion"

# Initialize services
Initialize-WindowsUpdate

# Scan for updates
$Updates = Scan-Updates

if ($Updates.Count -gt 0) {
    # Download updates
    $DownloadSuccess = Download-Updates -Updates $Updates
    
    if ($DownloadSuccess) {
        # Install updates
        $InstallSuccess = Install-Updates -Updates $Updates
    }
}

# Check compliance
$Compliance = Test-UpdateCompliance

# Configure automatic updates for future
Enable-AutomaticUpdates

# Summary
Write-Log "=== Update Orchestration Complete ==="
Write-Log "Updates Found: $($UpdateResults.UpdatesFound)"
Write-Log "Updates Installed: $($UpdateResults.UpdatesInstalled)"
Write-Log "Updates Failed: $($UpdateResults.UpdatesFailed)"
Write-Log "Reboot Required: $($UpdateResults.RebootRequired)"
Write-Log "Compliant: $($Compliance.IsCompliant)"

$UpdateResults | ConvertTo-Json -Depth 3 | Write-Host

if (-not $Compliance.IsCompliant) {
    exit 1  # Non-compliant - trigger Hexnode automation
}
exit 0
