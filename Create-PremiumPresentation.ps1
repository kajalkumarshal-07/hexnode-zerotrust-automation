<#
.SYNOPSIS
    Hexnode Zero Trust - Premium PowerPoint Presentation Generator
.DESCRIPTION
    Creates a professional, image-rich PowerPoint presentation.
.NOTES
    Output: C:\Users\SAM\Desktop\Hexnode-ZeroTrust-Premium.pptx
#>

# Load required assemblies
Add-Type -AssemblyName System.Drawing

# Image download directory
$ImageDir = "C:\Users\SAM\Hexnode-ZeroTrust\Images"
if (-not (Test-Path $ImageDir)) {
    New-Item -ItemType Directory -Path $ImageDir -Force | Out-Null
}

function Download-Image {
    param(
        [string]$Url,
        [string]$FileName,
        [string]$Category = "general"
    )
    
    $CategoryPath = Join-Path $ImageDir $Category
    if (-not (Test-Path $CategoryPath)) {
        New-Item -ItemType Directory -Path $CategoryPath -Force | Out-Null
    }
    
    $FilePath = Join-Path $CategoryPath $FileName
    
    if (-not (Test-Path $FilePath)) {
        try {
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($Url, $FilePath)
            Write-Host "  Downloaded: $FileName" -ForegroundColor Green
            return $FilePath
        } catch {
            Write-Host "  Failed to download: $FileName" -ForegroundColor Yellow
            return $null
        }
    } else {
        Write-Host "  Already exists: $FileName" -ForegroundColor Cyan
        return $FilePath
    }
}

function Create-PremiumPresentation {
    param(
        [string]$OutputPath = "C:\Users\SAM\Desktop\Hexnode-ZeroTrust-Premium.pptx"
    )
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  CREATING PREMIUM PRESENTATION" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Download images
    Write-Host "Step 1: Downloading images..." -ForegroundColor Yellow
    
    $images = @{
        "hexnode_logo" = Download-Image -Url "https://www.hexnode.com/wp-content/uploads/2023/08/Hexnode-Logo.png" -FileName "hexnode_logo.png" -Category "branding"
        "security_shield" = Download-Image -Url "https://img.icons8.com/ios-filled/50/00d9ff/shield.png" -FileName "shield.png" -Category "icons"
        "firewall_icon" = Download-Image -Url "https://img.icons8.com/ios-filled/50/00d9ff/firewall.png" -FileName "firewall.png" -Category "icons"
        "lock_icon" = Download-Image -Url "https://img.icons8.com/ios-filled/50/00d9ff/lock.png" -FileName "lock.png" -Category "icons"
        "network_icon" = Download-Image -Url "https://img.icons8.com/ios-filled/50/00d9ff/network.png" -FileName "network.png" -Category "icons"
        "checkmark_icon" = Download-Image -Url "https://img.icons8.com/ios-filled/50/00ff88/checkmark.png" -FileName "checkmark.png" -Category "icons"
        "alert_icon" = Download-Image -Url "https://img.icons8.com/ios-filled/50/ff6b6b/alert.png" -FileName "alert.png" -Category "icons"
        "update_icon" = Download-Image -Url "https://img.icons8.com/ios-filled/50/00d9ff/update.png" -FileName "update.png" -Category "icons"
        "provision_icon" = Download-Image -Url "https://img.icons8.com/ios-filled/50/00d9ff/add-user.png" -FileName "provision.png" -Category "icons"
        "splunk_logo" = Download-Image -Url "https://img.icons8.com/color/96/splunk.png" -FileName "splunk.png" -Category "siem"
        "carbonblack_logo" = Download-Image -Url "https://img.icons8.com/color/96/carbon-copy.png" -FileName "carbonblack.png" -Category "siem"
        "crowdstrike_logo" = Download-Image -Url "https://img.icons8.com/color/96/crowd.png" -FileName "crowdstrike.png" -Category "siem"
        "dark_tech_bg" = Download-Image -Url "https://images.unsplash.com/photo-1550751827-4bd374c3f58b?w=1920&q=80" -FileName "tech_bg.jpg" -Category "backgrounds"
        "circuit_bg" = Download-Image -Url "https://images.unsplash.com/photo-1518770660439-4636190af475?w=1920&q=80" -FileName "circuit_bg.jpg" -Category "backgrounds"
        "network_bg" = Download-Image -Url "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=1920&q=80" -FileName "network_bg.jpg" -Category "backgrounds"
    }
    
    Write-Host "`nStep 2: Creating PowerPoint presentation..." -ForegroundColor Yellow
    
    # Create PowerPoint application
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = $true
    
    # Create new presentation
    $presentation = $pptApp.Presentations.Add()
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540
    
    # Color scheme (RGB values)
    $primaryColor = 16761035    # Cyan
    $secondaryColor = 8947967   # Green
    $accentColor = 6750207      # Red
    $darkBg = 3026466           # Dark
    $lightText = 16777215       # White
    $grayText = 13421772        # Gray
    
    # Helper function to set slide background
    function Set-SlideBackground {
        param(
            [object]$Slide,
            [string]$ImagePath = $null,
            [int]$ColorValue = 3026466
        )
        
        $Slide.FollowMasterBackground = 0  # msoFalse
        
        if ($ImagePath -and (Test-Path $ImagePath)) {
            try {
                $Slide.Background.Fill.BackgroundImage = $ImagePath
                $Slide.Background.Fill.Transparency = 0.7
            } catch {
                $Slide.Background.Fill.Solid()
                $Slide.Background.Fill.ForeColor.RGB = $ColorValue
            }
        } else {
            $Slide.Background.Fill.Solid()
            $Slide.Background.Fill.ForeColor.RGB = $ColorValue
        }
    }
    
    # Helper function to add styled text box
    function Add-StyledTextBox {
        param(
            [object]$Slide,
            [string]$Text,
            [int]$X = 50,
            [int]$Y = 50,
            [int]$Width = 860,
            [int]$Height = 60,
            [int]$FontSize = 36,
            [int]$FontColor = 16777215,
            [bool]$Bold = $true,
            [int]$Alignment = 2
        )
        
        $textBox = $Slide.Shapes.AddTextbox(1, $X, $Y, $Width, $Height)
        $textBox.TextFrame.TextRange.Text = $Text
        $textBox.TextFrame.TextRange.Font.Size = $FontSize
        $textBox.TextFrame.TextRange.Font.Bold = $Bold
        $textBox.TextFrame.TextRange.Font.Color.RGB = $FontColor
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = $Alignment
        $textBox.TextFrame.WordWrap = $true
        
        return $textBox
    }
    
    # Helper function to add image
    function Add-Image {
        param(
            [object]$Slide,
            [string]$ImagePath,
            [int]$X = 100,
            [int]$Y = 100,
            [int]$Width = 200,
            [int]$Height = 200
        )
        
        if ($ImagePath -and (Test-Path $ImagePath)) {
            try {
                $Slide.Shapes.AddPicture($ImagePath, $false, $true, $X, $Y, $Width, $Height) | Out-Null
            } catch {
                Write-Host "  Warning: Could not add image" -ForegroundColor Yellow
            }
        }
    }
    
    # Helper function to add colored shape
    function Add-ColoredShape {
        param(
            [object]$Slide,
            [int]$ShapeType = 1,
            [int]$X,
            [int]$Y,
            [int]$Width,
            [int]$Height,
            [int]$FillColor,
            [string]$Text = "",
            [int]$FontSize = 14,
            [int]$TextColor = 16777215
        )
        
        $shape = $Slide.Shapes.AddShape($ShapeType, $X, $Y, $Width, $Height)
        $shape.Fill.Solid()
        $shape.Fill.ForeColor.RGB = $FillColor
        $shape.Line.Visible = $false
        
        if ($Text) {
            $shape.TextFrame.TextRange.Text = $Text
            $shape.TextFrame.TextRange.Font.Size = $FontSize
            $shape.TextFrame.TextRange.Font.Color.RGB = $TextColor
            $shape.TextFrame.TextRange.Font.Bold = $true
            $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $shape.TextFrame.VerticalAnchor = 3
        }
        
        return $shape
    }
    
    # ============================================
    # SLIDE 1: Title Slide
    # ============================================
    Write-Host "  Creating Slide 1: Title..." -ForegroundColor Green
    
    $slide1 = $presentation.Slides.Add(1, 12)
    Set-SlideBackground -Slide $slide1 -ImagePath $images["dark_tech_bg"]
    
    Add-Image -Slide $slide1 -ImagePath $images["hexnode_logo"] -X 380 -Y 30 -Width 200 -Height 80
    Add-StyledTextBox -Slide $slide1 -Text "HEXNODE ZERO TRUST" -X 50 -Y 150 -Width 860 -Height 80 -FontSize 54 -FontColor 16761035
    Add-StyledTextBox -Slide $slide1 -Text "AUTOMATION SUITE" -X 50 -Y 240 -Width 860 -Height 60 -FontSize 40 -FontColor 8947967
    Add-StyledTextBox -Slide $slide1 -Text "Enterprise Security Automation Framework" -X 50 -Y 320 -Width 860 -Height 40 -FontSize 20 -FontColor 13421772
    Add-ColoredShape -Slide $slide1 -X 0 -Y 480 -Width 960 -Height 60 -FillColor 16761035 -Text "Powered by Hexnode UEM | REST API | AI Automation" -FontSize 14 -TextColor 3026466
    
    # ============================================
    # SLIDE 2: Problem Statement
    # ============================================
    Write-Host "  Creating Slide 2: Problem Statement..." -ForegroundColor Green
    
    $slide2 = $presentation.Slides.Add(2, 12)
    Set-SlideBackground -Slide $slide2 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide2 -Text "THE CHALLENGE" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    Add-Image -Slide $slide2 -ImagePath $images["alert_icon"] -X 430 -Y 100 -Width 100 -Height 100
    
    Add-ColoredShape -Slide $slide2 -X 50 -Y 220 -Width 280 -Height 120 -FillColor 6750207 -Text "207 Days - Avg Time to Detect Threats" -FontSize 14 -TextColor 16777215
    Add-ColoredShape -Slide $slide2 -X 340 -Y 220 -Width 280 -Height 120 -FillColor 6750207 -Text "73 Days - Avg Time to Contain" -FontSize 14 -TextColor 16777215
    Add-ColoredShape -Slide $slide2 -X 630 -Y 220 -Width 280 -Height 120 -FillColor 6750207 -Text "$4.45M - Avg Breach Cost" -FontSize 14 -TextColor 16777215
    
    $problemText = "Manual security processes are slow and error-prone`nLack of real-time compliance monitoring`nSiloed security tools without automation`nGrowing attack surface with remote workforce`nRegulatory compliance pressure (GDPR, HIPAA, SOC2)"
    Add-StyledTextBox -Slide $slide2 -Text $problemText -X 50 -Y 360 -Width 860 -Height 160 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 3: Solution Overview
    # ============================================
    Write-Host "  Creating Slide 3: Solution Overview..." -ForegroundColor Green
    
    $slide3 = $presentation.Slides.Add(3, 12)
    Set-SlideBackground -Slide $slide3 -ImagePath $images["circuit_bg"]
    
    Add-StyledTextBox -Slide $slide3 -Text "OUR SOLUTION" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16777215
    
    Add-ColoredShape -Slide $slide3 -X 50 -Y 120 -Width 200 -Height 150 -FillColor 3026466 -Text "ZERO TRUST - Never Trust, Always Verify" -FontSize 14 -TextColor 16761035
    Add-ColoredShape -Slide $slide3 -X 270 -Y 120 -Width 200 -Height 150 -FillColor 3026466 -Text "AUTOMATED - Real-time Response" -FontSize 14 -TextColor 8947967
    Add-ColoredShape -Slide $slide3 -X 490 -Y 120 -Width 200 -Height 150 -FillColor 3026466 -Text "INTEGRATED - SIEM + UEM + API" -FontSize 14 -TextColor 16761035
    Add-ColoredShape -Slide $slide3 -X 710 -Y 120 -Width 200 -Height 150 -FillColor 3026466 -Text "SCALABLE - 10 to 10,000+ Devices" -FontSize 14 -TextColor 8947967
    
    $benefitText = "Automated compliance checking and remediation`nReal-time vulnerability scanning and patching`nSIEM integration for security orchestration`nZero Trust architecture implementation"
    Add-StyledTextBox -Slide $slide3 -Text $benefitText -X 50 -Y 300 -Width 860 -Height 150 -FontSize 18 -FontColor 16777215 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 4: Architecture Diagram
    # ============================================
    Write-Host "  Creating Slide 4: Architecture..." -ForegroundColor Green
    
    $slide4 = $presentation.Slides.Add(4, 12)
    Set-SlideBackground -Slide $slide4 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide4 -Text "ARCHITECTURE OVERVIEW" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    
    Add-ColoredShape -Slide $slide4 -X 50 -Y 150 -Width 200 -Height 180 -FillColor 6750207 -Text "SIEM ALERTS - Splunk, Carbon Black, CrowdStrike" -FontSize 12 -TextColor 16777215
    Add-StyledTextBox -Slide $slide4 -Text "-->" -X 260 -Y 210 -Width 60 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide4 -X 330 -Y 150 -Width 250 -Height 180 -FillColor 16761035 -Text "HEXNODE UEM - REST API Gateway, Automation Engine, Policy Management" -FontSize 12 -TextColor 3026466
    Add-StyledTextBox -Slide $slide4 -Text "-->" -X 590 -Y 210 -Width 60 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide4 -X 660 -Y 150 -Width 250 -Height 180 -FillColor 8947967 -Text "DEVICES - Windows, macOS, Linux, Endpoints" -FontSize 12 -TextColor 3026466
    
    Add-ColoredShape -Slide $slide4 -X 150 -Y 370 -Width 660 -Height 100 -FillColor 3026466 -Text "AUTOMATED REMEDIATION: Quarantine | Patch | Enforce | Report" -FontSize 16 -TextColor 16761035
    
    # ============================================
    # SLIDE 5: Feature - Compliance
    # ============================================
    Write-Host "  Creating Slide 5: Compliance Feature..." -ForegroundColor Green
    
    $slide5 = $presentation.Slides.Add(5, 12)
    Set-SlideBackground -Slide $slide5 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide5 -Text "FEATURE: COMPLIANCE CHECK" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    Add-Image -Slide $slide5 -ImagePath $images["checkmark_icon"] -X 430 -Y 100 -Width 100 -Height 100
    
    Add-ColoredShape -Slide $slide5 -X 50 -Y 220 -Width 180 -Height 100 -FillColor 8947967 -Text "Firewall Status" -FontSize 14 -TextColor 3026466
    Add-ColoredShape -Slide $slide5 -X 240 -Y 220 -Width 180 -Height 100 -FillColor 8947967 -Text "Disk Encryption" -FontSize 14 -TextColor 3026466
    Add-ColoredShape -Slide $slide5 -X 430 -Y 220 -Width 180 -Height 100 -FillColor 8947967 -Text "Defender Health" -FontSize 14 -TextColor 3026466
    Add-ColoredShape -Slide $slide5 -X 620 -Y 220 -Width 180 -Height 100 -FillColor 8947967 -Text "Update Status" -FontSize 14 -TextColor 3026466
    
    $featureText = "Real-time device compliance monitoring`nJailbreak/Root detection`nSecurity service verification`nAutomatic quarantine for non-compliant devices"
    Add-StyledTextBox -Slide $slide5 -Text $featureText -X 50 -Y 350 -Width 860 -Height 150 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 6: Feature - Vulnerability
    # ============================================
    Write-Host "  Creating Slide 6: Vulnerability Feature..." -ForegroundColor Green
    
    $slide6 = $presentation.Slides.Add(6, 12)
    Set-SlideBackground -Slide $slide6 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide6 -Text "FEATURE: VULNERABILITY SCANNING" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    Add-Image -Slide $slide6 -ImagePath $images["security_shield"] -X 430 -Y 100 -Width 100 -Height 100
    
    Add-ColoredShape -Slide $slide6 -X 50 -Y 220 -Width 200 -Height 80 -FillColor 6750207 -Text "CRITICAL: +10 pts" -FontSize 16 -TextColor 16777215
    Add-ColoredShape -Slide $slide6 -X 270 -Y 220 -Width 200 -Height 80 -FillColor 13369548 -Text "HIGH: +7 pts" -FontSize 16 -TextColor 16777215
    Add-ColoredShape -Slide $slide6 -X 490 -Y 220 -Width 200 -Height 80 -FillColor 16776960 -Text "MEDIUM: +4 pts" -FontSize 16 -TextColor 3026466
    Add-ColoredShape -Slide $slide6 -X 710 -Y 220 -Width 200 -Height 80 -FillColor 8947967 -Text "LOW: +1 pt" -FontSize 16 -TextColor 3026466
    
    $vulnText = "CVE detection and patch verification`nSoftware inventory and version auditing`nSuspicious process detection`nMalicious network connection monitoring"
    Add-StyledTextBox -Slide $slide6 -Text $vulnText -X 50 -Y 330 -Width 860 -Height 150 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 7: Feature - Firewall
    # ============================================
    Write-Host "  Creating Slide 7: Firewall Feature..." -ForegroundColor Green
    
    $slide7 = $presentation.Slides.Add(7, 12)
    Set-SlideBackground -Slide $slide7 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide7 -Text "FEATURE: FIREWALL ENFORCEMENT" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    Add-Image -Slide $slide7 -ImagePath $images["firewall_icon"] -X 430 -Y 100 -Width 100 -Height 100
    
    Add-ColoredShape -Slide $slide7 -X 100 -Y 220 -Width 760 -Height 60 -FillColor 6750207 -Text "BLOCKED PORTS: 3389 (RDP), 445 (SMB), 135, 137, 138, 139" -FontSize 14 -TextColor 16777215
    
    $fwText = "Windows Firewall profile management`nMalicious IP blacklisting`nSecurity service enforcement`nUnnecessary service disabling`nFirewall logging and monitoring"
    Add-StyledTextBox -Slide $slide7 -Text $fwText -X 50 -Y 310 -Width 860 -Height 180 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 8: Feature - OS Updates
    # ============================================
    Write-Host "  Creating Slide 8: OS Updates Feature..." -ForegroundColor Green
    
    $slide8 = $presentation.Slides.Add(8, 12)
    Set-SlideBackground -Slide $slide8 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide8 -Text "FEATURE: OS UPDATES AUTOMATION" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    Add-Image -Slide $slide8 -ImagePath $images["update_icon"] -X 430 -Y 100 -Width 100 -Height 100
    
    Add-ColoredShape -Slide $slide8 -X 50 -Y 220 -Width 180 -Height 80 -FillColor 8947967 -Text "SCAN" -FontSize 18 -TextColor 3026466
    Add-StyledTextBox -Slide $slide8 -Text "-->" -X 240 -Y 230 -Width 40 -Height 60 -FontSize 36 -FontColor 16761035
    Add-ColoredShape -Slide $slide8 -X 290 -Y 220 -Width 180 -Height 80 -FillColor 8947967 -Text "DOWNLOAD" -FontSize 18 -TextColor 3026466
    Add-StyledTextBox -Slide $slide8 -Text "-->" -X 480 -Y 230 -Width 40 -Height 60 -FontSize 36 -FontColor 16761035
    Add-ColoredShape -Slide $slide8 -X 530 -Y 220 -Width 180 -Height 80 -FillColor 8947967 -Text "INSTALL" -FontSize 18 -TextColor 3026466
    Add-StyledTextBox -Slide $slide8 -Text "-->" -X 720 -Y 230 -Width 40 -Height 60 -FontSize 36 -FontColor 16761035
    Add-ColoredShape -Slide $slide8 -X 770 -Y 220 -Width 140 -Height 80 -FillColor 8947967 -Text "VERIFY" -FontSize 18 -TextColor 3026466
    
    $updateText = "Automated Windows Update orchestration`nCritical patch prioritization`nBackground download and install`nReboot requirement detection`nMaintenance window support"
    Add-StyledTextBox -Slide $slide8 -Text $updateText -X 50 -Y 330 -Width 860 -Height 150 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 9: Feature - Provisioning
    # ============================================
    Write-Host "  Creating Slide 9: Provisioning Feature..." -ForegroundColor Green
    
    $slide9 = $presentation.Slides.Add(9, 12)
    Set-SlideBackground -Slide $slide9 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide9 -Text "FEATURE: BULK PROVISIONING" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    Add-Image -Slide $slide9 -ImagePath $images["provision_icon"] -X 430 -Y 100 -Width 100 -Height 100
    
    Add-ColoredShape -Slide $slide9 -X 50 -Y 220 -Width 200 -Height 100 -FillColor 16761035 -Text "CSV IMPORT - Device Inventory" -FontSize 14 -TextColor 3026466
    Add-StyledTextBox -Slide $slide9 -Text "-->" -X 260 -Y 240 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide9 -X 310 -Y 220 -Width 200 -Height 100 -FillColor 16761035 -Text "GROUP ASSIGNMENT - Policy Application" -FontSize 14 -TextColor 3026466
    Add-StyledTextBox -Slide $slide9 -Text "-->" -X 520 -Y 240 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide9 -X 570 -Y 220 -Width 200 -Height 100 -FillColor 16761035 -Text "APP DEPLOYMENT - Welcome Messages" -FontSize 14 -TextColor 3026466
    
    $provText = "CSV-based device inventory import`nAutomatic device group assignment`nPolicy application by location/department`nScalable from 10 to 10,000+ devices"
    Add-StyledTextBox -Slide $slide9 -Text $provText -X 50 -Y 350 -Width 860 -Height 150 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 10: Feature - SIEM Integration
    # ============================================
    Write-Host "  Creating Slide 10: SIEM Integration..." -ForegroundColor Green
    
    $slide10 = $presentation.Slides.Add(10, 12)
    Set-SlideBackground -Slide $slide10 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide10 -Text "FEATURE: SIEM INTEGRATION" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    
    Add-Image -Slide $slide10 -ImagePath $images["splunk_logo"] -X 150 -Y 120 -Width 150 -Height 150
    Add-Image -Slide $slide10 -ImagePath $images["carbonblack_logo"] -X 400 -Y 120 -Width 150 -Height 150
    Add-Image -Slide $slide10 -ImagePath $images["crowdstrike_logo"] -X 650 -Y 120 -Width 150 -Height 150
    
    Add-ColoredShape -Slide $slide10 -X 50 -Y 300 -Width 280 -Height 60 -FillColor 6750207 -Text "CRITICAL: Wipe Device" -FontSize 14 -TextColor 16777215
    Add-ColoredShape -Slide $slide10 -X 340 -Y 300 -Width 280 -Height 60 -FillColor 13369548 -Text "HIGH: Quarantine" -FontSize 14 -TextColor 16777215
    Add-ColoredShape -Slide $slide10 -X 630 -Y 300 -Width 280 -Height 60 -FillColor 16776960 -Text "MEDIUM: Remediate" -FontSize 14 -TextColor 3026466
    
    $siemText = "Splunk HEC webhook integration`nCarbon Black EDR alert handling`nCrowdStrike Falcon XDR support`nReal-time alert processing"
    Add-StyledTextBox -Slide $slide10 -Text $siemText -X 50 -Y 390 -Width 860 -Height 120 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 11: Security Benefits
    # ============================================
    Write-Host "  Creating Slide 11: Security Benefits..." -ForegroundColor Green
    
    $slide11 = $presentation.Slides.Add(11, 12)
    Set-SlideBackground -Slide $slide11 -ImagePath $images["network_bg"]
    
    Add-StyledTextBox -Slide $slide11 -Text "SECURITY BENEFITS" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16777215
    
    Add-ColoredShape -Slide $slide11 -X 50 -Y 120 -Width 280 -Height 150 -FillColor 3026466 -Text "ZERO TRUST - Never Trust, Always Verify, Least Privilege" -FontSize 14 -TextColor 16761035
    Add-ColoredShape -Slide $slide11 -X 340 -Y 120 -Width 280 -Height 150 -FillColor 3026466 -Text "REAL-TIME - 24/7 Monitoring, Threat Detection, Compliance Checks" -FontSize 14 -TextColor 8947967
    Add-ColoredShape -Slide $slide11 -X 630 -Y 120 -Width 280 -Height 150 -FillColor 3026466 -Text "AUTOMATED - Less than 5 Min Response, Auto-Fix Issues, Self-Healing" -FontSize 14 -TextColor 16761035
    
    Add-ColoredShape -Slide $slide11 -X 100 -Y 320 -Width 200 -Height 100 -FillColor 8947967 -Text "85% Faster Response" -FontSize 18 -TextColor 3026466
    Add-ColoredShape -Slide $slide11 -X 380 -Y 320 -Width 200 -Height 100 -FillColor 8947967 -Text "99.9% Uptime" -FontSize 18 -TextColor 3026466
    Add-ColoredShape -Slide $slide11 -X 660 -Y 320 -Width 200 -Height 100 -FillColor 8947967 -Text "100% Compliance" -FontSize 18 -TextColor 3026466
    
    # ============================================
    # SLIDE 12: ROI & Value
    # ============================================
    Write-Host "  Creating Slide 12: ROI & Value..." -ForegroundColor Green
    
    $slide12 = $presentation.Slides.Add(12, 12)
    Set-SlideBackground -Slide $slide12 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide12 -Text "ROI & VALUE PROPOSITION" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    
    Add-ColoredShape -Slide $slide12 -X 50 -Y 120 -Width 280 -Height 150 -FillColor 8947967 -Text "60% Fewer Security Incidents" -FontSize 18 -TextColor 3026466
    Add-ColoredShape -Slide $slide12 -X 340 -Y 120 -Width 280 -Height 150 -FillColor 16761035 -Text "40% IT Workload Reduction" -FontSize 18 -TextColor 3026466
    Add-ColoredShape -Slide $slide12 -X 630 -Y 120 -Width 280 -Height 150 -FillColor 6750207 -Text "85% Faster Threat Response" -FontSize 18 -TextColor 16777215
    
    $valueText = "Reduced mean time to detect (MTTD) from days to minutes`nReduced mean time to respond (MTTR) from hours to seconds`nLower compliance audit costs with automated reporting`nImproved security posture with continuous monitoring"
    Add-StyledTextBox -Slide $slide12 -Text $valueText -X 50 -Y 300 -Width 860 -Height 180 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 13: Implementation Roadmap
    # ============================================
    Write-Host "  Creating Slide 13: Roadmap..." -ForegroundColor Green
    
    $slide13 = $presentation.Slides.Add(13, 12)
    Set-SlideBackground -Slide $slide13 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide13 -Text "IMPLEMENTATION ROADMAP" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    
    Add-ColoredShape -Slide $slide13 -X 50 -Y 120 -Width 200 -Height 120 -FillColor 16761035 -Text "PHASE 1 - Week 1-2 - Setup and Config" -FontSize 12 -TextColor 3026466
    Add-StyledTextBox -Slide $slide13 -Text "-->" -X 260 -Y 150 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide13 -X 310 -Y 120 -Width 200 -Height 120 -FillColor 8947967 -Text "PHASE 2 - Week 3-4 - Script Deploy" -FontSize 12 -TextColor 3026466
    Add-StyledTextBox -Slide $slide13 -Text "-->" -X 520 -Y 150 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide13 -X 570 -Y 120 -Width 200 -Height 120 -FillColor 16761035 -Text "PHASE 3 - Week 5-6 - Automation" -FontSize 12 -TextColor 3026466
    Add-StyledTextBox -Slide $slide13 -Text "-->" -X 780 -Y 150 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide13 -X 830 -Y 120 -Width 100 -Height 120 -FillColor 8947967 -Text "PHASE 4 - Week 7-8 - SIEM" -FontSize 10 -TextColor 3026466
    
    $phaseText = "Phase 1: Configure Hexnode API, create quarantine groups`nPhase 2: Upload and test all automation scripts`nPhase 3: Create automations with triggers and schedules`nPhase 4: Integrate with SIEM and test alert handling`nPhase 5: Production deployment and monitoring"
    Add-StyledTextBox -Slide $slide13 -Text $phaseText -X 50 -Y 280 -Width 860 -Height 200 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 14: Quick Start
    # ============================================
    Write-Host "  Creating Slide 14: Quick Start..." -ForegroundColor Green
    
    $slide14 = $presentation.Slides.Add(14, 12)
    Set-SlideBackground -Slide $slide14 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide14 -Text "QUICK START GUIDE" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    
    Add-ColoredShape -Slide $slide14 -X 50 -Y 120 -Width 200 -Height 120 -FillColor 16761035 -Text "STEP 1 - Configure API" -FontSize 18 -TextColor 3026466
    Add-StyledTextBox -Slide $slide14 -Text "-->" -X 260 -Y 150 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide14 -X 310 -Y 120 -Width 200 -Height 120 -FillColor 8947967 -Text "STEP 2 - Upload Scripts" -FontSize 18 -TextColor 3026466
    Add-StyledTextBox -Slide $slide14 -Text "-->" -X 520 -Y 150 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide14 -X 570 -Y 120 -Width 200 -Height 120 -FillColor 16761035 -Text "STEP 3 - Create Automation" -FontSize 18 -TextColor 3026466
    Add-StyledTextBox -Slide $slide14 -Text "-->" -X 780 -Y 150 -Width 40 -Height 60 -FontSize 36 -FontColor 8947967
    Add-ColoredShape -Slide $slide14 -X 830 -Y 120 -Width 100 -Height 120 -FillColor 8947967 -Text "STEP 4 - Enable" -FontSize 14 -TextColor 3026466
    
    $instrText = "Edit hexnode_config.json with your API credentials`nUpload .ps1 scripts to Hexnode Custom Scripts`nCreate automations with time/activity triggers`nEnable automations and monitor results"
    Add-StyledTextBox -Slide $slide14 -Text $instrText -X 50 -Y 280 -Width 860 -Height 180 -FontSize 16 -FontColor 13421772 -Bold $false -Alignment 1
    
    # ============================================
    # SLIDE 15: Next Steps
    # ============================================
    Write-Host "  Creating Slide 15: Next Steps..." -ForegroundColor Green
    
    $slide15 = $presentation.Slides.Add(15, 12)
    Set-SlideBackground -Slide $slide15 -ColorValue 3026466
    
    Add-StyledTextBox -Slide $slide15 -Text "NEXT STEPS" -X 50 -Y 30 -Width 860 -Height 60 -FontSize 40 -FontColor 16761035
    
    $nextText = "Schedule technical deep-dive demonstration`nIdentify pilot group for proof of concept`nReview Hexnode API documentation`nPlan integration with existing SIEM infrastructure`nDefine compliance requirements and policies`nEstablish success metrics and KPIs"
    Add-StyledTextBox -Slide $slide15 -Text $nextText -X 50 -Y 120 -Width 860 -Height 250 -FontSize 18 -FontColor 13421772 -Bold $false -Alignment 1
    
    Add-ColoredShape -Slide $slide15 -X 200 -Y 400 -Width 560 -Height 80 -FillColor 8947967 -Text "Ready to Start? Let us Schedule a Demo!" -FontSize 24 -TextColor 3026466
    
    # ============================================
    # SLIDE 16: Thank You
    # ============================================
    Write-Host "  Creating Slide 16: Thank You..." -ForegroundColor Green
    
    $slide16 = $presentation.Slides.Add(16, 12)
    Set-SlideBackground -Slide $slide16 -ImagePath $images["dark_tech_bg"]
    
    Add-Image -Slide $slide16 -ImagePath $images["hexnode_logo"] -X 380 -Y 30 -Width 200 -Height 80
    Add-StyledTextBox -Slide $slide16 -Text "THANK YOU" -X 50 -Y 150 -Width 860 -Height 80 -FontSize 54 -FontColor 16761035
    Add-StyledTextBox -Slide $slide16 -Text "Questions and Discussion" -X 50 -Y 250 -Width 860 -Height 60 -FontSize 28 -FontColor 8947967
    Add-ColoredShape -Slide $slide16 -X 200 -Y 380 -Width 560 -Height 80 -FillColor 16761035 -Text "Hexnode Zero Trust Automation Suite" -FontSize 20 -TextColor 3026466
    
    # Save presentation
    Write-Host "`nStep 3: Saving presentation..." -ForegroundColor Yellow
    
    $presentation.SaveAs($OutputPath)
    Write-Host "  Presentation saved to: $OutputPath" -ForegroundColor Green
    
    # Cleanup
    $presentation.Close()
    $pptApp.Quit()
    
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "  PREMIUM PRESENTATION CREATED!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  File: $OutputPath" -ForegroundColor Cyan
    Write-Host "  Slides: 16" -ForegroundColor Cyan
    Write-Host "  Features: Real images, Professional design" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Green
}

# Run the presentation creator
Create-PremiumPresentation
