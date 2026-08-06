<#
.SYNOPSIS
    Hexnode Zero Trust Automation Suite - PowerPoint Presentation Generator
.DESCRIPTION
    Creates a professional PowerPoint presentation for pitching the Zero Trust solution.
.NOTES
    Requires PowerPoint to be installed on the system.
    Output: C:\Users\SAM\Desktop\Hexnode-ZeroTrust-Pitch.pptx
#>

function Create-ZeroTrustPresentation {
    param(
        [string]$OutputPath = "C:\Users\SAM\Desktop\Hexnode-ZeroTrust-Pitch.pptx"
    )
    
    # Create PowerPoint application
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = [Microsoft.Office.Interop.PowerPoint.MsoTriState]::msoTrue
    
    # Create new presentation
    $presentation = $pptApp.Presentations.Add()
    
    # Slide dimensions (widescreen 16:9)
    $slideWidth = 13.333 * 72  # 960 points
    $slideHeight = 7.5 * 72   # 540 points
    
    # Color scheme
    $primaryColor = [System.Drawing.Color]::FromArgb(0, 217, 255)    # Cyan
    $secondaryColor = [System.Drawing.Color]::FromArgb(0, 255, 136)  # Green
    $darkBg = [System.Drawing.Color]::FromArgb(26, 26, 46)           # Dark blue
    $accentColor = [System.Drawing.Color]::FromArgb(255, 107, 107)   # Red accent
    $white = [System.Drawing.Color]::White
    $lightGray = [System.Drawing.Color]::FromArgb(200, 200, 200)
    
    # Helper function to add a slide
    function Add-Slide {
        param(
            [int]$Layout = 12,  # ppLayoutBlank
            [string]$Title,
            [string]$Subtitle = ""
        )
        $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, $Layout)
        
        # Set background color
        $slide.FollowMasterBackground = [Microsoft.Office.Interop.PowerPoint.MsoTriState]::msoFalse
        $slide.Background.Fill.Solid()
        $slide.Background.Fill.ForeColor.RGB = [System.Drawing.Color]::ToArgb($darkBg)
        
        # Add title if provided
        if ($Title) {
            $titleBox = $slide.Shapes.AddTextbox(1, 50, 30, 860, 60)
            $titleBox.TextFrame.TextRange.Text = $Title
            $titleBox.TextFrame.TextRange.Font.Size = 36
            $titleBox.TextFrame.TextRange.Font.Bold = $true
            $titleBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($primaryColor)
        }
        
        # Add subtitle if provided
        if ($Subtitle) {
            $subBox = $slide.Shapes.AddTextbox(1, 50, 100, 860, 40)
            $subBox.TextFrame.TextRange.Text = $Subtitle
            $subBox.TextFrame.TextRange.Font.Size = 18
            $subBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($lightGray)
        }
        
        return $slide
    }
    
    # Helper function to add text box
    function Add-TextBox {
        param(
            [object]$Slide,
            [string]$Text,
            [int]$X = 50,
            [int]$Y = 150,
            [int]$Width = 860,
            [int]$Height = 300,
            [int]$FontSize = 16,
            [System.Drawing.Color]$Color = $lightGray,
            [bool]$Bold = $false
        )
        $textBox = $Slide.Shapes.AddTextbox(1, $X, $Y, $Width, $Height)
        $textBox.TextFrame.TextRange.Text = $Text
        $textBox.TextFrame.TextRange.Font.Size = $FontSize
        $textBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($Color)
        $textBox.TextFrame.TextRange.Font.Bold = $Bold
        $textBox.TextFrame.WordWrap = $true
        return $textBox
    }
    
    # Helper function to add bullet points
    function Add-BulletPoints {
        param(
            [object]$Slide,
            [string[]]$Points,
            [int]$X = 50,
            [int]$Y = 150,
            [int]$Width = 860,
            [int]$Height = 350,
            [int]$FontSize = 16
        )
        $textBox = $Slide.Shapes.AddTextbox(1, $X, $Y, $Width, $Height)
        $textBox.TextFrame.WordWrap = $true
        
        for ($i = 0; $i -lt $Points.Count; $i++) {
            if ($i -eq 0) {
                $textBox.TextFrame.TextRange.Text = "• $($Points[$i])"
            } else {
                $textBox.TextFrame.TextRange.InsertAfter("`n• $($Points[$i])")
            }
        }
        
        $textBox.TextFrame.TextRange.Font.Size = $FontSize
        $textBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($lightGray)
        
        return $textBox
    }
    
    # Helper function to add colored box
    function Add-ColorBox {
        param(
            [object]$Slide,
            [string]$Text,
            [int]$X,
            [int]$Y,
            [int]$Width,
            [int]$Height,
            [System.Drawing.Color]$FillColor,
            [System.Drawing.Color]$TextColor = $white,
            [int]$FontSize = 14
        )
        $shape = $Slide.Shapes.AddShape(1, $X, $Y, $Width, $Height)  # Rectangle
        $shape.Fill.Solid()
        $shape.Fill.ForeColor.RGB = [System.Drawing.Color]::ToArgb($FillColor)
        $shape.Line.Visible = $false
        $shape.TextFrame.TextRange.Text = $Text
        $shape.TextFrame.TextRange.Font.Size = $FontSize
        $shape.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($TextColor)
        $shape.TextFrame.TextRange.Font.Bold = $true
        $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # Center
        $shape.TextFrame.VerticalAnchor = 3  # Middle
        return $shape
    }
    
    Write-Host "Creating Hexnode Zero Trust Pitch Presentation..." -ForegroundColor Green
    
    # ============================================
    # SLIDE 1: Title Slide
    # ============================================
    $slide1 = Add-Slide -Title ""
    
    # Main title
    $titleBox = $slide1.Shapes.AddTextbox(1, 50, 120, 860, 100)
    $titleBox.TextFrame.TextRange.Text = "HEXNODE ZERO TRUST"
    $titleBox.TextFrame.TextRange.Font.Size = 54
    $titleBox.TextFrame.TextRange.Font.Bold = $true
    $titleBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($primaryColor)
    
    # Subtitle
    $subBox = $slide1.Shapes.AddTextbox(1, 50, 230, 860, 60)
    $subBox.TextFrame.TextRange.Text = "AUTOMATION SUITE"
    $subBox.TextFrame.TextRange.Font.Size = 40
    $subBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($secondaryColor)
    
    # Tagline
    $tagBox = $slide1.Shapes.AddTextbox(1, 50, 310, 860, 40)
    $tagBox.TextFrame.TextRange.Text = "Complete Enterprise Security Automation Framework"
    $tagBox.TextFrame.TextRange.Font.Size = 20
    $tagBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($lightGray)
    
    # Bottom bar
    Add-ColorBox -Slide $slide1 -Text "Powered by Hexnode UEM | REST API | AI Automation" -X 0 -Y 480 -Width 960 -Height 60 -FillColor $primaryColor -TextColor $darkBg -FontSize 14
    
    # ============================================
    # SLIDE 2: Executive Summary
    # ============================================
    $slide2 = Add-Slide -Title "EXECUTIVE SUMMARY"
    
    Add-BulletPoints -Slide $slide2 -Points @(
        "Complete Zero Trust security automation for enterprise endpoints"
        "Automated compliance checking, vulnerability scanning, and remediation"
        "Real-time SIEM integration with Splunk, Carbon Black, CrowdStrike"
        "85% faster threat response with < 5 minute remediation"
        "100% policy compliance across all managed devices"
        "Seamless integration with existing Hexnode UEM infrastructure"
        "ROI: Reduce security incidents by 60% and IT workload by 40%"
    ) -Y 130
    
    # ============================================
    # SLIDE 3: Problem Statement
    # ============================================
    $slide3 = Add-Slide -Title "THE CHALLENGE"
    
    Add-BulletPoints -Slide $slide3 -Points @(
        "Manual security processes are slow and error-prone"
        "Average time to detect threats: 207 days (IBM 2024)"
        "Average time to contain breaches: 73 days"
        "Lack of real-time compliance monitoring"
        "Siloed security tools without automation"
        "Growing attack surface with remote workforce"
        "Regulatory compliance pressure (GDPR, HIPAA, SOC2)"
        "IT teams overwhelmed with routine security tasks"
    ) -Y 130
    
    # ============================================
    # SLIDE 4: Solution Overview
    # ============================================
    $slide4 = Add-Slide -Title "OUR SOLUTION"
    
    Add-BulletPoints -Slide $slide4 -Points @(
        "Zero Trust Architecture: Never trust, always verify"
        "Automated Compliance: Real-time device health monitoring"
        "Intelligent Remediation: AI-powered threat response"
        "Unified Management: Single pane of glass via Hexnode UEM"
        "API-First Design: Integrate with any security tool"
        "Scalable: From 10 to 10,000+ devices"
    ) -Y 130
    
    # Key metrics box
    Add-ColorBox -Slide $slide4 -Text "85% Faster Response" -X 50 -Y 400 -Width 200 -Height 60 -FillColor $secondaryColor -TextColor $darkBg -FontSize 16
    Add-ColorBox -Slide $slide4 -Text "99.9% Uptime" -X 280 -Y 400 -Width 200 -Height 60 -FillColor $primaryColor -TextColor $darkBg -FontSize 16
    Add-ColorBox -Slide $slide4 -Text "100% Compliance" -X 510 -Y 400 -Width 200 -Height 60 -FillColor $accentColor -TextColor $white -FontSize 16
    
    # ============================================
    # SLIDE 5: Architecture
    # ============================================
    $slide5 = Add-Slide -Title "ARCHITECTURE OVERVIEW"
    
    # SIEM box
    Add-ColorBox -Slide $slide5 -Text "SIEM ALERTS`nSplunk | Carbon Black | CrowdStrike" -X 50 -Y 180 -Width 200 -Height 100 -FillColor $accentColor -TextColor $white -FontSize 12
    
    # Arrow
    $arrow1 = $slide5.Shapes.AddTextbox(1, 260, 210, 50, 40)
    $arrow1.TextFrame.TextRange.Text = "→"
    $arrow1.TextFrame.TextRange.Font.Size = 36
    $arrow1.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($secondaryColor)
    
    # Hexnode box
    Add-ColorBox -Slide $slide5 -Text "HEXNODE UEM`nAPI Gateway | Automation Engine" -X 320 -Y 180 -Width 220 -Height 100 -FillColor $primaryColor -TextColor $darkBg -FontSize 12
    
    # Arrow
    $arrow2 = $slide5.Shapes.AddTextbox(1, 550, 210, 50, 40)
    $arrow2.TextFrame.TextRange.Text = "→"
    $arrow2.TextFrame.TextRange.Font.Size = 36
    $arrow2.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($secondaryColor)
    
    # Devices box
    Add-ColorBox -Slide $slide5 -Text "DEVICES`nWindows | macOS | Linux" -X 610 -Y 180 -Width 200 -Height 100 -FillColor $secondaryColor -TextColor $darkBg -FontSize 12
    
    # Remediation box
    Add-ColorBox -Slide $slide5 -Text "AUTOMATED REMEDIATION`nQuarantine | Patch | Enforce | Report" -X 200 -Y 320 -Width 500 -Height 80 -FillColor $darkBg -TextColor $primaryColor -FontSize 14
    
    # ============================================
    # SLIDE 6: Core Features - Compliance
    # ============================================
    $slide6 = Add-Slide -Title "FEATURE: COMPLIANCE CHECK"
    
    Add-BulletPoints -Slide $slide6 -Points @(
        "Real-time device compliance monitoring"
        "Checks: Firewall, Encryption, Defender, Updates"
        "Jailbreak/Root detection"
        "Security service verification"
        "Automatic quarantine for non-compliant devices"
        "Custom compliance rules engine"
    ) -Y 130
    
    Add-ColorBox -Slide $slide6 -Text "Auto-Quarantine Non-Compliant Devices" -X 200 -Y 400 -Width 500 -Height 60 -FillColor $secondaryColor -TextColor $darkBg -FontSize 16
    
    # ============================================
    # SLIDE 7: Core Features - Vulnerability
    # ============================================
    $slide7 = Add-Slide -Title "FEATURE: VULNERABILITY SCANNING"
    
    Add-BulletPoints -Slide $slide7 -Points @(
        "CVE detection and patch verification"
        "Software inventory and version auditing"
        "Suspicious process detection"
        "Malicious network connection monitoring"
        "Risk scoring (Critical/High/Medium/Low)"
        "Automated vulnerability remediation"
    ) -Y 130
    
    Add-ColorBox -Slide $slide7 -Text "Risk Score: 0-100 Scale" -X 200 -Y 400 -Width 500 -Height 60 -FillColor $primaryColor -TextColor $darkBg -FontSize 16
    
    # ============================================
    # SLIDE 8: Core Features - Firewall
    # ============================================
    $slide8 = Add-Slide -Title "FEATURE: FIREWALL ENFORCEMENT"
    
    Add-BulletPoints -Slide $slide8 -Points @(
        "Windows Firewall profile management"
        "Dangerous port blocking (RDP, SMB, etc.)"
        "Malicious IP blacklisting"
        "Security service enforcement"
        "Unnecessary service disabling"
        "Firewall logging and monitoring"
    ) -Y 130
    
    Add-ColorBox -Slide $slide8 -Text "Block: 3389, 445, 135, 137, 138, 139" -X 200 -Y 400 -Width 500 -Height 60 -FillColor $accentColor -TextColor $white -FontSize 16
    
    # ============================================
    # SLIDE 9: Core Features - OS Updates
    # ============================================
    $slide9 = Add-Slide -Title "FEATURE: OS UPDATES AUTOMATION"
    
    Add-BulletPoints -Slide $slide9 -Points @(
        "Automated Windows Update orchestration"
        "Critical patch prioritization"
        "Background download and install"
        "Reboot requirement detection"
        "Update compliance reporting"
        "Maintenance window support"
    ) -Y 130
    
    Add-ColorBox -Slide $slide9 -Text "Auto-Patch Critical Vulnerabilities" -X 200 -Y 400 -Width 500 -Height 60 -FillColor $secondaryColor -TextColor $darkBg -FontSize 16
    
    # ============================================
    # SLIDE 10: Core Features - Provisioning
    # ============================================
    $slide10 = Add-Slide -Title "FEATURE: BULK PROVISIONING"
    
    Add-BulletPoints -Slide $slide10 -Points @(
        "CSV-based device inventory import"
        "Automatic device group assignment"
        "Policy application by location/department"
        "App deployment automation"
        "Welcome message and onboarding"
        "Scalable from 10 to 10,000+ devices"
    ) -Y 130
    
    Add-ColorBox -Slide $slide10 -Text "Enterprise-Scale Device Onboarding" -X 200 -Y 400 -Width 500 -Height 60 -FillColor $primaryColor -TextColor $darkBg -FontSize 16
    
    # ============================================
    # SLIDE 11: Core Features - SIEM
    # ============================================
    $slide11 = Add-Slide -Title "FEATURE: SIEM INTEGRATION"
    
    Add-BulletPoints -Slide $slide11 -Points @(
        "Splunk HEC webhook integration"
        "Carbon Black EDR alert handling"
        "CrowdStrike Falcon XDR support"
        "Generic webhook endpoint"
        "Severity-based automated response"
        "Real-time alert processing"
    ) -Y 130
    
    Add-ColorBox -Slide $slide11 -Text "Critical: Wipe | High: Quarantine | Medium: Remediate" -X 100 -Y 400 -Width 700 -Height 60 -FillColor $darkBg -TextColor $primaryColor -FontSize 14
    
    # ============================================
    # SLIDE 12: Integration Ecosystem
    # ============================================
    $slide12 = Add-Slide -Title "INTEGRATION ECOSYSTEM"
    
    Add-ColorBox -Slide $slide12 -Text "Splunk" -X 50 -Y 180 -Width 150 -Height 80 -FillColor $accentColor -TextColor $white -FontSize 16
    Add-ColorBox -Slide $slide12 -Text "Carbon Black" -X 220 -Y 180 -Width 150 -Height 80 -FillColor $accentColor -TextColor $white -FontSize 16
    Add-ColorBox -Slide $slide12 -Text "CrowdStrike" -X 390 -Y 180 -Width 150 -Height 80 -FillColor $accentColor -TextColor $white -FontSize 16
    Add-ColorBox -Slide $slide12 -Text "HEXNODE API" -X 580 -Y 180 -Width 200 -Height 80 -FillColor $primaryColor -TextColor $darkBg -FontSize 18
    
    Add-BulletPoints -Slide $slide12 -Points @(
        "REST API for custom integrations"
        "Webhook support for real-time alerts"
        "PowerShell and Python SDKs"
        "CI/CD pipeline integration"
        "Ticketing system integration (ServiceNow, Jira)"
    ) -Y 300
    
    # ============================================
    # SLIDE 13: Security Benefits
    # ============================================
    $slide13 = Add-Slide -Title "SECURITY BENEFITS"
    
    Add-ColorBox -Slide $slide13 -Text "ZERO TRUST`nNever Trust, Always Verify" -X 50 -Y 180 -Width 250 -Height 120 -FillColor $primaryColor -TextColor $darkBg -FontSize 14
    Add-ColorBox -Slide $slide13 -Text "REAL-TIME`n24/7 Monitoring" -X 330 -Y 180 -Width 250 -Height 120 -FillColor $secondaryColor -TextColor $darkBg -FontSize 14
    Add-ColorBox -Slide $slide13 -Text "AUTOMATED`n< 5 Min Response" -X 610 -Y 180 -Width 250 -Height 120 -FillColor $accentColor -TextColor $white -FontSize 14
    
    Add-BulletPoints -Slide $slide13 -Points @(
        "Least Privilege Access Control"
        "Micro-segmentation of Network"
        "Continuous Security Validation"
        "Automated Threat Response"
        "Compliance as Code"
    ) -Y 330
    
    # ============================================
    # SLIDE 14: ROI & Value Proposition
    # ============================================
    $slide14 = Add-Slide -Title "ROI & VALUE PROPOSITION"
    
    Add-ColorBox -Slide $slide14 -Text "60%`nFewer Security Incidents" -X 50 -Y 180 -Width 250 -Height 120 -FillColor $secondaryColor -TextColor $darkBg -FontSize 14
    Add-ColorBox -Slide $slide14 -Text "40%`nIT Workload Reduction" -X 330 -Y 180 -Width 250 -Height 120 -FillColor $primaryColor -TextColor $darkBg -FontSize 14
    Add-ColorBox -Slide $slide14 -Text "85%`nFaster Threat Response" -X 610 -Y 180 -Width 250 -Height 120 -FillColor $accentColor -TextColor $white -FontSize 14
    
    Add-BulletPoints -Slide $slide14 -Points @(
        "Reduced mean time to detect (MTTD) from days to minutes"
        "Reduced mean time to respond (MTTR) from hours to seconds"
        "Lower compliance audit costs with automated reporting"
        "Improved security posture with continuous monitoring"
        "Better resource utilization with automated routine tasks"
    ) -Y 330
    
    # ============================================
    # SLIDE 15: Implementation Roadmap
    # ============================================
    $slide15 = Add-Slide -Title "IMPLEMENTATION ROADMAP"
    
    Add-ColorBox -Slide $slide15 -Text "PHASE 1`nWeek 1-2`nSetup & Config" -X 50 -Y 180 -Width 180 -Height 100 -FillColor $primaryColor -TextColor $darkBg -FontSize 12
    Add-ColorBox -Slide $slide15 -Text "PHASE 2`nWeek 3-4`nScript Deploy" -X 250 -Y 180 -Width 180 -Height 100 -FillColor $secondaryColor -TextColor $darkBg -FontSize 12
    Add-ColorBox -Slide $slide15 -Text "PHASE 3`nWeek 5-6`nAutomation" -X 450 -Y 180 -Width 180 -Height 100 -FillColor $primaryColor -TextColor $darkBg -FontSize 12
    Add-ColorBox -Slide $slide15 -Text "PHASE 4`nWeek 7-8`nSIEM Integ" -X 650 -Y 180 -Width 180 -Height 100 -FillColor $accentColor -TextColor $white -FontSize 12
    
    Add-BulletPoints -Slide $slide15 -Points @(
        "Phase 1: Configure Hexnode API, create quarantine groups"
        "Phase 2: Upload and test all automation scripts"
        "Phase 3: Create automations with triggers and schedules"
        "Phase 4: Integrate with SIEM and test alert handling"
        "Phase 5: Production deployment and monitoring"
    ) -Y 310
    
    # ============================================
    # SLIDE 16: Technical Requirements
    # ============================================
    $slide16 = Add-Slide -Title "TECHNICAL REQUIREMENTS"
    
    Add-BulletPoints -Slide $slide16 -Points @(
        "Hexnode UEM subscription with API access"
        "Windows 10/11 devices (managed by Hexnode)"
        "PowerShell 5.1+ (for Windows scripts)"
        "Python 3.8+ (for API integrations)"
        "Network connectivity to Hexnode API endpoints"
        "SIEM system (optional: Splunk, Carbon Black, CrowdStrike)"
        "Admin privileges on target devices"
    ) -Y 130
    
    Add-ColorBox -Slide $slide16 -Text "Cloud-Based | No Additional Hardware Required" -X 200 -Y 420 -Width 500 -Height 60 -FillColor $secondaryColor -TextColor $darkBg -FontSize 16
    
    # ============================================
    # SLIDE 17: Quick Start
    # ============================================
    $slide17 = Add-Slide -Title "QUICK START GUIDE"
    
    Add-ColorBox -Slide $slide17 -Text "STEP 1`nConfigure API" -X 50 -Y 180 -Width 180 -Height 100 -FillColor $primaryColor -TextColor $darkBg -FontSize 14
    Add-ColorBox -Slide $slide17 -Text "STEP 2`nUpload Scripts" -X 250 -Y 180 -Width 180 -Height 100 -FillColor $secondaryColor -TextColor $darkBg -FontSize 14
    Add-ColorBox -Slide $slide17 -Text "STEP 3`nCreate Automation" -X 450 -Y 180 -Width 180 -Height 100 -FillColor $primaryColor -TextColor $darkBg -FontSize 14
    Add-ColorBox -Slide $slide17 -Text "STEP 4`nEnable & Run" -X 650 -Y 180 -Width 180 -Height 100 -FillColor $accentColor -TextColor $white -FontSize 14
    
    Add-BulletPoints -Slide $slide17 -Points @(
        "Edit hexnode_config.json with your API credentials"
        "Upload .ps1 scripts to Hexnode Custom Scripts"
        "Create automations with time/activity triggers"
        "Enable automations and monitor results"
    ) -Y 310
    
    # ============================================
    # SLIDE 18: Next Steps
    # ============================================
    $slide18 = Add-Slide -Title "NEXT STEPS"
    
    Add-BulletPoints -Slide $slide18 -Points @(
        "Schedule technical deep-dive demonstration"
        "Identify pilot group for proof of concept"
        "Review Hexnode API documentation"
        "Plan integration with existing SIEM infrastructure"
        "Define compliance requirements and policies"
        "Establish success metrics and KPIs"
    ) -Y 130
    
    Add-ColorBox -Slide $slide18 -Text "Ready to Start? Let's Schedule a Demo!" -X 200 -Y 420 -Width 500 -Height 60 -FillColor $secondaryColor -TextColor $darkBg -FontSize 18
    
    # ============================================
    # SLIDE 19: Thank You
    # ============================================
    $slide19 = Add-Slide -Title ""
    
    $titleBox = $slide19.Shapes.AddTextbox(1, 50, 150, 860, 80)
    $titleBox.TextFrame.TextRange.Text = "THANK YOU"
    $titleBox.TextFrame.TextRange.Font.Size = 54
    $titleBox.TextFrame.TextRange.Font.Bold = $true
    $titleBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($primaryColor)
    
    $subBox = $slide19.Shapes.AddTextbox(1, 50, 250, 860, 60)
    $subBox.TextFrame.TextRange.Text = "Questions & Discussion"
    $subBox.TextFrame.TextRange.Font.Size = 28
    $subBox.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::ToArgb($secondaryColor)
    
    Add-ColorBox -Slide $slide19 -Text "Hexnode Zero Trust Automation Suite" -X 200 -Y 380 -Width 500 -Height 60 -FillColor $primaryColor -TextColor $darkBg -FontSize 18
    
    # Save presentation
    $presentation.SaveAs($OutputPath)
    Write-Host "Presentation saved to: $OutputPath" -ForegroundColor Green
    
    # Cleanup
    $presentation.Close()
    $pptApp.Quit()
    
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null
    
    Write-Host "Presentation created successfully!" -ForegroundColor Green
    Write-Host "File: $OutputPath" -ForegroundColor Yellow
}

# Create the presentation
Create-ZeroTrustPresentation
