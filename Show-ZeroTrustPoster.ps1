<#
.SYNOPSIS
    Hexnode Zero Trust Automation Suite - Visual Poster
.DESCRIPTION
    Displays a marketing-style ASCII art poster in PowerShell console.
.NOTES
    Press any key to exit.
#>

function Show-ZeroTrustPoster {
    Clear-Host
    
    $TC = "Cyan"
    $HC = "Yellow"
    $WC = "White"
    $AC = "Green"
    $MC = "Magenta"
    $BC = "DarkCyan"
    
    Write-Host ""
    Write-Host "  +================================================================================+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *       ████████╗██╗  ██╗███████╗    ██████╗  ██████╗                            *" -ForegroundColor $TC
    Write-Host "  *       ╚══██╔══╝██║  ██║██╔════╝    ╚════██╗██╔═████╗                           *" -ForegroundColor $TC
    Write-Host "  *          ██║   ███████║███████╗     █████╔╝██║██╔██║                           *" -ForegroundColor $TC
    Write-Host "  *          ██║   ██╔══██║╚════██║    ██╔═══╝ ████╔╝██║                           *" -ForegroundColor $TC
    Write-Host "  *          ██║   ██║  ██║███████║    ███████╗╚██████╔╝                           *" -ForegroundColor $TC
    Write-Host "  *          ╚═╝   ╚═╝  ╚═╝╚══════╝    ╚══════╝ ╚═════╝                            *" -ForegroundColor $TC
    Write-Host "  *                            AUTOMATION SUITE                                     *" -ForegroundColor $AC
    Write-Host "  *             Complete Enterprise Security Automation Framework                   *" -ForegroundColor $WC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                           ARCHITECTURE FLOW                                      *" -ForegroundColor $HC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *      [SIEM ALERTS]          [HEXNODE UEM]          [DEVICES]                  *" -ForegroundColor $WC
    Write-Host "  *      +-----------+          +-----------+          +-----------+               *" -ForegroundColor $BC
    Write-Host "  *      .  Splunk   .          .   API     .          .  Windows  .               *" -ForegroundColor $MC
    Write-Host "  *      .  Carbon   .---------.  Gateway  .---------.  macOS    .               *" -ForegroundColor $AC
    Write-Host "  *      .  Black    .          .           .          .  Linux    .               *" -ForegroundColor $WC
    Write-Host "  *      .  Crowd    .          .  Auto     .          .  Endpoints.               *" -ForegroundColor $WC
    Write-Host "  *      .  Strike   .          .  Engine   .          .           .               *" -ForegroundColor $WC
    Write-Host "  *      +-----------+          +-----------+          +-----------+               *" -ForegroundColor $BC
    Write-Host "  *             |                     |                     |                     *" -ForegroundColor $BC
    Write-Host "  *             +---------------------+---------------------+                     *" -ForegroundColor $BC
    Write-Host "  *                                   |                                           *" -ForegroundColor $BC
    Write-Host "  *                                   v                                           *" -ForegroundColor $BC
    Write-Host "  *      +------------------------------------------------------------+          *" -ForegroundColor $BC
    Write-Host "  *      |                    AUTOMATED REMEDIATION                   |          *" -ForegroundColor $AC
    Write-Host "  *      |  +----------+  +----------+  +----------+  +----------+    |          *" -ForegroundColor $BC
    Write-Host "  *      |  |QUARANTINE |  | FIREWALL |  |  PATCH   |  | ENFORCE  |    |          *" -ForegroundColor $MC
    Write-Host "  *      |  |  DEVICE   |  |  BLOCK   |  |  UPDATE  |  |  POLICY  |    |          *" -ForegroundColor $WC
    Write-Host "  *      |  +----------+  +----------+  +----------+  +----------+    |          *" -ForegroundColor $BC
    Write-Host "  *      +------------------------------------------------------------+          *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                             CORE FEATURES                                       *" -ForegroundColor $HC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *  +---------------------------+  +---------------------------+  +--------------+ *" -ForegroundColor $BC
    Write-Host "  *  . [1] COMPLIANCE CHECK      .  . [2] VULNERABILITY SCAN    .  . [3] AUTO FIX . *" -ForegroundColor $AC
    Write-Host "  *  .   - Firewall Status       .  .   - CVE Detection         .  .  - Quarantine. *" -ForegroundColor $WC
    Write-Host "  *  .   - Encryption            .  .   - Software Audit        .  .  - Isolate   . *" -ForegroundColor $WC
    Write-Host "  *  .   - Defender Health       .  .   - Risk Scoring          .  .  - Remediate . *" -ForegroundColor $WC
    Write-Host "  *  +---------------------------+  +---------------------------+  +--------------+ *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *  +---------------------------+  +---------------------------+  +--------------+ *" -ForegroundColor $BC
    Write-Host "  *  . [4] FIREWALL CONTROL      .  . [5] OS UPDATES            .  . [6] PROVISION. *" -ForegroundColor $AC
    Write-Host "  *  .   - Port Blocking         .  .   - Auto-Download         .  .  - CSV Import. *" -ForegroundColor $WC
    Write-Host "  *  .   - IP Blacklist          .  .   - Auto-Install          .  .  - Policies  . *" -ForegroundColor $WC
    Write-Host "  *  .   - Service Mgmt          .  .   - Compliance Check      .  .  - Groups    . *" -ForegroundColor $WC
    Write-Host "  *  +---------------------------+  +---------------------------+  +--------------+ *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                         INTEGRATION ECOSYSTEM                                   *" -ForegroundColor $HC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *                        +------------------------+                              *" -ForegroundColor $BC
    Write-Host "  *                        |    HEXNODE REST API     |                              *" -ForegroundColor $AC
    Write-Host "  *                        +-----------+--------------+                              *" -ForegroundColor $BC
    Write-Host "  *                                    |                                             *" -ForegroundColor $BC
    Write-Host "  *          +-------------------------+-------------------------+                  *" -ForegroundColor $BC
    Write-Host "  *          |                         |                         |                  *" -ForegroundColor $BC
    Write-Host "  *          v                         v                         v                  *" -ForegroundColor $BC
    Write-Host "  *   +--------------+          +--------------+          +--------------+            *" -ForegroundColor $BC
    Write-Host "  *   |   SPLUNK     |          | CARBON BLACK |          | CROWDSTRIKE  |            *" -ForegroundColor $MC
    Write-Host "  *   |   ---------- |          |   ---------- |          |   ---------- |            *" -ForegroundColor $BC
    Write-Host "  *   |   - HEC      |          |   - EDR      |          |   - Falcon   |            *" -ForegroundColor $WC
    Write-Host "  *   |   - Alerts   |          |   - Alerts   |          |   - Alerts   |            *" -ForegroundColor $WC
    Write-Host "  *   |   - SIEM     |          |   - EPP      |          |   - XDR      |            *" -ForegroundColor $WC
    Write-Host "  *   +--------------+          +--------------+          +--------------+            *" -ForegroundColor $BC
    Write-Host "  *          |                         |                         |                  *" -ForegroundColor $BC
    Write-Host "  *          +-------------------------+-------------------------+                  *" -ForegroundColor $BC
    Write-Host "  *                                    |                                             *" -ForegroundColor $BC
    Write-Host "  *                                    v                                             *" -ForegroundColor $BC
    Write-Host "  *                        +------------------------+                              *" -ForegroundColor $BC
    Write-Host "  *                        |  AUTOMATED REMEDIATION  |                              *" -ForegroundColor $AC
    Write-Host "  *                        +------------------------+                              *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                           SECURITY BENEFITS                                      *" -ForegroundColor $HC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *  +------------------------------------------------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *  |                                                                        |   *" -ForegroundColor $BC
    Write-Host "  *  |   [A] ZERO TRUST         [B] REAL-TIME          [C] AUTOMATED         |   *" -ForegroundColor $AC
    Write-Host "  *  |       ARCHITECTURE           MONITORING              RESPONSE          |   *" -ForegroundColor $AC
    Write-Host "  *  |   --------------------     ------------------     ------------------  |   *" -ForegroundColor $BC
    Write-Host "  *  |   * Never Trust             * 24/7 Scanning       * < 5 min           |   *" -ForegroundColor $WC
    Write-Host "  *  |   * Always Verify           * Threat Detect       * Remediation        |   *" -ForegroundColor $WC
    Write-Host "  *  |   * Least Privilege         * Compliance          * Auto-Fix           |   *" -ForegroundColor $WC
    Write-Host "  *  |                                                                        |   *" -ForegroundColor $BC
    Write-Host "  *  |   +-----------------+  +-----------------+  +-----------------+        |   *" -ForegroundColor $BC
    Write-Host "  *  |   | 85% FASTER      |  | 99.9% UPTIME    |  | 100% COMPLIANT  |        |   *" -ForegroundColor $MC
    Write-Host "  *  |   | Threat Response  |  | Security        |  | Policy Enforce  |        |   *" -ForegroundColor $WC
    Write-Host "  *  |   +-----------------+  +-----------------+  +-----------------+        |   *" -ForegroundColor $BC
    Write-Host "  *  |                                                                        |   *" -ForegroundColor $BC
    Write-Host "  *  +------------------------------------------------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                          QUICK START GUIDE                                       *" -ForegroundColor $HC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *  +------------------------------------------------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *  |                                                                        |   *" -ForegroundColor $BC
    Write-Host "  *  |   STEP 1             STEP 2             STEP 3             STEP 4      |   *" -ForegroundColor $AC
    Write-Host "  *  |   +-----+            +-----+            +-----+            +-----+     |   *" -ForegroundColor $BC
    Write-Host "  *  |   |  1  |----------->|  2  |----------->|  3  |----------->|  4  |     |   *" -ForegroundColor $AC
    Write-Host "  *  |   +-----+            +-----+            +-----+            +-----+     |   *" -ForegroundColor $BC
    Write-Host "  *  |      |                  |                  |                  |         |   *" -ForegroundColor $BC
    Write-Host "  *  |      v                  v                  v                  v         |   *" -ForegroundColor $BC
    Write-Host "  *  |   +--------+        +--------+        +--------+        +--------+      |   *" -ForegroundColor $BC
    Write-Host "  *  |   |CONFIGURE|       | UPLOAD |       | CREATE |       | DEPLOY  |      |   *" -ForegroundColor $MC
    Write-Host "  *  |   |  API     |       | SCRIPTS|       |AUTOMAT |       |  AND   |      |   *" -ForegroundColor $WC
    Write-Host "  *  |   |  KEY     |       |   TO   |       | IONS   |       |  RUN   |      |   *" -ForegroundColor $WC
    Write-Host "  *  |   +--------+        +--------+        +--------+        +--------+      |   *" -ForegroundColor $BC
    Write-Host "  *  |      |                  |                  |                  |         |   *" -ForegroundColor $BC
    Write-Host "  *  |      v                  v                  v                  v         |   *" -ForegroundColor $BC
    Write-Host "  *  |   Edit config       Upload to         Set triggers       Enable        |   *" -ForegroundColor $WC
    Write-Host "  *  |   .json file        Hexnode UEM       and schedules      and run       |   *" -ForegroundColor $WC
    Write-Host "  *  |                                                                        |   *" -ForegroundColor $BC
    Write-Host "  *  +------------------------------------------------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                          SCRIPT INVENTORY                                        *" -ForegroundColor $HC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *  +--------------------------------------+-----------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *  |  SCRIPT NAME                         |  PURPOSE                          |   *" -ForegroundColor $BC
    Write-Host "  *  |--------------------------------------|-----------------------------------|   *" -ForegroundColor $BC
    Write-Host "  *  |  Device-Compliance-Quarantine.ps1    |  Auto-quarantine non-compliant   |   *" -ForegroundColor $AC
    Write-Host "  *  |  Vulnerability-Scanner.ps1           |  Scan for CVEs and threats       |   *" -ForegroundColor $AC
    Write-Host "  *  |  Enforce-FirewallServices.ps1        |  Block ports and enforce policy  |   *" -ForegroundColor $AC
    Write-Host "  *  |  Orchestrate-OSUpdates.ps1           |  Automate Windows patches        |   *" -ForegroundColor $AC
    Write-Host "  *  |  Bulk-Provision-Devices.ps1          |  CSV-based device onboarding     |   *" -ForegroundColor $AC
    Write-Host "  *  |  hexnode_security_api.py             |  Python API client               |   *" -ForegroundColor $AC
    Write-Host "  *  |  siem_webhook_server.py              |  SIEM integration server         |   *" -ForegroundColor $AC
    Write-Host "  *  +--------------------------------------+-----------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +--------------------------------------------------------------------------------+" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  *   +------------------------------------------------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *   |                                                                    |   |   *" -ForegroundColor $BC
    Write-Host "  *   |        READY TO IMPLEMENT ZERO TRUST AUTOMATION?                   |   |   *" -ForegroundColor $MC
    Write-Host "  *   |                                                                    |   |   *" -ForegroundColor $BC
    Write-Host "  *   |   Location: C:\Users\SAM\Hexnode-ZeroTrust                         |   |   *" -ForegroundColor $WC
    Write-Host "  *   |   Config:   Config\hexnode_config.json                            |   |   *" -ForegroundColor $WC
    Write-Host "  *   |   Start:    .\Invoke-ZeroTrustOrchestration.ps1 -RunAll           |   |   *" -ForegroundColor $WC
    Write-Host "  *   |                                                                    |   |   *" -ForegroundColor $BC
    Write-Host "  *   +------------------------------------------------------------------------+   *" -ForegroundColor $BC
    Write-Host "  *                                                                                *" -ForegroundColor $BC
    Write-Host "  +================================================================================+" -ForegroundColor $BC
    Write-Host ""
    Write-Host "    Hexnode Zero Trust Automation Suite v1.0" -ForegroundColor $AC
    Write-Host "    Enterprise Security Automation Framework" -ForegroundColor $WC
    Write-Host ""
    Write-Host "    Press any key to exit..." -ForegroundColor $WC
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Show-ZeroTrustPoster
