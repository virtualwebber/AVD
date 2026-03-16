<#
.SYNOPSIS
    Registers a scheduled task to run Set-AudioDeviceSettings.ps1 on device arrival.

.DESCRIPTION
    Creates a scheduled task triggered by Event ID 112 from the
    Microsoft-Windows-DeviceSetupManager/Admin log. This event fires whenever a
    device container has been fully serviced (e.g. USB headset plugged in via
    USB passthrough), at which point the device properties are written to the
    MMDevices registry and ready to be configured.

    WHY EVENT ID 112?
        When a USB device is connected (physically or via AVD USB passthrough),
        Windows Device Setup Manager (DSM) handles driver installation and
        property configuration. Event ID 112 is logged when DSM has finished
        servicing the device container — meaning all drivers are loaded and
        the device's registry entries (including MMDevices Audio) are populated.
        This is the ideal moment to apply our audio settings.

    WHY NOT WMI PERMANENT SUBSCRIPTIONS?
        The original approach used a WMI permanent event subscription with a
        CommandLineEventConsumer. This works in environments without security
        restrictions, but many enterprise environments block CommandLineEventConsumer
        via Attack Surface Reduction (ASR) rules (specifically rule
        d1e49aac-8f56-4280-b9ba-993a6d77406c — "Block process creations originating
        from PSExec and WMI commands"). The scheduled task approach avoids this
        restriction entirely. CreateWMI_Audio_Sub.ps1 remains in the repo for
        environments where WMI subscriptions are allowed.

.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Run once, elevated, during image build or on first boot.
    The task survives reboots and runs as SYSTEM.
    Log:     C:\_source\logs\RegisterAudioTask_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

# Name of the scheduled task as it appears in Task Scheduler
$taskName   = "AudioDeviceSettings"

# Path to the script that configures audio device settings (disable enhancements).
# This is what the scheduled task will execute when a device arrival event fires.
$scriptPath = "C:\_source\Set-AudioDeviceSettings.ps1"

# Local log directory — always written to
$logDir     = "C:\_source\logs"

# Optional UNC file share for centralised logging.
# Set to $null or "" to disable file share logging.
$logShare   = "\\cukavdukwprod01.file.core.windows.net\profiles\logs"

# Log file name includes computername for multi-host identification
$logName    = "${env:COMPUTERNAME}_RegisterAudioTask_$(Get-Date -Format 'yyyyMMdd').log"
$logFile    = Join-Path $logDir $logName

# Devices to exclude from triggering the scheduled task.
# Event ID 112 fires for ALL device types (printers, USB drives, etc.).
# Add device names here to prevent unnecessary script execution.
# Names must match the Prop_DeviceName field in the event data exactly.
$excludeDevices = @(
    "Remote Audio"         # RDP audio redirection — persists across sessions, not a real USB device
    # To add more devices, put each on a new line with a COMMA after the previous entry:
    # "Remote Audio",      # <-- note the comma at the end
    # "PDF Architect 9",   # <-- comma here too
    # "Some Other Device"  # <-- NO comma on the last entry
)

# ============================================================
# FUNCTIONS
# ============================================================

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped log entry to local file, optional file share, and console.
    .DESCRIPTION
        Logs are written to the local $logDir first (always), then optionally to
        $logShare if configured. If the share is unreachable, the error is silently
        caught so the script continues without interruption.
    #>
    param([string]$Message, [string]$Level = "INFO")

    # Ensure the local log directory exists
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Format: [timestamp] [LEVEL] message
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"

    # Always write to local log file
    Add-Content -Path $logFile -Value $entry

    # Optionally write to file share — silently skip if unreachable
    if ($logShare) {
        try {
            if (-not (Test-Path $logShare)) {
                New-Item -ItemType Directory -Path $logShare -Force -ErrorAction Stop | Out-Null
            }
            Add-Content -Path (Join-Path $logShare $logName) -Value $entry -ErrorAction Stop
        }
        catch {
            # Silently skip if the share is unreachable
        }
    }

    # Also output to console for interactive use
    Write-Host $entry
}

# ============================================================
# MAIN
# ============================================================

Write-Log "========================================"
Write-Log "Scheduled task registration started"
Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log "========================================"

# ----------------------------------------------------------
# STEP 1: Clean up any existing task with the same name.
# This ensures a fresh registration every time, avoiding
# stale triggers or misconfigured actions from previous runs.
# ----------------------------------------------------------
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Log "Removing existing task: $taskName"
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Log "  OK  Existing task removed"
} else {
    Write-Log "  No existing task found"
}

# ----------------------------------------------------------
# STEP 2: Define the scheduled task components.
# ----------------------------------------------------------
Write-Log "Creating scheduled task: $taskName"

# ACTION: Launch PowerShell to execute the audio settings script.
# -NonInteractive  — no prompts (runs unattended as SYSTEM)
# -WindowStyle Hidden — no console window visible to users
# -ExecutionPolicy Bypass — avoids policy restrictions on the script
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# TRIGGER: Event ID 112 from Microsoft-Windows-DeviceSetupManager/Admin.
# This fires when a device container has been fully serviced by DSM.
# The XPath filter excludes devices listed in $excludeDevices so the task
# only fires for real USB devices, not RDP redirections or other noise.
# We use the CIM class MSFT_TaskEventTrigger to create a native event trigger,
# as New-ScheduledTaskTrigger does not support event-based triggers directly.

# Build XPath exclusion clauses from the $excludeDevices array.
# Each entry becomes: Data[@Name='Prop_DeviceName'] != 'device name'
# Multiple exclusions are joined with " and " so all must be true.
$exclusions = ($excludeDevices | ForEach-Object { "Data[@Name='Prop_DeviceName'] != '$_'" }) -join " and "
$xpathQuery = "*[System[EventID=112] and EventData[$exclusions]]"

Write-Log "XPath query: $xpathQuery"

$triggerClass = Get-CimClass -ClassName "MSFT_TaskEventTrigger" -Namespace "Root/Microsoft/Windows/TaskScheduler"
$trigger = New-CimInstance -CimClass $triggerClass -ClientOnly
$trigger.Subscription = @"
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-DeviceSetupManager/Admin">
    <Select Path="Microsoft-Windows-DeviceSetupManager/Admin">$xpathQuery</Select>
  </Query>
</QueryList>
"@
$trigger.Enabled = $true

# PRINCIPAL: Run as SYSTEM with highest privileges.
# SYSTEM is required because the MMDevices registry keys are owned by
# AudioEndpointBuilder and need elevated access to modify.
$principal = New-ScheduledTaskPrincipal `
    -UserId "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

# SETTINGS: Ensure the task runs reliably in all conditions.
# -AllowStartIfOnBatteries   — run even on battery (laptops/tablets)
# -DontStopIfGoingOnBatteries — don't kill the task if power is removed
# -MultipleInstances Queue    — queue runs if multiple devices arrive rapidly
# -StartWhenAvailable         — run missed triggers if the machine was busy
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances Queue `
    -StartWhenAvailable

# ----------------------------------------------------------
# STEP 3: Register the scheduled task.
# -Force overwrites if it somehow still exists after cleanup.
# ----------------------------------------------------------
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Configures audio device settings when a new device is installed (Event ID 112)" `
    -Force | Out-Null

# ----------------------------------------------------------
# STEP 4: Verify the task was created successfully.
# ----------------------------------------------------------
$verifyTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($verifyTask) {
    Write-Log "  OK  Scheduled task created"
    Write-Log "========================================"
    Write-Log "Scheduled task registered successfully"
    Write-Log "  Task:    $taskName"
    Write-Log "  Script:  $scriptPath"
    Write-Log "  Trigger: Event ID 112 (DeviceSetupManager)"
    Write-Log "  Run as:  SYSTEM (highest privileges)"
    Write-Log "========================================"
} else {
    Write-Log "FAIL  Scheduled task was not created" -Level "WARN"
    Write-Log "========================================"
}
