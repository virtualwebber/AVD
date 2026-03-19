<#
.SYNOPSIS
    Registers a scheduled task to run Remove-AudioDeviceKeys.ps1 on device disconnection.

.DESCRIPTION
    Creates a scheduled task triggered by Event ID 1010 from the
    Microsoft-Windows-Kernel-PnP/Device Management log. This event fires whenever
    a device is surprise-removed (e.g. USB headset unplugged or AVD USB passthrough
    disconnected), at which point the device's registry state changes to NotPresent
    and the stale keys can be cleaned up.

    WHY EVENT ID 1010?
        When a USB device is disconnected (physically or via AVD USB passthrough),
        the Kernel Plug and Play manager logs Event ID 1010 -- "Device has been
        surprise removed as it is reported as missing on the bus." This fires
        immediately on unplug and is enabled by default in all Windows installations.
        Event ID 1011 (device reported as failing) is also included as a fallback.

    The Remove-AudioDeviceKeys.ps1 script handles filtering -- it only removes
    audio device keys with DeviceState = 4, so triggering on all USB removals
    is harmless (the script simply exits if there is nothing to clean up).

.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Run once, elevated, during image build or on first boot.
    The task survives reboots and runs as SYSTEM.
    Log:     C:\_source\logs\RegisterAudioRemovalTask_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

# Name of the scheduled task as it appears in Task Scheduler
$taskName   = "AudioDeviceRemoval"

# Path to the script that removes stale audio device registry keys.
# This is what the scheduled task will execute when a device removal event fires.
$scriptPath = "C:\_source\Remove-AudioDeviceKeys.ps1"

# Local log directory -- always written to
$logDir     = "C:\_source\logs"

# Optional UNC file share for centralised logging.
# Set to $null or "" to disable file share logging.
$logShare   = "\\fileserver.domain.com\logs\audio"

# Log file name includes computername for multi-host identification
$logName    = "${env:COMPUTERNAME}_RegisterAudioRemovalTask_$(Get-Date -Format 'yyyyMMdd').log"
$logFile    = Join-Path $logDir $logName

# ============================================================
# FUNCTIONS
# ============================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $logFile -Value $entry
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

# ACTION: Launch PowerShell to execute the removal script.
# -NonInteractive  -- no prompts (runs unattended as SYSTEM)
# -WindowStyle Hidden -- no console window visible to users
# -ExecutionPolicy Bypass -- avoids policy restrictions on the script
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# TRIGGER: Event ID 1010 (or 1011) from Microsoft-Windows-Kernel-PnP/Device Management.
# Event 1010 fires when a device is surprise-removed (missing on bus).
# Event 1011 fires when a device is surprise-removed (reported failing).
# NOTE: Windows Event Log XPath does not support starts-with(), so we cannot filter
# by DeviceInstanceId here. The removal script handles audio-only filtering itself.
# We use the CIM class MSFT_TaskEventTrigger to create a native event trigger,
# as New-ScheduledTaskTrigger does not support event-based triggers directly.

$xpathQuery = "*[System[(EventID=1010 or EventID=1011)]]"

Write-Log "XPath query: $xpathQuery"

$triggerClass = Get-CimClass -ClassName "MSFT_TaskEventTrigger" -Namespace "Root/Microsoft/Windows/TaskScheduler"
$trigger = New-CimInstance -CimClass $triggerClass -ClientOnly
$trigger.Subscription = @"
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-Kernel-PnP/Device Management">
    <Select Path="Microsoft-Windows-Kernel-PnP/Device Management">$xpathQuery</Select>
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
# -AllowStartIfOnBatteries   -- run even on battery (laptops/tablets)
# -DontStopIfGoingOnBatteries -- don't kill the task if power is removed
# -MultipleInstances Queue    -- queue runs if multiple devices are removed rapidly
# -StartWhenAvailable         -- run missed triggers if the machine was busy
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
    -Description "Removes stale audio device registry keys when a USB device is disconnected (Kernel-PnP Event 1010/1011)" `
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
    Write-Log "  Trigger: Event ID 1010/1011 (Kernel-PnP device removal)"
    Write-Log "  Log:     Microsoft-Windows-Kernel-PnP/Device Management"
    Write-Log "  Run as:  SYSTEM (highest privileges)"
    Write-Log "========================================"
} else {
    Write-Log "FAIL  Scheduled task was not created" -Level "WARN"
    Write-Log "========================================"
}
