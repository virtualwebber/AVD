<#
.SYNOPSIS
    Registers a scheduled task to run Set-AudioDeviceSettings.ps1 on device arrival.
.DESCRIPTION
    Creates a scheduled task triggered by Event ID 112 from the
    Microsoft-Windows-DeviceSetupManager/Admin log. This event fires whenever a
    device container has been fully serviced (e.g. USB headset plugged in via
    USB passthrough), at which point the device properties are written to the
    MMDevices registry and ready to be configured.

    This replaces the WMI permanent subscription approach in environments where
    CommandLineEventConsumer is blocked by security policy (e.g. ASR rules).
.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Run once, elevated, during image build or on first boot.
    Log:     C:\_source\logs\RegisterAudioTask_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

$taskName   = "AudioDeviceSettings"
$scriptPath = "C:\_source\Set-AudioDeviceSettings.ps1"
$logDir     = "C:\_source\logs"
$logShare   = "\\fileserver.domain.com\logs\audio"
$logName    = "${env:COMPUTERNAME}_RegisterAudioTask_$(Get-Date -Format 'yyyyMMdd').log"
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

# Remove existing task if present
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Log "Removing existing task: $taskName"
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Log "  OK  Existing task removed"
} else {
    Write-Log "  No existing task found"
}

# Create the scheduled task
Write-Log "Creating scheduled task: $taskName"

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# Trigger on Event ID 112 from DeviceSetupManager — fires when a device
# container has been fully serviced and properties written to the registry.
$trigger = New-ScheduledTaskTrigger -AtLogOn
$triggerXml = @"
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-DeviceSetupManager/Admin">
    <Select Path="Microsoft-Windows-DeviceSetupManager/Admin">*[System[EventID=112]]</Select>
  </Query>
</QueryList>
"@

$principal = New-ScheduledTaskPrincipal `
    -UserId "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances Queue `
    -StartWhenAvailable

# Register with a placeholder trigger, then update with the event trigger via XML
$task = Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Configures audio device settings when a new device is installed (Event ID 112)" `
    -Force

# Replace the placeholder trigger with the event-based trigger
$taskDef = $task | Get-ScheduledTask
$eventTrigger = New-Object -ComObject "Schedule.Service"
$eventTrigger.Connect()
$folder   = $eventTrigger.GetFolder("\")
$taskObj  = $folder.GetTask($taskName)
$taskXml  = $taskObj.Xml

# Inject the event trigger XML
$taskXml = $taskXml -replace '(?s)<Triggers>.*?</Triggers>', @"
<Triggers>
  <EventTrigger>
    <Enabled>true</Enabled>
    <Subscription>$triggerXml</Subscription>
  </EventTrigger>
</Triggers>
"@

$folder.RegisterTask($taskName, $taskXml, 6, $null, $null, 5) | Out-Null

# Verify
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
