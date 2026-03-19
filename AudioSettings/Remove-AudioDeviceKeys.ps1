<#
.SYNOPSIS
    Removes stale (disconnected) audio device registry keys from Render and Capture hives.

.DESCRIPTION
    Scans the MMDevices Audio Render and Capture registry hives for devices with
    DeviceState = 4 (NotPresent/Unplugged) and deletes their entire GUID key.

    Devices whose friendly name contains "Remote Audio" are protected and never
    deleted -- these are the RDP/AVD audio redirection endpoints.

    Intended to be triggered by a scheduled task when a USB audio device is
    disconnected (Kernel-PnP Event ID 1010), ensuring stale device entries are
    cleaned up immediately on unplug.

    Can also be run manually, elevated, for testing or one-off remediation.

    RESTART AUDIOENDPOINTBUILDER
       Optionally restarts the AudioEndpointBuilder service after cleanup.
       This clears the in-memory name counter so re-plugging a device gives
       "Headset" instead of "2 - Headset". WARNING: this will momentarily
       interrupt all active audio streams (including Remote Audio).
       Disabled by default -- enable via $RestartAudioService = $true below.

    MUTEX
       A named global mutex ensures only one instance of this script runs at a
       time; subsequent firings exit immediately without doing any work.

.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Run as SYSTEM via scheduled task, or elevated for manual testing.
    Log:     C:\_source\logs\AudioDeviceRemoval_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

$renderPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
$capturePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"

$logDir      = "C:\_source\logs"
$logShare    = "\\fileserver.domain.com\logs\audio"
$logName     = "${env:COMPUTERNAME}_AudioDeviceRemoval_$(Get-Date -Format 'yyyyMMdd').log"
$logFile     = Join-Path $logDir $logName

# Restart AudioEndpointBuilder after cleanup to reset the device name counter.
# Prevents "2 - Headset" naming on re-plug. WARNING: momentarily interrupts
# all active audio streams including Remote Audio.
$RestartAudioService = $false

# Friendly names containing any of these strings are protected and never deleted.
# Remote Audio is the RDP/AVD audio redirection endpoint.
$protectedDevicePatterns = @(
    "Remote Audio"
)

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

function Remove-StaleDevices {
    param(
        [string]$HivePath,
        [string]$HiveLabel
    )

    Write-Log "Scanning $HiveLabel hive for stale devices: $HivePath"

    $devices = Get-ChildItem $HivePath -ErrorAction SilentlyContinue

    if (-not $devices) {
        Write-Log "  No devices found in $HiveLabel hive"
        return
    }

    $removedCount   = 0
    $skippedCount   = 0
    $protectedCount = 0

    foreach ($device in $devices) {
        $deviceGuid = $device.PSChildName
        $propsPath  = Join-Path $device.PSPath "Properties"

        # DeviceState is a DWORD directly under the device GUID key.
        # Values: 1 = Active, 4 = NotPresent/Unplugged, 8 = Disabled
        $deviceState = (Get-ItemProperty -Path $device.PSPath -Name "DeviceState" -ErrorAction SilentlyContinue).DeviceState

        if ($null -eq $deviceState) {
            Write-Log "  SKIP $deviceGuid -- DeviceState not found" -Level "WARN"
            $skippedCount++
            continue
        }

        # Only target devices that are NotPresent (4)
        if ($deviceState -ne 4) {
            $skippedCount++
            continue
        }

        # Read friendly name for logging and protection check
        $friendlyName = (Get-ItemProperty -Path $propsPath -ErrorAction SilentlyContinue) |
                        Select-Object -ExpandProperty "{a45c254e-df1c-4efd-8020-67d146a850e0},2" -ErrorAction SilentlyContinue
        $label = if ($friendlyName) { $friendlyName } else { $deviceGuid }

        # Check if this device matches a protected pattern
        $isProtected = $false
        foreach ($pattern in $protectedDevicePatterns) {
            if ($friendlyName -and $friendlyName -like "*$pattern*") {
                $isProtected = $true
                break
            }
        }

        if ($isProtected) {
            Write-Log "  PROTECTED $label (matches protected pattern -- skipping)"
            $protectedCount++
            continue
        }

        # Delete the entire device GUID key
        Write-Log "  Removing stale device: $label (GUID: $deviceGuid, DeviceState: $deviceState)"

        try {
            Remove-Item -Path $device.PSPath -Recurse -Force -ErrorAction Stop
            Write-Log "  OK  Removed: $label"
            $removedCount++
        }
        catch {
            Write-Log "  FAIL Could not remove $label -- $($_.Exception.Message)" -Level "WARN"
        }
    }

    Write-Log "Cleanup $HiveLabel summary: $removedCount removed, $protectedCount protected, $skippedCount skipped"
}

# ============================================================
# MAIN
# ============================================================

# Mutex prevents concurrent execution if multiple removal events fire in quick succession.
$mutexName = "Global\AudioDeviceRemoval"
$mutex     = New-Object System.Threading.Mutex($false, $mutexName)

if (-not $mutex.WaitOne(0)) {
    Write-Host "Another instance is already running -- exiting."
    exit 0
}

try {
    # Brief delay to let Windows finish updating DeviceState in the registry after unplug.
    Start-Sleep -Seconds 3

    # Early exit -- quick check whether any device in either hive has DeviceState = 4.
    # This avoids the overhead of full logging and scanning when a non-audio device
    # was unplugged (the task fires for all Kernel-PnP removals, not just audio).
    $hasStale = $false
    foreach ($hive in @($renderPath, $capturePath)) {
        foreach ($dev in (Get-ChildItem $hive -ErrorAction SilentlyContinue)) {
            $state = (Get-ItemProperty -Path $dev.PSPath -Name "DeviceState" -ErrorAction SilentlyContinue).DeviceState
            if ($state -eq 4) { $hasStale = $true; break }
        }
        if ($hasStale) { break }
    }
    if (-not $hasStale) {
        exit 0
    }

    Write-Log "========================================"
    Write-Log "Audio device removal cleanup started"
    Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "========================================"

    # Query the most recent Kernel-PnP Event ID 1010 to log which device triggered this run.
    try {
        $triggerEvent = Get-WinEvent -FilterHashtable @{
            LogName   = "Microsoft-Windows-Kernel-PnP/Device Management"
            Id        = 1010
            StartTime = (Get-Date).AddSeconds(-30)
        } -MaxEvents 1 -ErrorAction Stop
        if ($triggerEvent) {
            $eventXml   = [xml]$triggerEvent.ToXml()
            $deviceId   = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq "DeviceInstanceId" }).'#text'
            Write-Log "Triggered by device removal: $deviceId"
        }
    }
    catch {
        Write-Log "No recent device removal event found (manual run?)"
    }

    Remove-StaleDevices -HivePath $renderPath  -HiveLabel "Render"
    Remove-StaleDevices -HivePath $capturePath -HiveLabel "Capture"

    if ($RestartAudioService) {
        Write-Log "Restarting AudioEndpointBuilder service (name counter reset)"
        try {
            Restart-Service -Name AudioEndpointBuilder -Force -ErrorAction Stop
            Write-Log "  OK  AudioEndpointBuilder restarted"
        }
        catch {
            Write-Log "  FAIL Could not restart AudioEndpointBuilder -- $($_.Exception.Message)" -Level "WARN"
        }
    }

    Write-Log "========================================"
    Write-Log "Audio device removal cleanup completed"
    Write-Log "========================================"
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
