<#
.SYNOPSIS
    Disables audio enhancements for all MMDevice audio endpoints.

.DESCRIPTION
    Intended to be called by a scheduled task or WMI event subscription on device arrival,
    ensuring USB headsets are always configured correctly without requiring manual intervention.

    Can also be run manually, elevated, for testing or one-off remediation.

    ENHANCEMENTS OFF
       Windows stores the enhancement toggle state in FxProperties, not Properties.
       The key {1da5d803},5 = dword:1 is what the Sound Settings GUI writes when the user
       unchecks "Enable audio enhancements". Confirmed by live registry capture before/after
       toggling both the Voice Clarity (capture) and Headset Earphone (render) toggles.
       Absent = enhancements ON. Present with value 1 = enhancements OFF.

    MUTEX
       WMI event subscriptions can fire multiple times per device insertion. A named global
       mutex ensures only one instance of this script runs at a time; subsequent firings exit
       immediately without doing any work.

.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Log:     C:\_source\logs\AudioDeviceSettings_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

$renderPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
$capturePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"
$logDir      = "C:\_source\logs"
$logShare    = "\\cukavdukwprod01.file.core.windows.net\profiles\logs"
$logName     = "${env:COMPUTERNAME}_AudioDeviceSettings_$(Get-Date -Format 'yyyyMMdd').log"
$logFile     = Join-Path $logDir $logName

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

function Set-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        [object]$Value,
        [string]$Type,
        [string]$Description
    )

    # Retries up to 5 times with a 500ms gap. AudioEndpointBuilder may briefly re-lock
    # keys after device arrival events, so a short retry loop avoids transient failures.
    $maxAttempts = 5
    $attempt     = 0

    while ($attempt -lt $maxAttempts) {
        try {
            if (-not (Test-Path $Path)) {
                New-Item -Path $Path -Force | Out-Null
            }
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
            Write-Log "  OK  $Description"
            return
        }
        catch {
            $attempt++
            if ($attempt -ge $maxAttempts) {
                Write-Log "  FAIL $Description -- $($_.Exception.Message)" -Level "WARN"
            }
            else {
                Start-Sleep -Milliseconds 500
            }
        }
    }
}

function Set-DeviceSettings {
    param([string]$HivePath, [string]$HiveLabel)

    Write-Log "Processing $HiveLabel hive: $HivePath"

    $devices = Get-ChildItem $HivePath -ErrorAction SilentlyContinue

    if (-not $devices) {
        Write-Log "No devices found in $HiveLabel hive" -Level "WARN"
        return
    }

    foreach ($device in $devices) {
        # FxProperties — APO (Audio Processing Object) effect chain settings
        $propsPath   = Join-Path $device.PSPath "Properties"
        $fxPropsPath = Join-Path $device.PSPath "FxProperties"

        # Read friendly name from Properties for logging. Key {a45c254e},2 = PKEY_Device_FriendlyName.
        $friendlyName = (Get-ItemProperty -Path $propsPath -ErrorAction SilentlyContinue) |
                        Select-Object -ExpandProperty "{a45c254e-df1c-4efd-8020-67d146a850e0},2" -ErrorAction SilentlyContinue
        $label = if ($friendlyName) { $friendlyName } else { $device.PSChildName }

        Write-Log "Configuring device: $label"

        # --- ENHANCEMENTS OFF ---
        # FxProperties\{1da5d803},5 is the master enhancement enable/disable toggle.
        # This is what Sound Settings writes when the user unchecks "Enable audio enhancements".
        # Value 1 = disabled. Absent = enabled (Windows default on device arrival).
        # Confirmed from live registry capture on Plantronics Blackwire 325.1 — same key
        # controls both render (Headset Earphone) and capture (Headset Microphone).
        Set-RegistryValue -Path $fxPropsPath `
            -Name "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5" `
            -Value 1 -Type DWord `
            -Description "Disable audio enhancements [{1da5d803},5 in FxProperties]"
    }
}

# ============================================================
# MAIN
# ============================================================

# Mutex prevents concurrent execution if WMI fires multiple events per device insertion.
# WaitOne(0) = non-blocking acquire. If another instance holds the mutex, exit immediately.
$mutexName = "Global\AudioDeviceSettings"
$mutex     = New-Object System.Threading.Mutex($false, $mutexName)

if (-not $mutex.WaitOne(0)) {
    exit 0
}

try {
    # Early exit — if no audio devices exist in either hive, exit silently.
    # This prevents unnecessary logging when Event ID 112 fires for non-audio
    # devices (monitors, printers, USB drives, etc.).
    $renderDevices  = Get-ChildItem $renderPath  -ErrorAction SilentlyContinue
    $captureDevices = Get-ChildItem $capturePath -ErrorAction SilentlyContinue
    if (-not $renderDevices -and -not $captureDevices) {
        exit 0
    }

    Write-Log "========================================"
    Write-Log "Audio device settings script started"
    Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "========================================"

    # Query the most recent Event ID 112 to log which device triggered this run.
    # Only looks at events from the last 30 seconds to ensure we log the device
    # that actually triggered this execution, not a stale event from hours ago.
    try {
        $triggerEvent = Get-WinEvent -FilterHashtable @{
            LogName   = "Microsoft-Windows-DeviceSetupManager/Admin"
            Id        = 112
            StartTime = (Get-Date).AddSeconds(-30)
        } -MaxEvents 1 -ErrorAction Stop
        if ($triggerEvent) {
            $eventXml   = [xml]$triggerEvent.ToXml()
            $deviceName = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq "Prop_DeviceName" }).'#text'
            Write-Log "Triggered by device: $deviceName"
        }
    }
    catch {
        # No recent Event ID 112 found — script was likely run manually
        Write-Log "No recent device arrival event found (manual run?)"
    }

    Set-DeviceSettings -HivePath $renderPath  -HiveLabel "Render"
    Set-DeviceSettings -HivePath $capturePath -HiveLabel "Capture"

    Write-Log "========================================"
    Write-Log "Audio device settings script completed"
    Write-Log "========================================"
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}