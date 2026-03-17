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
# PARAMETERS
# ============================================================

param(
    [ValidateSet("DVD", "Telephone", "TapePlayer")]
    [string]$OutputQuality = "Telephone",

    [ValidateSet("DVD", "Telephone", "TapePlayer")]
    [string]$InputQuality = "TapePlayer"
)

# ============================================================
# CONFIGURATION
# ============================================================

$renderPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
$capturePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"

# Target sample rates for each quality preset.
# Instead of hardcoding full WAVEFORMATEXTENSIBLE byte arrays (which assume 2 channels),
# we read the device's existing format and only patch the sample rate. This preserves the
# native channel count (mono mics, stereo headphones, etc.) and all other device-specific
# fields like bit depth, format tag, channel mask, and sub-format GUID.
$sampleRates = @{
    "DVD"        = 48000   # 48 kHz — DVD quality
    "Telephone"  = 8000    #  8 kHz — Telephone quality
    "TapePlayer" = 16000   # 16 kHz — Tape player quality
}
$logDir      = "C:\_source\logs"
$logShare    = "\\fileserver.domain.com\logs\audio"
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
        [string]$Description,
        [switch]$Create    # When set, creates the key/value if they don't exist.
                           # Use for settings like enhancements toggle where absent = ON.
                           # Omit for format values where absent = driver never set it.
    )

    if (-not (Test-Path $Path)) {
        if ($Create) {
            # Create the key — needed for settings like FxProperties\{1da5d803},5
            # which must be explicitly created to disable enhancements.
            New-Item -Path $Path -Force | Out-Null
        }
        else {
            # Skip if the registry key doesn't exist — not all audio devices expose
            # every property key (e.g. some lack FxProperties entirely).
            Write-Log "  SKIP $Description (key does not exist)" -Level "WARN"
            return
        }
    }

    $current = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $current) {
        if (-not $Create) {
            # Skip if the named value doesn't exist — the device driver never created it,
            # so forcing it could cause unexpected behaviour.
            Write-Log "  SKIP $Description (value does not exist)" -Level "WARN"
            return
        }
        # If -Create is set, fall through to write the value for the first time.
    }
    elseif ($current.$Name -eq $Value) {
        # Already set correctly — nothing to do.
        Write-Log "  --  $Description (already set)"
        return
    }

    # Retries up to 5 times with a 500ms gap. AudioEndpointBuilder may briefly re-lock
    # keys after device arrival events, so a short retry loop avoids transient failures.
    $maxAttempts = 5
    $attempt     = 0

    while ($attempt -lt $maxAttempts) {
        try {
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

function Get-PatchedAudioFormat {
    <#
    .SYNOPSIS
        Reads the device's existing WAVEFORMATEXTENSIBLE blob and patches only the sample rate.

    .DESCRIPTION
        Audio format properties store an 8-byte DEVPROPERTY header followed by a
        WAVEFORMATEXTENSIBLE (or WAVEFORMATEX) structure. Multiple properties must
        all agree on the sample rate — see $formatProperties in Set-DeviceSettings.

        WAVEFORMATEXTENSIBLE layout (offsets relative to start of blob, including 8-byte header):
          Offset  0-7  : DEVPROPERTY header (type + flags)
          Offset  8-9  : wFormatTag        (0xFFFE = WAVE_FORMAT_EXTENSIBLE)
          Offset 10-11 : nChannels          — 1 = mono, 2 = stereo (PRESERVED)
          Offset 12-15 : nSamplesPerSec     — sample rate in Hz    (PATCHED)
          Offset 16-19 : nAvgBytesPerSec    — nSamplesPerSec * nBlockAlign (RECALCULATED)
          Offset 20-21 : nBlockAlign         — nChannels * wBitsPerSample / 8 (PRESERVED)
          Offset 22-23 : wBitsPerSample      — typically 16 (PRESERVED)
          Offset 24+   : cbSize, wValidBitsPerSample, dwChannelMask, SubFormat GUID (PRESERVED)

        By reading the existing blob and only changing offsets 12-19, we preserve the device's
        native channel count, bit depth, channel mask, and sub-format GUID. This means mono
        microphones stay mono and stereo headphones stay stereo.

    .PARAMETER PropsPath
        Registry path to the device's Properties key.

    .PARAMETER FormatValueName
        The registry value name containing the format blob (e.g. "{f19f064d-...},0").

    .PARAMETER TargetSampleRate
        The desired sample rate in Hz (e.g. 48000 for DVD, 8000 for Telephone, 16000 for TapePlayer).

    .OUTPUTS
        [byte[]] The patched format blob ready to write back, or $null if the value
        doesn't exist or the blob is too short to be a valid WAVEFORMATEXTENSIBLE.
    #>
    param(
        [string]$PropsPath,
        [string]$FormatValueName,
        [int]$TargetSampleRate
    )

    # Read the existing format blob from the registry.
    $props = Get-ItemProperty -Path $PropsPath -Name $FormatValueName -ErrorAction SilentlyContinue
    if ($null -eq $props) {
        # Value doesn't exist — device driver never wrote a format. Nothing to patch.
        return $null
    }

    # Get the raw byte array. Clone it so we don't modify the registry cache in memory.
    [byte[]]$bytes = $props.$FormatValueName.Clone()

    # Minimum valid size: 8-byte header + 18-byte WAVEFORMATEX = 26 bytes.
    # A full WAVEFORMATEXTENSIBLE is 8 + 40 = 48 bytes, but we only need up to offset 21.
    if ($bytes.Length -lt 26) {
        Write-Log "  SKIP Audio format blob too short ($($bytes.Length) bytes)" -Level "WARN"
        return $null
    }

    # --- Read current values for logging ---

    # nChannels: 2 bytes at offset 10 (little-endian).
    # Mono = 1, Stereo = 2. This is preserved as-is.
    $nChannels = [BitConverter]::ToUInt16($bytes, 10)

    # nSamplesPerSec: 4 bytes at offset 12 (little-endian). Current sample rate.
    $currentRate = [BitConverter]::ToUInt32($bytes, 12)

    # nBlockAlign: 2 bytes at offset 20 (little-endian).
    # This is nChannels * wBitsPerSample / 8 (e.g. 2ch * 16bit / 8 = 4, 1ch * 16bit / 8 = 2).
    # Preserved as-is — it depends on channels and bit depth, not sample rate.
    $nBlockAlign = [BitConverter]::ToUInt16($bytes, 20)

    Write-Log "  Current format: ${nChannels}ch, ${currentRate} Hz, blockAlign=${nBlockAlign}"

    # If the sample rate is already correct, return $null to signal no change needed.
    if ($currentRate -eq $TargetSampleRate) {
        return $null
    }

    # --- Patch the sample rate (offset 12-15) ---
    # Convert the target sample rate to 4 little-endian bytes and write into the blob.
    $rateBytes = [BitConverter]::GetBytes([uint32]$TargetSampleRate)
    $rateBytes.CopyTo($bytes, 12)

    # --- Recalculate nAvgBytesPerSec (offset 16-19) ---
    # nAvgBytesPerSec = nSamplesPerSec * nBlockAlign.
    # This must stay consistent with the new sample rate, otherwise Windows may reject
    # the format or miscalculate buffer sizes.
    $avgBytesPerSec = [uint32]($TargetSampleRate * $nBlockAlign)
    $avgBytes = [BitConverter]::GetBytes($avgBytesPerSec)
    $avgBytes.CopyTo($bytes, 16)

    Write-Log "  Patched format: ${nChannels}ch, ${TargetSampleRate} Hz, avgBytes=${avgBytesPerSec}"

    return ,$bytes
}

function Set-DeviceSettings {
    param([string]$HivePath, [string]$HiveLabel, [string]$Quality)

    Write-Log "Processing $HiveLabel hive: $HivePath"

    $devices = Get-ChildItem $HivePath -ErrorAction SilentlyContinue

    if (-not $devices) {
        Write-Log "No devices found in $HiveLabel hive" -Level "WARN"
        return
    }

    # Look up the target sample rate for the chosen quality preset.
    $targetRate = $sampleRates[$Quality]

    foreach ($device in $devices) {
        # FxProperties — APO (Audio Processing Object) effect chain settings.
        $propsPath   = Join-Path $device.PSPath "Properties"
        $fxPropsPath = Join-Path $device.PSPath "FxProperties"

        # Read friendly name from Properties for logging.
        # Key {a45c254e},2 = PKEY_Device_FriendlyName.
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
            -Description "Disable audio enhancements [{1da5d803},5 in FxProperties]" `
            -Create

        # --- AUDIO FORMAT ---
        # Windows stores the audio format in multiple registry properties that must all
        # agree. Patching only one causes the capture stream to fail silently (mic shows
        # volume but records silence). Each blob is read individually and only the sample
        # rate is patched — channels, bit depth, and all other fields are preserved.
        $formatProperties = @(
            @{ Name = "{f19f064d-082c-4e27-bc73-6882a1bb8e4c},0"; Label = "DeviceFormat" }   # PKEY_AudioEngine_DeviceFormat — primary format used by the audio engine
            @{ Name = "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04},0"; Label = "OEMFormat" }      # PKEY_AudioEngine_OEMFormat — OEM/driver default format
            @{ Name = "{3d6e1656-72d5-4661-8d01-10f69e406c60},3"; Label = "Format3" }        # Additional format property written by AudioEndpointBuilder
            @{ Name = "{624f56de-fd24-473e-814a-de40aacaed16},3"; Label = "Format4" }        # Additional format property written by AudioEndpointBuilder
        )

        foreach ($fmt in $formatProperties) {
            $patchedFormat = Get-PatchedAudioFormat -PropsPath $propsPath `
                -FormatValueName $fmt.Name -TargetSampleRate $targetRate

            if ($null -ne $patchedFormat) {
                Set-RegistryValue -Path $propsPath `
                    -Name $fmt.Name `
                    -Value $patchedFormat -Type Binary `
                    -Description "Set $($fmt.Label) to $Quality quality ($targetRate Hz) [$($fmt.Name)]"
            }
            else {
                Write-Log "  --  $($fmt.Label) unchanged (already correct or not present)"
            }
        }
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

    Set-DeviceSettings -HivePath $renderPath  -HiveLabel "Render"  -Quality $OutputQuality
    Set-DeviceSettings -HivePath $capturePath -HiveLabel "Capture" -Quality $InputQuality

    Write-Log "========================================"
    Write-Log "Audio device settings script completed"
    Write-Log "========================================"
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}