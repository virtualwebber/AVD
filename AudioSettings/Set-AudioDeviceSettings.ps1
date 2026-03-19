<#
.SYNOPSIS
    Disables audio enhancements, sets sample rates, and optionally controls exclusive mode
    for all audio devices.

.DESCRIPTION
    Intended to be called by a scheduled task or WMI event subscription on device arrival,
    ensuring USB headsets are always configured correctly without requiring manual intervention.

    Can also be run manually, elevated, for testing or one-off remediation.

    WHAT THIS SCRIPT DOES
       For every audio device found in the Render (output) and Capture (input) hives:
         1. Disables audio enhancements (always)
         2. Sets exclusive mode checkboxes (if parameters provided)
         3. Patches the sample rate in all format properties (if parameters provided)

    HOW IT IS TRIGGERED
       A scheduled task (registered by Register-AudioDeviceTask.ps1) watches for
       Event ID 112 in the DeviceSetupManager/Admin log. This event fires when a
       device container has been fully serviced (drivers loaded, registry populated).
       The script can also be run manually for testing.

    REGISTRY STRUCTURE
       Audio devices live under:
         HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{GUID}
         HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{GUID}
       Each device GUID key contains:
         - DeviceState (DWORD) -- 1 = Active, 4 = NotPresent, 8 = Disabled
         - Properties\         -- device settings (format, name, exclusive mode, etc.)
         - FxProperties\       -- audio processing object (APO) effect chain settings

    ENHANCEMENTS OFF
       Windows stores the enhancement toggle state in FxProperties, not Properties.
       The key {1da5d803},5 = dword:1 is what the Sound Settings GUI writes when the user
       unchecks "Enable audio enhancements". Confirmed by live registry capture before/after
       toggling both the Voice Clarity (capture) and Headset Earphone (render) toggles.
       Absent = enhancements ON. Present with value 1 = enhancements OFF.

    EXCLUSIVE MODE
       Two separate DWORD values under Properties control the Advanced-tab checkboxes:
         {b3f8fa53},3 = "Allow applications to take exclusive control of this device"
         {b3f8fa53},4 = "Give exclusive mode applications priority"
       Each is independently 1 = ticked (ON), 0 = unticked (OFF).
       These are NOT a bitmask -- each checkbox is controlled by its own registry value.
       Applies to both render (output) and capture (input) devices.

    SAMPLE RATE
       Windows stores audio format data as WAVEFORMATEXTENSIBLE binary blobs in multiple
       registry properties. All four format properties must agree on the sample rate,
       otherwise the capture stream can fail silently (mic shows volume but records silence).
       The script reads each blob, patches only the sample rate bytes, and preserves
       everything else (channels, bit depth, channel mask, sub-format GUID).

    MUTEX
       WMI event subscriptions and scheduled tasks can fire multiple times per device
       insertion. A named global mutex ensures only one instance of this script runs at
       a time; subsequent firings exit immediately without doing any work.

.PARAMETER OutputRate
    Target sample rate (Hz) for render (output) devices.
    Valid values: 8000, 11025, 16000, 22050, 32000, 44100, 48000.
    Default $null = do not change the sample rate.
    8000 Hz (Telephone quality) minimises bandwidth over AVD USB redirection.

.PARAMETER InputRate
    Target sample rate (Hz) for capture (input) devices.
    Valid values: 8000, 11025, 16000, 22050, 32000, 44100, 48000.
    Default $null = do not change the sample rate.

.PARAMETER AllowExclusive
    Controls "Allow applications to take exclusive control of this device".
    $true = ticked (ON), $false = unticked (OFF), $null = leave untouched.

.PARAMETER ExclusivePriority
    Controls "Give exclusive mode applications priority".
    $true = ticked (ON), $false = unticked (OFF), $null = leave untouched.

.EXAMPLE
    # Disable enhancements only (no format or exclusive mode changes)
    .\Set-AudioDeviceSettings.ps1

.EXAMPLE
    # Disable enhancements and set both output/input to 8000 Hz
    .\Set-AudioDeviceSettings.ps1 -OutputRate 8000 -InputRate 8000

.EXAMPLE
    # Disable enhancements, set 48000 Hz, allow exclusive but disable priority
    .\Set-AudioDeviceSettings.ps1 -OutputRate 48000 -InputRate 48000 -AllowExclusive 1 -ExclusivePriority 0

.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Log:     C:\_source\logs\AudioDeviceSettings_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# PARAMETERS
# ============================================================

param(
    # Target sample rate (Hz) for render (output) devices. The shared-mode audio engine
    # handles resampling internally, so any standard rate works regardless of the device's
    # native hardware capabilities. Default $null = leave untouched (do not change format).
    [ValidateSet($null, 8000, 11025, 16000, 22050, 32000, 44100, 48000)]
    [Nullable[int]]$OutputRate = $null,

    # Target sample rate (Hz) for capture (input) devices.
    # Default $null = leave untouched (do not change format).
    [ValidateSet($null, 8000, 11025, 16000, 22050, 32000, 44100, 48000)]
    [Nullable[int]]$InputRate = $null,

    # Controls "Allow applications to take exclusive control of this device" on the
    # Advanced tab in Sound Settings. Registry value {b3f8fa53},3 (DWORD):
    #   1 ($true)  = ticked (applications CAN take exclusive control)
    #   0 ($false) = unticked (applications CANNOT take exclusive control)
    #   $null      = leave untouched (do not modify this setting)
    [Nullable[bool]]$AllowExclusive = 1,

    # Controls "Give exclusive mode applications priority" on the Advanced tab
    # in Sound Settings. Registry value {b3f8fa53},4 (DWORD):
    #   1 ($true)  = ticked (exclusive apps get priority over shared-mode apps)
    #   0 ($false) = unticked (no priority given)
    #   $null      = leave untouched (do not modify this setting)
    [Nullable[bool]]$ExclusivePriority = 0
)

# ============================================================
# CONFIGURATION
# ============================================================

# Registry paths to the MMDevices audio hives.
# Render = output devices (speakers, headset earphone)
# Capture = input devices (microphones, headset microphone)
$renderPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
$capturePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"

# Local log directory -- logs are always written here.
$logDir      = "C:\_source\logs"

# Optional UNC file share for centralised logging across multiple hosts.
# Set to $null or "" to disable file share logging.
$logShare    = "\\fileserver.domain.com\logs\audio"

# Log file name includes computername so logs from multiple hosts don't overwrite each other.
# Daily rolling -- one file per day per host.
$logName     = "${env:COMPUTERNAME}_AudioDeviceSettings_$(Get-Date -Format 'yyyyMMdd').log"
$logFile     = Join-Path $logDir $logName

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

    # Ensure the local log directory exists (creates on first run)
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Format: [timestamp] [LEVEL] message
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"

    # Always write to local log file
    Add-Content -Path $logFile -Value $entry

    # Optionally write to file share -- silently skip if unreachable
    if ($logShare) {
        try {
            if (-not (Test-Path $logShare)) {
                New-Item -ItemType Directory -Path $logShare -Force -ErrorAction Stop | Out-Null
            }
            Add-Content -Path (Join-Path $logShare $logName) -Value $entry -ErrorAction Stop
        }
        catch {
            # Silently skip if the share is unreachable (network issues, permissions, etc.)
        }
    }

    # Also output to console for interactive/manual use
    Write-Host $entry
}

function Set-RegistryValue {
    <#
    .SYNOPSIS
        Writes a single registry value with retry logic and optional key creation.
    .DESCRIPTION
        Handles three scenarios:
          1. Key/value exists -- update it if different, skip if already correct
          2. Key/value doesn't exist + Create flag -- create it (e.g. enhancements toggle)
          3. Key/value doesn't exist + no Create flag -- skip (don't force values the driver never set)

        Retries up to 5 times with 500ms gaps because AudioEndpointBuilder may briefly
        re-lock keys immediately after device arrival events.
    #>
    param(
        [string]$Path,           # Full registry path to the key
        [string]$Name,           # Registry value name (e.g. "{b3f8fa53},3")
        [object]$Value,          # Value to write
        [string]$Type,           # Registry type: DWord, Binary, String, etc.
        [string]$Description,    # Human-readable description for logging
        [switch]$Create          # When set, creates the key/value if they don't exist.
                                 # Use for settings like enhancements toggle where absent = ON.
                                 # Omit for format values where absent = driver never set it.
    )

    # Check if the registry key exists
    if (-not (Test-Path $Path)) {
        if ($Create) {
            # Create the key -- needed for settings like FxProperties\{1da5d803},5
            # which must be explicitly created to disable enhancements.
            New-Item -Path $Path -Force | Out-Null
        }
        else {
            # Skip if the registry key doesn't exist -- not all audio devices expose
            # every property key (e.g. some lack FxProperties entirely).
            Write-Log "  SKIP $Description (key does not exist)" -Level "WARN"
            return
        }
    }

    # Check if the named value already exists
    $current = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $current) {
        if (-not $Create) {
            # Skip if the named value doesn't exist -- the device driver never created it,
            # so forcing it could cause unexpected behaviour.
            Write-Log "  SKIP $Description (value does not exist)" -Level "WARN"
            return
        }
        # If -Create is set, fall through to write the value for the first time.
    }
    elseif ($current.$Name -eq $Value) {
        # Already set correctly -- nothing to do.
        Write-Log "  --  $Description (already set)"
        return
    }

    # Retry loop: AudioEndpointBuilder may briefly re-lock keys after device arrival
    # events, causing "access denied" errors. A short retry loop handles this.
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
        all agree on the sample rate -- see $formatProperties in Set-DeviceSettings.

        WAVEFORMATEXTENSIBLE layout (offsets relative to start of blob, including 8-byte header):
          Offset  0-7  : DEVPROPERTY header (type + flags)
          Offset  8-9  : wFormatTag        (0xFFFE = WAVE_FORMAT_EXTENSIBLE)
          Offset 10-11 : nChannels          -- 1 = mono, 2 = stereo (PRESERVED)
          Offset 12-15 : nSamplesPerSec     -- sample rate in Hz    (PATCHED)
          Offset 16-19 : nAvgBytesPerSec    -- nSamplesPerSec * nBlockAlign (RECALCULATED)
          Offset 20-21 : nBlockAlign         -- nChannels * wBitsPerSample / 8 (PRESERVED)
          Offset 22-23 : wBitsPerSample      -- typically 16 (PRESERVED)
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
        # Value doesn't exist -- device driver never wrote a format. Nothing to patch.
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
    # Preserved as-is -- it depends on channels and bit depth, not sample rate.
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
    <#
    .SYNOPSIS
        Configures all audio devices in a single hive (Render or Capture).
    .DESCRIPTION
        Iterates every device GUID under the given hive path and applies:
          1. Disable audio enhancements (always)
          2. Set exclusive mode checkboxes (if parameters are not $null)
          3. Patch sample rate in all format properties (if TargetRate is not $null)
    #>
    param(
        [string]$HivePath,                    # Full registry path to the hive (Render or Capture)
        [string]$HiveLabel,                   # "Render" or "Capture" -- used in log messages
        [Nullable[int]]$TargetRate,           # Target sample rate in Hz ($null = skip format changes)
        [Nullable[bool]]$AllowExclusive,      # {b3f8fa53},3 -- Allow exclusive control ($null = skip)
        [Nullable[bool]]$ExclusivePriority    # {b3f8fa53},4 -- Give exclusive mode priority ($null = skip)
    )

    Write-Log "Processing $HiveLabel hive: $HivePath"

    # Get all device GUID subkeys under this hive
    $devices = Get-ChildItem $HivePath -ErrorAction SilentlyContinue

    if (-not $devices) {
        Write-Log "No devices found in $HiveLabel hive" -Level "WARN"
        return
    }

    foreach ($device in $devices) {
        # Each device GUID key has two important subkeys:
        #   Properties\   -- device settings (format blobs, friendly name, exclusive mode)
        #   FxProperties\ -- audio processing object (APO) settings (enhancements toggle)
        $propsPath   = Join-Path $device.PSPath "Properties"
        $fxPropsPath = Join-Path $device.PSPath "FxProperties"

        # Read friendly name from Properties for logging.
        # {a45c254e-df1c-4efd-8020-67d146a850e0},2 = PKEY_Device_FriendlyName
        # This is what appears in Sound Settings (e.g. "Headset Earphone (Jabra SPEAK 510 USB)")
        $friendlyName = (Get-ItemProperty -Path $propsPath -ErrorAction SilentlyContinue) |
                        Select-Object -ExpandProperty "{a45c254e-df1c-4efd-8020-67d146a850e0},2" -ErrorAction SilentlyContinue
        $label = if ($friendlyName) { $friendlyName } else { $device.PSChildName }

        Write-Log "Configuring device: $label"

        # ----------------------------------------------------------
        # 1. ENHANCEMENTS OFF (always applied)
        # ----------------------------------------------------------
        # FxProperties\{1da5d803},5 is the master enhancement enable/disable toggle.
        # This is what Sound Settings writes when the user unchecks "Enable audio enhancements".
        #   Value 1 = enhancements DISABLED
        #   Absent  = enhancements ENABLED (Windows default on device arrival)
        # Confirmed from live registry capture on Plantronics Blackwire 325.1 -- same key
        # controls both render (Headset Earphone) and capture (Headset Microphone).
        # We use -Create because this key may not exist yet (first time the device is seen).
        Set-RegistryValue -Path $fxPropsPath `
            -Name "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5" `
            -Value 1 -Type DWord `
            -Description "Disable audio enhancements [{1da5d803},5 in FxProperties]" `
            -Create

        # ----------------------------------------------------------
        # 2. EXCLUSIVE MODE (only if parameters are not $null)
        # ----------------------------------------------------------
        # Two separate DWORD values under Properties control the Advanced-tab checkboxes.
        # These are NOT a bitmask -- each is an independent 0/1 value.
        #   {b3f8fa53},3 = "Allow applications to take exclusive control of this device"
        #   {b3f8fa53},4 = "Give exclusive mode applications priority"
        # We use -Create because these values should always be set when specified.
        if ($null -ne $AllowExclusive) {
            Set-RegistryValue -Path $propsPath `
                -Name "{b3f8fa53-0004-438e-9003-51a46e139bfc},3" `
                -Value ([int][bool]$AllowExclusive) -Type DWord `
                -Description "Allow exclusive control -> $(if ($AllowExclusive) {'ON'} else {'OFF'}) [{b3f8fa53},3]" `
                -Create
        }
        if ($null -ne $ExclusivePriority) {
            Set-RegistryValue -Path $propsPath `
                -Name "{b3f8fa53-0004-438e-9003-51a46e139bfc},4" `
                -Value ([int][bool]$ExclusivePriority) -Type DWord `
                -Description "Exclusive mode priority -> $(if ($ExclusivePriority) {'ON'} else {'OFF'}) [{b3f8fa53},4]" `
                -Create
        }

        # ----------------------------------------------------------
        # 3. AUDIO FORMAT (only if TargetRate is not $null)
        # ----------------------------------------------------------
        if ($null -ne $TargetRate) {
            Write-Log "  Target sample rate: $TargetRate Hz"

            # Windows stores the audio format in multiple registry properties that must ALL
            # agree on the sample rate. Patching only one causes the capture stream to fail
            # silently (mic shows volume in the mixer but records silence).
            #
            # Each property contains a WAVEFORMATEXTENSIBLE binary blob. We read each one
            # individually and only patch the sample rate bytes -- channels, bit depth, and
            # all other fields are preserved. See Get-PatchedAudioFormat for blob layout.
            #
            # The four format properties:
            $formatProperties = @(
                @{ Name = "{f19f064d-082c-4e27-bc73-6882a1bb8e4c},0"; Label = "DeviceFormat" }   # PKEY_AudioEngine_DeviceFormat -- primary format used by the audio engine
                @{ Name = "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04},0"; Label = "OEMFormat" }      # PKEY_AudioEngine_OEMFormat -- OEM/driver default format
                @{ Name = "{3d6e1656-72d5-4661-8d01-10f69e406c60},3"; Label = "Format3" }        # Additional format property written by AudioEndpointBuilder
                @{ Name = "{624f56de-fd24-473e-814a-de40aacaed16},3"; Label = "Format4" }        # Additional format property written by AudioEndpointBuilder
            )

            foreach ($fmt in $formatProperties) {
                # Read the existing blob and patch the sample rate
                $patchedFormat = Get-PatchedAudioFormat -PropsPath $propsPath `
                    -FormatValueName $fmt.Name -TargetSampleRate $TargetRate

                if ($null -ne $patchedFormat) {
                    # Write the patched blob back to the registry
                    Set-RegistryValue -Path $propsPath `
                        -Name $fmt.Name `
                        -Value $patchedFormat -Type Binary `
                        -Description "Set $($fmt.Label) to $TargetRate Hz [$($fmt.Name)]"
                }
                else {
                    # $null means either already correct or value doesn't exist
                    Write-Log "  --  $($fmt.Label) unchanged (already correct or not present)"
                }
            }
        }
    }
}

# ============================================================
# MAIN
# ============================================================

# Mutex prevents concurrent execution if WMI or the scheduled task fires multiple
# events per device insertion. WaitOne(0) = non-blocking acquire. If another instance
# holds the mutex, exit immediately rather than queueing up duplicate work.
$mutexName = "Global\AudioDeviceSettings"
$mutex     = New-Object System.Threading.Mutex($false, $mutexName)

if (-not $mutex.WaitOne(0)) {
    Write-Host "Another instance is already running -- exiting."
    exit 0
}

try {
    # Early exit -- if no audio devices exist in either hive, there is nothing to configure.
    # This avoids unnecessary logging and event log queries on machines with no audio hardware.
    $renderDevices  = Get-ChildItem $renderPath  -ErrorAction SilentlyContinue
    $captureDevices = Get-ChildItem $capturePath -ErrorAction SilentlyContinue
    if (-not $renderDevices -and -not $captureDevices) {
        Write-Host "No audio devices found in Render or Capture hives -- exiting."
        exit 0
    }

    Write-Log "========================================"
    Write-Log "Audio device settings script started"
    Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "========================================"

    # ----------------------------------------------------------
    # Log which device triggered this run (if called by scheduled task)
    # ----------------------------------------------------------
    # Query the most recent Event ID 112 from DeviceSetupManager to identify the device
    # that was just plugged in. Only looks at events from the last 30 seconds to ensure
    # we log the device that actually triggered this execution, not a stale event.
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
        # No recent Event ID 112 found -- script was likely run manually, not by scheduled task
        Write-Log "No recent device arrival event found (manual run?)"
    }

    # ----------------------------------------------------------
    # Log current parameter values
    # ----------------------------------------------------------
    Write-Log "Output sample rate: $(if ($null -ne $OutputRate) { "$OutputRate Hz" } else { 'unchanged' })"
    Write-Log "Input sample rate:  $(if ($null -ne $InputRate) { "$InputRate Hz" } else { 'unchanged' })"
    Write-Log "Allow exclusive:    $(if ($null -ne $AllowExclusive) { if ($AllowExclusive) {'ON'} else {'OFF'} } else { 'unchanged' })"
    Write-Log "Exclusive priority: $(if ($null -ne $ExclusivePriority) { if ($ExclusivePriority) {'ON'} else {'OFF'} } else { 'unchanged' })"

    # ----------------------------------------------------------
    # Apply settings to all devices in both hives
    # ----------------------------------------------------------
    # Render = output devices (speakers, headset earphone)
    Set-DeviceSettings -HivePath $renderPath  -HiveLabel "Render"  -TargetRate $OutputRate -AllowExclusive $AllowExclusive -ExclusivePriority $ExclusivePriority

    # Capture = input devices (microphones, headset microphone)
    Set-DeviceSettings -HivePath $capturePath -HiveLabel "Capture" -TargetRate $InputRate  -AllowExclusive $AllowExclusive -ExclusivePriority $ExclusivePriority

    Write-Log "========================================"
    Write-Log "Audio device settings script completed"
    Write-Log "========================================"
}
finally {
    # Always release the mutex, even if an error occurred, so the next run isn't blocked.
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
