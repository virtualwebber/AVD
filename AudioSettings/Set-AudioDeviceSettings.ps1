<#
.SYNOPSIS
    Disables audio enhancements and sets telephone quality for all MMDevice audio endpoints.

.DESCRIPTION
    Intended to be called by a WMI permanent event subscription on AudioEndpoint device arrival,
    ensuring USB headsets used by Alvaria contact centre software are always configured correctly
    without requiring manual intervention.

    Can also be run manually, elevated, for testing or one-off remediation.

    Two settings are applied to every render and capture device found in the MMDevices hive:

    1. ENHANCEMENTS OFF
       Windows stores the enhancement toggle state in FxProperties, not Properties.
       The key {1da5d803},5 = dword:1 is what the Sound Settings GUI writes when the user
       unchecks "Enable audio enhancements". Confirmed by live registry capture before/after
       toggling both the Voice Clarity (capture) and Headset Earphone (render) toggles.
       Absent = enhancements ON. Present with value 1 = enhancements OFF.

    2. TELEPHONE QUALITY FORMAT
       Windows stores the selected audio format as PROPVARIANT-wrapped WAVEFORMATEXTENSIBLE
       blobs in the device Properties subkey. Four keys are written by the GUI when Telephone
       quality is selected — all confirmed from live registry capture:
         {f19f064d},0  DeviceFormat — the active format used by the audio engine
         {e4870e26},0  OEMFormat    — the format reported as the device's native capability
         {3d6e1656},3  Mirror of OEMFormat (written redundantly by Windows, included for parity)
         {624f56de},3  Mirror of OEMFormat (written redundantly by Windows, included for parity)

    WHY TELEPHONE QUALITY?
       Contact centre workloads require telephone quality audio (8kHz) to match the audio
       profile expected by the Alvaria contact centre software. Setting this on device arrival
       ensures the correct format is always applied without requiring manual configuration.

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
$logFile     = Join-Path $logDir "AudioDeviceSettings_$(Get-Date -Format 'yyyyMMdd').log"

# ============================================================
# TELEPHONE QUALITY FORMAT BLOBS
#
# These are PROPVARIANT-wrapped WAVEFORMATEXTENSIBLE structures,
# exactly as written by the Windows Sound Settings GUI.
#
# PROPVARIANT header (8 bytes):
#   41 00 00 00  = VT_BLOB (variant type blob)
#   01 00 00 00  = blob flags
#
# WAVEFORMATEXTENSIBLE fields follow the header:
#   wFormatTag          WORD    0xFFFE = WAVE_FORMAT_EXTENSIBLE
#   nChannels           WORD    channel count
#   nSamplesPerSec      DWORD   sample rate in Hz
#   nAvgBytesPerSec     DWORD   nSamplesPerSec * nBlockAlign
#   nBlockAlign         WORD    nChannels * wBitsPerSample / 8
#   wBitsPerSample      WORD    bits per sample (container size)
#   cbSize              WORD    22 = size of extension fields
#   wValidBitsPerSample WORD    actual valid bits (may differ from container)
#   dwChannelMask       DWORD   0x00000003 = FRONT_LEFT | FRONT_RIGHT
#   SubFormat GUID      16 bytes
#     PCM:        01 00 00 00 00 00 10 00 80 00 00 AA 00 38 9B 71
#     IEEE_FLOAT: 03 00 00 00 00 00 10 00 80 00 00 AA 00 38 9B 71
# ============================================================

# DeviceFormat — the format the Windows audio engine uses for this device.
# Telephone quality: 8000Hz, 16-bit, stereo, PCM (KSDATAFORMAT_SUBTYPE_PCM).
# nAvgBytesPerSec = 8000 * 4 = 32000. nBlockAlign = 2ch * 2 bytes = 4.
$deviceFormat = [byte[]](
    0x41, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,  # PROPVARIANT header
    0xfe, 0xff,                                        # wFormatTag  = WAVE_FORMAT_EXTENSIBLE
    0x02, 0x00,                                        # nChannels   = 2 (stereo)
    0x40, 0x1f, 0x00, 0x00,                            # nSamplesPerSec = 8000 Hz
    0x00, 0x7d, 0x00, 0x00,                            # nAvgBytesPerSec = 32000
    0x04, 0x00,                                        # nBlockAlign = 4
    0x10, 0x00,                                        # wBitsPerSample = 16
    0x16, 0x00,                                        # cbSize = 22
    0x10, 0x00,                                        # wValidBitsPerSample = 16
    0x03, 0x00, 0x00, 0x00,                            # dwChannelMask = FL | FR
    0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00,   # SubFormat = KSDATAFORMAT_SUBTYPE_PCM
    0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71
)

# OEMFormat — the format reported as the device's native/preferred capability.
# Windows uses a higher bit depth internally for OEM format even at telephone quality.
# Telephone quality: 8000Hz, 32-bit, stereo, IEEE float (KSDATAFORMAT_SUBTYPE_IEEE_FLOAT).
# nAvgBytesPerSec = 8000 * 8 = 64000. nBlockAlign = 2ch * 4 bytes = 8.
$oemFormat = [byte[]](
    0x41, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,  # PROPVARIANT header
    0xfe, 0xff,                                        # wFormatTag  = WAVE_FORMAT_EXTENSIBLE
    0x02, 0x00,                                        # nChannels   = 2 (stereo)
    0x40, 0x1f, 0x00, 0x00,                            # nSamplesPerSec = 8000 Hz
    0x00, 0xfa, 0x00, 0x00,                            # nAvgBytesPerSec = 64000
    0x08, 0x00,                                        # nBlockAlign = 8
    0x20, 0x00,                                        # wBitsPerSample = 32
    0x16, 0x00,                                        # cbSize = 22
    0x20, 0x00,                                        # wValidBitsPerSample = 32
    0x03, 0x00, 0x00, 0x00,                            # dwChannelMask = FL | FR
    0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00,   # SubFormat = KSDATAFORMAT_SUBTYPE_IEEE_FLOAT
    0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71
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
        # Each device has two subkeys we care about:
        #   Properties   — device metadata and format settings
        #   FxProperties — APO (Audio Processing Object) effect chain settings
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

        # --- TELEPHONE QUALITY: DeviceFormat ---
        # {f19f064d},0 = PKEY_AudioEngine_DeviceFormat
        # This is the format the Windows audio engine uses when mixing for this device.
        # Setting to 8kHz ensures the correct format for Alvaria contact centre workloads.
        Set-RegistryValue -Path $propsPath `
            -Name "{f19f064d-082c-4e27-bc73-6882a1bb8e4c},0" `
            -Value $deviceFormat -Type Binary `
            -Description "DeviceFormat = telephone quality 8kHz/16-bit/stereo [{f19f064d},0]"

        # --- TELEPHONE QUALITY: OEMFormat ---
        # {e4870e26},0 = PKEY_AudioEngine_OEMFormat
        # This is the format reported as the device's native/preferred capability.
        # Windows uses 32-bit float internally for OEMFormat even at telephone quality.
        Set-RegistryValue -Path $propsPath `
            -Name "{e4870e26-3cc5-4cd2-ba46-ca0a9a70ed04},0" `
            -Value $oemFormat -Type Binary `
            -Description "OEMFormat = telephone quality 8kHz/32-bit/stereo [{e4870e26},0]"

        # --- TELEPHONE QUALITY: OEMFormat mirror keys ---
        # Windows writes the OEMFormat blob to two additional keys when format is changed
        # via the GUI. Exact purpose unknown but included for full parity with what the
        # GUI writes. Both confirmed from live registry capture.
        Set-RegistryValue -Path $propsPath `
            -Name "{3d6e1656-2e50-4c4c-8d85-d0acae3c6c68},3" `
            -Value $oemFormat -Type Binary `
            -Description "OEMFormat mirror [{3d6e1656},3]"

        Set-RegistryValue -Path $propsPath `
            -Name "{624f56de-fd24-473e-814a-de40aacaed16},3" `
            -Value $oemFormat -Type Binary `
            -Description "OEMFormat mirror [{624f56de},3]"
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
    Write-Log "========================================"
    Write-Log "Audio device settings script started"
    Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "========================================"

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