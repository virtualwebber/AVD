# AVD RemoteFX USB Device Filtering
## Client-Side USB Device Picker Configuration

---

## Overview

This configuration reduces noise in the RemoteFX USB device picker shown to users when connecting to Azure Virtual Desktop sessions. By default, all USB devices on the local machine are presented to the user — including cameras, fingerprint readers, storage devices, and network adapters. This solution filters out unwanted device classes while ensuring audio devices (headsets, speakerphones etc.) remain available for redirection.

---

## How It Works

There are two client-side registry filters that work together:

| Filter | Registry Key | Purpose |
|---|---|---|
| **Select filter** | `UsbSelectDeviceByInterfaces` | Only show devices that match at least one listed interface GUID |
| **Block filter** | `UsbBlockDeviceBySetupClasses` | Hide devices that match any listed setup class GUID |

Both filters must be **enabled simultaneously** (`= 1`) for either to take effect. If only one is enabled, the filtering stack does not activate correctly.

The select list uses the **generic USB interface GUID** (`{a5dcbf10-6530-11d2-901f-00c04fb951ed}`) to allow all USB devices through, then the **block list does the actual filtering** by removing unwanted device classes.

> ⚠️ Do **not** add `{36FC9E60-C465-11CF-8056-444553540000}` (USB Controllers) to the block list. Although commonly referenced for blocking USB storage, this GUID covers **all composite USB devices** including headsets, and will silently remove all audio devices from the picker.

---

## Prerequisites

### 1. Group Policy — RemoteFX USB Redirection (Required)

The following GPO setting **must** be applied via Group Policy. Setting the registry key directly is not sufficient — the Terminal Services Client Side Extension (CSE) performs additional initialisation (inserting `TsUsbFlt` into USB class UpperFilters and starting the RemoteFX USB stack) that only happens when the policy is applied through GPO.

**Path:**
```
Computer Configuration → Policies → Administrative Templates → Windows Components
→ Remote Desktop Services → Remote Desktop Connection Client
→ RemoteFX USB Device Redirection
→ Allow RDP redirection of other supported RemoteFX USB devices from this computer
```

**Setting:** Enabled  
**Access Rights:** Administrators and Users

> ℹ️ This setting is available in Intune but **does not work correctly** when applied that way. You must use Group Policy.

### 2. Session Host Configuration

Plug and Play redirection must be enabled on the session hosts:

**Path:**
```
Computer Configuration → Policies → Administrative Templates → Windows Components
→ Remote Desktop Services → Remote Desktop Session Host
→ Device and Resource Redirection
→ Do not allow supported Plug and Play device redirection
```

**Setting:** Disabled

### 3. Host Pool RDP Property

The host pool must be configured to allow USB redirection. In the Azure Portal under **Host Pool → RDP Properties → Device Redirection**, set USB device redirection to:

```
Redirect all USB devices that are not already redirected by another high-level redirection
```

Or via PowerShell:
```powershell
Update-AzWvdHostPool -Name "<hostpool>" -ResourceGroupName "<rg>" `
    -CustomRdpProperty "usbdevicestoredirect:s:*"
```

---

## Verification

After the GPO has applied and the machine has rebooted, verify the RemoteFX USB stack is correctly initialised:

```powershell
# TsUsbFlt should be present in UpperFilters
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}" | 
    Select-Object UpperFilters

# Expected output: {TsUsbFlt}

# TsUsbFlt service state (STOPPED is normal - it starts on demand)
sc.exe query TsUsbFlt
```

If `TsUsbFlt` is **not** in UpperFilters, the GPO has not been applied correctly or the machine needs a reboot.

---

## Deployment Script

Run the following script as **Administrator** on the local client machine. This script is idempotent — it clears and rewrites the filter keys cleanly on each run.

```powershell
$clientRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client"
$selectRegPath = "$clientRegPath\UsbSelectDeviceByInterfaces"
$blockRegPath  = "$clientRegPath\UsbBlockDeviceBySetupClasses"

# Create keys if missing
foreach ($path in @($clientRegPath, $selectRegPath, $blockRegPath)) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

# Enable RemoteFX USB redirection
Set-ItemProperty -Path $clientRegPath -Name "fUsbRedirectionEnableMode"         -Value 2 -Type DWord

# Enable both filters (both must be 1 for filtering stack to activate)
Set-ItemProperty -Path $clientRegPath -Name "fEnableUsbBlockDeviceBySetupClass" -Value 1 -Type DWord
Set-ItemProperty -Path $clientRegPath -Name "fEnableUsbSelectDeviceByInterface" -Value 1 -Type DWord
Write-Host "Flags set" -ForegroundColor Green

# Clear existing subkey values before rewriting
Remove-Item -Path $selectRegPath -Force -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path $blockRegPath  -Force -Recurse -ErrorAction SilentlyContinue
New-Item -Path $selectRegPath -Force | Out-Null
New-Item -Path $blockRegPath  -Force | Out-Null

# Select list
# Uses generic USB GUID to allow all devices through - block list handles the filtering
$selectGuids = @{
    "1000" = "{6994AD04-93EF-11D0-A3CC-00A0C9223196}"  # USB audio
    "1001" = "{65E8773D-8F56-11D0-A3B9-00A0C9223196}"  # KS audio streaming
    "1002" = "{65E8773E-8F56-11D0-A3B9-00A0C9223196}"  # KS audio topology
    "1003" = "{4d1e55b2-f16f-11cf-88cb-001111000030}"  # HID (headset call control buttons)
    "1004" = "{a5dcbf10-6530-11d2-901f-00c04fb951ed}"  # Generic USB - allows all through
}

foreach ($entry in $selectGuids.GetEnumerator()) {
    Set-ItemProperty -Path $selectRegPath -Name $entry.Key -Value $entry.Value -Type String
    Write-Host "Select [$($entry.Key)] $($entry.Value)" -ForegroundColor Cyan
}

# Block list
# Note: {36FC9E60} (USB Controllers) intentionally excluded - it covers composite USB
# devices including headsets and will silently block all audio devices if added
$blockGuids = @{
    "1000" = "{53D29EF7-377C-4D14-864B-EB3A85769359}"  # Biometric (fingerprint readers)
    "1001" = "{CA3E7AB9-B4C3-4AE6-8251-579EF933890F}"  # Camera
    "1002" = "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}"  # Image / WIA
    "1003" = "{4D36E967-E325-11CE-BFC1-08002BE10318}"  # Disk drives
    "1004" = "{50DD5230-BA8A-11D1-BF5D-0000F805F530}"  # Smartcard
    "1005" = "{E0CBF06C-CD8B-4647-BB8A-263B43F0F974}"  # Bluetooth
    "1006" = "{4D36E972-E325-11CE-BFC1-08002BE10318}"  # Network adapters
}

foreach ($entry in $blockGuids.GetEnumerator()) {
    Set-ItemProperty -Path $blockRegPath -Name $entry.Key -Value $entry.Value -Type String
    Write-Host "Block  [$($entry.Key)] $($entry.Value)" -ForegroundColor Yellow
}

Write-Host "`nDone - restart Windows App or MSTSC to test" -ForegroundColor Green
```

---

## Registry Reference

All keys sit under:
```
HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client
```

| Value | Type | Data | Purpose |
|---|---|---|---|
| `fUsbRedirectionEnableMode` | DWORD | `2` | Enables RemoteFX USB redirection |
| `fEnableUsbBlockDeviceBySetupClass` | DWORD | `1` | Activates the block filter |
| `fEnableUsbSelectDeviceByInterface` | DWORD | `1` | Activates the select filter |
| `fEnableUsbNoAckIsochWriteToDevice` | DWORD | `80` | Isochronous audio tuning (Windows App default, do not change) |

### Select List — `...\UsbSelectDeviceByInterfaces`

| Index | GUID | Class |
|---|---|---|
| 1000 | `{6994AD04-93EF-11D0-A3CC-00A0C9223196}` | USB Audio |
| 1001 | `{65E8773D-8F56-11D0-A3B9-00A0C9223196}` | KS Audio Streaming |
| 1002 | `{65E8773E-8F56-11D0-A3B9-00A0C9223196}` | KS Audio Topology |
| 1003 | `{4d1e55b2-f16f-11cf-88cb-001111000030}` | HID |
| 1004 | `{a5dcbf10-6530-11d2-901f-00c04fb951ed}` | Generic USB (allows all through) |

### Block List — `...\UsbBlockDeviceBySetupClasses`

> Setup class GUIDs are defined by Microsoft: [System-Defined Device Setup Classes Available to Vendors](https://learn.microsoft.com/en-us/windows-hardware/drivers/install/system-defined-device-setup-classes-available-to-vendors)

| Index | GUID | Class |
|---|---|---|
| 1000 | `{53D29EF7-377C-4D14-864B-EB3A85769359}` | Biometric |
| 1001 | `{CA3E7AB9-B4C3-4AE6-8251-579EF933890F}` | Camera |
| 1002 | `{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}` | Image / WIA |
| 1003 | `{4D36E967-E325-11CE-BFC1-08002BE10318}` | Disk Drives |
| 1004 | `{50DD5230-BA8A-11D1-BF5D-0000F805F530}` | Smartcard |
| 1005 | `{E0CBF06C-CD8B-4647-BB8A-263B43F0F974}` | Bluetooth |
| 1006 | `{4D36E972-E325-11CE-BFC1-08002BE10318}` | Network Adapters |

---

## Testing

1. Plug in the USB audio device **before** opening MSTSC or Windows App
2. Open MSTSC → **Show Options → Local Resources → More**
3. Expand **Other supported RemoteFX USB devices**
4. Audio devices should appear; cameras, fingerprint readers, storage etc. should not

> ℹ️ The **Other supported RemoteFX USB devices** category only appears in MSTSC when `fUsbRedirectionEnableMode = 2` is correctly applied via GPO. If the category is missing, check TsUsbFlt is in UpperFilters (see Verification above).

---

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---|---|---|
| No devices show at all | Both filter flags = 0 | Set both flags to 1 |
| RemoteFX USB category missing in MSTSC | GPO not applied / TsUsbFlt not in UpperFilters | Verify GPO, reboot |
| Audio devices missing with filters on | `{36FC9E60}` in block list | Remove that GUID |
| Select filter enabled but empty | Blocks all devices | Add GUIDs or disable select filter |
| Flags reset after reboot | GPO overwriting registry | Check GPO isn't conflicting |

---

*Tested on: Windows 11, Windows App, MSTSC — Azure Virtual Desktop*
