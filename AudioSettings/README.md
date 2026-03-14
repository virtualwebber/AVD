# AVD Audio Settings

PowerShell scripts for automatically configuring audio device settings on Azure Virtual Desktop (AVD) session hosts. Designed for environments using USB passthrough, ensuring USB headsets (e.g. Plantronics Blackwire 325.1) are always configured correctly without manual intervention.

## What it does

When a USB headset is connected via AVD USB passthrough, the audio enhancements are automatically disabled on all audio endpoints (both render and capture). This prevents audio quality issues caused by Windows audio enhancements being enabled by default on device arrival.

## Scripts

### Core Scripts

| Script                           | Purpose                                                                   | Run as |
| -------------------------------- | ------------------------------------------------------------------------- | ------ |
| `Set-AudioDeviceSettings.ps1`    | Disables audio enhancements on all MMDevice audio endpoints               | SYSTEM |
| `Set-AudioDevicePermissions.ps1` | Takes ownership of MMDevices registry keys and grants SYSTEM full control | SYSTEM |

### Trigger Scripts (choose one)

| Script                         | Purpose                                                                                | Run as           |
| ------------------------------ | -------------------------------------------------------------------------------------- | ---------------- |
| `Register-AudioDeviceTask.ps1` | **Recommended.** Registers a scheduled task triggered by Event ID 112 (device arrival) | Admin (elevated) |
| `CreateWMI_Audio_Sub.ps1`      | Registers a WMI permanent event subscription for device arrival                        | Admin (elevated) |

### Utility Scripts

| Script                           | Purpose                                      | Run as           |
| -------------------------------- | -------------------------------------------- | ---------------- |
| `CreateWMI_Audio_Sub_Delete.ps1` | Removes the WMI permanent event subscription | Admin (elevated) |

---

## Script Details

### Set-AudioDeviceSettings.ps1

The main remediation script. Iterates through all devices under the MMDevices Audio registry hive (both Render and Capture) and disables audio enhancements by setting `FxProperties\{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5` to `1` (DWORD).

- **Registry paths:** `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render` and `...\Capture`
- **Mutex:** Uses `Global\AudioDeviceSettings` to prevent concurrent execution if multiple device events fire simultaneously
- **Retry logic:** Retries registry writes up to 5 times with 500ms delay (AudioEndpointBuilder may briefly lock keys after device arrival)

### Set-AudioDevicePermissions.ps1

Takes ownership of the MMDevices Audio registry keys and grants SYSTEM full control. Required because these keys are owned by `NT SERVICE\AudioEndpointBuilder` by default, and even SYSTEM cannot write to them without first taking ownership.

- **P/Invoke:** Uses Win32 `AdjustTokenPrivileges` API to enable `SeTakeOwnershipPrivilege` and `SeRestorePrivilege`
- **Inheritance:** Sets `ContainerInherit + ObjectInherit` so all child device subkeys are covered
- **Run once:** Only needs to run once — ownership persists until AudioEndpointBuilder recreates the keys
- **Mutex:** Uses `Global\AudioDevicePermissions` to prevent concurrent execution

### Register-AudioDeviceTask.ps1 (Recommended Trigger)

Creates a Windows Scheduled Task that triggers `Set-AudioDeviceSettings.ps1` whenever a device is installed. Uses Event ID 112 from `Microsoft-Windows-DeviceSetupManager/Admin`, which fires when a device container has been fully serviced and its properties are written to the registry.

- **Trigger:** Event ID 112 — device container serviced (covers USB plug-in, USB passthrough connection)
- **Runs as:** SYSTEM with highest privileges
- **Multiple instances:** Queued (handles rapid device connections)
- **Idempotent:** Removes any existing task before creating a new one

**Why this over WMI?** Many enterprise environments block `CommandLineEventConsumer` via Attack Surface Reduction (ASR) rules. The scheduled task approach avoids this restriction entirely.

### CreateWMI_Audio_Sub.ps1 (Alternative Trigger)

Registers a WMI permanent event subscription that monitors for new `Win32_PnPEntity` objects with `PNPClass = 'AudioEndpoint'`. When detected, it launches `Set-AudioDeviceSettings.ps1` via a `CommandLineEventConsumer`.

- **WQL Query:** `SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity' AND TargetInstance.PNPClass = 'AudioEndpoint'`
- **Polling interval:** 2 seconds
- **Clean slate:** Removes all existing subscription components before creating new ones
- **Note:** Will not work in environments where ASR rule `d1e49aac-8f56-4280-b9ba-993a6d77406c` is enabled

### CreateWMI_Audio_Sub_Delete.ps1

Utility script to remove the WMI permanent event subscription created by `CreateWMI_Audio_Sub.ps1`. Removes all three components (binding, filter, consumer) and verifies removal.

---

## Deployment

### Prerequisites

- Windows 11 AVD session host
- USB passthrough enabled (not RDP audio redirection)
- PowerShell execution policy allows script execution
- Administrator/SYSTEM access

### Installation Steps

1. Copy scripts to `C:\_source\` on the session host

2. Run permissions script (elevated):

   ```powershell
   Set-AudioDevicePermissions.ps1
   ```

3. Register the trigger (elevated) — choose one:

   ```powershell
   # Option A: Scheduled task (recommended for enterprise/ASR environments)
   Register-AudioDeviceTask.ps1
   ```

   ```powershell
   # Option B: WMI subscription (for environments without ASR restrictions)
   CreateWMI_Audio_Sub.ps1
   ```

4. Test by plugging in a USB headset and checking `C:\_source\logs\`

### Logging

All scripts write to `C:\_source\logs\` with daily rolling log files. Log filenames are prefixed with the computer name for multi-host identification.

| Script                           | Log file pattern                                          |
| -------------------------------- | --------------------------------------------------------- |
| `Set-AudioDeviceSettings.ps1`    | `<COMPUTERNAME>_AudioDeviceSettings_yyyyMMdd.log`         |
| `Set-AudioDevicePermissions.ps1` | `<COMPUTERNAME>_AudioRegistryPermissions_yyyyMMdd.log`    |
| `Register-AudioDeviceTask.ps1`   | `<COMPUTERNAME>_RegisterAudioTask_yyyyMMdd.log`           |
| `CreateWMI_Audio_Sub.ps1`        | `<COMPUTERNAME>_AudioWMISubscription_yyyyMMdd.log`        |

Logs are also written to a configurable UNC file share (`$logShare` variable in each script) for centralised collection. If the share is unreachable, logging continues locally without error.
