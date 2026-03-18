#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Configures client-side RemoteFX USB device filtering for Azure Virtual Desktop.

.DESCRIPTION
    This script configures the client-side USB device picker shown to users when
    connecting to AVD sessions via MSTSC or Windows App. It reduces picker noise
    by blocking unwanted device classes (cameras, fingerprint readers, storage etc.)
    while ensuring all audio devices (headsets, speakerphones etc.) remain visible.

    HOW THE FILTERING WORKS
    -----------------------
    Two registry filters work in tandem under:
    HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client

        UsbSelectDeviceByInterfaces   - Only show devices matching these interface GUIDs
        UsbBlockDeviceBySetupClasses  - Hide devices matching these setup class GUIDs

    IMPORTANT: Both filters must be enabled (set to 1) simultaneously. If only one
    is active, the filtering stack does not initialise and no filtering occurs.

    The select list uses the generic USB interface GUID to allow ALL devices through.
    The block list then removes the unwanted classes. This approach is necessary because
    composite USB devices (like headsets) do not reliably expose specific interface GUIDs
    at the parent device level, so a strict allow-list approach will silently exclude them.

    CRITICAL: DO NOT add {36FC9E60-C465-11CF-8056-444553540000} to the block list.
    Although widely referenced for blocking USB storage, this GUID covers the USB
    Controllers setup class which includes ALL composite USB devices. Adding it will
    silently remove all headsets and audio devices from the picker.

    GPO REQUIREMENT — fUsbRedirectionEnableMode
    -------------------------------------------
    The value fUsbRedirectionEnableMode = 2 MUST be applied via Group Policy, not
    set directly in the registry. This is because the Terminal Services Client Side
    Extension (CSE) — which only runs when the policy is applied through GPO — performs
    critical background initialisation that a plain registry write cannot replicate:

        1. Inserts TsUsbFlt into the UpperFilters of the USB controller class
           (HKLM\SYSTEM\CurrentControlSet\Control\Class\{36FC9E60...})
        2. Registers and starts the TsUsbFlt kernel filter driver
        3. Initialises the RemoteFX USB redirection stack in the RDP client

    Without this GPO being applied correctly, TsUsbFlt will not be present in
    UpperFilters, the RemoteFX USB category will not appear in MSTSC or Windows App,
    and the UsbBlockDeviceBySetupClasses / UsbSelectDeviceByInterfaces keys will
    have no effect regardless of their values.

    GPO PATH:
    Computer Configuration > Policies > Administrative Templates > Windows Components
    > Remote Desktop Services > Remote Desktop Connection Client
    > RemoteFX USB Device Redirection
    > Allow RDP redirection of other supported RemoteFX USB devices from this computer
    Setting : Enabled
    Access  : Administrators and Users

    NOTE: This setting is available in Intune but does NOT work correctly when
    deployed that way. Group Policy must be used.

    VERIFY GPO IS APPLIED
    ---------------------
    After GPO has applied and the machine has rebooted, confirm with:

        Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}" |
            Select-Object UpperFilters
        # Expected: {TsUsbFlt}

        sc.exe query TsUsbFlt
        # STATE: STOPPED is normal - driver starts on demand

    If TsUsbFlt is NOT in UpperFilters, the GPO has not applied correctly or the
    machine needs a reboot.

.NOTES
    Author      : Andrew Webber
    Requires    : Administrator rights
    Tested On   : Windows 11, MSTSC, Windows App, Azure Virtual Desktop
    Important   : fUsbRedirectionEnableMode must be set via GPO - see description above
                  Run this script AFTER the GPO has been applied and machine rebooted
#>

# ============================================================
# REGISTRY PATH DEFINITIONS
# ============================================================

$clientRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client"
$selectRegPath = "$clientRegPath\UsbSelectDeviceByInterfaces"
$blockRegPath  = "$clientRegPath\UsbBlockDeviceBySetupClasses"


# ============================================================
# ENSURE REGISTRY KEYS EXIST
# ============================================================

Write-Host "`n[1/4] Checking registry keys..." -ForegroundColor White

foreach ($path in @($clientRegPath, $selectRegPath, $blockRegPath)) {
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
        Write-Host "      Created: $path" -ForegroundColor DarkGray
    } else {
        Write-Host "      Exists:  $path" -ForegroundColor DarkGray
    }
}


# ============================================================
# SET PARENT FLAGS
#
# fUsbRedirectionEnableMode
#   Value 2 enables RemoteFX USB redirection on the client.
#   NOTE: This script sets the registry value as a fallback reference only.
#   This value MUST be applied and initialised via GPO for the RemoteFX USB
#   stack (TsUsbFlt) to be correctly loaded. See script header for full details.
#
# fEnableUsbBlockDeviceBySetupClass
#   Activates the UsbBlockDeviceBySetupClasses subkey filtering.
#   Must be 1 for block list to take effect.
#
# fEnableUsbSelectDeviceByInterface
#   Activates the UsbSelectDeviceByInterfaces subkey filtering.
#   Must be 1 for select list to take effect.
#   Both this AND fEnableUsbBlockDeviceBySetupClass must be 1 simultaneously
#   for either filter to function.
#
# fEnableUsbNoAckIsochWriteToDevice
#   Controls isochronous write acknowledgement behaviour for USB audio devices.
#   Value 80 (0x50) is the Windows App default. Do not modify without testing.
# ============================================================

Write-Host "`n[2/4] Setting parent flags..." -ForegroundColor White

# Set by this script for reference - must also be applied via GPO (see header)
Set-ItemProperty -Path $clientRegPath -Name "fUsbRedirectionEnableMode"         -Value 2  -Type DWord
Write-Host "      fUsbRedirectionEnableMode         = 2  (NOTE: must also be set via GPO)" -ForegroundColor Yellow

Set-ItemProperty -Path $clientRegPath -Name "fEnableUsbBlockDeviceBySetupClass" -Value 1  -Type DWord
Write-Host "      fEnableUsbBlockDeviceBySetupClass  = 1  (block filter active)" -ForegroundColor Green

Set-ItemProperty -Path $clientRegPath -Name "fEnableUsbSelectDeviceByInterface" -Value 1  -Type DWord
Write-Host "      fEnableUsbSelectDeviceByInterface  = 1  (select filter active)" -ForegroundColor Green

Set-ItemProperty -Path $clientRegPath -Name "fEnableUsbNoAckIsochWriteToDevice" -Value 80 -Type DWord
Write-Host "      fEnableUsbNoAckIsochWriteToDevice  = 80 (isochronous audio default)" -ForegroundColor DarkGray


# ============================================================
# CLEAR EXISTING FILTER SUBKEYS
# Ensures a clean known state on every run - no stale entries
# ============================================================

Write-Host "`n[3/4] Clearing existing filter entries..." -ForegroundColor White

Remove-Item -Path $selectRegPath -Force -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path $blockRegPath  -Force -Recurse -ErrorAction SilentlyContinue
New-Item -Path $selectRegPath -Force | Out-Null
New-Item -Path $blockRegPath  -Force | Out-Null
Write-Host "      Filter subkeys cleared and recreated" -ForegroundColor DarkGray


# ============================================================
# SELECT LIST — UsbSelectDeviceByInterfaces
#
# Controls which devices are SHOWN in the RemoteFX USB picker.
# A device must match at least one GUID in this list to appear.
#
# The generic USB GUID {a5dcbf10} is included to allow ALL USB devices
# through the select filter. This is intentional — the block list below
# handles the actual filtering. This approach is required because composite
# USB devices (headsets etc.) do not reliably expose specific interface GUIDs
# at the parent device level, so a strict allow-list silently excludes them.
# ============================================================

Write-Host "`n[4/4] Writing filter entries..." -ForegroundColor White
Write-Host "`n      SELECT LIST (devices allowed through to block list evaluation):" -ForegroundColor Cyan

$selectGuids = [ordered]@{
    "1000" = @{ Guid = "{6994AD04-93EF-11D0-A3CC-00A0C9223196}"; Desc = "USB Audio interface" }
    "1001" = @{ Guid = "{65E8773D-8F56-11D0-A3B9-00A0C9223196}"; Desc = "KS Audio Streaming interface" }
    "1002" = @{ Guid = "{65E8773E-8F56-11D0-A3B9-00A0C9223196}"; Desc = "KS Audio Topology interface" }
    "1003" = @{ Guid = "{4d1e55b2-f16f-11cf-88cb-001111000030}"; Desc = "HID (headset call control buttons)" }
    "1004" = @{ Guid = "{a5dcbf10-6530-11d2-901f-00c04fb951ed}"; Desc = "Generic USB — allows ALL devices through (block list does the filtering)" }
}

foreach ($entry in $selectGuids.GetEnumerator()) {
    Set-ItemProperty -Path $selectRegPath -Name $entry.Key -Value $entry.Value.Guid -Type String
    Write-Host ("      [{0}] {1,-45} {2}" -f $entry.Key, $entry.Value.Guid, $entry.Value.Desc) -ForegroundColor Cyan
}


# ============================================================
# BLOCK LIST — UsbBlockDeviceBySetupClasses
#
# Controls which devices are HIDDEN from the RemoteFX USB picker.
# Any device whose setup class matches a GUID here will not appear.
#
# IMPORTANT: {36FC9E60-C465-11CF-8056-444553540000} is intentionally
# NOT included in this list. Although commonly referenced for blocking
# USB storage, this GUID is the USB Controllers setup class and covers
# ALL composite USB devices — including headsets. Adding it will silently
# remove all audio devices from the picker.
#
# USB storage is already effectively blocked via {4D36E967} (Disk Drives).
# ============================================================

Write-Host "`n      BLOCK LIST (device classes hidden from picker):" -ForegroundColor Yellow

$blockGuids = [ordered]@{
    "1000" = @{ Guid = "{53D29EF7-377C-4D14-864B-EB3A85769359}"; Desc = "Biometric (fingerprint readers)" }
    "1001" = @{ Guid = "{CA3E7AB9-B4C3-4AE6-8251-579EF933890F}"; Desc = "Camera" }
    "1002" = @{ Guid = "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}"; Desc = "Image / WIA (scanners)" }
    "1003" = @{ Guid = "{4D36E967-E325-11CE-BFC1-08002BE10318}"; Desc = "Disk Drives (USB storage)" }
    "1004" = @{ Guid = "{50DD5230-BA8A-11D1-BF5D-0000F805F530}"; Desc = "Smartcard" }
    "1005" = @{ Guid = "{E0CBF06C-CD8B-4647-BB8A-263B43F0F974}"; Desc = "Bluetooth" }
    "1006" = @{ Guid = "{4D36E972-E325-11CE-BFC1-08002BE10318}"; Desc = "Network Adapters" }
}

foreach ($entry in $blockGuids.GetEnumerator()) {
    Set-ItemProperty -Path $blockRegPath -Name $entry.Key -Value $entry.Value.Guid -Type String
    Write-Host ("      [{0}] {1,-45} {2}" -f $entry.Key, $entry.Value.Guid, $entry.Value.Desc) -ForegroundColor Yellow
}


# ============================================================
# SUMMARY
# ============================================================

Write-Host "`n============================================================" -ForegroundColor White
Write-Host " Configuration applied successfully" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor White
Write-Host @"

 Next steps:
   1. Confirm GPO has applied (see script header for GPO path)
   2. Verify TsUsbFlt is in UpperFilters:
      Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}' | Select-Object UpperFilters
   3. Plug in USB audio device BEFORE opening MSTSC or Windows App
   4. Open MSTSC > Show Options > Local Resources > More
   5. Expand 'Other supported RemoteFX USB devices'
      - Audio devices should appear
      - Cameras, fingerprint readers, storage etc. should not

 If 'Other supported RemoteFX USB devices' category is missing:
   TsUsbFlt is not loaded - GPO has not applied correctly or reboot required

"@ -ForegroundColor Gray
