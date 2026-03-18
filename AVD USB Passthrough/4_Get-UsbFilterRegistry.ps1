<#
.SYNOPSIS
    Display the full current state of the USB filter registry configuration.

.DESCRIPTION
    Shows the current select list and block list contents with human-readable
    descriptions alongside each GUID. Useful for quickly confirming what is
    and isn't in the filter lists without manually reading the registry.

    Also warns if the dangerous {36FC9E60} GUID is found in the block list,
    as this will silently block all composite USB devices including headsets.

.NOTES
    Author  : Andrew Webber
    Context : AVD RemoteFX USB filtering troubleshooting
    No changes are made to the system - read only
#>

Write-Host "`n=== USB Filter Registry Dump ===" -ForegroundColor Cyan

$clientRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client"
$selectRegPath = "$clientRegPath\UsbSelectDeviceByInterfaces"
$blockRegPath  = "$clientRegPath\UsbBlockDeviceBySetupClasses"

# Known GUIDs for reference
$knownGuids = @{
    "{53D29EF7-377C-4D14-864B-EB3A85769359}" = "Biometric (fingerprint readers)"
    "{CA3E7AB9-B4C3-4AE6-8251-579EF933890F}" = "Camera"
    "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}" = "Image / WIA (scanners)"
    "{4D36E967-E325-11CE-BFC1-08002BE10318}" = "Disk Drives (USB storage)"
    "{36FC9E60-C465-11CF-8056-444553540000}" = "*** DANGEROUS *** USB Controllers / Composite - blocks ALL composite devices including headsets"
    "{50DD5230-BA8A-11D1-BF5D-0000F805F530}" = "Smartcard"
    "{E0CBF06C-CD8B-4647-BB8A-263B43F0F974}" = "Bluetooth"
    "{4D36E972-E325-11CE-BFC1-08002BE10318}" = "Network Adapters"
    "{6994AD04-93EF-11D0-A3CC-00A0C9223196}" = "USB Audio interface"
    "{65E8773D-8F56-11D0-A3B9-00A0C9223196}" = "KS Audio Streaming"
    "{65E8773E-8F56-11D0-A3B9-00A0C9223196}" = "KS Audio Topology"
    "{4D1E55B2-F16F-11CF-88CB-001111000030}" = "HID"
    "{A5DCBF10-6530-11D2-901F-00C04FB951ED}" = "Generic USB (allows all devices through)"
    "{4D36E96C-E325-11CE-BFC1-08002BE10318}" = "Sound / MEDIA class"
}

function Resolve-Guid($guid) {
    $upper = $guid.ToUpper()
    if ($knownGuids[$upper]) { return $knownGuids[$upper] }
    if ($knownGuids[$guid])  { return $knownGuids[$guid] }
    return "Unknown GUID"
}

# -------------------------------------------------------
# SELECT LIST
# -------------------------------------------------------
Write-Host "`n--- Select List (UsbSelectDeviceByInterfaces) ---" -ForegroundColor Cyan
Write-Host "    Devices must match at least one entry here to appear in the picker`n" -ForegroundColor DarkGray

$selectItems = Get-ItemProperty $selectRegPath -ErrorAction SilentlyContinue
if ($selectItems) {
    $selectItems.PSObject.Properties | Where-Object { $_.Name -match '^\d+$' } |
        Sort-Object { [int]$_.Name } | ForEach-Object {
            $desc = Resolve-Guid $_.Value
            Write-Host ("    [{0}] {1,-45} {2}" -f $_.Name, $_.Value, $desc) -ForegroundColor Cyan
        }
} else {
    Write-Host "    [EMPTY] No entries found - if select filter is enabled this will block ALL devices" -ForegroundColor Red
}

# -------------------------------------------------------
# BLOCK LIST
# -------------------------------------------------------
Write-Host "`n--- Block List (UsbBlockDeviceBySetupClasses) ---" -ForegroundColor Yellow
Write-Host "    Devices matching these setup class GUIDs will be hidden from the picker`n" -ForegroundColor DarkGray

$blockItems = Get-ItemProperty $blockRegPath -ErrorAction SilentlyContinue
$dangerousFound = $false

if ($blockItems) {
    $blockItems.PSObject.Properties | Where-Object { $_.Name -match '^\d+$' } |
        Sort-Object { [int]$_.Name } | ForEach-Object {
            $desc = Resolve-Guid $_.Value
            if ($_.Value -like "*36FC9E60*") {
                $dangerousFound = $true
                Write-Host ("    [{0}] {1,-45} {2}" -f $_.Name, $_.Value, $desc) -ForegroundColor Red
            } else {
                Write-Host ("    [{0}] {1,-45} {2}" -f $_.Name, $_.Value, $desc) -ForegroundColor Yellow
            }
        }

    if ($dangerousFound) {
        Write-Host "`n    [WARNING] {36FC9E60} found in block list!" -ForegroundColor Red
        Write-Host "              This GUID covers the USB Controllers setup class which includes" -ForegroundColor Red
        Write-Host "              ALL composite USB devices - headsets will be silently blocked." -ForegroundColor Red
        Write-Host "              Remove this entry from the block list." -ForegroundColor Red
    }
} else {
    Write-Host "    [EMPTY] No entries found" -ForegroundColor Red
}

Write-Host ""
