<#
.SYNOPSIS
    Find the actual interface class GUIDs that a specific USB device exposes.

.DESCRIPTION
    Standard PnP cmdlets often return empty results for composite USB devices.
    This script searches the DeviceClasses registry directly using the device
    VID/PID which is far more reliable.

    This is the key diagnostic for identifying which GUIDs to add to
    UsbSelectDeviceByInterfaces. Used to identify the Plantronics Blackwire
    interface GUIDs during AVD RemoteFX USB filtering troubleshooting.

.NOTES
    Author  : Andrew Webber
    Context : AVD RemoteFX USB filtering troubleshooting
    No changes are made to the system - read only

    Edit the $targetVidPid variable below to match your device.
    VID/PID can be found from Script 1 (1_Get-UsbDeviceClass.ps1) or Device Manager.

.EXAMPLE
    # Edit $targetVidPid in the script then run:
    .\2_Get-UsbDeviceInterfaces.ps1

.OUTPUT
    InterfaceGUID                            Device
    {65E8773D-8F56-11D0-A3B9-00A0C9223196}  ##?#USB#VID_047F&PID_C03A...
    {6994AD04-93EF-11D0-A3CC-00A0C9223196}  ##?#USB#VID_047F&PID_C03A...
    {a5dcbf10-6530-11d2-901f-00c04fb951ed}  ##?#USB#VID_047F&PID_C03A...
#>

# *** CHANGE THIS TO YOUR DEVICE VID/PID ***
$targetVidPid = "VID_047F&PID_C03A"   # Plantronics Blackwire 325.1

Write-Host "`n=== USB Device Interface GUID Lookup ===" -ForegroundColor Cyan
Write-Host "    Target device : $targetVidPid" -ForegroundColor DarkGray
Write-Host "    Searching HKLM:\SYSTEM\CurrentControlSet\Control\DeviceClasses...`n" -ForegroundColor DarkGray

$basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceClasses"

$results = Get-ChildItem $basePath | ForEach-Object {
    $guid = $_.PSChildName
    Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.PSChildName -like "*$targetVidPid*") {
            [PSCustomObject]@{
                InterfaceGUID = $guid
                Device        = $_.PSChildName
            }
        }
    }
}

if ($results) {
    $results | Format-Table -AutoSize

    Write-Host "`n    Unique interface GUIDs found for $targetVidPid`:" -ForegroundColor Green
    $results | Select-Object -ExpandProperty InterfaceGUID -Unique | ForEach-Object {
        Write-Host "      $_" -ForegroundColor Green
    }
    Write-Host "`n    Add relevant GUIDs to UsbSelectDeviceByInterfaces as needed." -ForegroundColor DarkGray
} else {
    Write-Host "    No results found for $targetVidPid" -ForegroundColor Red
    Write-Host "    Check the device is plugged in and the VID/PID is correct." -ForegroundColor Red
}
