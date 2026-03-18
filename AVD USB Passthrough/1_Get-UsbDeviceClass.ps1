<#
.SYNOPSIS
    Enumerate USB devices and their setup class name and GUID.

.DESCRIPTION
    Used to identify what setup class a device registers under so it can be
    added to the UsbBlockDeviceBySetupClasses or UsbSelectDeviceByInterfaces lists.

.PARAMETER VidPid
    Optional VID/PID filter e.g. "VID_047F&PID_C03A"

.PARAMETER DeviceName
    Optional friendly name filter e.g. "Plantronics"

.EXAMPLE
    # All USB devices
    .\1_Get-UsbDeviceClass.ps1

    # Filter by VID/PID
    .\1_Get-UsbDeviceClass.ps1 -VidPid "VID_047F&PID_C03A"

    # Filter by friendly name
    .\1_Get-UsbDeviceClass.ps1 -DeviceName "Plantronics"

.NOTES
    Author  : Andrew Webber
    Context : AVD RemoteFX USB filtering troubleshooting
    No changes are made to the system - read only
#>

param(
    [string]$VidPid,
    [string]$DeviceName
)

Write-Host "`n=== USB Device Class Lookup ===" -ForegroundColor Cyan

$devices = Get-PnpDevice | Where-Object { $_.InstanceId -like "USB\*" }

if ($VidPid)     { $devices = $devices | Where-Object { $_.InstanceId -like "*$VidPid*" } }
if ($DeviceName) { $devices = $devices | Where-Object { $_.FriendlyName -like "*$DeviceName*" } }

foreach ($dev in $devices) {
    $classGuid = (Get-PnpDeviceProperty -InstanceId $dev.InstanceId `
        -KeyName 'DEVPKEY_Device_ClassGuid' -ErrorAction SilentlyContinue).Data

    # Resolve class GUID to human-readable name via registry
    $className = $null
    if ($classGuid) {
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\$classGuid"
        $className = (Get-ItemProperty $regPath -ErrorAction SilentlyContinue).Class
    }

    [PSCustomObject]@{
        FriendlyName = $dev.FriendlyName
        InstanceId   = $dev.InstanceId
        Status       = $dev.Status
        Class        = $dev.Class
        ClassGuid    = $classGuid
        ClassName    = $className
    }
} | Format-List
