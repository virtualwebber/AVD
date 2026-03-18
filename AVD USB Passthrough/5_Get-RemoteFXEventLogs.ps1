<#
.SYNOPSIS
    Gather relevant event log entries and driver status into a single text file.

.DESCRIPTION
    Collects TsUsbFlt driver status, UpperFilters, Terminal Services client events,
    and TsUsb system events into C:\temp\rdp-usb-debug.txt for offline analysis.

.NOTES
    Author  : Andrew Webber
    Context : AVD RemoteFX USB filtering troubleshooting
    Output  : C:\temp\rdp-usb-debug.txt
    Requires: Administrator rights to read some event logs
#>

#Requires -RunAsAdministrator

Write-Host "`n=== Collecting RemoteFX USB Event Logs ===" -ForegroundColor Cyan

# Ensure output folder exists
New-Item -Path "C:\GitHub\VirtualWebber\Scripts\AVD\AVD USB Passthrough" -ItemType Directory -Force | Out-Null

$output = @()
$output += "RemoteFX USB Diagnostic Log"
$output += "Generated : $(Get-Date)"
$output += "Machine   : $env:COMPUTERNAME"
$output += "User      : $env:USERNAME"
$output += "=" * 60


# -------------------------------------------------------
# TsUsbFlt driver status
# -------------------------------------------------------
Write-Host "    Collecting TsUsbFlt driver status..." -ForegroundColor DarkGray
$output += "`n=== TsUsbFlt Driver Status ==="
$output += (sc.exe query TsUsbFlt | Out-String)


# -------------------------------------------------------
# UpperFilters
# -------------------------------------------------------
Write-Host "    Collecting UpperFilters..." -ForegroundColor DarkGray
$output += "=== USB Controller Class UpperFilters ==="
$output += (Get-ItemProperty `
    "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}" `
    -ErrorAction SilentlyContinue | Select-Object UpperFilters | Format-List | Out-String)


# -------------------------------------------------------
# Filter registry values
# -------------------------------------------------------
Write-Host "    Collecting filter registry values..." -ForegroundColor DarkGray
$clientRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client"
$output += "=== Filter Flag Registry Values ==="
$output += (Get-ItemProperty $clientRegPath -ErrorAction SilentlyContinue |
    Select-Object fUsbRedirectionEnableMode, fEnableUsbBlockDeviceBySetupClass,
                  fEnableUsbSelectDeviceByInterface, fEnableUsbNoAckIsochWriteToDevice |
    Format-List | Out-String)

$output += "=== UsbSelectDeviceByInterfaces ==="
$output += (Get-ItemProperty "$clientRegPath\UsbSelectDeviceByInterfaces" `
    -ErrorAction SilentlyContinue | Format-List | Out-String)

$output += "=== UsbBlockDeviceBySetupClasses ==="
$output += (Get-ItemProperty "$clientRegPath\UsbBlockDeviceBySetupClasses" `
    -ErrorAction SilentlyContinue | Format-List | Out-String)


# -------------------------------------------------------
# Terminal Services RDP client events
# -------------------------------------------------------
Write-Host "    Collecting Terminal Services client events..." -ForegroundColor DarkGray
$output += "=== Terminal Services RDPClient/Operational Events ==="
$output += (Get-WinEvent -LogName "Microsoft-Windows-TerminalServices-RDPClient/Operational" `
    -ErrorAction SilentlyContinue |
    Select-Object TimeCreated, Id, Message | Format-List | Out-String)


# -------------------------------------------------------
# TsUsb system events
# -------------------------------------------------------
Write-Host "    Collecting TsUsb system events..." -ForegroundColor DarkGray
$output += "=== System Events (TsUsb related) ==="
$output += (Get-WinEvent -LogName "System" -ErrorAction SilentlyContinue |
    Where-Object { $_.ProviderName -like "*TsUsb*" -or $_.Message -like "*TsUsb*" } |
    Select-Object TimeCreated, Id, ProviderName, Message | Format-List | Out-String)


# -------------------------------------------------------
# Write output
# -------------------------------------------------------
$outputPath = "C:\GitHub\VirtualWebber\Scripts\AVD\AVD USB Passthrough\rdp-usb-debug.txt"
$output | Out-File $outputPath -Encoding UTF8

Write-Host "`n    Saved to $outputPath" -ForegroundColor Green
Write-Host "    Share this file for offline analysis`n" -ForegroundColor DarkGray
