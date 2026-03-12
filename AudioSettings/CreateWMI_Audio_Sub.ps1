<#
.SYNOPSIS
    Registers a permanent WMI event subscription to call Set-AudioDeviceSettings.ps1
    whenever an audio endpoint device is connected.
.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Run once, elevated, during image build or on first boot.
    The subscription survives reboots and runs as SYSTEM.
#>

$filterName   = "AudioDeviceArrival"
$consumerName = "AudioDeviceArrivalConsumer"
$scriptPath   = "C:\_source\scripts\Set-AudioDeviceSettings.ps1"

# Remove any existing subscription to avoid duplicates
Get-WmiObject -Namespace "root\subscription" -Class "__FilterToConsumerBinding" -ErrorAction SilentlyContinue |
    Where-Object { $_.Filter -like "*$filterName*" } |
    Remove-WmiObject

Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter" |
    Where-Object { $_.Name -eq $filterName } |
    Remove-WmiObject

Get-WmiObject -Namespace "root\subscription" -Class "CommandLineEventConsumer" |
    Where-Object { $_.Name -eq $consumerName } |
    Remove-WmiObject

# 1. Event filter — fires when a new AudioEndpoint PnP device is created
$filter = Set-WmiInstance -Namespace "root\subscription" -Class "__EventFilter" -Arguments @{
    Name           = $filterName
    EventNamespace = "root\cimv2"
    QueryLanguage  = "WQL"
    Query          = "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity' AND TargetInstance.PNPClass = 'AudioEndpoint'"
}

# 2. Consumer — runs the remediation script as SYSTEM
$consumer = Set-WmiInstance -Namespace "root\subscription" -Class "CommandLineEventConsumer" -Arguments @{
    Name                = $consumerName
    CommandLineTemplate = "powershell.exe -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
}

# 3. Bind filter to consumer
Set-WmiInstance -Namespace "root\subscription" -Class "__FilterToConsumerBinding" -Arguments @{
    Filter   = $filter
    Consumer = $consumer
}

Write-Host "WMI subscription registered successfully."
Write-Host "Filter:   $filterName"
Write-Host "Consumer: $consumerName"
Write-Host "Script:   $scriptPath"