# Filter
Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter"

# Consumer
Get-WmiObject -Namespace "root\subscription" -Class "CommandLineEventConsumer"

# Binding
Get-WmiObject -Namespace "root\subscription" -Class "__FilterToConsumerBinding"


<#
.SYNOPSIS
    Removes the AudioDeviceArrival WMI permanent event subscription.
.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
#>

$filterName   = "AudioDeviceArrival"
$consumerName = "AudioDeviceArrivalConsumer"

# Remove binding first
Get-WmiObject -Namespace "root\subscription" -Class "__FilterToConsumerBinding" |
    Where-Object { $_.Filter -like "*$filterName*" } |
    Remove-WmiObject

# Remove filter
Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter" |
    Where-Object { $_.Name -eq $filterName } |
    Remove-WmiObject

# Remove consumer
Get-WmiObject -Namespace "root\subscription" -Class "CommandLineEventConsumer" |
    Where-Object { $_.Name -eq $consumerName } |
    Remove-WmiObject

Write-Host "WMI subscription removed."

# Verify
$remaining = Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter" |
             Where-Object { $_.Name -eq $filterName }

if ($remaining) {
    Write-Host "WARN: Filter still present -- check manually" -ForegroundColor Yellow
} else {
    Write-Host "Verified: No subscription entries remain." -ForegroundColor Green
}