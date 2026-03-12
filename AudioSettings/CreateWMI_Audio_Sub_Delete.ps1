<#
.SYNOPSIS
    Removes the AudioDeviceArrival WMI permanent event subscription.
.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Log:     C:\_source\logs\AudioWMISubscription_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

$filterName   = "AudioDeviceArrival"
$consumerName = "AudioDeviceArrivalConsumer"
$logDir       = "C:\_source\logs"
$logFile      = Join-Path $logDir "AudioWMISubscriptionRemoval_$(Get-Date -Format 'yyyyMMdd').log"

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

# ============================================================
# MAIN
# ============================================================

Write-Log "========================================"
Write-Log "WMI subscription debug / removal started"
Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log "========================================"

# Display current state
Write-Log "Querying current WMI subscription state..."

$filters  = Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter"
$consumers = Get-WmiObject -Namespace "root\subscription" -Class "CommandLineEventConsumer"
$bindings = Get-WmiObject -Namespace "root\subscription" -Class "__FilterToConsumerBinding"

Write-Log "  Filters:   $($filters.Count) found"
Write-Log "  Consumers: $($consumers.Count) found"
Write-Log "  Bindings:  $($bindings.Count) found"

# Remove binding first
Write-Log "Removing binding..."
Get-WmiObject -Namespace "root\subscription" -Class "__FilterToConsumerBinding" |
    Where-Object { $_.Filter -like "*$filterName*" } |
    Remove-WmiObject
Write-Log "  OK  Binding removed"

# Remove filter
Write-Log "Removing filter: $filterName"
Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter" |
    Where-Object { $_.Name -eq $filterName } |
    Remove-WmiObject
Write-Log "  OK  Filter removed"

# Remove consumer
Write-Log "Removing consumer: $consumerName"
Get-WmiObject -Namespace "root\subscription" -Class "CommandLineEventConsumer" |
    Where-Object { $_.Name -eq $consumerName } |
    Remove-WmiObject
Write-Log "  OK  Consumer removed"

# Verify
Write-Log "Verifying removal..."
$remaining = Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter" |
             Where-Object { $_.Name -eq $filterName }

if ($remaining) {
    Write-Log "Filter still present -- check manually" -Level "WARN"
} else {
    Write-Log "Verified: No subscription entries remain"
}

Write-Log "========================================"
Write-Log "WMI subscription removal completed"
Write-Log "========================================"
