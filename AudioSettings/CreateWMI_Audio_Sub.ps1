<#
.SYNOPSIS
    Registers a permanent WMI event subscription to call Set-AudioDeviceSettings.ps1
    whenever an audio endpoint device is connected.
.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Run once, elevated, during image build or on first boot.
    The subscription survives reboots and runs as SYSTEM.
    Log:     C:\_source\logs\AudioWMISubscription_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

$filterName   = "AudioDeviceArrival"
$consumerName = "AudioDeviceArrivalConsumer"
$scriptPath   = "C:\_source\Set-AudioDeviceSettings.ps1"
$logDir       = "C:\_source\logs"
$logShare     = "\\cukavdukwprod01.file.core.windows.net\profiles\fslogixrules\Logs"
$logName      = "${env:COMPUTERNAME}_AudioWMISubscription_$(Get-Date -Format 'yyyyMMdd').log"
$logFile      = Join-Path $logDir $logName

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
    if ($logShare) {
        try {
            if (-not (Test-Path $logShare)) {
                New-Item -ItemType Directory -Path $logShare -Force | Out-Null
            }
            Add-Content -Path (Join-Path $logShare $logName) -Value $entry
        }
        catch {
            # Silently skip if the share is unreachable
        }
    }
    Write-Host $entry
}

# ============================================================
# MAIN
# ============================================================

Write-Log "========================================"
Write-Log "WMI subscription registration started"
Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log "========================================"

# Remove any existing subscription to avoid duplicates
Write-Log "Removing any existing subscription components..."

Get-WmiObject -Namespace "root\subscription" -Class "__FilterToConsumerBinding" -ErrorAction SilentlyContinue |
    Where-Object { $_.Filter -like "*$filterName*" } |
    Remove-WmiObject
Write-Log "  Cleared existing bindings"

Get-WmiObject -Namespace "root\subscription" -Class "__EventFilter" |
    Where-Object { $_.Name -eq $filterName } |
    Remove-WmiObject
Write-Log "  Cleared existing filters"

Get-WmiObject -Namespace "root\subscription" -Class "CommandLineEventConsumer" |
    Where-Object { $_.Name -eq $consumerName } |
    Remove-WmiObject
Write-Log "  Cleared existing consumers"

# 1. Event filter — fires when a new AudioEndpoint PnP device is created
Write-Log "Creating event filter: $filterName"
$filter = Set-WmiInstance -Namespace "root\subscription" -Class "__EventFilter" -Arguments @{
    Name           = $filterName
    EventNamespace = "root\cimv2"
    QueryLanguage  = "WQL"
    Query          = "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity' AND TargetInstance.PNPClass = 'AudioEndpoint'"
}
Write-Log "  OK  Event filter created"

# 2. Consumer — runs the remediation script as SYSTEM
Write-Log "Creating consumer: $consumerName"
$consumer = Set-WmiInstance -Namespace "root\subscription" -Class "CommandLineEventConsumer" -Arguments @{
    Name                = $consumerName
    CommandLineTemplate = "powershell.exe -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
}
Write-Log "  OK  Consumer created (script: $scriptPath)"

# 3. Bind filter to consumer
Write-Log "Binding filter to consumer..."
Set-WmiInstance -Namespace "root\subscription" -Class "__FilterToConsumerBinding" -Arguments @{
    Filter   = $filter
    Consumer = $consumer
}
Write-Log "  OK  Binding created"

Write-Log "========================================"
Write-Log "WMI subscription registered successfully"
Write-Log "  Filter:   $filterName"
Write-Log "  Consumer: $consumerName"
Write-Log "  Script:   $scriptPath"
Write-Log "========================================"
