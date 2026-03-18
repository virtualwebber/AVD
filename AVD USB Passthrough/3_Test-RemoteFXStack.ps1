<#
.SYNOPSIS
    Verify that the RemoteFX USB stack is correctly initialised on the client machine.

.DESCRIPTION
    This is the first diagnostic to run when devices are not appearing in the
    RemoteFX USB picker. Checks three things:

        1. TsUsbFlt is present in USB controller class UpperFilters
        2. TsUsbFlt driver state
        3. Current filter flag values in the registry

    EXPECTED RESULTS:
        UpperFilters : {TsUsbFlt}   <- must be present
        TsUsbFlt STATE : STOPPED    <- normal, driver starts on demand

    IF TsUsbFlt IS NOT IN UPPERFILTERS:
        The RemoteFX GPO has not applied correctly or the machine needs a reboot.
        Setting fUsbRedirectionEnableMode via registry alone will NOT insert TsUsbFlt.
        The GPO must be applied via Group Policy for the CSE to initialise the stack.

        GPO PATH:
        Computer Configuration > Policies > Administrative Templates > Windows Components
        > Remote Desktop Services > Remote Desktop Connection Client
        > RemoteFX USB Device Redirection
        > Allow RDP redirection of other supported RemoteFX USB devices from this computer
        Setting : Enabled | Access : Administrators and Users

.NOTES
    Author  : Andrew Webber
    Context : AVD RemoteFX USB filtering troubleshooting
    No changes are made to the system - read only
#>

Write-Host "`n=== RemoteFX USB Stack Verification ===" -ForegroundColor Cyan

# -------------------------------------------------------
# CHECK 1: TsUsbFlt in UpperFilters
# -------------------------------------------------------
Write-Host "`n--- [1] USB Controller Class UpperFilters ---" -ForegroundColor White

$upperFilters = (Get-ItemProperty `
    "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}" `
    -ErrorAction SilentlyContinue).UpperFilters

if ($upperFilters -contains "TsUsbFlt") {
    Write-Host "    UpperFilters : $upperFilters" -ForegroundColor Green
    Write-Host "    [PASS] TsUsbFlt is present - RemoteFX USB stack is initialised" -ForegroundColor Green
} else {
    Write-Host "    UpperFilters : $upperFilters" -ForegroundColor Red
    Write-Host "    [FAIL] TsUsbFlt is NOT in UpperFilters" -ForegroundColor Red
    Write-Host "           The RemoteFX GPO has not applied correctly or machine needs a reboot" -ForegroundColor Red
    Write-Host "           Setting fUsbRedirectionEnableMode via registry alone is not sufficient" -ForegroundColor Red
}


# -------------------------------------------------------
# CHECK 2: TsUsbFlt driver state
# -------------------------------------------------------
Write-Host "`n--- [2] TsUsbFlt Driver Status ---" -ForegroundColor White

$svcStatus = sc.exe query TsUsbFlt
$svcStatus | ForEach-Object { Write-Host "    $_" }

if ($svcStatus -match "RUNNING") {
    Write-Host "    [INFO] TsUsbFlt is running" -ForegroundColor Green
} elseif ($svcStatus -match "STOPPED") {
    Write-Host "    [INFO] TsUsbFlt is stopped - this is normal, it starts on demand" -ForegroundColor Yellow
} else {
    Write-Host "    [WARN] Unexpected TsUsbFlt state" -ForegroundColor Red
}


# -------------------------------------------------------
# CHECK 3: Filter flag values
# -------------------------------------------------------
Write-Host "`n--- [3] Filter Flag Registry Values ---" -ForegroundColor White

$clientRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client"
$flags = Get-ItemProperty $clientRegPath -ErrorAction SilentlyContinue

$flagList = @(
    @{ Name = "fUsbRedirectionEnableMode";         Expected = 2;  Note = "Must be set via GPO not registry directly" }
    @{ Name = "fEnableUsbBlockDeviceBySetupClass";  Expected = 1;  Note = "Activates block filter" }
    @{ Name = "fEnableUsbSelectDeviceByInterface";  Expected = 1;  Note = "Activates select filter" }
    @{ Name = "fEnableUsbNoAckIsochWriteToDevice";  Expected = 80; Note = "Isochronous audio tuning - do not change" }
)

foreach ($flag in $flagList) {
    $val = $flags.($flag.Name)
    $status = if ($val -eq $flag.Expected) { "[OK]" } else { "[UNEXPECTED]" }
    $colour = if ($val -eq $flag.Expected) { "Green" } else { "Red" }
    Write-Host ("    {0,-45} = {1,-5} {2}  ({3})" -f $flag.Name, $val, $status, $flag.Note) -ForegroundColor $colour
}

# Warn if both filters aren't 1 simultaneously
if ($flags.fEnableUsbBlockDeviceBySetupClass -ne 1 -or $flags.fEnableUsbSelectDeviceByInterface -ne 1) {
    Write-Host "`n    [WARN] Both fEnableUsbBlockDeviceBySetupClass AND fEnableUsbSelectDeviceByInterface" -ForegroundColor Red
    Write-Host "           must be set to 1 simultaneously for filtering to activate" -ForegroundColor Red
}

Write-Host ""
