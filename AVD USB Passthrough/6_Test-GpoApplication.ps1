<#
.SYNOPSIS
    Check whether the RemoteFX USB GPO is being applied to this machine.

.DESCRIPTION
    Verifies the RemoteFX USB GPO is applying correctly by checking:
        1. Applied computer GPOs from gpresult
        2. RemoteFX specific GPO entries
        3. fUsbRedirectionEnableMode registry value
        4. Machine OU location

    If the GPO is not applying, the RemoteFX USB category will not appear
    in MSTSC or Windows App regardless of what registry values are set.
    Run gpupdate /force and reboot if the GPO is not showing as applied.

.NOTES
    Author  : Andrew Webber
    Context : AVD RemoteFX USB filtering troubleshooting
    No changes are made to the system - read only

    GPO that must be applied:
    Computer Configuration > Policies > Administrative Templates > Windows Components
    > Remote Desktop Services > Remote Desktop Connection Client
    > RemoteFX USB Device Redirection
    > Allow RDP redirection of other supported RemoteFX USB devices from this computer
    Setting : Enabled | Access Rights : Administrators and Users
#>

Write-Host "`n=== RemoteFX GPO Application Check ===" -ForegroundColor Cyan
Write-Host "    Machine : $env:COMPUTERNAME" -ForegroundColor DarkGray
Write-Host "    User    : $env:USERNAME`n" -ForegroundColor DarkGray


# -------------------------------------------------------
# CHECK 1: Applied computer GPOs
# -------------------------------------------------------
Write-Host "--- [1] Applied Computer GPOs ---" -ForegroundColor White
gpresult /scope computer /r 2>&1 | Select-String -Pattern "GPO|Applied" |
    ForEach-Object { Write-Host "    $_" }


# -------------------------------------------------------
# CHECK 2: RemoteFX specific GPO entry
# -------------------------------------------------------
Write-Host "`n--- [2] RemoteFX Specific GPO Entries ---" -ForegroundColor White
$gpoResult = gpresult /scope computer /v 2>&1 | Select-String -Context 3,3 "RemoteFX"

if ($gpoResult) {
    $gpoResult | ForEach-Object { Write-Host "    $_" -ForegroundColor Green }
} else {
    Write-Host "    [WARN] No RemoteFX entries found in gpresult output" -ForegroundColor Red
    Write-Host "           Possible causes:" -ForegroundColor Red
    Write-Host "             - GPO not linked to this machine's OU" -ForegroundColor Red
    Write-Host "             - GPO not yet applied - try: gpupdate /force then reboot" -ForegroundColor Red
    Write-Host "             - Security filtering preventing GPO from applying" -ForegroundColor Red
    Write-Host "             - Machine not in correct OU (see check 4 below)" -ForegroundColor Red
}


# -------------------------------------------------------
# CHECK 3: fUsbRedirectionEnableMode registry value
# -------------------------------------------------------
Write-Host "`n--- [3] fUsbRedirectionEnableMode Registry Value ---" -ForegroundColor White
$clientRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client"
$regVal = (Get-ItemProperty $clientRegPath -ErrorAction SilentlyContinue).fUsbRedirectionEnableMode

if ($regVal -eq 2) {
    Write-Host "    fUsbRedirectionEnableMode = $regVal  [OK]" -ForegroundColor Green
    Write-Host "    NOTE: Value being present does not guarantee GPO was the source." -ForegroundColor DarkGray
    Write-Host "          Verify TsUsbFlt is in UpperFilters using 3_Test-RemoteFXStack.ps1" -ForegroundColor DarkGray
} elseif ($null -eq $regVal) {
    Write-Host "    fUsbRedirectionEnableMode = NOT SET  [FAIL]" -ForegroundColor Red
    Write-Host "    GPO has not applied or has not been created yet" -ForegroundColor Red
} else {
    Write-Host "    fUsbRedirectionEnableMode = $regVal  [UNEXPECTED - should be 2]" -ForegroundColor Red
}


# -------------------------------------------------------
# CHECK 4: Machine OU location
# -------------------------------------------------------
Write-Host "`n--- [4] Machine OU Location ---" -ForegroundColor White
try {
    $ouPath = ([adsisearcher]"(cn=$env:COMPUTERNAME)").FindOne().Path
    if ($ouPath) {
        Write-Host "    $ouPath" -ForegroundColor Gray
        Write-Host "    Verify the RemoteFX GPO is linked to this OU or a parent OU" -ForegroundColor DarkGray
    } else {
        Write-Host "    Could not determine OU path" -ForegroundColor Yellow
    }
} catch {
    Write-Host "    Could not determine OU - machine may not be domain joined" -ForegroundColor Yellow
}


# -------------------------------------------------------
# SUMMARY
# -------------------------------------------------------
Write-Host "`n--- Summary ---" -ForegroundColor White

$upperFilters = (Get-ItemProperty `
    "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{36FC9E60-C465-11CF-8056-444553540000}" `
    -ErrorAction SilentlyContinue).UpperFilters

if ($upperFilters -contains "TsUsbFlt" -and $regVal -eq 2) {
    Write-Host "    [PASS] GPO appears to have applied correctly" -ForegroundColor Green
    Write-Host "           TsUsbFlt is in UpperFilters and fUsbRedirectionEnableMode = 2" -ForegroundColor Green
} elseif ($regVal -eq 2 -and $upperFilters -notcontains "TsUsbFlt") {
    Write-Host "    [FAIL] fUsbRedirectionEnableMode is set but TsUsbFlt is NOT in UpperFilters" -ForegroundColor Red
    Write-Host "           The registry value may have been set directly rather than via GPO" -ForegroundColor Red
    Write-Host "           The Terminal Services CSE has not initialised the RemoteFX USB stack" -ForegroundColor Red
    Write-Host "           Ensure GPO is applied and reboot the machine" -ForegroundColor Red
} else {
    Write-Host "    [FAIL] GPO does not appear to be applied correctly" -ForegroundColor Red
    Write-Host "           Run: gpupdate /force then reboot and re-run this script" -ForegroundColor Red
}

Write-Host ""
