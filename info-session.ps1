#Install-Module ps2exe
#Invoke-PS2EXE .\SessionInfo.ps1 .\SessionInfo.exe -noConsole -iconFile .\avd.ico

<#
.SYNOPSIS
    Display computer name and AVD image information
#>

$ProgressPreference = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms

$computerName = $env:COMPUTERNAME

# FASTEST: Direct .NET call for IP
try {
    $ipAddresses = [System.Net.Dns]::GetHostAddresses($env:COMPUTERNAME) | 
        Where-Object { $_.AddressFamily -eq 'InterNetwork' -and $_.ToString() -notlike '127.*' } |
        ForEach-Object { $_.ToString() }
    $ipInfo = if ($ipAddresses) { $ipAddresses -join ', ' } else { "No IP found" }
}
catch {
    $ipInfo = "No IP found"
}

# Registry
$regPath = 'HKLM:\Software\AVD_Image_Version'
if (Test-Path $regPath) {
    $regKey = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    $creationDate = $regKey.CreationDate
    $image = $regKey.Image
}
else {
    $creationDate = "Registry key not found"
    $image = "Registry key not found"
}

$message = "AVD Sessionhost: $computerName`nIP Address: $ipInfo`n`nImage Date: $creationDate`nImage: $image"

# Simplest TopMost approach
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true

$null = [System.Windows.Forms.MessageBox]::Show($form, $message, 'AVD Session Information', 'OK', 'None')

$form.Dispose()