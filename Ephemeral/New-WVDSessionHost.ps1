##############################
#    AVD Script Parameters   #
##############################
Param (        
    [Parameter(Mandatory=$true)]
        [string]$ProfilePath,
    [Parameter(Mandatory=$true)]
        [string]$RegistrationToken,
    [Parameter(Mandatory=$false)]
        [string]$Optimize = $true           
)


######################
#    AVD Variables   #
######################
$LocalAVDpath            = "c:\temp\AVD\"
$AVDBootURI              = 'https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RWrxrH'
$AVDAgentURI             = 'https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RWrmXv'
$FSLogixURI              = 'https://aka.ms/fslogix_download'
$FSInstaller             = 'FSLogixAppsSetup.zip'
$WVDAgentInstaller       = 'WVD-Agent.msi'
$WVDBootInstaller        = 'WVD-Bootloader.msi'
$Optimizations           = "All"


####################################
#    Test/Create Temp Directory    #
####################################
if((Test-Path c:\temp) -eq $false) {
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Create C:\temp Directory"
    Write-Host `
        -ForegroundColor Cyan `
        -BackgroundColor Black `
        "creating temp directory"
    New-Item -Path c:\temp -ItemType Directory
}
else {
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "C:\temp Already Exists"
    Write-Host `
        -ForegroundColor Yellow `
        -BackgroundColor Black `
        "temp directory already exists"
}
if((Test-Path $LocalAVDpath) -eq $false) {
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Create C:\temp\AVD Directory"
    Write-Host `
        -ForegroundColor Cyan `
        -BackgroundColor Black `
        "creating c:\temp\AVD directory"
    New-Item -Path $LocalAVDpath -ItemType Directory
}
else {
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "C:\temp\AVD Already Exists"
    Write-Host `
        -ForegroundColor Yellow `
        -BackgroundColor Black `
        "c:\temp\AVD directory already exists"
}
New-Item -Path c:\ -Name New-AVDSessionHost.log -ItemType File
Add-Content `
-LiteralPath C:\New-AVDSessionHost.log `
"
ProfilePath       = $ProfilePath
RegistrationToken = $RegistrationToken
Optimize          = $Optimize
"

Start-Transcript -Path "C:\temo\transcript.txt" -NoClobber

#################################
#    Download AVD Componants    #
#################################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Downloading AVD Boot Loader"
    Invoke-WebRequest -Uri $AVDBootURI -OutFile "$LocalAVDpath$WVDBootInstaller"
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Downloading FSLogix"
    Invoke-WebRequest -Uri $FSLogixURI -OutFile "$LocalAVDpath$FSInstaller"
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Downloading AVD Agent"
    Invoke-WebRequest -Uri $AVDAgentURI -OutFile "$LocalAVDpath$WVDAgentInstaller"


##############################
#    Prep for AVD Install    #
##############################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Unzip FSLogix"
Expand-Archive `
    -LiteralPath "C:\temp\AVD\$FSInstaller" `
    -DestinationPath "$LocalAVDpath\FSLogix" `
    -Force `
    -Verbose
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
cd $LocalAVDpath 
Add-Content -LiteralPath C:\New-AVDSessionHost.log "UnZip FXLogix Complete"

##############################
#    OS Specific Settings    #
##############################
$OS = (Get-WmiObject win32_operatingsystem).name
If(($OS) -match 'server') {
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Windows Server OS Detected"
    write-host -ForegroundColor Cyan -BackgroundColor Black "Windows Server OS Detected"
    If(((Get-WindowsFeature -Name RDS-RD-Server).installstate) -eq 'Installed') {
        "Session Host Role is already installed"
    }
    Else {
        "Installing Session Host Role"
        Install-WindowsFeature `
            -Name RDS-RD-Server `
            -Verbose `
            -LogPath "$LocalAVDpath\RdsServerRoleInstall.txt"
    }
    $AdminsKey = "SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
    $UsersKey = "SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
    $BaseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey("LocalMachine","Default")
    $SubKey = $BaseKey.OpenSubkey($AdminsKey,$true)
    $SubKey.SetValue("IsInstalled",0,[Microsoft.Win32.RegistryValueKind]::DWORD)
    $SubKey = $BaseKey.OpenSubKey($UsersKey,$true)
    $SubKey.SetValue("IsInstalled",0,[Microsoft.Win32.RegistryValueKind]::DWORD)    
}
Else {
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Windows Client OS Detected"
    write-host -ForegroundColor Cyan -BackgroundColor Black "Windows Client OS Detected"
    if(($OS) -match 'Windows 10') {
        write-host `
            -ForegroundColor Yellow `
            -BackgroundColor Black  `
            "Windows 10 detected...skipping to next step"
        Add-Content -LiteralPath C:\New-AVDSessionHost.log "Windows 10 Detected...skipping to next step"     
    }    
    else { 
    }
}


################################
#    Install AVD Componants    #
################################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Installing AVD Bootloader"
$bootloader_deploy_status = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList "/i $WVDBootInstaller", `
        "/quiet", `
        "/qn", `
        "/norestart", `
        "/passive", `
        "/l* $LocalAVDpath\AgentBootLoaderInstall.txt" `
    -Wait `
    -Passthru
$sts = $bootloader_deploy_status.ExitCode
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Installing AVD Bootloader Complete"
Write-Output "Installing RDAgentBootLoader on VM Complete. Exit code=$sts`n"
Wait-Event -Timeout 5
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Installing AVD Agent"
Write-Output "Installing RD Infra Agent on VM $AgentInstaller`n"
$agent_deploy_status = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList "/i $WVDAgentInstaller", `
        "/quiet", `
        "/qn", `
        "/norestart", `
        "/passive", `
        "REGISTRATIONTOKEN=$RegistrationToken", "/l* $LocalAVDpath\AgentInstall.txt" `
    -Wait `
    -Passthru
Add-Content -LiteralPath C:\New-AVDSessionHost.log "AVD Agent Install Complete"
Wait-Event -Timeout 5

<#
#########################
#    FSLogix Install    #
#########################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Installing FSLogix"
$fslogix_deploy_status = Start-Process `
    -FilePath "$LocalAVDpath\FSLogix\x64\Release\FSLogixAppsSetup.exe" `
    -ArgumentList "/install /quiet" `
    -Wait `
    -Passthru

#######################################
#    FSLogix User Profile Settings    #
#######################################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Configure FSLogix Profile Settings"
Push-Location 
Set-Location HKLM:\SOFTWARE\
New-Item `
    -Path HKLM:\SOFTWARE\FSLogix `
    -Name Profiles `
    -Value "" `
    -Force
New-Item `
    -Path HKLM:\Software\FSLogix\Profiles\ `
    -Name Apps `
    -Force
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "Enabled" `
    -Type "Dword" `
    -Value "1"
New-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "CCDLocations" `
    -Value "type=smb,connectionString=$ProfilePath" `
    -PropertyType MultiString `
    -Force
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "SizeInMBs" `
    -Type "Dword" `
    -Value "30000"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "IsDynamic" `
    -Type "Dword" `
    -Value "1"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "VolumeType" `
    -Type String `
    -Value "vhdx"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "FlipFlopProfileDirectoryName" `
    -Type "Dword" `
    -Value "1" 
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "SIDDirNamePattern" `
    -Type String `
    -Value "%username%%sid%"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "SIDDirNameMatch" `
    -Type String `
    -Value "%username%%sid%"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name DeleteLocalProfileWhenVHDShouldApply `
    -Type DWord `
    -Value 1
Pop-Location



##########################################
#    Enable Screen Capture Protection    #
##########################################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Enable Screen Capture Protection"
Push-Location 
Set-Location "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
New-ItemProperty `
    -Path .\ `
    -Name fEnableScreenCaptureProtection `
    -Value "1" `
    -PropertyType DWord `
    -Force
Pop-Location


##############################
#    Enable Azure AD Join    #
##############################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Enable Azure AD Join"
Push-Location 
Set-Location HKLM:\SOFTWARE\Microsoft
New-Item `
    -Path HKLM:\SOFTWARE\Microsoft `
    -Name RDInfraAgent `
    -Force
New-Item `
    -Path HKLM:\Software\Microsoft\RDInfraAgent `
    -Name AADJPrivate `
    -Force
Pop-Location
#>
<#
##############################################
#    AVD Optimizer (Virtual Desktop Team)    #
##############################################
If ($Optimize -eq $true) {  
    Write-Output "Optimizer selected"  
    ################################
    #    Download AVD Optimizer    #
    ################################
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Optimize Selected"
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Creating C:\Optimize folder"
    New-Item -Path C:\ -Name Optimize -ItemType Directory -ErrorAction SilentlyContinue
    $LocalPath = "C:\Optimize\"
    $AVDOptimizeURL = 'https://github.com/The-Virtual-Desktop-Team/Virtual-Desktop-Optimization-Tool/archive/refs/heads/main.zip'
    $AVDOptimizeInstaller = "Windows_10_VDI_Optimize-master.zip"
    Invoke-WebRequest `
        -Uri $AVDOptimizeURL `
        -OutFile "$Localpath$AVDOptimizeInstaller"


    ###############################
    #    Prep for AVD Optimize    #
    ###############################
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Optimize downloaded and extracted"
    Expand-Archive `
        -LiteralPath "C:\Optimize\Windows_10_VDI_Optimize-master.zip" `
        -DestinationPath "$Localpath" `
        -Force `
        -Verbose



    #################################
    #    Run AVD Optimize Script    #
    #################################
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Begining Optimize"
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force -Verbose
    .\Win10_VirtualDesktop_Optimize.ps1 -Optimizations $Optimizations -Restart -AcceptEULA -Verbose
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Optimization Complete"
}
else {
    Write-Output "Optimize not selected"
    Add-Content -LiteralPath C:\New-AVDSessionHost.log "Optimize NOT selected"    
}
#>
Stop-Transcript
##########################
#    Restart Computer    #
##########################
Add-Content -LiteralPath C:\New-AVDSessionHost.log "Process Complete - REBOOT"
Restart-Computer -Force 

