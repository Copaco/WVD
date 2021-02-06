##############################
#    WVD Script Functions    #
##############################

#download latest M365 apps
function Get-ODTUri {
    <#
        .SYNOPSIS
            Get Download URL of latest Office 365 Deployment Tool (ODT).
        .NOTES
        
        .LINK
           
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param ()

    $url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
    }
    catch {
        Throw "Failed to connect to ODT: $url with error $_."
        Break
    }
    finally {
        $ODTUri = $response.links | Where-Object {$_.outerHTML -like "*click here to download manually*"}
        Write-Output $ODTUri.href
    }
}

##############################
#    WVD Script Parameters   #
##############################
if (Test-Path c:\WVDAppInstallation.log){
    try {
    Remove-Item -Path C:\WVDAppInstallation.log -Force 
    }
    catch{
    
    }
   }

New-Item -Path c:\ -Name WVDAppInstallation.log -ItemType File


if (Test-Path c:\temp){
    try {
    Remove-Item -Path c:\temp -Force -Recurse
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Removed c:\temp"
    }
    catch{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Failed to remove c:\temp"
    }
   }

if (Test-Path c:\Optimize){
    try {
    Remove-Item -Path C:\Optimize -Force -Recurse
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Removed c:\Optimize"
    }
    catch{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Failed to remove c:\Optimize"
    }
   }



######################
#    WVD Variables   #
######################
$erroractionpreference    = 'silentlycontinue'
$LocalWVDpath            = "c:\temp\wvd\"
$ODTpath                 = "c:\temp\wvd\ODT"
$FSLogixURI              = 'https://aka.ms/fslogix_download'
$FSInstaller             = 'FSLogixAppsSetup.zip'
$NotePadURI              = 'https://notepad-plus-plus.org/repository/7.x/7.3.3/npp.7.3.3.Installer.exe'
$VScodeURI               = 'https://go.microsoft.com/fwlink/?Linkid=852157'
$TeamsURI                = 'https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true'
$TeamWebSocket           = 'https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RE4AQBt'
$OneDriveURI             = 'https://go.microsoft.com/fwlink/?linkid=844652'
$ProfilePath             = '\\copfslogix.file.core.windows.net\fslogix'



####################################
#    Test/Create Temp Directory    #
####################################
if((Test-Path c:\temp) -eq $false) {
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Create C:\temp Directory"
    Write-Host `
        -ForegroundColor Cyan `
        -BackgroundColor Black `
        "creating temp directory"
    New-Item -Path c:\temp -ItemType Directory
}
else {
    Add-Content -LiteralPath C:\WVDAppInstallation.log "C:\temp Already Exists"
    Write-Host `
        -ForegroundColor Yellow `
        -BackgroundColor Black `
        "temp directory already exists"
}
if((Test-Path $LocalWVDpath) -eq $false) {
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Create C:\temp\WVD Directory"
    Write-Host `
        -ForegroundColor Cyan `
        -BackgroundColor Black `
        "creating c:\temp\wvd directory"
    New-Item -Path $LocalWVDpath -ItemType Directory
}
else {
    Add-Content -LiteralPath C:\WVDAppInstallation.log "C:\temp\WVD Already Exists"
    Write-Host `
        -ForegroundColor Yellow `
        -BackgroundColor Black `
        "c:\temp\wvd directory already exists"
}

if((Test-Path $ODTpath) -eq $false) {
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Create C:\temp\WVD\ODT Directory"
    Write-Host `
        -ForegroundColor Cyan `
        -BackgroundColor Black `
        "creating c:\temp\wvd\odt directory"
    New-Item -Path $ODTpath -ItemType Directory
}
else {
    Add-Content -LiteralPath C:\WVDAppInstallation.log "C:\temp\WVD\ODT Already Exists"
    Write-Host `
        -ForegroundColor Yellow `
        -BackgroundColor Black `
        "c:\temp\wvd\odt directory already exists"
}

#################################
#    Download WVD Componants    #
#################################

Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading FSLogix"
    Invoke-WebRequest -Uri $FSLogixURI -OutFile "$LocalWVDpath$FSInstaller"

Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading NotePad++"
    Invoke-WebRequest -Uri $NotePadURI -OutFile "$LocalWVDpath\notepadplusplus.exe"

Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading VSCode"
    Invoke-WebRequest -Uri $VScodeURI -OutFile "$LocalWVDpath\VScode.exe"

Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading Teams"
    Invoke-WebRequest -Uri $TeamsURI -OutFile "$LocalWVDpath\Teams.msi"

Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading Teams WebSocket"
    Invoke-WebRequest -Uri $TeamWebSocket -OutFile "$LocalWVDpath\TeamsWebSocket.msi"

Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading OneDrive"
    Invoke-WebRequest -Uri $OneDriveURI -OutFile "$LocalWVDpath\OneDrive.exe"
    

#########################
#    FSLogix Install    #
#########################
Add-Content -LiteralPath C:\WVDAppInstallation.log "Unzip FSLogix"
Expand-Archive `
    -LiteralPath "C:\temp\wvd\$FSInstaller" `
    -DestinationPath "$LocalWVDpath\FSLogix" `
    -Force `
    -Verbose
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
cd $LocalWVDpath 
Add-Content -LiteralPath C:\WVDAppInstallation.log "UnZip FXLogix Complete"

Add-Content -LiteralPath C:\WVDAppInstallation.log "Installing FSLogix"
$fslogix_deploy_status = Start-Process `
    -FilePath "$LocalWVDpath\FSLogix\x64\Release\FSLogixAppsSetup.exe" `
    -ArgumentList "/install /quiet /norestart" `
    -Wait `
    -Passthru

#########################
#    FSLogix Configuration#
#########################
Add-Content -LiteralPath C:\WVDAppInstallation.log "Configure FSLogix Profile Settings"
Push-Location 
Set-Location HKLM:\SOFTWARE\
New-Item `
    -Path HKLM:\SOFTWARE\FSLogix `
    -Name Profiles `
    -Value "" `
    -Force
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "Enabled" `
    -Type "Dword" `
    -Value "1"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "DeleteLocalProfileWhenVHDShouldApply" `
    -Type "Dword" `
    -Value "1"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "IsDynamic" `
    -Type "Dword" `
    -Value "1"
    Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "KeepLocalDir" `
    -Type "Dword" `
    -Value "0"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "KeepLocalDir" `
    -Type "Dword" `
    -Value "0"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "ProfileType" `
    -Type "Dword" `
    -Value "0"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "SetTempToLocalPath" `
    -Type "Dword" `
    -Value "3"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "SizeInMBs" `
    -Type "Dword" `
    -Value "7530"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "VHDLocations" `
    -Type String `
    -Value $ProfilePath
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "VHDNameMatch" `
    -Type String `
    -Value "%userdomain%-%username%"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "VHDNamePattern" `
    -Type String `
    -Value "%userdomain%-%username%"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "VolumeType" `
    -Type String `
    -Value "VHDX"
Set-ItemProperty `
    -Path HKLM:\Software\FSLogix\Profiles `
    -Name "FlipFlopProfileDirectoryName" `
    -Type "Dword" `
    -Value "1" 
Pop-Location


#########################
#    Install VSCode     #
#########################

Add-Content -LiteralPath C:\WVDAppInstallation.log "Installing VsCode"
$vscode_deploy_status = Start-Process `
    -FilePath "$LocalWVDpath\VScode.exe" `
    -ArgumentList "/verysilent" 

#Start sleep
Start-Sleep -Seconds 40

#stop VScode
Get-Process -Name Code | Stop-Process -Force

#########################
#    NotePad ++         #
#########################

Add-Content -LiteralPath C:\WVDAppInstallation.log "Installing NotePad ++"
$Notepad_deploy_status = Start-Process `
    -FilePath "$LocalWVDpath\notepadplusplus.exe" `
    -ArgumentList "/S" 

#Start sleep
Start-Sleep -Seconds 10


#########################
#    M365 Apps          #
#########################

#download ODT 
$URL = $(Get-ODTUri)
Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading latest version Office 365 Deployment Tool"
Set-Location $ODTpath
Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile .\officedeploymenttool.exe
Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloaded latest version Office 365 Deployment Tool"

#extract ODT
$Version = (Get-Command .\officedeploymenttool.exe).FileVersionInfo.FileVersion

.\officedeploymenttool.exe /quiet /extract:.\$Version
start-sleep -s 5

#Installing Office
Set-Location .\$Version


$Unattendedxmluri = "https://raw.githubusercontent.com/Copaco/WVD/main/M365apps4business.xml"
$Unattendedxml = "M365apps4business.xml"
Invoke-WebRequest -Uri $Unattendedxmluri -OutFile $Unattendedxml -UseBasicParsing


$ODTDownload = "/download $Unattendedxml"
$ODTInstallation = "/configure $Unattendedxml"


try{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Downloading installation files installation M365 Apps"
    $DownloadExitCode = (Start-Process -FilePath setup.exe -ArgumentList $ODTDownload -Wait -PassThru).ExitCode
}catch{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Failed to download M365 Apps"
}
if($DownloadExitCode -eq 0){
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Sucessfully Downloaded M365 Apps"
}
else{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Error during download of"
}

try{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Starting with installation of M365 Apps"
    $InstallExitCode = (Start-Process -FilePath .\setup.exe -ArgumentList $ODTInstallation -Wait -PassThru).ExitCode
}catch{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Failed to install M365 Apps"
}
if($InstallExitCode -eq 0){
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Sucessfully Downloaded M365 Apps"
}
else{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Error during download of"
}


start-sleep -s 10
#########################
#    Teams  w optimalization#
#########################


# set required registry
$testRegistry = Test-Path 'HKLM:\SOFTWARE\Microsoft\Teams'

if ($testRegistry -eq $false) {
    try{
    New-Item -Path 'HKLM:\SOFTWARE\Microsoft\' -Name Teams
    New-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Teams' -Name "IsWVDEnvironment" -Value 1 -PropertyType "DWORD" -Force
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Added WVD Registry Key"
    }
    catch{
        Add-Content -LiteralPath C:\WVDAppInstallation.log "Unable to set Team WVD Registry Key"
    }
}
else{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "WVD Registry Key Already present"
}

Set-Location $LocalWVDpath

#Teams Audio Optimalization
$TeamsWebSocketArugment = '/q'
$TeamsWebSocketInstaller = "$LocalWVDpath\TeamsWebSocket.msi"

try{
    $TeamsWebSocketInstallationOutput = Start-Process -FilePath $TeamsWebSocketInstaller -ArgumentList $TeamsWebSocketArugment -Wait -PassThru
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Successfully installed Teams Websocket"
}
catch{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Unsuccessfully installed Teams Websocket"
    Add-Content -LiteralPath C:\WVDAppInstallation.log "$    $TeamsWebSocketInstallationOutput = Start-Process -FilePath $Teamsinstaller -ArgumentList $TeamsinstallArguments -Wait -PassThru
    "
}

#start installation Teams
$TeamsinstallArguments = 'OPTIONS="noAutoStart=true" ALLUSERS=1 ALLUSER=1'
$Teamsinstaller = "$LocalWVDpath\teams.msi"

try{
    $TeamsInstallationOutput = Start-Process -FilePath $Teamsinstaller -ArgumentList $TeamsinstallArguments -Wait -PassThru
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Successfully installed Teams"
}
catch{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Unsuccessfully installed Teams"
    Add-Content -LiteralPath C:\WVDAppInstallation.log "$TeamsInstallationOutput"
}


#########################
#    Onedrive           #
#########################

#Start installation OneDrive
$OneDriveArguments = "/ALLUSERS"
$OndriveInstaller = "$LocalWVDpath\OneDrive.exe"

try{
    $OneDriveInstallationOutput = Start-Process -FilePath $OndriveInstaller -ArgumentList $OneDriveArguments 
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Successfully installed OneDrive"
}
catch{
    Add-Content -LiteralPath C:\WVDAppInstallation.log "Unsuccessfully installed OneDrive"
    Add-Content -LiteralPath C:\WVDAppInstallation.log "$OneDriveInstallationOutput"
}

Start-Sleep -Seconds 120
#########################
#    CleanUP           #
#########################

# Clean up buildArtifacts directory
#Remove-Item -Path "C:\buildArtifacts\*" -Force -Recurse

# Delete the buildArtifacts directory
#Remove-Item -Path "C:\buildArtifacts" -Force

# Clean up temp directory
#Remove-Item -Path "C:\temp\*" -Force -Recurse

Add-Content -LiteralPath C:\WVDAppInstallation.log "Finished installing all components"
return 0
