<# 
.DESCRIPTION 
    Author: Barbara Forbes
    Version: 1.0 
 
    Install the latest version of Powershell Core on your machine, so you don't have to click through the setup 
 
.PARAMETER bit32
    Switch, Optional: If you want to use x84Version instead of x64, call this parameter. Default is x64
.PARAMETER Downloadlocation 
    String,Optional: Default download location is the location the script is started in. Logfile will be written to downloadlocation as well. 
.EXAMPLE 
    .\install-PScore.ps1 -Downloadlocation C:\temp\ -bit32 
    This command will download the x64-msi-file to c:temp and leave the logfile there. 
.NOTES 
    Must run as admin
    This runs in Powershell 5.1. Earlier verions have not been tested yet. 
    Used https://kevinmarquette.github.io/2016-10-21-powershell-installing-msi-files/ for the MSIexec-part

#>

[CmdletBinding()]
param(
    [parameter()]
    [switch]$bit32,

    [Parameter()]
    [string]$downloadlocation = $PSScriptRoot

)

write-Output "Script is starting"
#set TLS to 1.2 instead of 1.0
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#get filename of latest version through the URLs in the Releasepage
$url = "https://github.com/PowerShell/PowerShell/releases"
[string]$urlstring = (Invoke-WebRequest -Uri $url).content

#check for bit32-parameter and create right regex-pattern
if ($bit32){
    Write-Verbose "x86 version was requested"
    $regex = [regex] 'PowerShell-\d.\d.\d-win-x86.msi'
}
else{
    $regex = [regex] 'PowerShell-\d.\d.\d-win-x64.msi'
}

#match regex to filename
$urlstring -match $regex | out-null
$filename = $matches[0]
Write-Output "Last version found: $filename"

#check downloadlocation. If user has not typed an \ at the end, add it.
if ($downloadlocation -notlike "*\"){
    $downloadlocation = $downloadlocation+"\"
}
#check if downloadlocation exists. If not, create it.
if (-not (Test-Path $downloadlocation) ){
    Write-Verbose "Filelocation does not exist, it is being created"
    New-Item -Path $downloadlocation -ItemType Directory
}


#download file
$downloadurl = "$url/download/v6.1.0/$filename"
$output = $downloadlocation+$filename
Write-Verbose "starting download"
Invoke-WebRequest -Uri $downloadurl -OutFile $output
Write-Output "download is finished"

#change downloadlocation to right path
Set-Location $downloadlocation


#install msi and create log for errors
$Date = Get-Date -Format yyyyMMddTHHmmss
$logFile = '{0}-{1}.log' -f $filename,$Date
$MSIArguments = @(
    "/i"
    ('"{0}"' -f $filename)
    "/qn"
    "/norestart"
    "/L*v"
    $logFile
)
#start installation
Write-Verbose "starting installation"
Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow 
Write-Output "installation has finished"

#check if installation succeeded 
if ($bit32){
    try{ 
        Get-Item "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{E3F4E1B9-A800-4C4A-BFA4-8891384E8D48}" -ErrorAction stop  | out-null
    }
    catch{
        Write-Output "Something went wrong, Powershell Core x86 was not installed. Please check the log for errors"
        return
    }
Write-Output "installation has succeeded, you can find powershell Core x86 in the start menu"
}
else {
    try{ 
        Get-Item "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{5B7A41F8-E132-45BE-92D5-48543F89372F}" -ErrorAction stop | out-null
    }
    catch{
        Write-Output "Something went wrong, Powershell Core x64 was not installed. Please check the log for errors"
        return
    }
Write-Output "installation has succeeded, you can find powershell Core x64 in the start menu"
}   
