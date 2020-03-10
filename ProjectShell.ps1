###############
### Section Parameters
### These params can be specified in the command line
###############

param (
    # $instance "dev"
    # To which stage do we want to deploy
    [string]$instance = "dev"
)

###############
### Section Variables
### Set urls to the Dev and Prod Sites
###############

$SharePointDevSiteUrl = "https://<tenant>.sharepoint.com/sites/XXX"
$SharePointProdSiteUrl = "https://<tenant>.sharepoint.com/sites/XXX"


###############
### Section Modules
### Install and import Powershell modules
###############

# Make all npm project specific commands avaliable in Path
$env:Path = "$($PSScriptRoot)\node_modules\.bin;" + $env:Path

# Install Powershell modules if not exists
if (-not (Test-Path "$($PSScriptRoot)\ps_modules" -PathType Container) ) {
    New-Item "$($PSScriptRoot)\ps_modules" -ItemType Directory
    # Download additional modules 
    Save-Module -Name SharePointPnPPowerShellOnline -Path "$($PSScriptRoot)\ps_modules" 
}

# Import addtional modules
Import-Module "$($PSScriptRoot)\ps_modules\SharePointPnPPowerShellOnline" -DisableNameChecking


###############
### Prepare for history data
### 
###############

if (-not (Test-Path "$($PSScriptRoot)\history" -PathType Container) ) {
    New-Item "$($PSScriptRoot)\history" -ItemType Directory
}


###############
### Section Connection
### Connect to the SharePoint Instance
###############

if ($instance -eq "dev") {
    Connect-PnPOnline -Url $SharePointDevSiteUrl -UseWebLogin
    # Register History
    Register-EngineEvent PowerShell.Exiting -SupportEvent -Action {
        "# $($env:UserName) $(get-date)" | out-file  "$($PSScriptRoot)\history\dev.ps1.log" -append 
        get-history | select CommandLine | ft -HideTableHeaders | out-file  "$($PSScriptRoot)\history\dev.ps1.txt" -append 
    } | out-null
    
}
else {
    $instance = "prod"
    Connect-PnPOnline -Url $SharePointProdSiteUrl -UseWebLogin
    Register-EngineEvent PowerShell.Exiting -SupportEvent -Action {
        "# $($env:UserName) $(get-date)" | out-file  "$($PSScriptRoot)\history\prod.ps1.txt" -append 
        get-history | select CommandLine | ft -HideTableHeaders | out-file  "$($PSScriptRoot)\history\prod.ps1.log" -append 
    } | out-null
}
Set-PnPTraceLog -On -Level Debug



###############
### History
### Register persistent powershell handlers
###############

# Register Powershell history Path


###############
### Section Window
### Format the appearance of the powershell window
###############

$web = Get-PnPWeb
$title = "[$instance shell] $($web.ServerRelativeUrl)"
if ($watch -eq $true) {
    $title = "[$instance watch] $($web.ServerRelativeUrl)"
}
$host.ui.RawUI.WindowTitle = $title
Write-Host $title
