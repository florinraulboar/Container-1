<#
.SYNOPSIS
  <Overview of script>
.DESCRIPTION
  <Brief description of script>
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
.NOTES
  Version:        1.0
  Author:         <Name>
  Creation Date:  <Date>
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
#$ErrorActionPreference = "SilentlyContinue"

param(
     [Parameter(
        Mandatory=$True, 
        Position = 0,
        HelpMessage="Enter the name of your Office 365 organization, example: contosotoycompany")]
     [ValidateNotNullorEmpty()]
     [string]$orgName
    )


#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
#$sScriptVersion = "1.0"

#Log File Info

#Check if path already exists
OutputInYellow "Creating path C:\Users\Public\SharePoint_LogFiles"
If(Test-Path -Path 'C:\Users\Public\SharePoint_LogFiles'){OutputInYellow "Path already exists!"}
Else {New-Item -Path 'C:\Users\Public\SharePoint_LogFiles' -ItemType Directory;
      OutputInGreen "Created!"} #Create a folder for SharePoint's LogFiles

#Check if files already exists
# If file doesn't exists => create file

#1.LogFile_DisallowInfectedFileDownload.txt
OutputInYellow "Creating LogFile_DisallowInfectedFileDownload.txt"
$fileToCheck = "C:\Users\Public\SharePoint_LogFiles\LogFile_DisallowInfectedFileDownload.txt"
if(Test-Path $fileToCheck -PathType leaf) {OutputInYellow "File already exists!"}
else {New-Item C:\Users\Public\SharePoint_LogFiles\LogFile_DisallowInfectedFileDownload.txt; 
      Set-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_DisallowInfectedFileDownload.txt' -Value 'LogFile Informations--------';  
      OutputInGreen "Created!"} 


#2.LogFile_MonitorExternalSharingInvitations.txt
OutputInYellow "Creating LogFile_MonitorExternalSharingInvitations.txt"
$fileToCheck = "C:\Users\Public\SharePoint_LogFiles\LogFile_MonitorExternalSharingInvitations.txt"
if(Test-Path $fileToCheck -PathType leaf) {OutputInYellow "File already exists!"}
else {New-Item C:\Users\Public\SharePoint_LogFiles\LogFile_MonitorExternalSharingInvitations.txt; 
      Set-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_MonitorExternalSharingInvitations.txt' -Value 'LogFile Informations--------';  
      OutputInGreen "Created!"} # If file doesn't exist => create file


#3.SetExpirationForSharePointOnlineAnonymousLinks.txt
OutputInYellow "Creating LogFile_SetExpirationForSharePointOnlineAnonymousLinks.txt"
$fileToCheck = "C:\Users\Public\SharePoint_LogFiles\LogFile_SetExpirationForSharePointOnlineAnonymousLinks.txt"
if(Test-Path $fileToCheck -PathType leaf) {OutputInYellow "File already exists!"}
else {New-Item C:\Users\Public\SharePoint_LogFiles\LogFile_SetExpirationForSharePointOnlineAnonymousLinks.txt; 
      Set-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_MonitorExternalSharingInvitations.txt' -Value 'LogFile Informations--------';
      OutputInGreen "Created!"} # If file doesn't exist => create file



#-----------------------------------------------------------[Functions]------------------------------------------------------------

#Set green color for output
function OutputInGreen
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Text
    )

    Write-Host -ForegroundColor Green "$Text`n"
}

#Set red color for output
function OutputInRed
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Text
    )

    Write-Host -ForegroundColor Red "$Text`n"
}

#Set yellow color for output
function OutputInYellow
{

    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Text
    )

    Write-Host -ForegroundColor Yellow "$Text`n"
}


#Connection for the SharePoint tenant
Function Connect-Tenant{
    $siteUrl = "https://$orgName-admin.sharepoint.com"
    Write-Host "$(Get-Date) connect to tenant admin $($siteUrl)"
    Connect-SPOService -Url $siteUrl 
    OutputInGreen "Connected!"
}

#Set DisallowInfectedFileDownload
Function DisallowInfectedFileDownload{
  $Activate_DisallowDownload = Read-Host "Set DisallowInfectedFileDownload (True or False)"
  if ($Activate_DisallowDownload -eq "True"){Set-SPOTenant -DisallowInfectedFileDownload $True;
      Add-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_DisallowInfectedFileDownload.txt' -Value "$orgName => set DisallowInfectedFileDownload -$Activate_DisallowDownload => [$(Get-Date)])"}
  elseif ($Activate_DisallowDownload -eq "False"){Set-SPOTenant -DisallowInfectedFileDownload $False;
      Add-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_DisallowInfectedFileDownload.txt' -Value "$orgName => set DisallowInfectedFileDownload -$Activate_DisallowDownload => [$(Get-Date)])"}
  else {OutputInRed "***No matches!***"} 
}

#Set MonitorExternalSharingInvitations
Function MonitorExternalSharingInvitations{
  $Activate_SharingInvitations = Read-Host "Set BccExternalSharingInvitations (True or False)"
  if ($Activate_SharingInvitations -eq "True") {
  $Mailbox_List = Read-Host "Set BccExternalSharingInvitationsList to the address of your mailbox";
  Set-SPOTenant -BccExternalSharingInvitations $True -BccExternalSharingInvitationsList $Mailbox_List;
  Add-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_MonitorExternalSharingInvitations.txt' -Value "$orgName => set MonitorExternalSharingInvitations -$Activate_SharingInvitations ; mailbox: $Mailbox_List => [$(Get-Date)]"}
  elseif ($Activate_SharingInvitations -eq "False"){Set-SPOTenant -BccExternalSharingInvitations $False -BccExternalSharingInvitationsList "{}";
  Add-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_MonitorExternalSharingInvitations.txt' -Value "$orgName => set MonitorExternalSharingInvitations -$Activate_SharingInvitations ; mailbox: $Mailbox_List => [$(Get-Date)]"}
  else {OutputInRed "***No matches!***"}

}

#Set SetExpirationForSharePointOnlineAnonymousLinks
Function SetExpiration{
  $Days_to_expire = Read-Host "Set expire in days [30]"
  Set-SPOTenant -SharingCapability ExternalUserAndGuestSharing -RequireAnonymousLinksExpireInDays $Days_to_expire
  Add-Content -Path 'C:\Users\Public\SharePoint_LogFiles\LogFile_SetExpirationForSharePointOnlineAnonymousLinks.txt' -Value "$orgName => Set days to expire to $Days_to_expire => [$(Get-Date)]"
}


Function SharePointPolicy{
 
$Intrebare = Read-Host "Do you want to set a policy? Y or N"

if ($Intrebare -eq "Y"){$Intrebare = Read-Host "Choose the number of the policy you want to set: 1.DisallowInfectedFileDownload 2.MonitorExternalSharingIvitations 3.SetExpirationForSharePointOnlineAnonymousLinks"}

Switch ($Intrebare){
  "1"{DisallowInfectedFileDownload 
      OutputInYellow "You have chosen to set a policy for SharePointOnline named DisallowInfectedFileDownload..."
      OutputInGreen "Done! [$(Get-Date)]"
      SharePointPolicy}
  "2"{MonitorExternalSharingInvitations 
      OutputInYellow "You have chosen to set a policy for SharePointOnline named MonitorExternalSharingInvitations..."
      OutputInGreen "Done! [$(Get-Date)]"
      SharePointPolicy}
  "3"{SetExpiration 
      OutputInYellow "You have chosen to set a policy for SharePointOnline named SetExpirationForSharePointOnlineAnonymousLinks..."
      OutputInGreen "Done! [$(Get-Date)]"
      SharePointPolicy}
  "4"{Write-Host"Exit"}
  Default {OutputInRed "***No matches!***" }
  }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#-ScriptVersion $sScriptVersion

try{ 
    Connect-Tenant
}
catch { 
    write-host $_.Exception.Message
}

try{ 
    SharePointPolicy
}
catch { 
    write-host $_.Exception.Message
}