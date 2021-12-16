Clear-Host

Import-Module .\Modules\Add-ADObject.psm1
Import-Module .\Modules\Edit-ADObject.psm1
Import-Module .\Modules\Remove-ADObject.psm1

# Adding colors to output to help differentiate between messages
$Global:FGTitle = [System.ConsoleColor]::Magenta
$Global:FGLabel = [System.ConsoleColor]::Cyan
$Global:FGHeader = [System.ConsoleColor]::White
$Global:FGSuccess = [System.ConsoleColor]::Green
$Global:FGNotice = [System.ConsoleColor]::Yellow
$Global:FGWarning = [System.ConsoleColor]::Red
$Global:BGBold = [System.ConsoleColor]::Black
$Global:FGInput = [System.ConsoleColor]::Black
$Global:BGInput = [System.ConsoleColor]::Gray

# Pull list of available Distribution Points from MEMCM Site Server. If MEMCM Console not installed, use hardcoded list of servers
if ($ENV:SMS_ADMIN_UI_PATH -eq $null) {
  $SMS_ADMIN_UI_PATH = Get-ChildItem -Path "C:\Program Files (x86)" -Include *ConfigurationManager.psd1* -File -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Directory
  if ($SMS_ADMIN_UI_PATH.Count -gt 1) {
    $Global:SMS_ADMIN_UI_PATH = $SMS_ADMIN_UI_PATH | Sort-Object Length -Descending | Select-Object -First 1
  }
}
else {
  $Global:SMS_ADMIN_UI_PATH = "$ENV:SMS_ADMIN_UI_PATH\.."
}

if ($Global:SMS_ADMIN_UI_PATH -eq $null) {
  $Global:DistributionPoints = @(
    "\\SCCM-01.EDUBEAR.NET",
    "\\SCCM-DP-CHHS-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-CNAS-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-COAL-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-COB-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-COMB-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-COMN-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-CORE-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-ITAC-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-LICN-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-OUTR-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-REDIP01.MISSOURISTATE.EDU",
    "\\SCCM-DP-RLIFE01.MISSOURISTATE.EDU",
    "\\SCCM-DP-SCUF-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-ULIB-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-USIT-01.MISSOURISTATE.EDU",
    "\\SCCM-DP-WP-01.MISSOURISTATE.EDU"
  )
}
else {
  if((Get-Module ConfigurationManager) -eq $null) {
      Import-Module "$($Global:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"
  }
  if((Get-PSDrive -Name "MSU" -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
      New-PSDrive -Name "MSU" -PSProvider CMSite -Root "SCCM-01.MISSOURISTATE.EDU"
      Clear-Host
  }
  $Global:LocalLocation = Get-Location
  Set-Location "MSU:\"
  $Global:DistributionPoints = Get-CMDistributionPoint | Sort-Object -Property NetworkOSPath | Select-Object -ExpandProperty NetworkOSPath
  Set-Location $LocalLocation
  Remove-PSDrive -Name "MSU" -Force
}

# Make sure the local machine has RSAT installed and can use the ActiveDirectory module.
if (Get-Module -ListAvailable -Name ActiveDirectory) {
  Import-Module ActiveDirectory
}
else {
  Write-Host "`nYou do not have Active Directory on this machine." -ForegroundColor $FGWarning -BackgroundColor $BGBold
  Write-Host "This script heavily relies on the Active Directory Module and cannot function without it." -ForegroundColor $FGWarning -BackgroundColor $BGBold
  Write-Host "Please install the Remote Server Administration Tools (RSAT) package on your machine to use this script.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
  Write-Host "Press Enter to quit.`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
  Read-Host
  Exit
}

$Global:StartingOU = "OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
# $Global:StartingOU = "OU=CORE,OU=LABS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
$Global:CurrentLocation = Get-Location | Select-Object -ExpandProperty Path
$Global:LogFilePath = $CurrentLocation + "\ADComputer_Log.CSV"
$Global:ComputersAdded = @()
$Global:ComputersEdited = @()
$Global:ComputersDeleted = @()

# Array of words Microsoft won't allow a computer name to be, AKA Reserved Words
$Global:ReservedWords = @(
  "ANONYMOUS","AUTHENTICATED USERS","BATCH","BUILTIN","CREATOR GROUP","CREATOR GROUP SERVER","CREATOR OWNER",
  "CREATOR OWNER SERVER","DIALUP","DIGEST AUTH","INTERACTIVE","INTERNET","LOCAL","LOCAL SYSTEM","NETWORK",
  "NETWORK SERVICE","NT AUTHORITY","NT DOMAIN","NTLM AUTH","NULL","PROXY","REMOTE INTERACTIVE","RESTRICTED",
  "SCHANNEL AUTH","SELF","SERVER","SERVICE","SYSTEM","TERMINAL SERVER","THIS ORGANIZATION","USERS","WORLD"
)

# Regular Expression representing the different formats a MAC address is traditionally represented in
$Global:ValidMACPatterns = @(
  '^([0-9a-f]{2}:){5}([0-9a-f]{2})$'
  '^([0-9a-f]{2}-){5}([0-9a-f]{2})$'
  '^([0-9a-f]{2} ){5}([0-9a-f]{2})$'
  '^([0-9a-f]{2}\.){5}([0-9a-f]{2})$'
  '^([0-9a-f]{12})$'
)

# The needed headers for the CSV file when importing Computer Objects from a CSV file
$Global:CorrectCSVHeaders = @(
  "COMPUTERNAME",
  "COMPUTERMAC",
  "COMPUTERDESC",
  "COMPUTEROU",
  "DEPLOYMENTSERVER"
)

# Write-Host "`nMissouri State University Active Directory Suite" -ForegroundColor $FGTitle -BackgroundColor $BGBold
# Write-Host "This script will allow you to add, edit, or delete Computer AD objects the SGF.EDUBEAR.NET domain."
# Write-Host "To get started, please select one of the options listed below."
#
# Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
# Write-Host "- Enter '1' to ADD Computers to the domain."
# Write-Host "- Enter '2' to EDIT an existing Computer on the domain."
# Write-Host "- Enter '3' to REMOVE an existing Computer from the domain."
# Write-Host "- Press Enter to QUIT.`n"

# Logic for handling user's input
[bool]$Valid = $False
do {
  # Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
  # $Selection = Read-Host
  $Selection = 1
  if ($Selection -eq "") {
    Exit
  }
  elseif ($Selection -eq 1) {
    $Valid = $True
    Add-ADObject
  }
  elseif ($Selection -eq 2) {
    $Valid = $True
    Edit-ADObject
  }
  elseif ($Selection -eq 3) {
    $Valid = $True
    Remove-ADObject
  }
  else {
    Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
  }
} while ($Valid -eq $False)
