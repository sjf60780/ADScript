Clear-Host

# Function to color outputs
function Global:Format-Color([hashtable]$Colors = @{}, [switch]$SimpleMatch)
{
	$Lines = ($input | Out-String) -replace "`r", "" -split "`n"
	ForEach($Line in $Lines)
  {
		$Color = ''
		ForEach($Pattern in $Colors.Keys)
    {
			if(!$SimpleMatch -and $Line -match $Pattern) { $Color = $Colors[$Pattern] }
			elseif ($SimpleMatch -and $Line -like $Pattern) { $Color = $Colors[$Pattern] }
		}
		if($Color)
    {
			Write-Host -ForegroundColor $Color $Line
		}
    else
    {
			Write-Host $Line
		}
	}
}

function Global:Show-ADForest([string]$Title = "", [string]$Mode = "", [string]$CurrentOU = "", [switch]$ShowComputers)
{
	$OUPath = $StartingOU
	[bool]$Done = $FALSE
	do
	{
		Clear-Host
		Write-Host "`n$($Title)" -ForegroundColor $FGTitle -BackgroundColor $BGBold
		if ($CurrentOU)
		{
			Write-Host "`nComputer Object's Current OU Path: " -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
			Write-Host $CurrentOU
		}
		Write-Host "`nDefault AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
		Write-Host "`t$StartingOU"
		if ($OUPath -ne $StartingOU)
		{
			Write-Host "Parent  AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
			Write-Host "`t$ParentOU"
		}
		Write-Host "Current AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
		Write-Host "`t$OUPath`n"

		# List all (if any) computer objects in current location
		if ($ShowComputers)
		{
			[Array]$ComputerObjects = Get-ADComputer -Filter 'Name -like "*"' -SearchBase $OUPath -SearchScope OneLevel -Properties Description
			if (!$ComputerObjects)
			{
				Write-Host "This OU does not contain any computer objects.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
			}
			else
			{
				Write-Host "List of Computer Objects" -ForegroundColor $FGHeader -BackgroundColor $BGBold
				ForEach ($Computer in $ComputerObjects)
				{
					Write-Host "$($Computer.Name)"
				}
				Write-Host
			}
		}

		# Get all Child OUs from current OU
		Write-Host "List of Child OUs" -ForegroundColor $FGHeader -BackgroundColor $BGBold -NoNewLine
		[Array]$ChildOUs = Get-ADOrganizationalUnit -Filter 'Name -like "*"' -SearchBase $OUPath -SearchScope OneLevel -Properties Description

		if (!$ChildOUs)
		{
			Write-Host "`nThis OU does not contain any child OUs.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
		}
		else
		{
			$DisplayTable = @()
			For ([int]$i = 0; $i -lt $ChildOUs.Count; $i++)
			{
				$TableRow =
				[PSCustomObject]@{
					ID = $i + 1
					NAME = $ChildOUs[$i].Name
				}
				$DisplayTable += $TableRow
			}
			if (!$ChildOUs)
			{
				Write-Host "This OU does not contain any child OUs.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
			}
			else
			{
				$DisplayTable | Format-Table | Out-Host
			}
		}

		# List available options for current OU
		Write-Host "OPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
		if ($Mode -eq "Add")
		{
			Write-Host "- Enter '+' to add a Computer Object at current location."
		}
		if ($ComputerObjects -and $Mode -eq "Edit")
		{
			Write-Host "- Enter '+' to edit a Computer Object at current location."
		}
		if ($ComputerObjects -and $Mode -eq "Remove")
		{
			Write-Host "- Enter '+' to remove a Computer Object at current location."
		}
		if ($OUPath -ne $StartingOU)
		{
			Write-Host "- Enter '<' to return to Parent OU location."
			Write-Host "- Enter '^' to return to Default OU location."
		}
		if ($ChildOUs)
		{
			Write-Host "- Enter the ID number to move to the respective OU location."
		}
		Write-Host

		# Logic for handling user's input
		[bool]$Valid = $FALSE
		do
		{
			Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
			$Selection = Read-Host
			if ($Selection -eq '+' -and $Mode -in ("Add", "Remove"))
			{
				try
	      {
	        New-ADComputer -Name "TEST-AD-PERMISSIONS" -SamAccountName "TEST-AD-PERMISSIONS" -Path $OUPath
	        Remove-ADComputer -Identity "TEST-AD-PERMISSIONS" -Confirm:$FALSE
	        $Done = $TRUE
	        $Valid = $TRUE
					Return $OUPath
	      }
	      catch
	      {
	        Write-Host "`nYou do not have permissions to add/remove computer objects to this OU. Please select another location.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
	      }
			}
			elseif ($Selection -eq '+' -and $Mode -eq "Edit")
			{
				try
				{
					#Try modifying the description. If failed, no permissions
					$OriginalDesc = $ComputerObjects[0].Description
					Set-ADComputer -Identity $ComputerObjects[0].Name -Description "Testing AD Permissions"
					Set-ADComputer -Identity $ComputerObjects[0].Name -Description $OriginalDesc
					$Done = $TRUE
					$Valid = $TRUE
					Return $OUPath
				}
				catch
				{
					Write-Host "`nYou do not have permissions to edit computer objects in this OU. Please select another location.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
				}
			}
			elseif ($Selection -eq '<')
			{
				$OUPath = $ParentOU
				$Valid = $TRUE
				if ($OUPath -eq $StartingOU)
				{
					$ParentOU = $StartingOU
				}
				else
				{
					$ParentOU = ($ParentOU -split ',', 2)[1]
				}
			}
			elseif ($Selection -eq '^')
			{
				$OUPath = $StartingOU
				$ParentOU = $StartingOU
				$Valid = $TRUE
			}
			elseif ($Selection -match "^\d+$" -and [int]$Selection -in 1..$ChildOUs.Count)
			{
				$ParentOU = $OUPath
				$Valid = $TRUE
				$OUPath = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $OUPath -SearchScope OneLevel | Select-Object -First $([int]$Selection) | Select-Object -Last 1
			}
			elseif ($Selection -eq "")
			{
				Write-Host "`nYou did not enter a selection. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
			}
			else
			{
				Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
			}
		} while (!$Valid)
		Clear-Host
	} while (!$Done)
}

function Global:Search-ADForest([string]$SearchObject = "", [array]$SearchKeywords)
{
	[System.Collections.ArrayList]$PotentialResults = @()
	if ($SearchObject -eq "OrganizationalUnit")
	{
		# If the user entered the full DistinguishedName of the OU
		if ($SearchKeywords.Count -eq 1 -and $SearchKeywords -match "^OU=.+$")
		{
			$PotentialResults += Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $StartingOU | Where-Object {$_.DistinguishedName -match $SearchKeywords} | Select-Object -ExpandProperty DistinguishedName
		}
		# If the user entered the full Canonical Name of the OU
		elseif ($SearchKeywords.Count -eq 1 -and $SearchKeywords -match "^SGF.EDUBEAR.NET.+$")
		{
			$PotentialResults += Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $StartingOU -Properties CanonicalName | Where-Object {$_.CanonicalName -match $SearchKeywords} | Select-Object -ExpandProperty DistinguishedName
		}
		else
		{
			ForEach ($Keyword in $SearchKeywords)
			{
				# Lists every OU on the AD that has the given keyword
				$ValidOUPath = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $StartingOU -Properties CanonicalName -PipelineVariable OU |
				Where-Object {$OU.Name -like $Keyword -or $OU.DistinguishedName -match $Keyword -or $OU.CanonicalName -match $Keyword} |
				Select-Object -ExpandProperty DistinguishedName

				# Performing an Intersection of all keywords to narrow down potential OU choices
				if ($PotentialResults.Count -eq 0 -and $ValidOUPath -ne "")
				{
					$PotentialResults += $ValidOUPath
				}
				else
				{
					$PotentialResults = $PotentialResults | Where-Object {$ValidOUPath -contains $_}
				}
			}
		}
	}

	if ($SearchObject -eq "ComputerObject")
	{
		# If the user entered a MAC Address
		if ($SearchKeywords -match ($ValidMACPatterns -join '|'))
		{
			[guid]$ComputerGUID = "00000000-0000-0000-0000-$SearchKeywords"
			$ADComputers = Get-ADComputer -Filter "Name -like '*' -and netbootGUID -like '*'" -SearchBase $StartingOU -Properties netbootGUID | Where-Object {($_.netbootGUID -join "") -eq ($ComputerGUID.ToByteArray() -join "")}
			if ($ADComputers -ne $null)
			{
				ForEach ($Computer in $ADComputers)
				{
					$PotentialResults += $Computer.Name
				}
			}
		}
		else
		{
			ForEach ($Keyword in $SearchKeywords)
			{
				# Lists every Computer Object on the AD that has the given keyword
				$ValidComputerObject = Get-ADComputer -Filter "Name -like '*'" -SearchBase $StartingOU -PipelineVariable Computer | Where-Object {$Computer.Name -match $Keyword} | Select-Object -ExpandProperty Name

				# Performing an Intersection of all keywords to narrow down potential OU choices
				if ($PotentialResults.Count -eq 0 -and $ValidComputerObject -ne "")
				{
					$PotentialResults += $ValidComputerObject
				}
				else
				{
					$PotentialResults = $PotentialResults | Where-Object {$ValidComputerObject -contains $_}
				}
			}
		}
	}

	# To ensure that there is at least one OU or Computer Object available to choose from
	if ($PotentialResults.Count -gt 0)
	{
		# Sorting the OUs pulled from the input into hierarchical order
		if ($SearchObject -eq "OrganizationalUnit")
		{
			[System.Collections.ArrayList]$SortedResults = @()
			$TotalObjects = $PotentialResults.Count
			do
			{
				ForEach ($OU in $PotentialResults)
				{
					$ParentOU = ($OU -split ',', 2)[1]
					if ($PotentialResults -notcontains $ParentOU)
					{
						$Index = 0..($SortedResults.Count - 1) | Where-Object {$SortedResults[$_] -eq $ParentOU}
						if ($Index -ne $null)
						{
							if ($SortedResults.Count -eq 0)
							{
								$SortedResults.Add($OU) | Out-Null
							}
							else
							{
								$SortedResults.Insert($Index + 1, $OU)
							}
						}
						else
						{
							$SortedResults.Insert(0, $OU)
						}
					}
				}
				ForEach ($OU in $SortedResults)
				{
					$PotentialResults = @($PotentialResults | Where-Object {$_ -ne $OU})
				}
			} while ($SortedResults.Count -ne $TotalObjects)
			$PotentialResults = $SortedResults
		}
		else
		{
			$PotentialResults.Sort()
		}

		[array]$DisplayTable = @()
		For ([int]$i = 0; $i -lt $PotentialResults.Count; $i++)
		{
			if ($SearchObject -eq "OrganizationalUnit")
			{
				$Name = Get-ADOrganizationalUnit -Identity $PotentialResults[$i] | Select-Object -ExpandProperty Name
				$TableRow =
				[PSCustomObject]@{
					ID = $i+1
					NAME = $Name
					DISTINGUISHEDNAME = $PotentialResults[$i]
				}
			}
			if ($SearchObject -eq "ComputerObject")
			{
				$ComputerProperties = Get-ADComputer -Identity $PotentialResults[$i] -Properties netbootGUID
				$DistinguishedName = $ComputerProperties.DistinguishedName
				$MAC = ($ComputerProperties.netbootGUID | ForEach ToString X2) -join ""
				$ShortOU = ($DistinguishedName -Split ',' | Where-Object {$_ -like '*OU*'}).SubString(3)
				[array]::Reverse($ShortOU)
				$ShortOU = $ShortOU -join "/"

				$TableRow =
				[PSCustomObject]@{
					SEL = "[ ]"
					ID = $i+1
					NAME = $PotentialResults[$i]
					MAC = $MAC.SubString(20)
					OUPATH = $ShortOU
				}
			}
			$DisplayTable += $TableRow
		}
		Return $DisplayTable
	}
	else
	{
		Return $null
	}
}

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
if ($ENV:SMS_ADMIN_UI_PATH -eq $null)
{
  $SMS_ADMIN_UI_PATH = Get-ChildItem -Path "C:\Program Files (x86)" -Include *ConfigurationManager.psd1* -File -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Directory
  if ($SMS_ADMIN_UI_PATH.Count -gt 1)
  {
    $Global:SMS_ADMIN_UI_PATH = $SMS_ADMIN_UI_PATH | Sort-Object Length -Descending | Select-Object -First 1
  }
}
else
{
  $Global:SMS_ADMIN_UI_PATH = "$ENV:SMS_ADMIN_UI_PATH\.."
}

if ($Global:SMS_ADMIN_UI_PATH -eq $null)
{
  $Global:DistributionPoints =
  @(
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
else
{
  if((Get-Module ConfigurationManager) -eq $null)
  {
      Import-Module "$($Global:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"
  }
  if((Get-PSDrive -Name "MSU" -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null)
  {
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
if (Get-Module -ListAvailable -Name ActiveDirectory)
{
  Import-Module ActiveDirectory
}
else
{
  Write-Host "`nYou do not have Active Directory on this machine." -ForegroundColor $FGWarning -BackgroundColor $BGBold
  Write-Host "This script heavily relies on the Active Directory Module and cannot function without it." -ForegroundColor $FGWarning -BackgroundColor $BGBold
  Write-Host "Please install the Remote Server Administration Tools (RSAT) package on your machine to use this script.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
  Write-Host "Press ENTER to quit.`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
  Read-Host
  Exit
}

$Global:StartingOU = "OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
$Global:CurrentLocation = Get-Location | Select-Object -ExpandProperty Path
$Global:LogFilePath = $CurrentLocation + "\ADComputer_Log.CSV"
$Global:ComputerArchive = @()

# Array of words Microsoft won't allow a computer name to be, AKA Reserved Words
$Global:ReservedWords =
@(
  "ANONYMOUS","AUTHENTICATED USERS","BATCH","BUILTIN","CREATOR GROUP","CREATOR GROUP SERVER","CREATOR OWNER",
  "CREATOR OWNER SERVER","DIALUP","DIGEST AUTH","INTERACTIVE","INTERNET","LOCAL","LOCAL SYSTEM","NETWORK",
  "NETWORK SERVICE","NT AUTHORITY","NT DOMAIN","NTLM AUTH","NULL","PROXY","REMOTE INTERACTIVE","RESTRICTED",
  "SCHANNEL AUTH","SELF","SERVER","SERVICE","SYSTEM","TERMINAL SERVER","THIS ORGANIZATION","USERS","WORLD"
)

# Regular Expression representing the different formats a MAC address is traditionally represented in
$Global:ValidMACPatterns =
@(
  '^([0-9a-f]{2}:){5}([0-9a-f]{2})$'
  '^([0-9a-f]{2}-){5}([0-9a-f]{2})$'
  '^([0-9a-f]{2} ){5}([0-9a-f]{2})$'
  '^([0-9a-f]{2}\.){5}([0-9a-f]{2})$'
  '^([0-9a-f]{12})$'
)

[bool]$Running = $TRUE
do
{
	if ($Selection -eq $NULL -or $Selection -notmatch "[1-5]")
	{
		Clear-Host
		Write-Host "`nMissouri State University Active Directory Suite" -ForegroundColor $FGTitle -BackgroundColor $BGBold
		Write-Host "`nThis script will allow you to add, edit, or delete Computer AD objects the SGF.EDUBEAR.NET domain."
		Write-Host "To get started, please select one of the options listed below."

		Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
		Write-Host "- Enter '1' to SEARCH for an existing Computer on the domain."
		Write-Host "- Enter '2' to ADD a new Computer to the domain."
		Write-Host "- Enter '3' to EDIT an existing Computer on the domain."
		Write-Host "- Enter '4' to REMOVE an existing Computer from the domain."
		Write-Host "- Enter '5' to UPDATE the Active Directory with a CSV file."
		Write-Host "- Press ENTER to QUIT.`n"
	}

	# Logic for handling user's input
	[bool]$Valid = $FALSE
	do
	{
		if ($Selection -eq $NULL -or $Selection -notmatch "[1-5]")
		{
		  Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
		  $Selection = Read-Host
		}
	  if ($Selection -eq "")
	  {
	    Exit
	  }
	  elseif ($Selection -eq 1)
	  {
	    $Valid = $TRUE
			Import-Module .\Modules\Search-ADObject.psm1
	    Search-ADObject
	  }
		elseif ($Selection -eq 2)
		{
			$Valid = $TRUE
			Import-Module .\Modules\Add-ADObject.psm1
			Add-ADObject
		}
	  elseif ($Selection -eq 3)
	  {
	    $Valid = $TRUE
			Import-Module .\Modules\Edit-ADObject.psm1
	    Edit-ADObject
	  }
	  elseif ($Selection -eq 4)
	  {
	    $Valid = $TRUE
			Import-Module .\Modules\Remove-ADObject.psm1
	    Remove-ADObject
	  }
		elseif ($Selection -eq 5)
		{
			$Valid = $TRUE
			Import-Module .\Modules\Update-ADObject.psm1
			Update-ADObject
		}
	  else
	  {
	    Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
	  }
	} while (!$Valid)

	Clear-Host
	Write-Host "`nMissouri State University Active Directory Suite`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
	Write-Host "Any operations that you have made will be logged in a CSV file."
	Write-Host "The current log file location is: " -NoNewLine
	Write-Host $LogFilePath -ForegroundColor $FGLabel -BackgroundColor $BGBold
	Write-Host "You just finished using the " -NoNewLine
	switch ($Selection)
	{
		1 { Write-Host "Search-ADObject" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine}
		2 { Write-Host "Add-ADObject" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine}
		3 { Write-Host "Edit-ADObject" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine}
		4 { Write-Host "Remove-ADObject" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine}
		5 { Write-Host "Update-ADObject" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine}
		Default {}
	}
	Write-Host " function of this Active Directory Suite."
	Write-Host "If you would like to use another function of this suite, please enter one of the following options below.`n"

	Write-Host "OPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
	Write-Host "- Enter '1' to SEARCH for an existing Computer on the domain."
	Write-Host "- Enter '2' to ADD a new Computer to the domain."
	Write-Host "- Enter '3' to EDIT an existing Computer on the domain."
	Write-Host "- Enter '4' to REMOVE an existing Computer from the domain."
	Write-Host "- Enter '5' to UPDATE the Active Directory Domain from a CSV file."
	Write-Host "- Press ENTER to QUIT.`n"

	$Valid = $FALSE
	do
	{
		Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
		$Selection = Read-Host
		if ($Selection -eq "")
		{
			# Export everything done during this session to an excel file
			Get-Process Excel -ErrorAction SilentlyContinue | Select-Object -Property ID | Stop-Process
			Start-Sleep -Milliseconds 300
			$ComputerArchive | Export-CSV -Path $LogFilePath -NoTypeInformation -Append
			Exit
		}
		elseif ($Selection -match "[1-5]")
		{
			$Valid = $TRUE
			$Running = $TRUE
		}
		else
		{
			Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
		}
	} while (!$Valid)
} while ($Running)
