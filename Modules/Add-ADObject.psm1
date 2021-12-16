function Add-ADObject {
  <#
  .SYNOPSIS
  Computer Object creation script with the capability of adding objects with a CSV file or
  by entering the attributes manually. Once the machines have been added to the Active Directory site,
  a log file is created in the same directory as this script.

  .DESCRIPTION
  Example CSV file:

  FUNCTION,COMPUTERNAME,COMPUTERMAC,COMPUTERDESC,COMPUTEROU,DEPLOYMENTSERVER
  PSSH-PRESTAGE-00,K4L9PAE60EA0,Optiplex 780 for Kevin,"OU=USIT,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET",SCCM-DP-SCUF-01.missouristate.edu
  PSSH-PRESTAGE-01,K4L9PAE60EA1,Laptop for Tim Hurt,"OU=CORE,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET",SCCM-DP-CORE-01.missouristate.edu

  FUNCTION          - A string value representing the function to use with the data. Options are ADD, EDIT, REMOVE.
  COMPUTERNAME      - A string value representing the Object's name.
  COMPUTERMAC       - A string value representing the Object's MAC Address.
      *This value can have the following delimiters and still function: ':', '-', '.', ' '
      **EX: 1A2B3C4D5F67, 1A:2B:3C:4D:5F:67, 1A-2B-3C-4D-5F-67, 1A.2B.3C.4D.5F.67, and 1A 2B 3C 4D 5F 67 are all valid
  COMPUTERDESC      - A string value representing the Object's description.
  COMPUTEROU        - A string value representing the desired Object's OU location.
      *This value can work with the full Distinguished Name (DN) surrounded by "", Canonical Name (CN), or a few select keywords to represent the OU location.
      **For better results, use the DN or CN as the keywords option can find multiple OUs with said keyword in it.
      ***EX: For the "OU=USIT,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET" OU, these are some acceptable entries...
      "OU=USIT,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
      SGF.EDUBEAR.NET/CUSTOM/COMPUTERS/WORKSTATIONS/USIT
      Computers Workstations USIT
  DEPLOYMENTSERVER  - A string value representing the Object's netbootMachineFilePath server, also known as a Deployment Server or Distribution Point.
      *This value can work with the full server name or the unique keyword of the server.
      **This script attempts to connect to the SCCM-01 site server to pull all available Distribution Points. In the event that the MEMCM Powershell module cannot
          be found, this script will use a hardcoded list of Distribution Points that was pulled on 12/12/2021.
      ***EX: For the SCCM-DP-SCUF-01.missouristate.edu, you can use the value SCCM-DP-SCUF-01.missouristate.edu or SCUF.
  #>
  Clear-Host

  Write-Host "`nAdding Active Directory Computer Objects" -ForegroundColor $FGTitle -BackgroundColor $BGBold
  Write-Host "This script will allow you to add new Computer AD objects to the SGF.EDUBEAR.NET domain."
  Write-Host "It also has the capability to add multiple objects from a CSV file as well!"
  Write-Host "To get started, please select one of the options listed below."

  Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
  Write-Host "- Enter '1' to manually navigate to the desired OU path."
  Write-Host "- Enter '2' to find OU using DistinguishedName, CanonicalName, or select Keywords."
  Write-Host "- Enter '3' to pass a CSV file through the script. (Useful when adding multiple machines)"
  Write-Host "- Press Enter to quit.`n"

  # Logic for handling user's input
  do {
    Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
    $Method = Read-Host
    if ($Method -eq "") {
      Exit
    }
    elseif ($Method -eq 1 -or $Method -eq 2) {
      $OUPath = $StartingOU
      $Confirmation = 0

      do {
      # Logic for manually navigating to desired OU
        if ($Method -eq 1 -and ($Repeat -eq 2 -or $Confirmation -eq 1 -or $Confirmation -eq 0)) {

          [bool]$Done = $False
          do {
            Clear-Host
            Write-Host "`nFinding Computer OU Location Manually" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "`nDefault AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            Write-Host "`t$StartingOU"
            if ($OUPath -ne $StartingOU) {
              Write-Host "Parent  AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
              Write-Host "`t$ParentOU"
            }
            Write-Host "Current AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            Write-Host "`t$OUPath`n"

            # Get all Child OUs from current OU
            $ChildOUs = Get-ADOrganizationalUnit -Filter 'Name -like "*"' -SearchBase $OUPath -SearchScope OneLevel
            if (!$ChildOUs) {
              Write-Host "This OU does not contain any child OUs.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
            }
            else {
              Write-Host "ID`tNAME" -ForegroundColor $FGHeader -BackgroundColor $BGBold
              [int]$Index = 1
              ForEach ($OU in $ChildOUs) {
                Write-Host "$Index`t$($OU.Name)"
                $Index += 1
              }
              Write-Host
            }

            # List available options for current OU
            Write-Host "`OPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
            if ($Confirmation -eq 1) {
              Write-Host "- Enter '+' to select this OU for your Computer Object."
            }
            else {
              Write-Host "- Enter '+' to create a Computer Object at current location."
            }
            if ($OUPath -ne $StartingOU) {
              Write-Host "- Enter '^' to return to Default OU location."
              Write-Host "- Enter '<' to return to Parent OU location."
            }
            if ($ChildOUs) {
              Write-Host "- Enter the ID number to move to the respective OU location."
            }
            Write-Host

            # Logic for handling user's input
            [bool]$Valid = $False
            do {
              Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
              $Selection = Read-Host
              if ($Selection -eq '+') {
                try {
                  New-ADComputer -Name "TEST-AD-PERMISSIONS" -SamAccountName "TEST-AD-PERMISSIONS" -Path $OUPath
                  Remove-ADComputer -Identity "TEST-AD-PERMISSIONS" -Confirm:$False
                  $Done = $True
                  $Valid = $True
                }
                catch {
                  Write-Host "`nYou do not have permissions to add computer objects to this OU. Please select another location.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                }
              }
              elseif ($Selection -eq '^') {
                $OUPath = $StartingOU
                $ParentOU = $StartingOU
                $Valid = $True
              }
              elseif ($Selection -eq '<') {
                $OUPath = $ParentOU
                $Valid = $True
                if ($OUPath -eq $StartingOU) {
                  $ParentOU = $StartingOU
                }
                else {
                  $ParentOU = ($ParentOU -split ',', 2)[1]
                }
              }
              elseif ($Selection -match "^\d+$" -and [int]$Selection -in 1..$ChildOUs.Count) {
                $ParentOU = $OUPath
                $Valid = $True
                $OUPath = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $OUPath -SearchScope OneLevel | Select-Object -First $([int]$Selection) | Select-Object -Last 1
              }
              elseif ($Selection -eq "") {
                Write-Host "`nYou did not enter a selection. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
              else {
                Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
            } while (!$Valid)
            Clear-Host
          } while (!$Done)
        }

        # Logic for entering Keywords to find desired OU
        elseif ($Method -eq 2 -and ($Repeat -eq 2 -or $Confirmation -eq 1 -or $Confirmation -eq 0)) {

          [bool]$Done = $False
          do {
            Clear-Host
            Write-Host "`nFinding OU using Distinguished Name (DN), Canonical Name (CN), or select Keywords (KW)`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "When entering any of these values, please use the following formats..." -ForegroundColor $FGNotice -BackgroundColor $BGBold
            Write-Host "- DN: OU=USIT,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
            Write-Host "- CN: SGF.EDUBEAR.NET/CUSTOM/COMPUTERS/WORKSTATIONS/USIT"
            Write-Host "- KW: USIT WORKSTATIONS`n"
            $PotentialOUs = @()

            [bool]$Valid = $False
            do {
              Write-Host "Enter the DN, CN, or KWs for the desired OU:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
              $EnteredOU = Read-Host
              $OUKeywords = $EnteredOU.ToUpper() -split " "

              # If the user entered the full DistinguishedName of the OU
              if ($OUKeywords.Count -eq 1 -and $OUKeywords -match "^OU=.+$") {
                $ValidOUPath = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $StartingOU | Where-Object {$_.DistinguishedName -match $OUKeywords} | Select-Object -ExpandProperty DistinguishedName

                if ($ValidOUPath.Count -eq 0) {
                  Write-Host "`nThe Distinguished Name you entered could not find an OU Object. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  Write-Host "If the Distinguished Name belongs to an existing OU, you may not have access to add objects to it.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold

                }
                else {
                  $PotentialOUs += $ValidOUPath
                  $Valid = $True
                }
              }
              # If the user entered the full Canonical Name of the OU
              elseif ($OUKeywords.Count -eq 1 -and $OUKeywords -match "^SGF.EDUBEAR.NET.+$") {
                $ValidOUPath = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $StartingOU -Properties CanonicalName | Where-Object {$_.CanonicalName -match $OUKeywords} | Select-Object -ExpandProperty DistinguishedName

                if ($ValidOUPath.Count -eq 0) {
                  Write-Host "`nThe Canonical Name you entered could not find an OU Object. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  Write-Host "If the Canonical Name belongs to an existing OU, you may not have access to add objects to it.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
                }
                else {
                  $PotentialOUs += $ValidOUPath
                  $Valid = $True
                }
              }
              else {
                ForEach ($Keyword in $OUKeywords) {
                  # Lists every OU on the AD that has the given keyword
                  $ValidOUPath = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $StartingOU -Properties CanonicalName -PipelineVariable OU |
                  Where-Object {$OU.Name -like $Keyword -or $OU.DistinguishedName -match $Keyword -or $OU.CanonicalName -match $Keyword} |
                  Select-Object -ExpandProperty DistinguishedName

                  # Performing an Intersection of all keywords to narrow down potential OU choices
                  if ($PotentialOUs.Count -eq 0 -and $ValidOUPath -ne "") {
                    $PotentialOUs += $ValidOUPath
                  }
                  else {
                    $PotentialOUs = $PotentialOUs | Where-Object {$ValidOUPath -contains $_}
                  }
                }
                if ($PotentialOUs.Count -eq 0) {
                  Write-Host "`nThe Keyword(s) you entered could not find an OU object. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  Write-Host "If the Keyword(s) represent an existing OU, you may not have access to add objects to it." -ForegroundColor $FGNotice -BackgroundColor $BGBold
                  Write-Host "You might try putting in fewer Keywords as the search tries to find OUs with every Keyword entered.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
                }
                else {
                  $Valid = $True
                }
              }
            } while (!$Valid)

            # To ensure that there is at least one OU available to choose from
            if ($PotentialOUs.Count -gt 0) {
              [System.Collections.ArrayList]$SortedOUs = @()
              $TotalOUs = $PotentialOUs.Count
              $PotentialOUs = $PotentialOUs | Sort-Object -Descending

              # Sorting the OUs pulled from the input into hierarchical order
              do {
                ForEach ($OU in $PotentialOUs) {
                  $ParentOU = ($OU -split ',', 2)[1]
                  if ($PotentialOUs -notcontains $ParentOU) {
                    $Index = 0..($SortedOUs.Count - 1) | Where-Object {$SortedOUs[$_] -eq $ParentOU}
                    if ($Index -ne $null) {
                      if ($SortedOUs.Count -eq 0) {
                        $SortedOUs.Add($OU) | Out-Null
                      }
                      else {
                        $SortedOUs.Insert($Index + 1, $OU)
                      }
                    }
                    else {
                      $SortedOUs.Insert(0, $OU)
                    }
                  }
                }
                ForEach ($OU in $SortedOUs) {
                  $PotentialOUs = $PotentialOUs | Where-Object {$_ -ne $OU}
                }
              } while ($SortedOUs.Count -ne $TotalOUs)
              $PotentialOUs = $SortedOUs

              $DisplayTable = @()
              For ([int]$i = 0; $i -lt $PotentialOUs.Count; $i++) {
                $Name = Get-ADOrganizationalUnit -Identity $PotentialOUs[$i] | Select-Object -ExpandProperty Name
                $TableRow = [PSCustomObject]@{
                  ID = $i+1
                  NAME = $Name
                  DISTINGUISHEDNAME = $PotentialOUs[$i]
                }
                $DisplayTable += $TableRow
              }

              Clear-Host
              Write-Host "`nFinding OU using Distinguished Name (DN), Canonical Name (CN), or select Keywords (KW)`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
              Write-Host "Found $($DisplayTable.Count) potential OU path(s).`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold

              $DisplayTable | Format-Table

              Write-Host "`OPTIONS" -ForegroundColor $FGInput -BackgroundColor $BGInput
              Write-Host "- Enter the ID number to select the respective OU location."
              Write-Host "- Enter '^' to try again if the desired OU is not listed.`n"

              # Logic for handling user input
              $Valid = $False
              do {
                Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
                $Selection = Read-Host
                if ($Selection -match "^\d+$" -and [int]$Selection -in 1..$DisplayTable.Count) {
                  $OUPath = $DisplayTable[$Selection - 1].DISTINGUISHEDNAME
                  try {
                    New-ADComputer -Name "TEST-AD-PERMISSIONS" -SamAccountName "TEST-AD-PERMISSIONS" -Path $OUPath
                    Remove-ADComputer -Identity "TEST-AD-PERMISSIONS" -Confirm:$False
                    $Done = $True
                    $Valid = $True
                  }
                  catch {
                    Write-Host "`nYou do not have permissions to add computer objects to this OU. Please select another location.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  }
                }
                elseif ($Selection -eq '^') {
                  $Done = $False
                  $Valid = $True
                  Clear-Host
                }
                else {
                  Write-Host "`nYou did not enter a selection. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                }
              } while (!$Valid)
            }
          } while (!$Done)
        }

        # Get Computer Name Loop
        if ($Confirmation -eq 2 -or $Confirmation -eq 0 -or $Confirmation -eq "") {

          Clear-Host
          Write-Host "`nAdding Computer Object Attributes: Computer Name" -ForegroundColor $FGTitle -BackgroundColor $BGBold
          Write-Host "`nCURRENT VALUES" -ForegroundColor $FGHeader -BackgroundColor $BGBold
          Write-Host "Computer OU  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $OUPath"
          Write-Host "Computer Name:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerName"
          Write-Host "Computer MAC :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $MACAddress"
          Write-Host "Computer Desc:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerDescription"
          Write-Host "Computer DP  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $DeploymentServer`n"

          # Check if the entered Computer Name is already tied to a Computer Object on the AD
          $Done = $False
          do {
            Write-Host "Enter a Name for the new Computer Object:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $ComputerName = Read-Host
            if ($ComputerName -eq "") {
              Write-Host "`nYou did not enter a Computer Name. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($ComputerName -match '^\.|([\\/:\*\?"<>\|])+') {
              Write-Host `n'Computer Name cannot start with a . and\or contain the following: \ / : * ? " < > |'`n -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($ComputerName.Length -notin 1..15) {
              Write-Host "`nComputer Name must be between 1 and 15 characters.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($ComputerName.ToUpper() -in $ReservedWords) {
              Write-Host "`n$($ComputerName.ToUpper()) is a Reserved Word and can't be a Computer Name." -ForegroundColor $FGWarning -BackgroundColor $BGBold
              Write-Host "The following words are Reserved Words, as deemed by Microsoft:" -ForegroundColor $FGNotice -BackgroundColor $BGBold
              Write-Host ($ReservedWords -join ", ")`n -ForegroundColor $FGNotice -BackgroundColor $BGBold
            }
            else {
              $ComputerName = $ComputerName.ToUpper()
              $CheckName = Get-ADObject -Filter "ObjectClass -eq 'computer' -and Name -eq '$($ComputerName)'"
              if ($CheckName -eq $null) {
                $Done = $True
              }
              else {
                Write-Host "`nThere is already a Computer Object with that Name:`n$($CheckName.DistinguishedName)`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
            }
          } while (!$Done)
        }

        # Get Computer MAC Address Loop
        if ($Confirmation -eq 3 -or $Confirmation -eq 0 -or $Confirmation -eq "") {

          Clear-Host
          Write-Host "`nAdding Computer Object Attributes: Computer MAC Address" -ForegroundColor $FGtitle -BackgroundColor $BGBold
          Write-Host "`nCURRENT VALUES" -ForegroundColor $FGHeader -BackgroundColor $BGBold
          Write-Host "Computer OU  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $OUPath"
          Write-Host "Computer Name:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerName"
          Write-Host "Computer MAC :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $MACAddress"
          Write-Host "Computer Desc:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerDescription"
          Write-Host "Computer DP  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $DeploymentServer`n"

          Write-Host "When entering the MAC Address, please use the following formats..." -ForegroundColor $FGNotice -BackgroundColor $BGBold
          Write-Host "- ABCDEF123456"
          Write-Host "- AB CD EF 12 34 56"
          Write-Host "- AB:CD:EF:12:34:56"
          Write-Host "- AB-CD-EF-12-34-56`n"

          # Check if the given MAC Address is already tied to a Computer Object on the AD
          $Done = $False
          do {
            Write-Host "Enter the MAC Address for the Computer Object (Ex. 1A2B3C4D5E6F):`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $EnteredMAC = Read-Host
            if ($EnteredMAC -eq "") {
              Write-Host "`nYou did not enter a MAC Address. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($EnteredMAC -match ($ValidMACPatterns -join '|')) {
              $MACAddress = $EnteredMAC.ToUpper() -replace '\W'
              [guid]$NetbootGUID = "00000000-0000-0000-0000-$MACAddress"
              $CheckMAC = Get-ADComputer -Filter {netbootGUID -like $NetbootGUID}
              if ($CheckMAC -eq $null) {
                $Done = $True
              }
              else {
                Write-Host "`nThere is already a Computer Object with that MAC Address:`n$($CheckMAC.DistinguishedName)`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
            }
            else {
              Write-Host "`nThe entered MAC Address is not properly formatted." -ForegroundColor $FGWarning -BackgroundColor $BGBold
              Write-Host "A properly formatted MAC Address is a series of 12 Hexadecimal characters (0-9, A-F)" -ForegroundColor $FGNotice -BackgroundColor $BGBold
              Write-Host "The characters can be delimited (or not) at every two characters by a ':', '-', '.', or ' '.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
            }
          } while (!$Done)
        }

        # Get Computer Description Loop
        if ($Confirmation -eq 4 -or $Confirmation -eq 0 -or $Confirmation -eq "") {

          Clear-Host
          Write-Host "`nAdding Computer Object Attributes: Computer Description" -ForegroundColor $FGTitle -BackgroundColor $BGBold
          Write-Host "`nCURRENT VALUES" -ForegroundColor $FGHeader -BackgroundColor $BGBold
          Write-Host "Computer OU  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $OUPath"
          Write-Host "Computer Name:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerName"
          Write-Host "Computer MAC :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $MACAddress"
          Write-Host "Computer Desc:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerDescription"
          Write-Host "Computer DP  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $DeploymentServer`n"

          # Add Computer's Description
          Write-Host "Enter the Computer Object's Description Ex. 'Boomer Bear (CARR 0123)':`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
          $ComputerDescription = Read-Host
        }

        # Get Computer Deployment Server Loop
        if ($Confirmation -eq 5 -or $Confirmation -eq 0 -or $Confirmation -eq "") {

          Clear-Host
          Write-Host "`nAdding Computer Object Attributes: Deployment Server" -ForegroundColor $FGTitle -BackgroundColor $BGBold
          Write-Host "`nCURRENT VALUES" -ForegroundColor $FGHeader -BackgroundColor $BGBold
          Write-Host "Computer OU  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $OUPath"
          Write-Host "Computer Name:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerName"
          Write-Host "Computer MAC :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $MACAddress"
          Write-Host "Computer Desc:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $ComputerDescription"
          Write-Host "Computer DP  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
          Write-Host " $DeploymentServer"

          # Select appropriate Deployment Server
          $DisplayTable = @()
          For ([int]$i = 0; $i -lt $DistributionPoints.Count; $i++) {
            $TableRow = [PSCustomObject]@{
              ID = $i+1
              DISTRIBUTIONPOINT = $DistributionPoints[$i].ToUpper()
            }
            $DisplayTable += $TableRow
          }
          $DisplayTable | Format-Table

          # Logic for handling user input
          $Done = $False
          do {
            Write-Host "`Enter the ID of the Distribution Point for this Computer Object:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $Selection = Read-Host
            if ($Selection -eq "") {
              Write-Host "`nYou did not enter a selection. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold

            }
            elseif ($Selection -match "^\d+$" -and [int]$Selection -in 1..$DisplayTable.Count) {
              $DistributionPoint = $DisplayTable[$Selection - 1].DISTRIBUTIONPOINT
              $NetbootMFP = $DistributionPoints | Where-Object {$_ -match $DistributionPoint}
              $DeploymentServer = $NetbootMFP.ToUpper() -replace "\\"
              $Done = $True
            }
            else {
              Write-Host "`nYour selection doesn't exist. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
          } while (!$Done)
        }

        # Confirm or change all the values given before creating the computer object
        Clear-Host
        Write-Host "`nThese are the Computer Object values you have entered." -ForegroundColor $FGTitle -BackgroundColor $BGBold
        Write-Host "`nCURRENT VALUES" -ForegroundColor $FGHeader -BackgroundColor $BGBold
        Write-Host "Computer OU  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
        Write-Host " $OUPath"
        Write-Host "Computer Name:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
        Write-Host " $ComputerName"
        Write-Host "Computer MAC :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
        Write-Host " $MACAddress"
        Write-Host "Computer Desc:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
        Write-Host " $ComputerDescription"
        Write-Host "Computer DP  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
        Write-Host " $DeploymentServer`n"

        Write-Host "OPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
        Write-Host "- Enter '1' to change the Current OU Path."
        Write-Host "- Enter '2' to change the Computer Name."
        Write-Host "- Enter '3' to change the Computer MAC."
        Write-Host "- Enter '4' to change the Computer Description."
        Write-Host "- Enter '5' to change the Computer Distribution Point."
        Write-Host "- Press Enter to confirm these values.`n"

        # Logic for handling user input
        $Valid = $False
        do {
          Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
          $Confirmation = Read-Host
          if ($Confirmation -in 1..5 -or $Confirmation -eq "") {
            Write-Host $Confirmation, $Selection
            $Valid = $True
          }
          else {
            Write-Host "`nYour selection doesn't exist. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
          }
        } while (!$Valid)

        if ($Confirmation -eq "") {
          New-ADComputer -Name $ComputerName -SamAccountName $ComputerName -Description $ComputerDescription -OtherAttributes @{'netbootGUID' = $NetbootGUID; 'netbootMachineFilePath' = $DeploymentServer} -Path $OUPath

          # This check is to ensure that the machine was actually added to the AD
          $CheckAddition = Get-ADComputer -Identity $ComputerName
          if ($CheckAddition) {

            # Object Information to export to CSV
            $ComputerLog = [PSCustomObject]@{
              COMPUTERNAME = $ComputerName
              OUPATH = $OUPath
              ADDEDBY = $([Environment]::UserName)
              DATEADDED = Get-Date
              SUCCESSFUL = $TRUE
              ERRORS = $ErrorLog
            }

            Get-Process Excel -ErrorAction SilentlyContinue | Select-Object -Property ID | Stop-Process
            Start-Sleep -Milliseconds 300
            $ComputerLog | Export-CSV -Path $LogFilePath -NoTypeInformation -Append

            Write-Host "`nThe computer $ComputerName has been added to the Active Directory successfully.`n" -ForegroundColor $FGSuccess -BackgroundColor $BGBold
          }
          else {
            Write-Host "`nSomething went wrong and the Computer Object was unable to be added to the Active Directory. Please close the script and try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            Write-Host "If this continues to happen, please contact your system administrator.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
          }

          Write-Host "OPTIONS" -BackgroundColor $BGBold
          if ($CheckAddition) {
            Write-Host "- Enter '1' to add another Computer Object to the same OU location."
            Write-Host "- Enter '2' to add another Computer Object, but search for a different OU location."
          }
          Write-Host "- Press Enter to quit.`n"

          # Logic for handling user input
          $Valid = $False
          do {
            Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $Repeat = Read-Host

            if ($Repeat -eq 1 -or $Repeat -eq 2) {
              $Valid = $True

              # Reset Values for next Computer Object
              $ComputerName = ''
              $MACAddress = ''
              $ComputerDescription = ''
              $DeploymentServer = ''
              if ($Repeat -eq 2) {
                $OUPath = $StartingOU
                $Confirmation = 0
              }
            }
            elseif ($Repeat -eq "") {
              $Valid = $True
              Exit
            }
            else {
              Write-Host "Your selection doesn't exist. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
          } while (!$Valid)
        }
      } while ($Repeat -eq 1 -or $Repeat -eq 2 -or $Confirmation -in 1..5)
    }

    # Automated Side of the script pulling values from CSV File
    elseif ($Method -eq 3) {
      Clear-Host
      Write-Host "`nImporting Computer Objects via external CSV File" -ForegroundColor $FGTitle -BackgroundColor $BGBold
      Write-Host "To add computer objects using a CSV file, the file must contain the following headers in order:"
      Write-Host "COMPUTERNAME`tCOMPUTERMAC`tCOMPUTERDESC`tCOMPUTEROU`tDEPLOYMENTSERVER" -ForegroundColor $FGNotice -BackgroundColor $BGBold
      Write-Host "`nOnce that is done, feel free to add as many computer objects as you need."
      Write-Host "After the import is completed, there will be a log file created at"
      Write-Host "$($LogFilePath)" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
      Write-Host " detailing whether the object was successfully created or not."

      $Valid = $False
      do {
        Write-Host "`nEnter the full path of the CSV file you wish to import:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
        $CSVFile = Read-Host
        if ($CSVFile) {
          $CSVPath = Test-Path $CSVFile -Include "*.csv"
          if ($CSVPath -eq $True) {
            $ValidHeaders = $True
            $FileHeaders = (Get-Content $CSVFile -TotalCount 1).Split(',')
            for ([int]$i = 0; $i -lt $FileHeaders.Count; $i++) {
              if ($FileHeaders[$i].Trim() -ne $CorrectCSVHeaders[$i]) {
                $ValidHeaders = $False
              }
            }
            if ($ValidHeaders -eq $True) {
              $Valid = $True
            }
          }
          if ($CSVPath -eq $False -or $ValidHeaders -eq $False) {
            Write-Host "`n$CSVFile is not a valid CSV file. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            if ($ValidHeaders -eq $False) {
              Write-Host "Make sure that your CSV File has the following headers in this order:" -ForegroundColor $FGNotice -BackgroundColor $BGBold
              Write-Host "COMPUTERNAME,COMPUTERMAC,COMPUTERDESC,COMPUTEROU,DEPLOYMENTSERVER" -ForegroundColor $FGLabel -BackgroundColor $BGBold
            }
          }
        }
      } while (!$Valid)

      # Validate CSVFile being passed as an argument
      if ($CSVPath -eq $True -and $ValidHeaders -eq $True) {
        $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
        [int]$TotalComputers = Import-CSV -Path $CSVFile | Measure-Object | Select-Object -expand Count
        [int]$FinishedComputers = 0
        [int]$Success = 0
        [int]$Failed = 0
        Clear-Host
        Write-Host "`nImporting AD Objects from: $CSVFile`n" -ForegroundColor $FGSuccess -BackgroundColor $BGBold

        Import-CSV -Path $CSVFile | ForEach-Object {

          $FinishedComputers += 1
          Write-Host "Importing Object $FinishedComputers of $TotalComputers" -ForegroundColor $FGNotice -BackgroundColor $BGBold

          $ErrorLog = ''
          $Computer = [PSCustomObject]@{
            Name = $_.COMPUTERNAME
            MAC = $_.COMPUTERMAC
            Desc = $_.COMPUTERDESC
            OU = $_.COMPUTEROU
            DP = $_.DEPLOYMENTSERVER
          }

          # Check if given value meets the requirements of a valid computer name
          # Then check if the given Computer Name is already tied to a Computer Object on the AD
          if ($Computer.Name -eq "") {
            $ErrorLog += "No Name given;"
          }
          elseif ($Computer.Name -match '^\.|([\\/:\*\?"<>\|])+') {
            $ErrorLog += "Computer name contains invalid character(s);"
          }
          elseif ($Computer.Name.Length -notin 1..15) {
            $ErrorLog += "Computer name doesn't meet length requirements;"
          }
          elseif ($Computer.Name.ToUpper() -in $ReservedWords) {
            $ErrorLog += "Computer name is a reserved word;"
          }
          else {
            $Computer.Name = $Computer.Name.ToUpper()
            $CheckName = Get-ADObject -Filter "ObjectClass -eq 'computer' -and Name -eq '$($Computer.Name)'"
            if ($CheckName -ne $null) {
              $ErrorLog += "'$($Computer.Name)' already exists in the AD;"
            }
          }

          # Check if the given MAC Address is already tied to a Computer Object on the AD
          if ($Computer.MAC -eq "") {
            $ErrorLog += "No MAC Address given;"
          }
          elseif ($Computer.MAC -match ($ValidMACPatterns -join '|')) {
            $MACAddress = $Computer.MAC.ToUpper() -replace '\W'
            [guid]$NetbootGUID = "00000000-0000-0000-0000-$MACAddress"
            $CheckMAC = Get-ADComputer -Filter {netbootGUID -eq $NetbootGUID}
            if ($CheckMAC -ne $null) {
              $ErrorLog += "MAC Address is tied to $($CheckMAC.Name);"
            }
          }
          else {
            $ErrorLog += "MAC Address is not properly formatted;"
          }

          # Check if given keywords can find an existing OU on the AD
          if ($Computer.OU -eq "") {
            $PotentialOUs = ""
            $ErrorLog += "No OU given;"
          }
          else {
            $PotentialOUs = @()
            $OUKeywords = $Computer.OU -split " "
            ForEach ($Keyword in $OUKeywords) {

              # Lists every OU on the AD that has the given keyword
              $ValidOUPath = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase $StartingOU -Properties CanonicalName -PipelineVariable OU |
              Where-Object {$OU.Name -like $Keyword -or $OU.DistinguishedName -match $Keyword -or $OU.CanonicalName -match $Keyword} |
              Select-Object -ExpandProperty DistinguishedName

              # Performing an Intersection of all keywords to narrow down potential OU choices
              if ($PotentialOUs.Count -eq 0 -and $ValidOUPath -ne "") {
                $PotentialOUs = $ValidOUPath
              }
              else {
                $PotentialOUs = $PotentialOUs | ?{$ValidOUPath -contains $_}
              }
            }

            if ($PotentialOUs.Count -gt 1) {
              $ErrorLog += "Couldn't narrow search down to one OU;"
            }
            elseif ($PotentialOUs.Count -eq 0) {
              $ErrorLog += "Could not find an OU you have access to matching those keywords;"
            }
          }

          # Check if entered DeploymentServer value can match to a Distribution Point
          if ($Computer.DP -eq "") {
            $ErrorLog += "No Deployment Server given;"
          }
          else {
            $NetbootMFP = $DistributionPoints | Where-Object {$_ -match $Computer.DP}
            if ($NetbootMFP -eq $null) {
              $ErrorLog += "The Deployment Server is not valid;"
            }
            else {
              $DeploymentServer = $NetbootMFP -replace '\\'
            }
          }

          $ComputerLog = [PSCustomObject]@{
            COMPUTERNAME = $Computer.Name
            OUPATH = $PotentialOUs
            ADDEDBY = $([Environment]::UserName)
            DATEADDED = "N/A"
            SUCCESSFUL = $False
            ERRORS = ''
          }

          if ($ErrorLog -eq "") {
            try {
              New-ADComputer -Name $Computer.Name -SamAccountName $Computer.Name -Description $Computer.Desc -OtherAttributes @{'netbootGUID' = $NetbootGUID; 'netbootMachineFilePath' = $DeploymentServer} -Path $PotentialOUs
              $CheckAddition = Get-ADComputer -Identity $Computer.Name
              if ($CheckAddition) {
                $Success += 1
                $ComputerLog.DATEADDED = Get-Date
                $ComputerLog.SUCCESSFUL = $True
              }
            }
            catch {
              $Failed += 1
              $ComputerLog.ERRORS = "Could not create object. You might not have Create Item permissions for OU."
            }
          }
          else {
            $Failed += 1
            $ErrorLog = $ErrorLog.SubString(0, $ErrorLog.Length - 1)
            $ComputerLog.ERRORS = $ErrorLog
          }
          $ComputersAdded += $ComputerLog
        }

        $StopWatch.Stop()
        Get-Process Excel -ErrorAction SilentlyContinue | Select-Object -Property ID | Stop-Process
        Start-Sleep -Milliseconds 300
        $ComputersAdded | Export-CSV -Path $LogFilePath -NoTypeInformation -Append
        Write-Host "`nImport of file $CSVFile is Finished" -ForegroundColor $FGSuccess -BackgroundColor $BGBold
        Write-Host "- $Success of $TotalComputers objects were imported successfully." -ForegroundColor $FGNotice -BackgroundColor $BGBold
        Write-Host "- $Failed of $TotalComputers objects were unable to be imported." -ForegroundColor $FGNotice -BackgroundColor $BGBold
        Write-Host "`nCheck the log file listed below for more information."-ForegroundColor $FGHeader -BackgroundColor $BGBold
        Write-Host "$LogFilePath" -ForegroundColor $FGLabel -BackgroundColor $BGBold
        Write-Host "`nThis operation took $($StopWatch.Elapsed.TotalSeconds) seconds."
        Write-Host "`nPress Enter to quit." -ForegroundColor $FGInput -BackgroundColor $BGInput
        Read-Host
      }
    }
    else {
      Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
    }
  } while ($Method -notin 1..3)
}