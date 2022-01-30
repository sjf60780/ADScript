function Add-ADObject
{
  Clear-Host
  Write-Host "`nAdding Active Directory Computer Objects`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
  Write-Host "This script will allow you to add new Computer AD objects to the SGF.EDUBEAR.NET domain."
  Write-Host "To get started, please select one of the options listed below."

  Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
  Write-Host "- Enter '1' to manually navigate to the desired OU path."
  Write-Host "- Enter '2' to find OU using DistinguishedName, CanonicalName, or select Keywords."
  Write-Host "- Press ENTER to return to the main menu.`n"

  # Logic for handling user's input
  do
  {
    Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
    $Method = Read-Host
    if ($Method -eq "")
    {
      Break
    }
    elseif ($Method -eq 1 -or $Method -eq 2)
    {
      $OUPath = $StartingOU
      $Confirmation = 0
      do
      {
        # Logic for manually navigating to desired OU
        if ($Method -eq 1 -and ($Confirmation -eq 1 -or $Confirmation -eq 0))
        {
          $OUPath = Show-ADForest "Finding Computer OU Location Manually" "Add"
        }

        # Logic for entering Keywords to find desired OU
        elseif ($Method -eq 2 -and ($Repeat -eq 2 -or $Confirmation -eq 1 -or $Confirmation -eq 0))
        {
          Clear-Host
          Write-Host "`nFinding OU using Distinguished Name (DN), Canonical Name (CN), or select Keywords (KW)`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
    			Write-Host "When entering any of these values, please use the following formats..." -ForegroundColor $FGNotice -BackgroundColor $BGBold
    			Write-Host "- DN: OU=USIT,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
    			Write-Host "- CN: SGF.EDUBEAR.NET/CUSTOM/COMPUTERS/WORKSTATIONS/USIT"
    			Write-Host "- KW: USIT WORKSTATIONS"

          [bool]$Done = $FALSE
          do
          {
      			Write-Host "`nEnter the DN, CN, or KWs for the desired OU:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
      			$EnteredKeywords = Read-Host
      			[array]$SearchKeywords = $EnteredKeywords.ToUpper() -split " "

            $SearchResults = Search-ADForest "OrganizationalUnit" $SearchKeywords

            if ($SearchResults -ne $NULL)
            {
              Write-Host "`nFound $($SearchResults[-1].ID) potential OU path(s)." -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine

              $SearchResults | Format-Table | Out-Host

              Write-Host "OPTIONS" -ForegroundColor $FGInput -BackgroundColor $BGInput
              Write-Host "- Enter '^' to try again if the desired OU is not listed."
              Write-Host "- Enter the ID number to select the respective OU location.`n"

              # Logic for handling user input
              $Valid = $FALSE
              do
              {
                Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
                $Selection = Read-Host
                if ($Selection -match "^\d+$" -and $Selection -in 1..$SearchResults.Count)
                {
                  $OUPath = $SearchResults[$Selection - 1].DISTINGUISHEDNAME
                  try
                  {
                    New-ADComputer -Name "TEST-AD-PERMISSIONS" -SamAccountName "TEST-AD-PERMISSIONS" -Path $OUPath
                    Remove-ADComputer -Identity "TEST-AD-PERMISSIONS" -Confirm:$FALSE
                  }
                  catch
                  {
                    Write-Host "`nYou do not have permissions to add computer objects to this OU. Please select another location.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  }
                  $Done = $TRUE
                  $Valid = $TRUE
                }
                elseif ($Selection -eq '^')
                {
                  Clear-Host
                  Write-Host "`nFinding OU using Distinguished Name (DN), Canonical Name (CN), or select Keywords (KW)`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            			Write-Host "When entering any of these values, please use the following formats..." -ForegroundColor $FGNotice -BackgroundColor $BGBold
            			Write-Host "- DN: OU=USIT,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
            			Write-Host "- CN: SGF.EDUBEAR.NET/CUSTOM/COMPUTERS/WORKSTATIONS/USIT"
            			Write-Host "- KW: USIT WORKSTATIONS`n"

                  $Done = $FALSE
                  $Valid = $TRUE
                }
                else
                {
                  Write-Host "`nYou did not enter a selection. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                }
              } while (!$Valid)
            }
            else
            {
              Write-Host "`nThe Keyword(s) you entered could not find a Computer Object. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
              Write-Host "* If the Keyword(s) represent an existing Computer Object, you may not have permissions to add objects to it." -ForegroundColor $FGNotice
              Write-Host "* You might try putting in fewer Keywords as the search tries to find Computer Objects with every Keyword entered." -ForegroundColor $FGNotice
            }
          } while (!$Done)
        }

        # Get Computer Name Loop
        if ($Confirmation -eq 2 -or $Confirmation -eq 0 -or $Confirmation -eq "")
        {
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
          Write-Host " $DeploymentServer"

          # Check if the entered Computer Name is already tied to a Computer Object on the AD
          $Done = $FALSE
          do
          {
            Write-Host "`nEnter a Name for the new Computer Object:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $ComputerName = Read-Host
            if ($ComputerName -eq "")
            {
              Write-Host "`nYou did not enter a Computer Name. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($ComputerName -match '^\.|([\\/:\*\?"<>\|])+')
            {
              Write-Host `n'Computer Name cannot start with a . and\or contain the following: \ / : * ? " < > |'`n -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($ComputerName.Length -notin 1..15)
            {
              Write-Host "`nComputer Name must be between 1 and 15 characters.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($ComputerName.ToUpper() -in $ReservedWords)
            {
              Write-Host "`n$($ComputerName.ToUpper()) is a Reserved Word and can't be a Computer Name." -ForegroundColor $FGWarning -BackgroundColor $BGBold
              Write-Host "The following words are Reserved Words, as deemed by Microsoft:" -ForegroundColor $FGNotice -BackgroundColor $BGBold
              Write-Host ($ReservedWords -join ", ")`n -ForegroundColor $FGNotice -BackgroundColor $BGBold
            }
            else
            {
              $ComputerName = $ComputerName.ToUpper()
              $CheckName = Get-ADObject -Filter "ObjectClass -eq 'computer' -and Name -eq '$($ComputerName)'"
              if ($CheckName -eq $NULL)
              {
                $Done = $TRUE
              }
              else
              {
                Write-Host "`nThere is already a Computer Object with that Name:`n$($CheckName.DistinguishedName)`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
            }
          } while (!$Done)
        }

        # Get Computer MAC Address Loop
        if ($Confirmation -eq 3 -or $Confirmation -eq 0 -or $Confirmation -eq "")
        {
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
          $Done = $FALSE
          do
          {
            Write-Host "Enter the MAC Address for the Computer Object (Ex. 1A2B3C4D5E6F):`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $EnteredMAC = Read-Host
            if ($EnteredMAC -eq "")
            {
              Write-Host "`nYou did not enter a MAC Address. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            elseif ($EnteredMAC -match ($ValidMACPatterns -join '|'))
            {
              $MACAddress = $EnteredMAC.ToUpper() -replace '\W'
              [guid]$NetbootGUID = "00000000-0000-0000-0000-$MACAddress"
              $CheckMAC = Get-ADComputer -Filter {netbootGUID -like $NetbootGUID}
              if ($CheckMAC -eq $NULL) {
                $Done = $TRUE
              }
              else
              {
                Write-Host "`nThere is already a Computer Object with that MAC Address:`n$($CheckMAC.DistinguishedName)`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
            }
            else
            {
              Write-Host "`nThe entered MAC Address is not properly formatted." -ForegroundColor $FGWarning -BackgroundColor $BGBold
              Write-Host "A properly formatted MAC Address is a series of 12 Hexadecimal characters (0-9, A-F)" -ForegroundColor $FGNotice -BackgroundColor $BGBold
              Write-Host "The characters can be delimited (or not) at every two characters by a ':', '-', '.', or ' '.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
            }
          } while (!$Done)
        }

        # Get Computer Description Loop
        if ($Confirmation -eq 4 -or $Confirmation -eq 0 -or $Confirmation -eq "")
        {
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
        if ($Confirmation -eq 5 -or $Confirmation -eq 0 -or $Confirmation -eq "")
        {
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
          Write-Host " $DeploymentServer`n"

          # Select appropriate Deployment Server
          $DeploymentTable = @()
          For ([int]$i = 0; $i -lt $DistributionPoints.Count; $i++)
          {
            $TableRow =
            [PSCustomObject]@{
              ID = $i+1
              DISTRIBUTIONPOINT = $DistributionPoints[$i].ToUpper()
            }
            $DeploymentTable += $TableRow
          }
          $DeploymentTable | Format-Table | Out-Host

          # Logic for handling user input
          $Done = $FALSE
          do
          {
            Write-Host "Enter the ID of the Distribution Point for this Computer Object:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $Selection = Read-Host
            if ($Selection -eq "")
            {
              Write-Host "`nYou did not enter a selection. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold

            }
            elseif ($Selection -match "^\d+$" -and $Selection -in 1..$DeploymentTable.Count)
            {
              $DistributionPoint = $DeploymentTable[$Selection - 1].DISTRIBUTIONPOINT
              $NetbootMFP = $DistributionPoints | Where-Object {$_ -match $DistributionPoint}
              $DeploymentServer = $NetbootMFP.ToUpper() -replace "\\"
              $Done = $TRUE
            }
            else
            {
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
        Write-Host "- Press ENTER to confirm these values.`n"

        # Logic for handling user input
        $Valid = $FALSE
        do
        {
          Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
          $Confirmation = Read-Host
          if ($Confirmation -in 1..5 -or $Confirmation -eq "")
          {
            $Valid = $TRUE
          }
          else
          {
            Write-Host "`nYour selection doesn't exist. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
          }
        } while (!$Valid)

        if ($Confirmation -eq "")
        {
          try
          {
            New-ADComputer -Name $ComputerName -SamAccountName $ComputerName -Description $ComputerDescription -OtherAttributes @{'netbootGUID' = $NetbootGUID; 'netbootMachineFilePath' = $DeploymentServer} -Path $OUPath

            # Object Information to export to CSV
            $ComputerLog =
            [PSCustomObject]@{
              ACTION = "ADD"
              COMPUTERNAME = $ComputerName
              OUPATH = $OUPath
              ACTIONSBY = $([Environment]::UserName)
              DATEADDED = Get-Date
              DATEMODIFIED = $NULL
              DATEDELETED = $NULL
              SUCCESSFUL = $TRUE
              CHANGES = $NULL
              ERRORS = $NULL
            }
            $Global:ComputerArchive += $ComputerLog

            Write-Host "`nThe computer $ComputerName has been added to the Active Directory successfully.`n" -ForegroundColor $FGSuccess -BackgroundColor $BGBold
          }
          catch
          {
            Write-Host "`nSomething went wrong and the Computer Object was unable to be added to the Active Directory. Please close the script and try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            Write-Host "If this continues to happen, please contact your system administrator.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
          }
        }

        Write-Host "OPTIONS" -BackgroundColor $BGBold
        if ($ComputerLog -ne $NULL) {
          Write-Host "- Enter '1' to add another Computer Object to the same OU location."
          Write-Host "- Enter '2' to add another Computer Object, but search for a different OU location."
        }
        Write-Host "- Press ENTER if you are finished adding Computer Objects.`n"

        # Logic for handling user input
        $Valid = $FALSE
        do {
          $ComputerLog = $NULL
          Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
          $Repeat = Read-Host

          if ($Repeat -match "[1-2]") {
            $Valid = $TRUE

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
            $Valid = $TRUE
            Break
          }
          else {
            Write-Host "Your selection doesn't exist. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
          }
        } while (!$Valid)
      } while ($Repeat -match "[1-2]" -or $Confirmation -in 1..5)
    }
    else
    {
      Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
    }
  } while ($Method -notin 1..2)
}
