function Edit-ADObject
{
  Clear-Host
  Write-Host "`nEditing Active Directory Computer Objects" -ForegroundColor $FGTitle -BackgroundColor $BGBold
  Write-Host "This script will allow you to edit existing Computer AD objects on the SGF.EDUBEAR.NET domain."
  Write-Host "To get started, please select one of the options listed below."

  Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
  Write-Host "- Enter '1' to find the desired object by navigating the AD site."
  Write-Host "- Enter '2' to enter keywords to search for the desired object."
  Write-Host "- Press ENTER to return to the main menu.`n"

  # Logic for handling user's input
  do {
    Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
    $Method = Read-Host
    if ($Method -eq "")
    {
      Break
    }
    elseif ($Method -eq 1 -or $Method -eq 2)
    {
      $OUPath = $StartingOU
      [Array]$EditObjects = @()

      # Logic for manually navigating to desired OU
      if ($Method -eq 1)
      {
        $OUPath = Show-ADForest "Navigating Active Directory Site for Computer Objects to Modify" "Edit" -ShowComputers

        $PotentialObjects = Get-ADComputer -Filter "Name -like '*'" -SearchBase $OUPath -SearchScope OneLevel -Properties netbootGuid, Description, netbootMachineFilePath

        $ComputerTable = @()
        For ([int]$i = 0; $i -lt $PotentialObjects.Count; $i++)
        {
          if ($PotentialObjects[$i].netbootGUID -ne $NULL)
          {
            $netbootGUID = ($PotentialObjects[$i].netbootGUID | ForEach-Object ToString X2) -join ""
            $ComputerMAC = $netbootGUID.SubString($netbootGUID.length - 12)
          }
          else
          {
            $ComputerMAC = $NULL
          }

          $ShortOU = ($PotentialObjects[$i].DistinguishedName -Split ',' | Where-Object {$_ -like '*OU*'}).SubString(3)
          [array]::Reverse($ShortOU)
          $ShortOU = $ShortOU -join "/"

          if ($PotentialObjects[$i].netbootMachineFilePath -ne $NULL)
          {
            $ShortDP = (($PotentialObjects[$i].netbootMachineFilePath -replace "01.MISSOURISTATE.EDU") -split "-")[2].ToUpper()
          }
          else
          {
            $ShortDP = $NULL
          }

          $TableRow =
          [PSCustomObject]@{
            EDIT = "[ ]"
            ID = $i + 1
            NAME = $PotentialObjects[$i].Name
            MAC = $ComputerMAC
            DP = $ShortDP
            DESC = $PotentialObjects[$i].Description
            OU = $ShortOU
          }
          $ComputerTable += $TableRow
        }

        # Select Computer Object if there are multiple available
        $Done = $FALSE
        do
        {
          Clear-Host
          Write-Host "`nNavigating Active Directory Site for Computer Objects to Modify" -ForegroundColor $FGTitle -BackgroundColor $BGBold
          Write-Host "`nCurrent AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
      		Write-Host "`t$OUPath`n" -NoNewLine
          Write-Host "`nFound $($ComputerTable[-1].ID) potential Computer Object(s)." -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine

          $ComputerTable | Format-Table | Format-Color @{'[X]' = 'Green'}

          # List available options for current OU
          Write-Host "OPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
          if ($ComputerTable.Count -gt 0)
          {
            Write-Host "- Enter '+' to edit currently selected Computer Object(s)"
            Write-Host "- Enter the ID number(s) to select/deselect a Computer Object(s) to edit"
            Write-Host "* To select multiple objects, enter all desired ID numbers separated by a ',' or space.`n" -ForegroundColor $FGNotice
          }

          # Logic for handling user's input
          [bool]$Valid = $FALSE
          do
          {
            Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $Selection = Read-Host
            if ($Selection -eq '+')
            {
              ForEach ($Computer in $ComputerTable)
              {
                if ($Computer.SEL -eq "[X]")
                {
                  $EditObjects += $Computer.NAME
                }
              }

              # Ensure at least one Computer Object was selected
              if ($EditObjects.Count -gt 0)
              {
                $Done = $TRUE
                $Valid = $TRUE
              }
              else
              {
                Write-Host "`nYou did not select a Computer Object." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                Write-Host "To select a Computer Object to edit, please enter an ID number listed above.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
                Start-Sleep 2
              }
            }
            elseif ($Selection -match "\d+")
            {
              [array]$Choices = $Selection -split {$_ -eq " " -or $_ -eq ","}
              ForEach ($ID in $Choices)
              {
                if ($ID -in 1..$ComputerTable.Count)
                {
                  if ($ComputerTable[$ID - 1].SEL -eq "[ ]")
                  {
                    $ComputerTable[$ID - 1].SEL = "[X]"
                  }
                  else
                  {
                    $ComputerTable[$ID - 1].SEL = "[ ]"
                  }
                  $Valid = $TRUE
                }
              }
            }
            elseif ($Selection -eq "")
            {
              Write-Host "`nYou did not enter a selection. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              Start-Sleep 2
            }
            else
            {
              Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              Start-Sleep 2
            }
          } while (!$Valid)
        } while (!$Done)
      }

      # Logic for entering Keywords to find desired Computer Object(s)
      elseif ($Method -eq 2)
      {
        [bool]$Done = $FALSE
        do
        {
          if ($EnteredKeywords -eq $NULL)
          {
            Clear-Host
            Write-Host "`nFinding Computer Object(s) using Name Keywords or MAC Address.`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "When entering any of these values, please use the following formats..." -ForegroundColor $FGNotice -BackgroundColor $BGBold
            Write-Host "- KW : CHEK 135"
            Write-Host "- MAC: ABCDEF123456, AB CD EF 12 34 56, AB:CD:EF:12:34:56, AB-CD-EF-12-34-56"
            do
            {
              Write-Host "`nEnter the Name or MAC Address of the desired Computer Object:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
              $EnteredKeywords = Read-Host
              [array]$SearchKeywords = $EnteredKeywords.ToUpper() -split " "

              $SearchResults = Search-ADForest "ComputerObject" $SearchKeywords

              if ($SearchResults -eq $NULL)
              {
                Write-Host "`nThe Keyword(s) you entered could not find a Computer Object. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                Write-Host "* If the Keyword(s) represent an existing Computer Object, you may not have permissions to add objects to it." -ForegroundColor $FGNotice
                Write-Host "* You might try putting in fewer Keywords as the search tries to find Computer Objects with every Keyword entered.`n" -ForegroundColor $FGNotice
              }
            } while ($SearchResults -eq $NULL)
          }

          if ($SearchResults[-1].ID -gt 1)
          {
            Clear-Host
            Write-Host "`nSearching for Computer Objects matching the search of:" -ForegroundColor $FGTitle -BackgroundColor $BGBold -NoNewLine
            Write-Host " $EnteredKeywords"
            Write-Host "`nFound $($SearchResults[-1].ID) potential Computer Object(s)." -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine

            $SearchResults | Format-Table SEL, ID, NAME, MAC, OUPATH | Format-Color @{'[X]' = 'Green'}

            Write-Host "OPTIONS" -ForegroundColor $FGInput -BackgroundColor $BGInput
            Write-Host "- Enter '+' to edit the currently selected Computer Objects"
            Write-Host "- Enter '^' to try again if the desired Computer Object is not listed."
            Write-Host "- Enter the ID number(s) to select/deselect the respective Computer Object(s)."
            Write-Host "* To select multiple objects, enter all desired ID numbers separated by a ',' or space.`n" -ForegroundColor $FGNotice

            # Logic for handling user input
            $Valid = $FALSE
            do
            {
              Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
              $Selection = Read-Host
              if ($Selection -eq '+')
              {
                ForEach ($Computer in $SearchResults)
                {
                  if ($Computer.SEL -eq "[X]")
                  {
                    $EditObjects += $Computer.NAME
                  }
                }

                # Ensure at least one Computer Object was selected
                if ($EditObjects.Count -gt 0)
                {
                  $Done = $TRUE
                  $Valid = $TRUE
                }
                else
                {
                  Write-Host "`nYou did not select a Computer Object." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  Write-Host "To select a Computer Object to edit, please enter at least one of the ID numbers listed above.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
                  Start-Sleep 2
                }
              }
              elseif ($Selection -eq '^')
              {
                $Done = $FALSE
                $Valid = $TRUE
                $EnteredKeywords = $NULL
              }
              elseif ($Selection -match "\d+")
              {
                [array]$Choices = $Selection -split {$_ -eq " " -or $_ -eq ","}
                ForEach ($ID in $Choices)
                {
                  if ($ID -in 1..$SearchResults.Count)
                  {
                    if ($SearchResults[$ID - 1].SEL -eq "[ ]")
                    {
                      $SearchResults[$ID - 1].SEL = "[X]"
                    }
                    else
                    {
                      $SearchResults[$ID - 1].SEL = "[ ]"
                    }
                  }
                  else
                  {
                    Write-Host "`nThe selected ID $($ID) does not exist." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                    Start-Sleep 2
                  }
                }
                $Valid = $TRUE
              }
              elseif ($Selection -eq "")
              {
                Write-Host "`nYou did not enter a selection. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                Start-Sleep 2
              }
              else
              {
                Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                Start-Sleep 2
              }
            } while (!$Valid)
          }
          elseif ($SearchResults[-1].ID -eq 1)
          {
            $Done = $TRUE
            $EditObjects += $SearchResults[0].NAME
          }
        } while (!$Done)
      }

      # Retrieve attributes for all computer objects to edit
      ForEach ($ObjectName in $EditObjects)
      {
        $ObjectInformation = Get-ADComputer -Identity $ObjectName -Properties Description, netbootGUID, netbootMachineFilePath

        $OUPath = ($ObjectInformation.DistinguishedName -Split ',', 2)[1]
        $ComputerName = $ObjectInformation.Name
        [String]$netbootGUID = ($ObjectInformation.netbootGUID | ForEach-Object ToString X2) -join ""
        $ComputerMAC = $netbootGUID.SubString($netbootGUID.length - 12)
        $ComputerDesc = $ObjectInformation.Description
        $ComputerDP = $ObjectInformation.netbootMachineFilePath

        $NewOU = $NewName = $NewMAC = $NewDescription = $NewDP = $NULL

        [System.Collections.ArrayList]$Changes = @()
        [bool]$Editing = $TRUE
        do
        {
          # Logic for handling user input
          $Valid = $FALSE
          do
          {
            # List commonly modified attributes for Computer Objects (OU Location, Computer Name, Description netbootGUID, Deployment Server)
            Clear-Host
            Write-Host "`nThese are the current attributes of the selected Computer Object.`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "Any prospective changes to the object will be listed in Green." -ForegroundColor $FGNotice -BackgroundColor $BGBold
            Write-Host "`nCURRENT VALUES" -ForegroundColor $FGHeader -BackgroundColor $BGBold

            Write-Host "Computer OU  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewOU -and $NewOU -ne $OUPath) { Write-Host " $NewOU" -ForegroundColor $FGSuccess }
            else { Write-Host " $OUPath" }

            Write-Host "Computer Name:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewName) { Write-Host " $NewName" -ForegroundColor $FGSuccess }
            else { Write-Host " $ComputerName" }

            Write-Host "Computer MAC :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewMAC) { Write-Host " $NewMAC" -ForegroundColor $FGSuccess }
            else { Write-Host " $ComputerMAC" }

            Write-Host "Computer Desc:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewDescription) { Write-Host " $NewDescription" -ForegroundColor $FGSuccess }
            else { Write-Host " $ComputerDesc" }

            Write-Host "Computer DP  :" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewDP) { Write-Host " $NewDP" -ForegroundColor $FGSuccess }
            else { Write-Host " $ComputerDP" }

            Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
            Write-Host "- Enter '1' to change the Current OU Path."
            Write-Host "- Enter '2' to change the Computer Name."
            Write-Host "- Enter '3' to change the Computer MAC."
            Write-Host "- Enter '4' to change the Computer Description."
            Write-Host "- Enter '5' to change the Computer Distribution Point."
            Write-Host "- Press ENTER to confirm changes."
            Write-Host "* To select multiple options, enter all desired option numbers separated by a ',' or space.`n" -ForegroundColor $FGNotice

            Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $Selection = Read-Host
            if ($Selection -eq "")
            {
              if ($NewOU -or $NewName -or $NewMAC -or $NewDescription -or $NewDP)
              {
                $Valid = $TRUE
                $Editing = $FALSE
              }
              else
              {
                Write-Host "It appears no changes have been made. To confirm, please enter '*'." -ForegroundColor $FGNotice -BackgroundColor $BGBold
                Write-Host "Otherwise, please enter one of the options listed above.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
                Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
                $ConfirmChanges = Read-Host
                if ($ConfirmChanges -eq '*')
                {
                  $Valid = $TRUE
                  $Editing = $FALSE
                }
              }
            }
            elseif ($Selection -match "\d+")
            {
              [array]$Choices = $Selection -split {$_ -eq " " -or $_ -eq ","}
              ForEach ($Option in $Choices)
              {
                if ($Option -notin $Changes -and $Option -match "[1-5]")
                {
                  $Changes.Add($Option) | Out-Null
                }
              }
              if ($Changes.Count -gt 0)
              {
                $Valid = $TRUE
              }
              else
              {
                Write-Host "`nYour selection doesn't exist. Please select at least one of the options above and try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
            }
            else
            {
              Write-Host "`nYour selection doesn't exist. Please select at least one of the options above and try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
          } while (!$Valid)

          # Logic for changing the Computer's OU location
          if (1 -in $Changes -and $Method -eq 1)
          {
            Clear-Host
            $NewOU = Show-ADForest "Edit OU for Computer Object by Searching for Location Manually" "Add" $OUPath
            if ($NewOU -eq $OUPath)
            {
              $NewOU = $FALSE
            }
          }

          elseif (1 -in $Changes -and $Method -eq 2)
          {
            do
            {
              Clear-Host
              Write-Host "`nEdit OU for Computer Object using Distinguished Name (DN), Canonical Name (CN), or select Keywords (KW)`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
              Write-Host "Computer Object's Current OU Path:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
              Write-Host " $OUPath"

        			Write-Host "`nWhen entering any of these values, please use the following formats..." -ForegroundColor $FGNotice -BackgroundColor $BGBold
        			Write-Host "- DN: OU=USIT,OU=WORKSTATIONS,OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"
        			Write-Host "- CN: SGF.EDUBEAR.NET/CUSTOM/COMPUTERS/WORKSTATIONS/USIT"
        			Write-Host "- KW: USIT WORKSTATIONS`n"

              do
              {
          			Write-Host "Enter the DN, CN, or KWs for the desired OU:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
          			$EnteredKeywords = Read-Host
          			$SearchKeywords = $EnteredKeywords.ToUpper() -split " "

                $SearchResults = Search-ADForest "OrganizationalUnit" $SearchKeywords

                if ($SearchResults -ne $NULL)
                {
                  Write-Host "`nFound $($SearchResults[-1].ID) potential OU path(s)." -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine

                  $SearchResults | Format-Table

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
                      $NewOU = $SearchResults[$Selection - 1].DISTINGUISHEDNAME
                      if ($NewOU -eq $OUPath) { $NewOU = $FALSE }
                      else
                      {
                        try
                        {
                          New-ADComputer -Name "TEST-AD-PERMISSIONS" -SamAccountName "TEST-AD-PERMISSIONS" -Path $NewOU
                          Remove-ADComputer -Identity "TEST-AD-PERMISSIONS" -Confirm:$FALSE
                          $Done = $TRUE
                          $Valid = $TRUE
                        }
                        catch
                        {
                          Write-Host "`nYou do not have permissions to add computer objects to this OU. Please select another location.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                          $NewOU = $NULL
                        }
                      }
                    }
                    elseif ($Selection -eq '^')
                    {
                      $Done = $FALSE
                      $Valid = $TRUE
                    }
                    else
                    {
                      Write-Host "`nYou did not enter a selection. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                    }
                  } while (!$Valid -and !$Done)
                }
                else
                {
                  Write-Host "`nThe Keyword(s) you entered could not find an Organizational Unit. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  Write-Host "* If the Keyword(s) represent an existing Organizational Unit, you may not have permissions to add objects to it." -ForegroundColor $FGNotice
                  Write-Host "* You might try putting in fewer Keywords as the search tries to find Organizational Units with every Keyword entered." -ForegroundColor $FGNotice
                }
              } while ($SearchResults -eq $NULL)
            } while ($NewOU -eq $NULL)
          }

          # Logic for changing the Computer's Name
          if (2 -in $Changes)
          {
            Clear-Host
            Write-Host "`nEdit Name for Computer Object" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "`nCurrent Computer Name:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewName) { Write-Host " $NewName"}
            else { Write-Host " $ComputerName"}

            do
            {
              Write-Host "`nEnter a new Name for the Computer Object:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
              $EnteredName = Read-Host
              if ($EnteredName -eq "")
              {
                Write-Host "`nYou did not enter a Computer Name. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
              elseif ($EnteredName -match '^\.|([\\/:\*\?"<>\|])+')
              {
                Write-Host `n'Computer Name cannot start with a . and\or contain the following: \ / : * ? " < > |'`n -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
              elseif ($EnteredName.Length -notin 1..15)
              {
                Write-Host "`nComputer Name must be between 1 and 15 characters.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
              elseif ($EnteredName.ToUpper() -in $ReservedWords)
              {
                Write-Host "`n$($EnteredName.ToUpper()) is a Reserved Word and can't be a Computer Name." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                Write-Host "The following words are Reserved Words, as deemed by Microsoft:" -ForegroundColor $FGNotice -BackgroundColor $BGBold
                Write-Host ($ReservedWords -join ", ")`n -ForegroundColor $FGNotice -BackgroundColor $BGBold
              }
              else
              {
                $EnteredName = $EnteredName.ToUpper()
                if ($EnteredName -eq $ComputerName) { $NewName = $FALSE }
                else
                {
                  $CheckName = Get-ADObject -Filter "ObjectClass -eq 'computer' -and Name -eq '$($EnteredName)'"
                  if ($CheckName -eq $NULL)
                  {
                    $NewName = $EnteredName
                  }
                  else
                  {
                    Write-Host "`nThere is already a Computer Object with that Name:`n$($CheckName.DistinguishedName)`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  }
                }
              }
            } while ($NewName -eq $NULL)
          }

          # Logic for changing the Computer's MAC
          if (3 -in $Changes)
          {
            Clear-Host
            Write-Host "`nEdit MAC Address for Computer Object" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "`nCurrent MAC Address:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewMAC) { Write-Host " $NewMAC" }
            else { Write-Host " $ComputerMAC" }

            do
            {
              Write-Host "`nEnter a new MAC Address for the Computer Object (Ex. 1A2B3C4D5E6F):`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
              $EnteredMAC = Read-Host
              if ($EnteredMAC -eq "")
              {
                Write-Host "`nYou did not enter a MAC Address. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
              elseif ($EnteredMAC -match ($ValidMACPatterns -join '|'))
              {
                $MACAddress = $EnteredMAC.ToUpper() -replace '\W'
                if ($MACAddress -eq $ComputerMAC) { $NewMAC = $FALSE }
                else
                {
                  [guid]$NetbootGUID = "00000000-0000-0000-0000-$MACAddress"
                  $CheckMAC = Get-ADComputer -Filter {netbootGUID -like $NetbootGUID}
                  if ($CheckMAC -eq $NULL)
                  {
                    $NewMAC = $MACAddress
                  }
                  else
                  {
                    Write-Host "`nThere is already a Computer Object with that MAC Address:`n$($CheckMAC.DistinguishedName)`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  }
                }
              }
              else
              {
                Write-Host "`nThe entered MAC Address is not properly formatted." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                Write-Host "A properly formatted MAC Address is a series of 12 Hexadecimal characters (0-9, A-F)" -ForegroundColor $FGNotice -BackgroundColor $BGBold
                Write-Host "The characters can be delimited (or not) at every two characters by a ':', '-', '.', or ' '.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
              }
            } while ($NewMAC -eq $NULL)
          }

          # Logic for changing the Computer's Description
          if (4 -in $Changes)
          {
            Clear-Host
            Write-Host "`nEdit Description for Computer Object" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "`nCurrent Description:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewDescription) { Write-Host " $NewDescription"}
            else { Write-Host " $ComputerDesc" }

            Write-Host "`nEnter a new Description for the Computer Object Ex. 'Boomer Bear (CARR 0123)':`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
            $NewDescription = Read-Host
            if ($NewDescription -eq $ComputerDesc) { $NewDescription = $FALSE }
          }

          # Logic for changing the Computer's Deployment Server
          if (5 -in $Changes)
          {
            Clear-Host
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

            Write-Host "`nEdit Distribution Point for Computer Object" -ForegroundColor $FGTitle -BackgroundColor $BGBold
            Write-Host "`nCurrent Distribution Point:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
            if ($NewDP) { Write-Host " $NewDP" }
            else { Write-Host " $ComputerDP" }

            $DeploymentTable | Format-Table

            # Logic for handling user input
            do
            {
              Write-Host "`nEnter the ID of the Distribution Point for this Computer Object:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
              $Selection = Read-Host
              if ($Selection -eq "")
              {
                Write-Host "`nYou did not enter a selection. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
              elseif ($Selection -match "^\d+$" -and $Selection -in 1..$DeploymentTable.Count)
              {
                $DistributionPoint = $DeploymentTable[$Selection - 1].DISTRIBUTIONPOINT
                $NetbootMFP = $DistributionPoints | Where-Object {$_ -match $DistributionPoint}
                $NewDP = $NetbootMFP.ToUpper() -replace "\\"
                if ($NewDP -eq $ComputerDP) { $NewDP = $FALSE }
              }
              else
              {
                Write-Host "`nYour selection doesn't exist. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
              }
            } while ($NewDP -eq $NULL)
          }

          if ($Changes.Count -gt 0)
          {
            $Changes.Clear()
          }
        } while ($Editing)

        #Perform changes and log them here
        $Edits = ""
        if ($NewMAC -or $NewDescription -or $NewDP)
        {
          $ReplaceValues = @{}
          if ($NewMAC)
          {
            $ReplaceValues["netbootGUID"] = $NetbootGUID
            $Edits += "MAC Address changed;"
          }
          if ($NewDescription)
          {
            $ReplaceValues["description"] = $NewDescription
            $Edits += "Description changed;"
          }
          if ($NewDP)
          {
            $ReplaceValues["netbootMachineFilePath"] = $NewDP
            $Edits += "Deployment Server changed;"
          }

          try
          {
            Set-ADComputer -Identity $ObjectInformation.DistinguishedName -Replace $ReplaceValues
            $ReplaceValues.Clear()
          }
          catch
          {
            Write-Host "`nSomething went wrong and the Computer Object was unable to be edited. Please close the script and try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            Write-Host "If this continues to happen, please contact your system administrator.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
          }
        }
        if ($NewOU -or $NewName)
        {
          $ObjectPlaceholder = Get-ADComputer -Identity $ComputerName -Properties Description, netbootGUID, netbootMachineFilePath
          if ($NewOU)
          {
            $OUPath = $NewOU
            $Edits += "OU Path changed;"
          }
          if ($NewName)
          {
            $ComputerName = $NewName
            $Edits += "Computer Name changed;"
          }

          try
          {
            Remove-ADComputer -Identity $ObjectPlaceholder.Name -Confirm:$FALSE
            New-ADComputer -Name $ComputerName -SamAccountName $ComputerName -Description $ObjectPlaceholder.Description -OtherAttributes @{'netbootGUID' = $ObjectPlaceholder.netbootGUID; 'netbootMachineFilePath' = $ObjectPlaceholder.netbootMachineFilePath} -Path $OUPath
          }
          catch
          {
            Write-Host "`nSomething went wrong and the Computer Object was unable to be edited. Please close the script and try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            Write-Host "If this continues to happen, please contact your system administrator.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
          }
        }

        # Object Information to export to CSV
        $ComputerLog =
        [PSCustomObject]@{
          ACTION = "MOD"
          COMPUTERNAME = $ComputerName
          OUPATH = $OUPath
          ACTIONSBY = $([Environment]::UserName)
          DATEADDED = $NULL
          DATEMODIFIED = Get-Date
          DATEDELETED = $NULL
          SUCCESSFUL = $TRUE
          CHANGES = $Edits
          ERRORS = $NULL
        }
        $Global:ComputerArchive += $ComputerLog
      }
    }
    else
    {
      Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
    }
  } while ($Method -notin 1..2)
}
