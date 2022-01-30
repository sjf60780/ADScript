function Remove-ADObject
{
  Clear-Host
  Write-Host "`nRemoving Active Directory Computer Objects" -ForegroundColor $FGTitle -BackgroundColor $BGBold
  Write-Host "This script will allow you to remove existing Computer AD objects from the SGF.EDUBEAR.NET domain."
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
      [Array]$RemoveObjects = @()

      # Logic for manually navigating to desired OU
      if ($Method -eq 1)
      {
        $OUPath = Show-ADForest "Navigating Active Directory Site for Computer Objects to Remove" "Remove" -ShowComputers

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
            REMOVE = "[ ]"
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
          Write-Host "`nNavigating Active Directory Site for Computer Objects to Remove" -ForegroundColor $FGTitle -BackgroundColor $BGBold
          Write-Host "`nCurrent AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
      		Write-Host "`t$OUPath`n" -NoNewLine
          Write-Host "`nFound $($ComputerTable[-1].ID) potential Computer Object(s).`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine

          $ComputerTable | Format-Table | Format-Color @{'[X]' = 'Green'}

          # List available options
          Write-Host "OPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
          if ($ComputerTable.Count -gt 0)
          {
            Write-Host "- Enter '+' to remove currently selected Computer Object(s)"
            Write-Host "- Enter the ID number(s) to select/deselect a Computer Object(s) to remove"
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
                  $RemoveObjects += $Computer.NAME
                }
              }

              # Ensure at least one Computer Object was selected
              if ($RemoveObjects.Count -gt 0)
              {
                $Done = $TRUE
                $Valid = $TRUE
              }
              else
              {
                Write-Host "`nYou did not select a Computer Object." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                Write-Host "To select a Computer Object to remove, please enter an ID number listed above.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
              }
            }
            elseif ($Selection -match "\d+")
            {
              [array]$Choices = $Selection -split {$_ -eq " " -or $_ -eq ","}
              ForEach ($ID in $Choices)
              {
                if ([int]$ID -in 1..$ComputerTable.Count)
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
            }
            else
            {
              Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            }
            Read-Host
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
                Write-Host "* You might try putting in fewer Keywords as the search tries to find Computer Objects with every Keyword entered." -ForegroundColor $FGNotice
              }
            } while ($SearchResults -eq $NULL)
          }

          if ($SearchResults[-1].ID -gt 1)
          {
            Clear-Host
            Write-Host "`nSearching for Computer Objects matching the search of:" -ForegroundColor $FGTitle -BackgroundColor $BGBold -NoNewLine
            Write-Host " $EnteredKeywords"
            Write-Host "`nFound $($SearchResults[-1].ID) potential Computer Object(s).`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine

            $SearchResults | Format-Table SEL, ID, NAME, OUPATH | Format-Color @{'[X]' = 'Green'}

            Write-Host "OPTIONS" -ForegroundColor $FGInput -BackgroundColor $BGInput
            Write-Host "- Enter '+' to remove the currently selected Computer Objects"
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
                    $RemoveObjects += $Computer.NAME
                  }
                }

                # Ensure at least one Computer Object was selected
                if ($RemoveObjects.Count -gt 0)
                {
                  $Done = $TRUE
                  $Valid = $TRUE
                }
                else
                {
                  Write-Host "`nYou did not select a Computer Object." -ForegroundColor $FGWarning -BackgroundColor $BGBold
                  Write-Host "To select a Computer Object to remove, please enter at least one of the ID numbers listed above.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
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
            $RemoveObjects += $SearchResults[0].NAME
          }
        } while (!$Done)
      }

      # List all the Computer Objects slated to be removed
      Clear-Host
      Write-Host "`nComputer Objects slated to be removed.`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
      Write-Host "Double-check these objects and confirm if you wish for ALL of them to be deleted." -ForegroundColor $FGNotice -BackgroundColor $BGBold

      $RemoveObjects | Format-Table

      Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
      Write-Host "- Enter '*' to confirm deletion of the listed Computer Object(s)."
      Write-Host "- Enter '<' to ignore these objects and go back to the main menu.`n"

      $Valid = $FALSE
      do
      {
        Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
        $Confirmation = Read-Host
        if ($Confirmation -eq '*')
        {
          $Valid = $TRUE
          $Remove = $TRUE
        }
        elseif ($Confirmation -eq '<')
        {
          Break
        }
        else
        {
          Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
        }
      } while (!$Valid)

      # Remove all Computer Objects confirmed to removed
      if ($Remove -eq $TRUE)
      {
        ForEach ($ObjectName in $RemoveObjects)
        {
          try
          {
            Remove-ADComputer -Identity $ObjectName -Confirm:$False
          }
          catch
          {
            Write-Host "`nSomething went wrong and the Computer Object was unable to be removed from the Active Directory. Please close the script and try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
            Write-Host "If this continues to happen, please contact your system administrator.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
            Start-Sleep 3
          }

          # Object Information to export to CSV
          $ComputerLog =
          [PSCustomObject]@{
            ACTION = "REM"
            COMPUTERNAME = $ObjectName
            OUPATH = $NULL
            ACTIONSBY = $([Environment]::UserName)
            DATEADDED = $NULL
            DATEMODIFIED = $NULL
            DATEDELETED = Get-Date
            SUCCESSFUL = $TRUE
            CHANGES = "$ObjectName was deleted."
            ERRORS = $NULL
          }
          $Global:ComputerArchive += $ComputerLog
        }
      }
    }
    else
    {
      Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
    }
  } while ($Method -notin 1..2)
}
