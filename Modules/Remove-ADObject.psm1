function Remove-ADObject {

  Clear-Host

  Write-Host "`nRemoving Active Directory Computer Objects" -ForegroundColor $FGTitle -BackgroundColor $BGBold
  Write-Host "This script will allow you to remove existing Computer AD objects from the SGF.EDUBEAR.NET domain."
  #Write-Host "It also has the capability to remove multiple objects from a CSV file as well!"
  Write-Host "To get started, please select one of the options listed below."

  Write-Host "`nOPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
  Write-Host "- Enter '1' to manually navigate the forest to find the desired Computer Object."
  # Write-Host "- Enter '2' to find a Computer Object by its Name or MAC Address."
  # Write-Host "- Enter '3' to pass a CSV file through the script. (Useful when removing multiple machines)"
  Write-Host "- Press Enter to quit.`n"

  # # Logic for handling user's input
  # do {
  #   Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
  #   $Selection = Read-Host
  #   if ($Selection -eq "") {
  #     Exit
  #   }
  #   elseif ($Selection -eq 1) {
  #     [bool]$Manual = $True
  #     [bool]$Auto = $False
  #   }
  #   # elseif ($Selection -eq 2) {
  #   #   [bool]$Manual = $False
  #   #   [bool]$Auto = $False
  #   # }
  #   # elseif ($Selection -eq 3) {
  #   #   [bool]$Manual = $False
  #   #   [bool]$Auto = $True
  #   # }
  #   else {
  #     Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
  #   }
  # } while ($Manual -eq $null -or $Auto -eq $null)

  $OUPath = $StartingOU

# Logic for manually navigating to desired OU
  do {
    Clear-Host
    Write-Host "`nFinding Computer's Location Manually" -ForegroundColor $FGTitle -BackgroundColor $BGBold
    Write-Host "`nDefault AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
    Write-Host "`t$StartingOU"
    if ($OUPath -ne $StartingOU) {
      Write-Host "Parent  AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
      Write-Host "`t$ParentOU"
    }
    Write-Host "Current AD Location:" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
    Write-Host "`t$OUPath`n"

    # List how many Computer Objects are in the current OU path
    Write-Host "Number of Computer Objects in $OUPath" -ForegroundColor $FGLabel -BackgroundColor $BGBold
    [bool]$Selectable = $True
    $NumObjects = Get-ADComputer -Filter 'Name -like "*"' -SearchBase $OUPath -SearchScope OneLevel | Measure-Object
    if ($NumObjects.Count -eq 0) {
      $Selectable = $False
      Write-Host "This OU does not contain any Computer Objects.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
    }
    elseif ($NumObjects.Count -eq 1) {
      Write-Host "There is $($NumObjects.Count) Computer Object in this OU.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
    }
    else {
      Write-Host "There are $($NumObjects.Count) Computer Objects in this OU.`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold
    }

    # Get all Child OUs from current OU
    Write-Host "List of Child OUs" -ForegroundColor $FGLabel -BackgroundColor $BGBold
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
    if ($Selectable -eq $False) {
      Write-Host "You cannot select this OU as there are no computer objects at this location." -ForegroundColor $FGNotice -BackgroundColor $BGBold
    }
    else {
      Write-Host "- Enter '+' to remove a Computer Object at the current location."
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
        $Done = $True
        $Valid = $True
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

  # List computer objects in selected OU available for deletion
  Write-Host "`nFinding Computer's Location Manually`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
  Write-Host "ID`tNAME" -ForegroundColor $FGHeader -BackgroundColor $BGBold
  $OUObjects = Get-ADComputer -Filter "Name -like '*'" -SearchBase $OUPath -SearchScope OneLevel
  $Index = 1
  ForEach ($Object in $OUObjects) {
    Write-Host $Index`t$Object.Name
    $Index += 1
  }
  Write-Host

  Write-Host "`OPTIONS" -ForegroundColor $FGHeader -BackgroundColor $BGBold
  Write-Host "- Enter the ID number to select the object you want to remove."
  Write-Host "- Enter '^' to search a different OU for computer objects.`n"

  # Logic for handling user input
  $Valid = $False
  do {
    Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
    $Selection = Read-Host
    if ($Selection -match "^\d+$" -and [int]$Selection -in 1..$PotentialComputers.Count) {
      $Valid = $True
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
