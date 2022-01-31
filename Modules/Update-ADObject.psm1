function Update-ADObject
{
  Clear-Host
  Write-Host "`nUpdating Active Directory via external CSV File`n" -ForegroundColor $FGTitle -BackgroundColor $BGBold
  Write-Host "This function allows you to Add new computer objects or Edit/Remove existing computer objects from Active Directory through a CSV file. You can use any CSV file, but the file MUST have the following as the first line:"
  Write-Host "MODE,COMPUTEROU,COMPUTERNAME,COMPUTERMAC,COMPUTERDESC,DEPLOYMENTSERVER,EDITS" -ForegroundColor $FGNotice -BackgroundColor $BGBold
  Write-Host "`nTo Add new objects:" -ForegroundColor $FGLabel -BackgroundColor $BGBold
  Write-Host "- You must put 'ADD' as the MODE."
  Write-Host "- Enter the new object's properties in COMPUTEROU, COMPUTERNAME, COMPUTERMAC, and DEPLOYMENTSERVER."
  Write-Host "* COMPUTERDESC is optional and doesn't need to be entered." -ForegroundColor $FGNotice
  Write-Host "EX: ADD,CHEK-0135,NEWCOMPUTER-01,ABCDEF123456,,CORE," -BackgroundColor $BGBold
  Write-Host "`nTo Edit existing objects:" -ForegroundColor $FGLabel -BackgroundColor $BGBold
  Write-Host "- You must put 'EDIT' as the MODE."
  Write-Host "- Either a COMPUTERNAME or COMPUTERMAC must be specified as the object to edit."
  Write-Host "- All changes must then be aligned with the EDITS header and contain at least one of the following (OU;NAME;MAC;DESC;DP)"
  Write-Host "* Each edit MUST be separated by a ';'" -ForegroundColor $FGNotice
  Write-Host "EX: EDIT,,NEWCOMPUTER-01,,,,NewOU;NewName;NewMac;NewDescription;NewDeploymentServer" -BackgroundColor $BGBold
  Write-Host "`nTo Remove existing objects:" -ForegroundColor $FGLabel -BackgroundColor $BGBold
  Write-Host "- You must put 'REMOVE' as the MODE."
  Write-Host "- Either a COMPUTERNAME or COMPUTERMAC must be specified as the object to remove."
  Write-Host "EX: REMOVE,,NEWCOMPUTER-01,,,," -BackgroundColor $BGBold

  Write-Host "`nAfter the update is completed, there will be a log file created at"
  Write-Host "$($LogFilePath)" -ForegroundColor $FGLabel -BackgroundColor $BGBold -NoNewLine
  Write-Host " detailing whether the AD was successfully updated with your action."

  # The needed headers for the CSV file when importing Computer Objects from a CSV file
  $CorrectCSVHeaders =
  @(
    "MODE",
    "COMPUTEROU",
    "COMPUTERNAME",
    "COMPUTERMAC",
    "COMPUTERDESC",
    "DEPLOYMENTSERVER",
    "EDITS"
  )

  $Valid = $FALSE
  do
  {
    Write-Host "`nEnter the full path of the CSV file you wish to import:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
    $CSVFile = Read-Host
    if ($CSVFile)
    {
      $CSVPath = Test-Path $CSVFile -Include "*.csv"
      if ($CSVPath)
      {
        $ValidHeaders = $TRUE
        $FileHeaders = (Get-Content $CSVFile -TotalCount 1).Split(',')
        for ([int]$i = 0; $i -lt $FileHeaders.Count; $i++)
        {
          if ($FileHeaders[$i].Trim() -ne $CorrectCSVHeaders[$i]) { $ValidHeaders = $FALSE }
        }
        if ($ValidHeaders -eq $TRUE) { $Valid = $TRUE }
        else
        {
          Write-Host "`n$CSVFile is not a valid CSV file. Please try again." -ForegroundColor $FGWarning -BackgroundColor $BGBold
          if ($ValidHeaders -eq $FALSE)
          {
            Write-Host "Make sure that your CSV File has the following headers in this order:" -ForegroundColor $FGNotice -BackgroundColor $BGBold
            Write-Host "MODE,COMPUTEROU,COMPUTERNAME,COMPUTERMAC,COMPUTERDESC,DEPLOYMENTSERVER,EDITS" -ForegroundColor $FGLabel -BackgroundColor $BGBold
          }
        }
      }
      else
      {
        Write-Host "`n$CSVFile is not a CSV file. Please enter a CSV file." -ForegroundColor $FGNotice -BackgroundColor $BGBold
      }
    }
  } while (!$Valid)

  # Validate CSVFile being passed as an argument
  if ($CSVPath -eq $TRUE -and $ValidHeaders -eq $TRUE)
  {
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    [int]$TotalComputers = Import-CSV -Path $CSVFile | Measure-Object | Select-Object -expand Count
    [int]$FinishedComputers = 0
    [int]$Success = 0
    [int]$Failed = 0
    Clear-Host
    Write-Host "`nUpdating AD Objects from: $CSVFile`n" -ForegroundColor $FGSuccess -BackgroundColor $BGBold

    Import-CSV -Path $CSVFile | ForEach-Object {
      $FinishedComputers += 1
      $ErrorLog = ''

      $ComputerLog =
      [PSCustomObject]@{
        ACTION = $NULL
        COMPUTERNAME = $NULL
        OUPATH = $NULL
        ACTIONSBY = $([Environment]::UserName)
        DATEADDED = $NULL
        DATEMODIFIED = $NULL
        DATEDELETED = $NULL
        SUCCESSFUL = $NULL
        CHANGES = $NULL
        ERRORS = $NULL
      }

      Write-Host "Updating Object $FinishedComputers of $TotalComputers" -ForegroundColor $FGNotice -BackgroundColor $BGBold

      # Adding computers from the CSV File
      if ($_.MODE.ToUpper() -eq "ADD")
      {
        $ComputerLog.ACTION = "ADD"

        # Check if given keywords can find an existing OU on the AD
        if ($_.COMPUTEROU -eq "") { $ErrorLog += "No OU/Keyword(s) given;" }
        else
        {
          $OUKeywords = $_.COMPUTEROU -split " "
          $PotentialOUs = Search-ADForest "OrganizationalUnit" $OUKeywords
          if ($PotentialOUs[-1].ID -gt 1) { $ErrorLog += "Could not narrow search down to one OU;" }
          elseif ($PotentialOUs -eq $NULL) { $ErrorLog += "Could not find an OU you have permissions to matching those keywords;" }
          else { $OUPath = $PotentialOUs.DistinguishedName }
        }

        # Check if given value meets the requirements of a valid computer name
        # Then check if the given Computer Name is already tied to a Computer Object on the AD
        if ($_.COMPUTERNAME -eq "") { $ErrorLog += "No Name given;" }
        elseif ($_.COMPUTERNAME -match '^\.|([\\/:\*\?"<>\|])+') { $ErrorLog += "Computer name contains invalid character(s);" }
        elseif ($_.COMPUTERNAME.Length -notin 1..15) { $ErrorLog += "Computer name doesn't meet length requirements;" }
        elseif ($_.COMPUTERNAME.ToUpper() -in $ReservedWords) { $ErrorLog += "Computer name is a reserved word;" }
        else
        {
          $ComputerName = $_.COMPUTERNAME.ToUpper()
          $CheckName = Get-ADObject -Filter "ObjectClass -eq 'computer' -and Name -eq '$($ComputerName)'"
          if ($CheckName -ne $NULL) { $ErrorLog += "'$($ComputerName)' already exists in the AD;" }
        }

        # Check if the given MAC Address is already tied to a Computer Object on the AD
        if ($_.COMPUTERMAC -eq "") { $ErrorLog += "No MAC Address given;" }
        elseif ($_.COMPUTERMAC -match ($ValidMACPatterns -join '|'))
        {
          $MACAddress = $_.COMPUTERMAC.ToUpper() -replace '\W'
          [guid]$NetbootGUID = "00000000-0000-0000-0000-$MACAddress"
          $CheckMAC = Get-ADComputer -Filter {netbootGUID -eq $NetbootGUID}
          if ($CheckMAC -ne $NULL) { $ErrorLog += "MAC Address is tied to $($CheckMAC.Name);" }
        }
        else { $ErrorLog += "MAC Address is not properly formatted;" }

        # Check if entered DeploymentServer value can match to a Distribution Point
        if ($_.DEPLOYMENTSERVER -eq "") { $ErrorLog += "No Deployment Server given;" }
        else
        {
          $DP = $_.DEPLOYMENTSERVER
          $NetbootMFP = $DistributionPoints | Where-Object {$_ -match $DP}
          if ($NetbootMFP -eq $NULL) { $ErrorLog += "The given Deployment Server is not valid;" }
          else { $DeploymentServer = $NetbootMFP -replace '\\' }
        }

        if ($ErrorLog -eq "")
        {
          try
          {
            New-ADComputer -Name $ComputerName -SamAccountName $ComputerName -Description $_.COMPUTERDESC -OtherAttributes @{'netbootGUID' = $NetbootGUID; 'netbootMachineFilePath' = $DeploymentServer} -Path $OUPath
            $CheckAddition = Get-ADComputer -Identity $ComputerName
            if ($CheckAddition)
            {
              $Success += 1
              $ComputerLog.COMPUTERNAME = $ComputerName
              $ComputerLog.OUPATH = $OUPath
              $ComputerLog.DATEADDED = Get-Date
              $ComputerLog.SUCCESSFUL = $TRUE
              $ComputerLog.CHANGES = "$($ComputerName) was added to $($OUPath)"
            }
          }
          catch
          {
            $Failed += 1
            $ComputerLog.ERRORS = "Could not create object. You might not have Create Item permissions for the OU."
          }
        }
        else
        {
          $Failed += 1
          $ErrorLog = $ErrorLog.SubString(0, $ErrorLog.Length - 1)
          $ComputerLog.ERRORS = $ErrorLog
        }
      }
      elseif ($_.MODE.ToUpper() -eq "EDIT")
      {
        $ComputerLog.ACTION = "MOD"

        #TODO Add the logic for detecting errors with editing objects from CSV
      }
      elseif ($_.MODE.ToUpper() -eq "REMOVE")
      {
        $ComputerLog.ACTION = "REM"

        if ($_.COMPUTERMAC -or $_.COMPUTERNAME)
        {
          if ($_.COMPUTERMAC)
          {
            $SearchResults = Search-ADForest "ComputerObject" $_.COMPUTERMAC
            if ($SearchResults -eq $NULL) { $ErrorLog += "Could not find a computer with the given MAC;"}
          }
          elseif ($_.COMPUTERNAME)
          {
            $SearchResults = Search-ADForest "ComputerObject" $_.COMPUTERNAME
            if ($SearchResults.Count -gt 0) { $ErrorLog += "Could not narrow search to one computer;" }
            elseif ($SearchResults.Count -eq 0) { $ErrorLog += "Could not find a computer matching the given keyword;"}
          }
        }
        else { $ErrorLog += "No Computer Name or MAC Address given;" }

        if ($ErrorLog -eq "")
        {
          try
          {
            Remove-ADComputer -Identity $SearchResults.Name -Confirm:$FALSE
            $Success += 1
            $ComputerLog.COMPUTERNAME = $SearchResults.Name
            $ComputerLog.DATEDELETED = Get-Date
            $ComputerLog.SUCCESSFUL = $TRUE
            $ComputerLog.CHANGES = "$($SearchResults.Name) was removed from the AD"
          }
          catch
          {
            $Failed += 1
            $ComputerLog.ERRORS = "Could not remove object. You might not have Remove Item permissions for the OU."
          }
        }
        else
        {
          $Failed += 1
          $ErrorLog = $ErrorLog.SubString(0, $ErrorLog.Length - 1)
          $ComputerLog.ERRORS = $ErrorLog
        }
      }
      else { $ErrorLog += "Mode must be ADD/EDIT/REMOVE;" }

      $ComputerArchive += $ComputerLog
    }

    $StopWatch.Stop()
    Get-Process Excel -ErrorAction SilentlyContinue | Select-Object -Property ID | Stop-Process
    Start-Sleep -Milliseconds 300
    $ComputerArchive | Export-CSV -Path $LogFilePath -NoTypeInformation -Append
    Write-Host "`nUpdate of file $CSVFile is Finished" -ForegroundColor $FGSuccess -BackgroundColor $BGBold
    Write-Host "- $Success of $TotalComputers objects were updated successfully." -ForegroundColor $FGNotice -BackgroundColor $BGBold
    Write-Host "- $Failed of $TotalComputers objects were unable to be updated." -ForegroundColor $FGNotice -BackgroundColor $BGBold
    Write-Host "`nCheck the log file listed below for more information."-ForegroundColor $FGHeader -BackgroundColor $BGBold
    Write-Host "$LogFilePath" -ForegroundColor $FGLabel -BackgroundColor $BGBold
    Write-Host "`nThis operation took $($StopWatch.Elapsed.TotalSeconds) seconds."
    Write-Host "`nPress Enter to quit." -ForegroundColor $FGInput -BackgroundColor $BGInput
    Read-Host
  }
}
