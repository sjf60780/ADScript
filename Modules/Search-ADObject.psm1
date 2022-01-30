function Search-ADObject
{
  $Done = $FALSE
  do
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

    if ($SearchResults[-1].ID -gt 1)
    {
      Clear-Host
      Write-Host "`nSearching for Computer Objects matching the search of:" -ForegroundColor $FGTitle -BackgroundColor $BGBold -NoNewLine
      Write-Host " $EnteredKeywords"
      Write-Host "`nFound $($SearchResults[-1].ID) potential Computer Object(s).`n" -ForegroundColor $FGNotice -BackgroundColor $BGBold -NoNewLine

      $SearchResults | Format-Table

      Write-Host "OPTIONS" -ForegroundColor $FGInput -BackgroundColor $BGInput
      Write-Host "- Enter '^' to start a new search."
      Write-Host "- Press ENTER if you are finished searching for Computer Objects.`n"

      # Logic for handling user input
      $Valid = $FALSE
      do
      {
        Write-Host "Enter your Selection:`n" -ForegroundColor $FGInput -BackgroundColor $BGInput -NoNewLine
        $Selection = Read-Host
        if ($Selection -eq '^')
        {
          $Valid = $TRUE
        }
        elseif ($Selection -eq '')
        {
          $Valid = $TRUE
          $Done = $TRUE
        }
        else
        {
          Write-Host "`nYour selection doesn't exist. Please try again.`n" -ForegroundColor $FGWarning -BackgroundColor $BGBold
          Start-Sleep 2
        }
      } while (!$Valid)
    }
  } while (!$Done)
}
