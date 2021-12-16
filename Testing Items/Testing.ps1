Import-Module ActiveDirectory

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

$StartingOU = "OU=COMPUTERS,OU=CUSTOM,DC=SGF,DC=EDUBEAR,DC=NET"

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

Read-Host
