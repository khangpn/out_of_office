param(
  [string]$mailBox,
  [string]$startDate,
  [string]$endDate,
  [string]$externalMessage,
  [string]$internalMessage
)


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

set-mailboxautoreplyconfiguration $mailBox -AutoReplyState scheduled -starttime $startDate -endtime $endDate -externalmessage $externalMessage -internalmessage $internalMessage
