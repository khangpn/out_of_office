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

Write-Host $mailBox
Write-Host $startDate
Write-Host $endDate
Write-Host $externalMessage
Write-Host $internalMessage

set-mailboxautoreplyconfiguration $mailBox -AutoReplyState scheduled -starttime $startTime -endtime $endTime -externalmessage $externalMessage -internalmessage $internalMessage
