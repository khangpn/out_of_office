param(
  [string]$mailBox,
  [string]$startDate,
  [string]$endDate,
  [string]$externalMessage,
  [string]$internalMessage,
  [string]$forwardTo
)


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

set-mailboxautoreplyconfiguration $mailBox -AutoReplyState scheduled -starttime $startDate -endtime $endDate -externalmessage $externalMessage -internalmessage $internalMessage
if ($forwardTo) {
  # Set-Mailbox -Identity $mailBox -DeliverToMailboxAndForward $true -ForwardingAddress $forwardTo
  # KnP: the script for setting forwarding rule goes here.
}