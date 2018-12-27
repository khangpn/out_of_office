echo off
REM This file launches an Exchange powershell script
REM To run this script as a scheduled task a domain user which is member of the appropriate Exchange management groups should be used.
REM Depending on the needed rights make the user member of 'Organization Management' group or less.
REM The user does not need to be member of 'Domain Admins' or Administrators groups.

set psScript=%1
set mailBox=%2
set externalMessage=%3
set internalMessage=%4
set startDate=%5
set endDate=%6
set forwardTo=%7

echo %psScript%
echo %mailBox%
echo %externalMessage%
echo %internalMessage%
echo %startDate%
echo %endDate%
echo %forwardTo%

REM C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -version 2.0 -command ". 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto; C:\Scripts\setOutOfOffice\set-OutOfOffice.ps1"
%windir%\System32\WindowsPowerShell\v1.0\powershell.exe -version 2.0 -executionpolicy unrestricted -command ". 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto;" 
%windir%\System32\WindowsPowerShell\v1.0\powershell.exe -version 2.0 -executionpolicy unrestricted -command ". '%psScript%' -mailBox '%mailBox%' -externalMessage '%externalMessage%' -internalMessage '%internalMessage%' -startDate '%startDate%' -endDate '%endDate%' -forwardTo '%forwardTo%'"
PAUSE