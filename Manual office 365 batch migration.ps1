#$IdentitiesToMigrateToOffice365 = #Array of users goes here

$ADUsers = $IdentitiesToMigrateToOffice365 | % { Get-TervisADUser -Identity $_ -IncludeMailboxProperties}
foreach ($ADUser in $ADUsers) { -not $ADUser.O365Mailbox -and $ADUser.ExchangeMailbox }
foreach ($ADUser in $ADUsers) { -not (Get-MsolUser -UserPrincipalName $ADUser.UserPrincipalName -ErrorAction SilentlyContinue) }
$ADUsers | Set-TervisMSOLUserLicense -License $License
New-Alias -Name Get-O365OutboundConnector -Value Get-OutboundConnector
New-Alias -Name New-O365MoveRequest -Value New-MoveRequest
New-Alias -Name Get-O365MoveRequest -Value Get-MoveRequest
New-Alias -Name Get-O365MoveRequestStatistics -Value Get-MoveRequestStatistics
New-Alias -Name Set-O365Mailbox -Value Set-Mailbox
New-Alias -Name Set-O365Clutter -Value Set-Clutter
New-Alias -Name Set-O365FocusedInbox -Value Set-FocusedInbox

        
$Office365DeliveryDomain = Get-MsolDomain | Where Name -Like "*.mail.onmicrosoft.com" | Select -ExpandProperty Name
$InternalMailServerPublicDNS = Get-O365OutboundConnector | Where Name -Match 'Outbound to' | Select -ExpandProperty SmartHosts
$OnPremiseCredential = Get-Credential -Message "tervis\ credential for local exchange server"
foreach ($ADUser in $ADUsers) { New-O365MoveRequest -Remote -RemoteHostName $InternalMailServerPublicDNS -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $Office365DeliveryDomain -identity $ADUser.UserPrincipalName -SuspendWhenReadyToComplete:$false }
foreach ($ADUser in $ADUsers) { Get-O365MoveRequestStatistics -Identity $ADUser.UserPrincipalName | Select StatusDetail,PercentComplete }
foreach ($ADUser in $ADUsers) { $ADUser.O365Mailbox -and -not $ADUser.ExchangeMailbox }

foreach ($ADUser in $ADUsers) { 
    Set-O365Mailbox $ADUser.UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00 
    Set-O365Clutter -Identity $ADUser.UserPrincipalName -Enable $false
    Set-O365FocusedInbox -Identity $ADUser.UserPrincipalName -FocusedInboxOn $false
    Enable-Office365MultiFactorAuthentication -UserPrincipalName $ADUser.UserPrincipalName
}


