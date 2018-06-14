function Install-TervisTechnicalServicesWindowsServer {
    param(
        [System.Management.Automation.PSCredential]$Office365Credential = $(get-credential -message "Please supply the credentials to access ExchangeOnline. Username must be in the form UserName@Domain.com"),
        [System.Management.Automation.PSCredential]$InternalCredential = $(get-credential -message "Please supply the credentials to access internal resources. Username must be in the form DOMAIN\username")
    )
    $Office365Credential | Export-Clixml $env:USERPROFILE\Office365EmailCredential.txt    
    $InternalCredential | Export-Clixml $env:USERPROFILE\OnPremiseExchangeCredential.txt
}

function New-TervisDistributionGroup {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Name,
        $Members
    )
    begin {
        Import-TervisExchangePSSession
        $ADOrganizationalUnit = Get-ADOrganizationalUnit -Filter { Name -eq "Distribution Group" }
    }
    process {    
        New-ExchangeDistributionGroup @PSBoundParameters -RequireSenderAuthenticationEnabled:$false -OrganizationalUnit $ADOrganizationalUnit.DistinguishedName
    }
    end {
        Sync-ADDomainControllers -Blocking
        Invoke-ADAzureSync
    }
}

function New-TervisWindowsUser {
    [CmdletBinding(DefaultParameterSetName="NewADUser")]
    param(
        [Parameter(Mandatory, ParameterSetName="NewADUser")]$GivenName,
        [Parameter(Mandatory, ParameterSetName="NewADUser")]$Surname,

        [Parameter(Mandatory)]
        [Parameter(ParameterSetName="UseExistingADUser")]
        [Parameter(ParameterSetName="NewADUser")]
        $SAMAccountName,

        [Parameter(Mandatory)]
        [Parameter(ParameterSetName="UseExistingADUser")]
        [Parameter(ParameterSetName="NewADUser")]
        $ManagerSAMAccountName,

        [Parameter(Mandatory, ParameterSetName="NewADUser")]$Department,
        [Parameter(Mandatory, ParameterSetName="NewADUser")]$Title,
        [Parameter(Mandatory, ParameterSetName="NewADUser")]$AccountPassword,
        [Parameter(ParameterSetName="NewADUser")]$Company = "Tervis",
        [Parameter(Mandatory)]$SAMAccountNameToBeLike,
        [switch]$UserHasTheirOwnDedicatedComputer,
        [Parameter(ParameterSetName="UseExistingADUser")][Switch]$UseExistingADUser
    )
    $AdDomainNetBiosName = (Get-ADDomain | Select-Object -ExpandProperty NetBIOSName).tolower()        
    $UserPrincipalName = "$SAMAccountName@$AdDomainNetBiosName.com"

    $ADUserParameters = @{
        Path = Get-ADUserOU -SAMAccountName $SAMAccountNameToBeLike
        Manager = Get-ADUser $ManagerSAMAccountName | Select-Object -ExpandProperty DistinguishedName   
    }
    
    $ADUser = try {Get-TervisADUser -Identity $SAMAccountName -IncludeMailboxProperties } catch {}
    if (-not $ADUser -and -not $UseExistingADUser){
        New-ADUser `
            -SamAccountName $SAMAccountName `
            -Name "$GivenName $Surname" `
            -GivenName $GivenName `
            -Surname $Surname `
            -UserPrincipalName $UserPrincipalName `
            -AccountPassword $AccountPassword `
            -ChangePasswordAtLogon $true `
            -Company $Company `
            -Department $Department `
            -Title $Title `
            -Enabled $true `
            @ADUserParameters
        
        $ADUser = Get-TervisADUser -Identity $SAMAccountName -IncludeMailboxProperties
        Sync-ADDomainControllers -Blocking
    } elseif (-not $ADUser -and $UseExistingADUser) {
        Throw "$SAMAccountName doesn't exist but `$UseExistingADUser switch used"
    }
    
    Copy-ADUserGroupMembership -Identity $SAMAccountNameToBeLike -DestinationIdentity $SAMAccountName
    
    if (-not $ADUser.O365Mailbox -and -not $ADUser.ExchangeMailbox -and -not $ADUser.ExchangeRemoteMailbox) {
        Enable-ExchangeRemoteMailbox -identity $ADUser.SamAccountName -RemoteRoutingAddress "$($ADUser.SamAccountName)@tervis0.mail.onmicrosoft.com"
        Sync-ADDomainControllers -Blocking
    }

    if (-not $ADUser.O365Mailbox -and $ADUser.ExchangeRemoteMailbox) {
        Invoke-ADAzureSync

        Connect-TervisMsolService
        While (-not (Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue)) {
            Start-Sleep 30
        }
        
        $License = if ($UserHasTheirOwnDedicatedComputer) { "E3" } else { "E1" }
        $ADUser | Set-TervisMSOLUserLicense -License $License
        
        While (-Not $ADUser.O365Mailbox) {
            Start-Sleep 60
        }
    }

    if ($ADUser.O365Mailbox -and -not $ADUser.ExchangeMailbox -and $ADUser.ExchangeRemoteMailbox) {

        $InExchangeOnlinePowerShellModuleShell = Connect-EXOPSSessionWithinExchangeOnlineShell
        if (-not $InExchangeOnlinePowerShellModuleShell) {
            Import-TervisOffice365ExchangePSSession
        } else {
            New-Alias -Name Get-O365OutboundConnector -Value Get-OutboundConnector
            New-Alias -Name New-O365MoveRequest -Value New-MoveRequest
            New-Alias -Name Get-O365MoveRequest -Value Get-MoveRequest
            New-Alias -Name Get-O365MoveRequestStatistics -Value Get-MoveRequestStatistics
            New-Alias -Name Set-O365Mailbox -Value Set-Mailbox
            New-Alias -Name Set-O365Clutter -Value Set-Clutter
            New-Alias -Name Set-O365FocusedInbox -Value Set-FocusedInbox
        }
        
        if ($UserHasTheirOwnDedicatedComputer) {
            Set-O365Mailbox $UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00 -LitigationHoldDuration 2555 -LitigationHoldEnabled $true
            Import-TervisExchangePSSession
            Enable-ExchangeRemoteMailbox $UserPrincipalName -Archive
        } else {
            Set-O365Mailbox $UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00
        }

        Set-O365Clutter -Identity $UserPrincipalName -Enable $false
        Set-O365FocusedInbox -Identity $UserPrincipalName -FocusedInboxOn $false
        Enable-Office365MultiFactorAuthentication -UserPrincipalName $UserPrincipalName
    }

    if ($ADUser.O365Mailbox -and $ADUser.ExchangeMailbox) {
        Throw "$($ADUser.SamAccountName) has both an Office 365 mailbox and an exchange mailbox"
    }
}

function New-TervisProductionUser {
    param(
        [parameter(mandatory)]$GivenName,
        [parameter(mandatory)]$SurName,
        [parameter(mandatory)]$SAMAccountName,
        [parameter(mandatory)]$AccountPassword
    )
    $AdDomainNetBiosName = (Get-ADDomain | Select-Object -ExpandProperty NetBIOSName).tolower()        
    $UserPrincipalName = "$SAMAccountName@$AdDomainNetBiosName.com"

    $Path = Get-ADOrganizationalUnit -Filter * | 
    Where-Object DistinguishedName -match "OU=Users,OU=Production Floor" |
    Select-Object -ExpandProperty DistinguishedName

    $ADUser = try {Get-TervisADUser -Identity $SAMAccountName -IncludeMailboxProperties} catch {}
    if (-not $ADUser){
        New-ADUser `
            -SamAccountName $SAMAccountName `
            -Name "$GivenName $Surname" `
            -GivenName $GivenName `
            -Surname $Surname `
            -UserPrincipalName $UserPrincipalName `
            -AccountPassword $AccountPassword `
            -ChangePasswordAtLogon $false `
            -Company "Tervis" `
            -Department "Production" `
            -Enabled $false `
            -Path $Path
        Set-ADUser -CannotChangePassword $true -PasswordNeverExpires $true -Identity $SAMAccountName
    }
}

function New-TervisContractor {
    [CmdletBinding()]
    Param(
        [parameter(mandatory)]$FirstName,
        [parameter(mandatory)]$LastName,
        [parameter(Mandatory)]$EmailAddress,
        [parameter(mandatory)]$ManagerUserName,
        [parameter(mandatory)]$Title,
        [parameter(Mandatory)]$Description
    )
    DynamicParam {
            $ParameterName = 'Company'
            $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ParameterAttribute.Mandatory = $true
            $ParameterAttribute.Position = 4
            $AttributeCollection.Add($ParameterAttribute)
            $arrSet = Get-TervisContractorDefinition -All | select Name -ExpandProperty Name
            $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
            $AttributeCollection.Add($ValidateSetAttribute)
            $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
            $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
            return $RuntimeParameterDictionary
    }
    begin {
        $Company = $PsBoundParameters[$ParameterName]
    }
    
    process {
        $UserName = Get-AvailableSAMAccountName -GivenName $FirstName -Surname $LastName        
        if ($UserName) {
    
            [string]$AdDomainNetBiosName = (Get-ADDomain | Select -ExpandProperty NetBIOSName).substring(0).tolower()
            [string]$DisplayName = $FirstName + ' ' + $LastName
            [string]$UserPrincipalName = $username + '@' + $AdDomainNetBiosName + '.com'
            [string]$LogonName = $AdDomainNetBiosName + '\' + $username
            [string]$Path = Get-ADUser $ManagerUserName | select distinguishedname -ExpandProperty distinguishedname | Get-ADObjectParentContainer
            $ManagerDN = Get-ADUser $ManagerUserName | Select -ExpandProperty DistinguishedName
            if ((Get-ADGroup -Filter {SamAccountName -eq $Company}) -eq $null ){
                New-ADGroup -Name $Company -GroupScope Universal -GroupCategory Security
            }
            $CompanySecurityGroup = Get-ADGroup -Identity $Company
            $PW= Get-TempPassword -MinPasswordLength 8 -MaxPasswordLength 12 -FirstChar abcdefghjkmnpqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ23456789
            $SecurePW = ConvertTo-SecureString $PW -asplaintext -force
    
            New-ADUser `
                -SamAccountName $Username `
                -Name $DisplayName `
                -GivenName $FirstName `
                -Surname $LastName `
                -UserPrincipalName $UserPrincipalName `
                -AccountPassword $SecurePW `
                -ChangePasswordAtLogon $true `
                -Path $Path `
                -Company $Company `
                -Department $Department `
                -Description $Description `
                -Title $Title `
                -Manager $ManagerDN `
                -Enabled $true

            Add-ADGroupMember $CompanySecurityGroup -Members $UserName
            Add-ADGroupMember "CiscoVPN" -Members $UserName
            Import-TervisExchangePSSession
            New-ExchangeMailContact -FirstName $FirstName -LastName $LastName -Name $DisplayName -ExternalEmailAddress $EmailAddress 
            
            New-PasswordStatePassword -PasswordListId 78 -Title $DisplayName -Username $LogonName -Password $SecurePW

            Send-TervisContractorWelcomeLetter -Name $DisplayName -EmailAddress $EmailAddress
        }
    }
}

function Move-MailboxToOffice365 {
    param(
        [parameter(mandatory)]$UserPrincipalName,
        [switch]$UserHasTheirOwnDedicatedComputer
    )
    $ADUser = Get-TervisADUser -Filter {UserPrincipalName -eq $UserPrincipalName} -IncludeMailboxProperties
    if (-not $ADUser) {throw "User not found in AD"}

    if (-not $ADUser.O365Mailbox -and $ADUser.ExchangeMailbox) {
        Invoke-ADAzureSync

        Connect-TervisMsolService
        While (-not (Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue)) {
            Start-Sleep 30
        }
        
        $License = if ($UserHasTheirOwnDedicatedComputer) { "E3" } else { "E1" }
        $ADUser | Set-TervisMSOLUserLicense -License $License
        Start-Sleep 300

        $InExchangeOnlinePowerShellModuleShell = Connect-EXOPSSessionWithinExchangeOnlineShell
        if (-not $InExchangeOnlinePowerShellModuleShell) {
            Import-TervisOffice365ExchangePSSession
        } else {
            New-Alias -Name Get-O365OutboundConnector -Value Get-OutboundConnector
            New-Alias -Name New-O365MoveRequest -Value New-MoveRequest
            New-Alias -Name Get-O365MoveRequest -Value Get-MoveRequest
            New-Alias -Name Get-O365MoveRequestStatistics -Value Get-MoveRequestStatistics
            New-Alias -Name New-O365MoveRequest -Value New-MoveRequest
            New-Alias -Name Set-O365Mailbox -Value Set-Mailbox
            New-Alias -Name Set-O365Clutter -Value Set-Clutter
            New-Alias -Name Set-O365FocusedInbox -Value Set-FocusedInbox
        }
        
        $Office365DeliveryDomain = Get-MsolDomain | Where Name -Like "*.mail.onmicrosoft.com" | Select -ExpandProperty Name
        $InternalMailServerPublicDNS = Get-O365OutboundConnector | Where Name -Match 'Outbound to' | Select -ExpandProperty SmartHosts
        $OnPremiseCredential = Import-Clixml $env:USERPROFILE\OnPremiseExchangeCredential.txt
        New-O365MoveRequest -Remote -RemoteHostName $InternalMailServerPublicDNS -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $Office365DeliveryDomain -identity $UserPrincipalName -SuspendWhenReadyToComplete:$false

        While (-Not ((Get-O365MoveRequest -Identity $UserPrincipalName).Status -match "Complete")) {
            Get-O365MoveRequestStatistics -Identity $UserPrincipalName | Select StatusDetail,PercentComplete
            Start-Sleep 60
        }
    }

    if ($ADUser.O365Mailbox -and -not $ADUser.ExchangeMailbox) {
        if ($UserHasTheirOwnDedicatedComputer) {
            Set-O365Mailbox $UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00 -LitigationHoldDuration 2555 -LitigationHoldEnabled $true
            Import-TervisExchangePSSession
            Enable-ExchangeRemoteMailbox $UserPrincipalName -Archive
        } else {
            Set-O365Mailbox $UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00
        }

        Set-O365Clutter -Identity $UserPrincipalName -Enable $false
        Set-O365FocusedInbox -Identity $UserPrincipalName -FocusedInboxOn $false
        Enable-Office365MultiFactorAuthentication -UserPrincipalName $UserPrincipalName
    }

    if ($ADUser.O365Mailbox -and $ADUser.ExchangeMailbox) {
        Throw "$($ADUser.SamAccountName) has both an Office 365 mailbox and an exchange mailbox"
    }
}

function Update-TervisSNMPConfiguration {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    Begin {
        $ConfigurationDetails = Get-PasswordstatePassword -ID 12
        $CommunityString = $ConfigurationDetails | Select -ExpandProperty Password
        $SNMPTrap = $ConfigurationDetails | Select -ExpandProperty URL
    }
    Process {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            If ((Get-WindowsFeature -Name SNMP-Service).Installed -ne "True") {
                Add-WindowsFeature SNMP-Service
            }
            If ((Get-WindowsFeature -Name SNMP-WMI-Provider).Installed -ne "True") {
                Add-WindowsFeature SNMP-WMI-Provider
            }
            if (-NOT (Test-Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers\")) {
                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers\" -Force | Out-Null
            }
            if (-NOT (Test-Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities\")) {
                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities\" -Force | Out-Null
            }
            if (-NOT ((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers\" -Name "1").1) -eq "Localhost") {
                New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers\" -Name "1" -Value "localhost" -PropertyType STRING -Force | Out-Null
            }
            if (-NOT ((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers\" -Name "2").2) -eq $Using:SNMPTrap) {
                New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers\" -Name "2" -Value $Using:SNMPTrap -PropertyType STRING -Force | Out-Null
            }
            if (-NOT ((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities\" -Name $Using:CommunityString).$Using:CommunityString) -eq $Using:CommunityString) {
                New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities\" -Name $Using:CommunityString -Value "4" -PropertyType DWORD -Force | Out-Null
            }
            Restart-Service SNMP
        }
    }
}

function New-TervisSharedMailBox {
    param(
        [parameter(mandatory)]$GivenName,
        [parameter(mandatory)]$SAMAccountName,
        [parameter(mandatory)]$Department,
        [parameter(mandatory)]$Surname
    )
     
    $SecurePW = (Get-PasswordstateRandomPassword) | select -ExpandProperty Password | ConvertTo-SecureString -asplaintext -force
    $AdDomainNetBiosName = (Get-ADDomain | Select-Object -ExpandProperty NetBIOSName).tolower()
    $UserPrincipalName = "$SAMAccountName@$AdDomainNetBiosName.com"
    $path = 'OU=Shared Mailbox,OU=Exchange,DC=tervis,DC=prv'

    New-ADUser `
            -SamAccountName $SAMAccountName `
            -Name $GivenName `
            -GivenName $GivenName `
            -Surname $Surname `
            -UserPrincipalName $UserPrincipalName `
            -Department $Department `
            -AccountPassword $SecurePW `
            -Path $path `
            -PasswordNeverExpires $true `
            -ChangePasswordAtLogon $false `
            -Enabled $true `
        
    Import-TervisExchangePSSession   
    Enable-TervisExchangeMailbox $UserPrincipalName
    Set-ExchangeMailbox -Identity $UserPrincipalName -Type “Shared”
     
    Import-TervisOffice365ExchangePSSession
    Connect-TervisMsolService

    $Office365DeliveryDomain = Get-MsolDomain | Where Name -Like "*.mail.onmicrosoft.com" | Select -ExpandProperty Name
    $InternalMailServerPublicDNS = Get-O365OutboundConnector | Where Name -Match 'Outbound to' | Select -ExpandProperty SmartHosts
    $OnPremiseCredential = Import-Clixml $env:USERPROFILE\OnPremiseExchangeCredential.txt
    New-O365MoveRequest -Remote -RemoteHostName $InternalMailServerPublicDNS -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $Office365DeliveryDomain -identity $UserPrincipalName -SuspendWhenReadyToComplete:$false

    While (-Not ((Get-O365MoveRequest -Identity $UserPrincipalName).Status -match "Complete")) {
        Get-O365MoveRequestStatistics -Identity $UserPrincipalName | Select StatusDetail,PercentComplete
        Start-Sleep 60
    }
}

function Invoke-WindowsAdminCenterGatewayProvision {
    param (
        $EnvironmentName = "Infrastructure"
    )
    Invoke-ApplicationProvision -ApplicationName WindowsAdminCenterGateway -EnvironmentName $EnvironmentName
    $Nodes = Get-TervisApplicationNode -ApplicationName WindowsAdminCenterGateway -EnvironmentName $EnvironmentName
}
