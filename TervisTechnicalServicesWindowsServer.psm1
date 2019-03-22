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
    param (
        [ValidateSet("Employee","Contractor")]
        [Parameter(Mandatory)]
        $Type,

        [Parameter(Mandatory)]
        $SAMAccountName,

        [Parameter(Mandatory)]        
        $ManagerSAMAccountName,
        
        [Parameter(Mandatory)]
        [System.Security.SecureString]$AccountPassword,

        $Company,
        [switch]$UserHasTheirOwnDedicatedComputer,
        [switch]$ADUserAccountCreationOnly
    )
    DynamicParam {
        $DynamicParameters = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        if ($Type -eq "Employee") {
            New-DynamicParameter -Name UseExistingADUser -Type Switch -Position 30 -ParameterSetName UseExistingADUser -Dictionary $DynamicParameters
            New-DynamicParameter -Name SAMAccountNameToBeLike -Type String -Position 29 -ParameterSetName NewADUser -Dictionary $DynamicParameters
        }
        if (-not $UseExistingADUser) {
            New-DynamicParameter -Name GivenName -Type String -Position 2 -ParameterSetName NewADUser -Dictionary $DynamicParameters
            New-DynamicParameter -Name Surname -Type String -Position 3 -ParameterSetName NewADUser -Dictionary $DynamicParameters
        }

        if (-not $UseExistingADUser -and $Type -eq "Employee") {
            New-DynamicParameter -Name Department -Type String -Position 4 -ParameterSetName NewADUser -Dictionary $DynamicParameters
            New-DynamicParameter -Name Title -Type String -Position 5 -ParameterSetName NewADUser -Dictionary $DynamicParameters
        }

        $DynamicParameters
    }
    process {
        $AdDomainNetBiosName = (Get-ADDomain | Select-Object -ExpandProperty NetBIOSName).tolower()        
        $UserPrincipalName = "$SAMAccountName@$AdDomainNetBiosName.com"
        
        $ADUser = try {Get-TervisADUser -Identity $SAMAccountName -IncludeMailboxProperties } catch {}
        if (-not $ADUser -and -not $UseExistingADUser) {
            $ADUserParameters = @{
                Path = $(
                    if ($SAMAccountNameToBeLike) {
                        Get-ADUserOU -SAMAccountName $SAMAccountNameToBeLike
                    } else {
                        Get-ADOrganizationalUnit -filter {Name -eq  "Company - Vendors"} |
                        Select-Object -ExpandProperty DistinguishedName
                    }
                )
                Manager = Get-ADUser $ManagerSAMAccountName | Select-Object -ExpandProperty DistinguishedName   
            }
    
            $ADUserParameters += $PSBoundParameters | 
            ConvertFrom-PSBoundParameters -Property SAMAccountName, GivenName, Surname, AccountPassword, Company, Department, Title -AsHashTable
    
            New-ADUser @ADUserParameters `
                -Name "$GivenName $Surname" `
                -UserPrincipalName $UserPrincipalName `
                -ChangePasswordAtLogon $true `
                -Enabled $true
            
            $ADUser = Get-TervisADUser -Identity $SAMAccountName -IncludeMailboxProperties
            Sync-ADDomainControllers -Blocking
        } elseif (-not $ADUser -and $UseExistingADUser) {
            Throw "$SAMAccountName doesn't exist but `$UseExistingADUser switch used"
        }
        
        if ($SAMAccountNameToBeLike) {
            Copy-ADUserGroupMembership -Identity $SAMAccountNameToBeLike -DestinationIdentity $SAMAccountName
        }
    
        if (-not $Contractor -and -not $ADUserAccountCreationOnly) {
            New-TervisMSOLUser -ADUser $ADUser -UserHasTheirOwnDedicatedComputer:$UserHasTheirOwnDedicatedComputer
        }
    }
}

function New-TervisContractor {
    [CmdletBinding()]
    Param(
        [parameter(mandatory)]$GivenName,
        [parameter(mandatory)]$SurName,
        [parameter(Mandatory)]$ExternalEmailAddress,
        [parameter(mandatory)]$ManagerSAMAccountName,
        [parameter(mandatory)]$Title
    )
    DynamicParam {
        New-DynamicParameter -Name Company -Mandatory -Position 4 -ValidateSet $(
            Get-TervisContractorDefinition -All | select Name -ExpandProperty Name
        )
    }
    begin {
        $Company = $PsBoundParameters.Company
    }
    process {
        New-TervisPerson @PsBoundParameters -Contractor
    }
}

function Move-MailboxToOffice365 {
    param(
        [parameter(mandatory)]$Identity,
        [switch]$UserHasTheirOwnDedicatedComputer
    )
    #https://github.com/MicrosoftDocs/office-docs-powershell/issues/1653
    # use $ADUser.UserPrincipalName instead of $UserPrincipalName to work around issue linked above
    $UserPrincipalName = $PSBoundParameters.UserPrincipalName
    $ADUser = Get-TervisADUser -Identity $Identity -IncludeMailboxProperties
    if (-not $ADUser) {throw "User not found in AD"}

    if (-not $ADUser.O365Mailbox -and $ADUser.ExchangeMailbox) {
        Invoke-ADAzureSync

        Connect-TervisMsolService
        While (-not (Get-MsolUser -UserPrincipalName $ADUser.UserPrincipalName -ErrorAction SilentlyContinue)) {
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
            New-Alias -Name Set-O365Mailbox -Value Set-Mailbox
            New-Alias -Name Set-O365Clutter -Value Set-Clutter
            New-Alias -Name Set-O365FocusedInbox -Value Set-FocusedInbox
        }
        
        $Office365DeliveryDomain = Get-MsolDomain | Where Name -Like "*.mail.onmicrosoft.com" | Select -ExpandProperty Name
        $InternalMailServerPublicDNS = Get-O365OutboundConnector | Where Name -Match 'Outbound to' | Select -ExpandProperty SmartHosts
        $OnPremiseCredential = Get-Credential -Message "tervis\ credential for local exchange server"
        New-O365MoveRequest -Remote -RemoteHostName $InternalMailServerPublicDNS -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $Office365DeliveryDomain -identity $ADUser.UserPrincipalName -SuspendWhenReadyToComplete:$false

        While (-Not ((Get-O365MoveRequest -Identity $ADUser.UserPrincipalName).Status -match "Complete")) {
            Get-O365MoveRequestStatistics -Identity $ADUser.UserPrincipalName | Select StatusDetail,PercentComplete
            Start-Sleep 60
        }
    }

    if ($ADUser.O365Mailbox -and -not $ADUser.ExchangeMailbox) {
        if ($UserHasTheirOwnDedicatedComputer) {
            Set-O365Mailbox $ADUser.UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00 -LitigationHoldDuration 2555 -LitigationHoldEnabled $true
            Import-TervisExchangePSSession
            Enable-ExchangeRemoteMailbox $ADUser.UserPrincipalName -Archive
        } else {
            Set-O365Mailbox $ADUser.UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00
        }

        Set-O365Clutter -Identity $ADUser.UserPrincipalName -Enable $false
        Set-O365FocusedInbox -Identity $ADUser.UserPrincipalName -FocusedInboxOn $false
        Enable-Office365MultiFactorAuthentication -UserPrincipalName $ADUser.UserPrincipalName
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

function Invoke-WindowsUpdateRepair {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        Set-Service -ComputerName $ComputerName -Name wuauserv -StartupType Disabled
        Restart-Computer -ComputerName $ComputerName -Force -Wait
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            dism /online /cleanup-image /restorehealth 
            sfc /scannow
        }
        Restart-Computer -ComputerName $ComputerName -Force -Wait
        Set-Service -ComputerName $ComputerName -Name wuauserv -StartupType Manual
    }
}
