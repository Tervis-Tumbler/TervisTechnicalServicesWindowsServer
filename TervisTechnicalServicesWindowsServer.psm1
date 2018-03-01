function Install-TervisTechnicalServicesWindowsServer {
    param(
        [System.Management.Automation.PSCredential]$Office365Credential = $(get-credential -message "Please supply the credentials to access ExchangeOnline. Username must be in the form UserName@Domain.com"),
        [System.Management.Automation.PSCredential]$InternalCredential = $(get-credential -message "Please supply the credentials to access internal resources. Username must be in the form DOMAIN\username"),
        [System.Management.Automation.PSCredential]$PasswordStateCredential = $(get-credential -message 'Enter "NewUser" in the username field. Enter the API key for PasswordState in the password field. This can be found under Administration > System Settings > API Keys')
    )
    $Office365Credential | Export-Clixml $env:USERPROFILE\Office365EmailCredential.txt    
    $InternalCredential | Export-Clixml $env:USERPROFILE\OnPremiseExchangeCredential.txt
    Initialize-PasswordStateRepository -ApiEndpoint 'https://passwordstate/api' -CredentialRepository 'C:\PasswordStateCreds'
    Export-PasswordStateApiKey -ApiKey $PasswordStateCredential
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

function _GetDefault {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Option = [string]::empty
    )

    $repo = (Join-Path -Path $env:USERPROFILE -ChildPath '.passwordstate')

    if (Test-Path -Path $repo -Verbose:$false) {

        $options = (Join-Path -Path $repo -ChildPath 'options.json')
        
        if (Test-Path -Path $options ) {
            $obj = Get-Content -Path $options | ConvertFrom-Json
            if ($options -ne [string]::empty) {
                return $obj.$Option
            } else {
                return $obj
            }
        } else {
            Write-Error -Message "Unable to find [$options]"
        }
    } else {
        Write-Error -Message "Undable to find PasswordState configuration folder at [$repo]"
    }
}

function Initialize-PasswordStateRepository {
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [string]$ApiEndpoint,

        [string]$CredentialRepository = (Join-Path -path $env:USERPROFILE -ChildPath '.passwordstate' -Verbose:$false)
    )

    # If necessary, create our repository under $env:USERNAME\.passwordstate
    $repo = (Join-Path -Path $env:USERPROFILE -ChildPath '.passwordstate')
    if (-not (Test-Path -Path $repo -Verbose:$false)) {
        Write-Debug -Message "Creating PasswordState configuration repository: $repo"
        New-Item -ItemType Directory -Path $repo -Verbose:$false | Out-Null
    } else {
        Write-Debug -Message "PasswordState configuration repository appears to already be created at [$repo]"
    }

    $options = @{
        api_endpoint = $ApiEndpoint
        credential_repository = $CredentialRepository
    }

    $json = $options | ConvertTo-Json -Verbose:$false
    Write-Debug -Message $json
    $json | Out-File -FilePath (Join-Path -Path $repo -ChildPath 'options.json') -Force -Confirm:$false -Verbose:$false
}

function Export-PasswordStateApiKey {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [pscredential[]]$ApiKey,

        [string]$Repository = (_GetDefault -Option 'credential_repository')
    )

    begin {
        if (-not (Test-Path -Path $Repository -Verbose:$false)) {
            Write-Verbose -Message "Creating PasswordState key repository: $Repository"
            New-Item -ItemType Directory -Path $Repository -Verbose:$false | Out-Null
        }
    }

    process {
        foreach ($item in $ApiKey) {
            $exportPath = Join-Path -path $Repository -ChildPath "$($item.Username).cred" -Verbose:$false
            Write-Verbose -Message "Exporting key [$($item.Username)] to $exportPath"
            $item.Password | ConvertFrom-SecureString -Verbose:$false | Out-File $exportPath -Verbose:$false
        }
    }
}

function Import-PasswordStateApiKey {
    [cmdletbinding()]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]]$Name,

        [string]$Repository = (_GetDefault -Option 'credential_repository')
    )

    begin {
        if (-not (Test-Path -Path $Repository -Verbose:$false)) {
            Write-Error -Message "PasswordState key repository does not exist!"
            break
        }
    }

    process {
        foreach ($item in $Name) {
            if ($Name -like "*.cred") {
                $keyPath = Join-Path -Path $Repository -ChildPath "$Name"
            } else {
                $keyPath = Join-Path -Path $Repository -ChildPath "$Name.cred"
            }
            
            if (-not (Test-Path -Path $keyPath)) {
                Write-Error -Message "Key file $keyPath not found!"
                break
            }

            $secPass = Get-Content -Path $keyPath | ConvertTo-SecureString
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Name, $secPass

            return $cred
        }
    }
}

function Get-PasswordStateList {
    [cmdletbinding()]
    param(
        [parameter(mandatory = $true)]
        [pscredential]$ApiKey,

        [parameter(mandatory = $true)]
        [int]$PasswordListId,

        [string]$Endpoint = (_GetDefault -Option 'api_endpoint'),

        [ValidateSet('json','xml')]
        [string]$Format = 'json',

        [switch]$UseV6Api
    )

    $headers = @{}
    $headers['Accept'] = "application/$Format"

    if (-Not $PSBoundParameters.ContainsKey('UseV6Api')) {
        $headers['APIKey'] = $ApiKey.GetNetworkCredential().password    
        $uri = ("$Endpoint/passwordlists/$PasswordListId" + "?format=$Format&QueryAll")
    } else {
        $uri = ("$Endpoint/passwordlists/$PasswordListId" + "?apikey=$($ApiKey.GetNetworkCredential().password)&format=$Format&QueryAll")
    }   

    $result = Invoke-RestMethod -Uri $uri -Method Get -ContentType "application/$Format" -Headers $headers
    return $result
}

function Get-PasswordStateAllLists {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [pscredential]$SystemApiKey,

        [string]$Endpoint = (_GetDefault -Option 'api_endpoint'),

        [ValidateSet('json','xml')]
        [string]$Format = 'json',

        [switch]$UseV6Api
    )

    $headers = @{}
    $headers['Accept'] = "application/$Format"

    if (-Not $PSBoundParameters.ContainsKey('UseV6Api')) {
        $headers['APIKey'] = $SystemApiKey.GetNetworkCredential().password    
        $uri = "$Endpoint/passwordlists?format=$Format"
    } else {
        $uri = "$Endpoint/passwordlists?apikey=$($SystemApiKey.GetNetworkCredential().password)&format=$Format"
    }  

    $result = Invoke-RestMethod -Uri $uri -Method Get -ContentType "application/$Format" -Headers $headers
    return $result
}

function New-PasswordStatePassword {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingUserNameAndPassWordParams', '')]
    [cmdletbinding(SupportsShouldProcess = $true)]
    param(
        [parameter(Mandatory)]
        [pscredential]$ApiKey,

        [parameter(Mandatory)]
        [int]$PasswordListId,

        [string]$Endpoint = (_GetDefault -Option 'api_endpoint'),

        [ValidateSet('json','xml')]
        [string]$Format = 'json',

        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory = $true,ParameterSetName = 'UsePassword')]
        [Parameter(Mandatory = $true,ParameterSetName = 'UsePasswordWithFile')]
        [securestring]$Password,

        [string]$Username,

        [string]$Description,

        [string]$GenericField1,
            
        [string]$GenericField2,

        [string]$GenericField3,

        [string]$GenericField4,

        [string]$GenericField5,

        [string]$GenericField6,

        [string]$GenericField7,

        [string]$GenericField8,

        [string]$GenericField9,

        [string]$GenericField10,

        [string]$Notes,

        [int]$AccountTypeID,

        [string]$Url,

        [string]$ExpiryDate,

        [bool]$AllowExport,

        [Parameter(Mandatory = $true,ParameterSetName = 'GenPassword')]
        [Parameter(Mandatory = $true,ParameterSetName = 'GenPasswordWithFile')]
        [switch]$GeneratePassword,

        [switch]$GenerateGenFieldPassword,

        [switch]$UseV6Api,

        [Parameter(Mandatory = $true,ParameterSetName = 'GenPasswordWithFile')]
        [Parameter(Mandatory = $true,ParameterSetName = 'UsePasswordWithFile')]
        [String]$DocumentPath,

        [Parameter(Mandatory = $true,ParameterSetName = 'GenPasswordWithFile')]
        [Parameter(Mandatory = $true,ParameterSetName = 'UsePasswordWithFile')]
        [String]$DocumentName,
            
        [Parameter(Mandatory = $true,ParameterSetName = 'GenPasswordWithFile')]
        [Parameter(Mandatory = $true,ParameterSetName = 'UsePasswordWithFile')]
        [String]$DocumentDescription
    )

    $headers = @{}
    $headers['Accept'] = "application/$Format"

    $request = '' | Select-Object -Property Title, PasswordListID, apikey
    $request.Title = $Title
    $request.PasswordListID = $PasswordListId
    $request.apikey = $($ApiKey.GetNetworkCredential().password)

    if ($null -ne $Password) {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $request | Add-Member -MemberType NoteProperty -Name Password -Value $UnsecurePassword
    }

    if ($PSBoundParameters.ContainsKey('Username')) {
       $request | Add-Member -MemberType NoteProperty -Name UserName -Value $Username
    }

    if ($PSBoundParameters.ContainsKey('Description')) {
       $request | Add-Member -MemberType NoteProperty -Name Description -Value $Description
    }
    if ($PSBoundParameters.ContainsKey('GenericField1')) {
       $request | Add-Member -MemberType NoteProperty -Name GenericField1 -Value $GenericField1
    }
    if ($PSBoundParameters.ContainsKey('GenericField2')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField2 -Value $GenericField2
    }
    if ($PSBoundParameters.ContainsKey('GenericField3')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField3 -Value $GenericField3
    }
    if ($PSBoundParameters.ContainsKey('GenericField4')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField4 -Value $GenericField4
    }
    if ($PSBoundParameters.ContainsKey('GenericField5')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField5 -Value $GenericField5
    }
    if ($PSBoundParameters.ContainsKey('GenericField6')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField6 -Value $GenericField6
    }
    if ($PSBoundParameters.ContainsKey('GenericField7')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField7 -Value $GenericField7
    }
    if ($PSBoundParameters.ContainsKey('GenericField8')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField8 -Value $GenericField8
    }
    if ($PSBoundParameters.ContainsKey('GenericField9')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField9 -Value $GenericField9
    }
    if ($PSBoundParameters.ContainsKey('GenericField10')) {
        $request | Add-Member -MemberType NoteProperty -Name GenericField10 -Value $GenericField10
    }
    if ($PSBoundParameters.ContainsKey('Notes')) {
        $request | Add-Member -MemberType NoteProperty -Name Notes -Value $Notes
    }
    if ($PSBoundParameters.ContainsKey('AccountTypeID')) {
        $request | Add-Member -MemberType NoteProperty -Name AccountTypeID -Value $AccountTypeID
    }
    if ($PSBoundParameters.ContainsKey('Url')) {
        $request | Add-Member -MemberType NoteProperty -Name Url -Value $Url
    }
    if ($GeneratePassword.IsPresent) {
        $request | Add-Member -MemberType NoteProperty -Name GeneratePassword -Value $true
    }
    if ($GenerateGenFieldPassword.IsPresent) {
        $request | Add-Member -MemberType NoteProperty -Name GenerateGenFieldPassword -Value $true
    }

    $uri = "$Endpoint/passwords"

    if (-Not $PSBoundParameters.ContainsKey('UseV6Api')) {
        $headers['APIKey'] = $ApiKey.GetNetworkCredential().password
    }
    else {
        $uri += "?apikey=$($ApiKey.GetNetworkCredential().password)"
    }

    $json = ConvertTo-Json -InputObject $request

    Write-Verbose -Message $json

    $output = @()

    $documentInfo = $null
    If ($DocumentPath) {
        $documentInfo = "Upload Document.`nDocumentPath : $DocumentPath`nDocumentName : $DocumentName`nDocument Description : $DocumentDescription"
    }

    If ($PSCmdlet.ShouldProcess("Creating new password entry: $Title `n$json`n$documentInfo")) {
        $result = Invoke-RestMethod -Uri $uri -Method Post -ContentType "application/$Format" -Headers $headers -Body $json
        $output += $result
      
        If ($DocumentPath) {
            $uri = "$Endpoint/document/password/$($result.PasswordID)?DocumentName=$([System.Web.HttpUtility]::UrlEncode($DocumentName))&DocumentDescription=$([System.Web.HttpUtility]::UrlEncode($DocumentDescription))"
            Write-Verbose  -Message $uri 

            $result = Invoke-RestMethod -Uri $uri -Method Post -InFile $DocumentPath -ContentType 'multipart/form-data' -Headers $headers 
            $output += $result
            
            return $output
        }
    }
}

function Get-TempPassword {
    <#
    .Synopsis
       Generates one or more complex passwords designed to fulfill the requirements for Active Directory
    .DESCRIPTION
       Generates one or more complex passwords designed to fulfill the requirements for Active Directory
    .EXAMPLE
       Get-TempPassword
       C&3SX6Kn

       Will generate one password with a length between 8  and 12 chars.
    .EXAMPLE
       Get-TempPassword -MinPasswordLength 8 -MaxPasswordLength 12 -Count 4
       7d&5cnaB
       !Bh776T"Fw
       9"C"RxKcY
       %mtM7#9LQ9h

       Will generate four passwords, each with a length of between 8 and 12 chars.
    .EXAMPLE
       Get-TempPassword -InputStrings abc, ABC, 123 -PasswordLength 4
       3ABa

       Generates a password with a length of 4 containing atleast one char from each InputString
    .EXAMPLE
       Get-TempPassword -InputStrings abc, ABC, 123 -PasswordLength 4 -FirstChar abcdefghijkmnpqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ
       3ABa

       Generates a password with a length of 4 containing atleast one char from each InputString that will start with a letter from 
       the string specified with the parameter FirstChar
    .OUTPUTS
       [String]
    .NOTES
       Written by Simon Wåhlin, blog.simonw.se
       I take no responsibility for any issues caused by this script.
    .FUNCTIONALITY
       Generates random passwords
    .LINK
       http://blog.simonw.se/powershell-generating-random-password-for-active-directory/
   
    #>
    [CmdletBinding(DefaultParameterSetName='FixedLength',ConfirmImpact='None')]
    [OutputType([String])]
    Param
    (
        # Specifies minimum password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='RandomLength')]
        [ValidateScript({$_ -gt 0})]
        [Alias('Min')] 
        [int]$MinPasswordLength = 8,
        
        # Specifies maximum password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='RandomLength')]
        [ValidateScript({
                if($_ -ge $MinPasswordLength){$true}
                else{Throw 'Max value cannot be lesser than min value.'}})]
        [Alias('Max')]
        [int]$MaxPasswordLength = 12,

        # Specifies a fixed password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='FixedLength')]
        [ValidateRange(1,2147483647)]
        [int]$PasswordLength = 8,
        
        # Specifies an array of strings containing charactergroups from which the password will be generated.
        # At least one char from each group (string) will be used.
        [String[]]$InputStrings = @('abcdefghijkmnpqrstuvwxyz', 'ABCEFGHJKLMNPQRSTUVWXYZ', '23456789', '!@#$%^&*+/=_-'),

        # Specifies a string containing a character group from which the first character in the password will be generated.
        # Useful for systems which requires first char in password to be alphabetic.
        [String] $FirstChar,
        
        # Specifies number of passwords to generate.
        [ValidateRange(1,2147483647)]
        [int]$Count = 1
    )
    Begin {
        Function Get-Seed{
            # Generate a seed for randomization
            $RandomBytes = New-Object -TypeName 'System.Byte[]' 4
            $Random = New-Object -TypeName 'System.Security.Cryptography.RNGCryptoServiceProvider'
            $Random.GetBytes($RandomBytes)
            [BitConverter]::ToUInt32($RandomBytes, 0)
        }
    }
    Process {
        For($iteration = 1;$iteration -le $Count; $iteration++){
            $Password = @{}
            # Create char arrays containing groups of possible chars
            [char[][]]$CharGroups = $InputStrings

            # Create char array containing all chars
            $AllChars = $CharGroups | ForEach-Object {[Char[]]$_}

            # Set password length
            if($PSCmdlet.ParameterSetName -eq 'RandomLength')
            {
                if($MinPasswordLength -eq $MaxPasswordLength) {
                    # If password length is set, use set length
                    $PasswordLength = $MinPasswordLength
                }
                else {
                    # Otherwise randomize password length
                    $PasswordLength = ((Get-Seed) % ($MaxPasswordLength + 1 - $MinPasswordLength)) + $MinPasswordLength
                }
            }

            # If FirstChar is defined, randomize first char in password from that string.
            if($PSBoundParameters.ContainsKey('FirstChar')){
                $Password.Add(0,$FirstChar[((Get-Seed) % $FirstChar.Length)])
            }
            # Randomize one char from each group
            Foreach($Group in $CharGroups) {
                if($Password.Count -lt $PasswordLength) {
                    $Index = Get-Seed
                    While ($Password.ContainsKey($Index)){
                        $Index = Get-Seed                        
                    }
                    $Password.Add($Index,$Group[((Get-Seed) % $Group.Count)])
                }
            }

            # Fill out with chars from $AllChars
            for($i=$Password.Count;$i -lt $PasswordLength;$i++) {
                $Index = Get-Seed
                While ($Password.ContainsKey($Index)){
                    $Index = Get-Seed                        
                }
                $Password.Add($Index,$AllChars[((Get-Seed) % $AllChars.Count)])
            }
            Write-Output -InputObject $(-join ($Password.GetEnumerator() | Sort-Object -Property Name | Select-Object -ExpandProperty Value))
        }
    }
}

function New-TervisWindowsUser {
    param(
        [parameter(mandatory)]$GivenName,
        [parameter(mandatory)]$Surname,
        [parameter(mandatory)]$SAMAccountName,
        [parameter(mandatory)]$ManagerSAMAccountName,
        [parameter(mandatory)]$Department,
        [parameter(mandatory)]$Title,
        [parameter(mandatory)]$AccountPassword,
        $Company = "Tervis",
        [parameter(mandatory)]$SAMAccountNameToBeLike,
        [switch]$UserHasTheirOwnDedicatedComputer
    )
    $AdDomainNetBiosName = (Get-ADDomain | Select-Object -ExpandProperty NetBIOSName).tolower()        
    $UserPrincipalName = "$SAMAccountName@$AdDomainNetBiosName.com"

    $ADUserParameters = @{
        Path = Get-ADUserOU -SAMAccountName $SAMAccountNameToBeLike
        Manager = Get-ADUser $ManagerSAMAccountName | Select-Object -ExpandProperty DistinguishedName   
    }
    
    $ADUser = try {Get-TervisADUser -Identity $SAMAccountName} catch {}
    if (-not $ADUser){
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
        
        $ADUser = Get-TervisADUser -Identity $SAMAccountName
        Sync-ADDomainControllers -Blocking
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

    $ADUser = try {Get-TervisADUser -Identity $SAMAccountName} catch {}
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
            
            $NewUserCredential = Import-PasswordStateApiKey -Name 'NewUser'
            New-PasswordStatePassword -ApiKey $NewUserCredential -PasswordListId 78 -Title $DisplayName -Username $LogonName -Password $SecurePW

            Send-TervisContractorWelcomeLetter -Name $DisplayName -EmailAddress $EmailAddress
        }
    }
}

function Move-MailboxToOffice365 {
    param(
        [parameter(mandatory)]$UserPrincipalName,
        [Switch]$EnableArchive = $False,
        [switch]$UserHasTheirOwnDedicatedComputer = $False
    )

    [String]$DisplayName = Get-ADUser $UserPrincipalName.Split('@')[0] | Select -ExpandProperty Name

    $Office365Credential = Import-Clixml $env:USERPROFILE\Office365EmailCredential.txt
    $OnPremiseCredential = Import-Clixml $env:USERPROFILE\OnPremiseExchangeCredential.txt

    Connect-ToTervisExchange
    Connect-TervisMsolService

    [string]$Office365DeliveryDomain = Get-MsolDomain | Where Name -Like "*.mail.onmicrosoft.com" | Select -ExpandProperty Name
    if ($UserHasTheirOwnDedicatedComputer) {
        $E3Licenses = Get-MsolAccountSku | Where {$_.AccountSkuID -like "*ENTERPRISEPACK"}
        [string]$License = $E3Licenses | Select -ExpandProperty AccountSkuId
        if ($E3Licenses.ConsumedUnits -ge $E3Licenses.ActiveUnits) {
            Throw "There are not any E3 licenses available to assign to this user."
        }
    } else {
        $E1Licenses = Get-MsolAccountSku | Where {$_.AccountSkuID -like "*STANDARDPACK"}
        [string]$License = $E1Licenses | Select -ExpandProperty AccountSkuId
        if ($E1Licenses.ConsumedUnits -ge $E1Licenses.ActiveUnits) {
            Throw "There are not any E1 licenses available to assign to this user."
        }
    }

    Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation 'US'
    Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $License

    Write-Verbose "Connect to Exchange Online"
    $Sessions = Get-PsSession
    $Connected = $false
    Foreach ($Session in $Sessions) {
        if ($Session.ComputerName -eq 'ps.outlook.com' -and $Session.ConfigurationName -eq 'Microsoft.Exchange' -and $Session.State -eq 'Opened') {
            $Connected = $true
        } elseif ($Session.ComputerName -eq 'ps.outlook.com' -and $Session.ConfigurationName -eq 'Microsoft.Exchange' -and $Session.State -eq 'Broken') {
            Remove-PSSession $Session
        }
    }
    if ($Connected -eq $false) {
        Write-Verbose "Connect to Exchange Online"
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $Office365Credential
        Import-PSSession $Session -Prefix 'O365' -DisableNameChecking -AllowClobber
    }

    [string]$InternalMailServerPublicDNS = Get-O365OutboundConnector | Where Name -Match 'Outbound to' | Select -ExpandProperty SmartHosts
    New-O365MoveRequest -Remote -RemoteHostName $InternalMailServerPublicDNS -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $Office365DeliveryDomain -identity $UserPrincipalName -SuspendWhenReadyToComplete:$false

    Write-Verbose "Migrating the mailbox"
    While (!((Get-O365MoveRequest $DisplayName).Status -eq 'Completed')) {
        Get-O365MoveRequestStatistics $UserPrincipalName | Select PercentComplete
        Start-Sleep 60
    }

    if ($UserHasTheirOwnDedicatedComputer) {
        Set-O365Mailbox $UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00 -LitigationHoldDuration 2555 -LitigationHoldEnabled $true
    } else {
        Set-O365Mailbox $UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00
    }
    Set-O365Clutter -Identity $UserPrincipalName -Enable $false
    Set-O365FocusedInbox -Identity $UserPrincipalName -FocusedInboxOn $false

    if (($EnableArchive -eq $True) -and ($UserHasTheirOwnDedicatedComputer -eq $False)) {
        Throw "In-place archive can only be enabled on mailboxes with an E3 license."
    }
    if (($EnableArchive -eq $True) -and ($UserHasTheirOwnDedicatedComputer)) {
        Enable-remoteMailbox $UserPrincipalName -Archive
    }
}

function Update-TervisSNMPConfiguration {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    Begin {
        $ConfigurationDetails = Get-PasswordstateEntryDetails -PasswordID 12
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
     
    $PW = Get-TempPassword -MinPasswordLength 8 -MaxPasswordLength 12 -FirstChar abcdefghjkmnpqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ23456789
    $SecurePW = ConvertTo-SecureString $PW -asplaintext -force

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
