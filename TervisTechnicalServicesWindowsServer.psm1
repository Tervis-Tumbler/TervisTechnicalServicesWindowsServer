Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

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

Function Get-TempPassword() {
    Param(
        [int]$length=8,
        [string[]]$sourcedata
    )

    For ($loop=1; $loop –le $length; $loop++) {
        $TempPassword+=($sourcedata | Get-Random)
    }
    Return $TempPassword
}

function New-TervisWindowsUser{
    param(
        [parameter(mandatory)]$FirstName,
        [parameter(mandatory)]$LastName,
        [parameter()]$MiddleInitial,
        [parameter(mandatory)]$Manager,
        [parameter(mandatory)]$Department,
        [parameter(mandatory)]$Title,
        [parameter(mandatory)]$SourceUser
    )

    [string]$FirstInitialLastName = $FirstName[0] + $LastName
    [string]$FirstNameLastInitial = $FirstName + $LastName[0]

    If (!(Get-ADUser -filter {sAMAccountName -eq $FirstInitialLastName})) {
        [string]$UserName = $FirstInitialLastName.substring(0).tolower()
        Write-Host "UserName is $UserName" -ForegroundColor Green
    } elseif (!(Get-ADUser -filter {sAMAccountName -eq $FirstNameLastInitial})) {
        [string]$UserName = $FirstNameLastInitial
        Write-Host 'First initial + last name is in use.' -ForegroundColor Red
        Write-Host "UserName is $UserName" -ForegroundColor Green
    } else {
        Write-Host 'First initial + last name is in use.' -ForegroundColor Red
        Write-Host 'First name + last initial is in use.' -ForegroundColor Red
        Write-Host 'You will need to manually define $UserName' -ForegroundColor Red
        $UserName = $null
    }

    If (!($UserName -eq $null)) {

        [string]$AdDomainNetBiosName = (Get-ADDomain | Select -ExpandProperty NetBIOSName).substring(0).tolower()
        [string]$Company = $AdDomainNetBiosName.substring(0,1).toupper()+$AdDomainNetBiosName.substring(1).tolower()
        [string]$DisplayName = $FirstName + ' ' + $LastName
        [string]$UserPrincipalName = $username + '@' + $AdDomainNetBiosName + '.com'
        [string]$LogonName = $AdDomainNetBiosName + '\' + $username
        [string]$Path = Get-ADUser $SourceUser -Properties distinguishedname,cn | select @{n='ParentContainer';e={$_.distinguishedname -replace '^.+?,(CN|OU.+)','$1'}} | Select -ExpandProperty ParentContainer
        $ManagerDN = Get-ADUser $Manager | Select -ExpandProperty DistinguishedName

        $ascii=$NULL;
        For ($a=48;$a –le 122;$a++) {$ascii+=,[char][byte]$a }
        $PW= Get-TempPassword –length 8 –sourcedata $ascii
        $SecurePW = ConvertTo-SecureString $PW -asplaintext -force

        $Office365Credential = Import-Clixml $env:USERPROFILE\Office365EmailCredential.txt
        $OnPremiseCredential = Import-Clixml $env:USERPROFILE\OnPremiseExchangeCredential.txt

        if ($MiddleInitial) {
            [string]$Initials = '-Initials ' + $MiddleInitial
        } else {
            $Initials = $null
        }

        New-ADUser `
            -SamAccountName $Username `
            -Name $DisplayName `
            -GivenName $FirstName `
            -Surname $LastName `
            $Initials `
            -UserPrincipalName $UserPrincipalName `
            -AccountPassword $SecurePW `
            -ChangePasswordAtLogon $true `
            -Path $Path `
            -Company $Company `
            -Department $Department `
            -Office $Department `
            -Description $Title `
            -Title $Title `
            -Manager $ManagerDN

        $NewUserCredential = Import-PasswordStateApiKey -Name 'NewUser'
        New-PasswordStatePassword -ApiKey $NewUserCredential -PasswordListId 78 -Title $DisplayName -Username $LogonName -Password $SecurePW

        Write-Verbose "Forcing a sync between domain controllers"
        $DC = Get-ADDomainController | Select -ExpandProperty HostName
        Invoke-Command -ComputerName $DC -ScriptBlock {repadmin /syncall}
        Start-Sleep 30
        [string]$MailboxDatabase = Get-MailboxDatabase | Where Name -NotLike "Temp*" | Select -Index 0 | Select -ExpandProperty Name
        Enable-Mailbox -Identity $UserPrincipalName -Database $MailboxDatabase

        $Groups = Get-ADUser $SourceUser -Properties MemberOf | Select -ExpandProperty MemberOf

        Foreach ($Group in $Groups) {
            Add-ADGroupMember -Identity $group -Members $UserName
        }
        
        Write-Verbose "Forcing a sync between domain controllers"
        $DC = Get-ADDomainController | select -ExpandProperty HostName
        Invoke-Command -ComputerName $DC -ScriptBlock {repadmin /syncall}
        Start-Sleep 30

        Write-Verbose 'Starting Sync From AD to Office 365 & Azure AD'
        Invoke-Command -ComputerName 'DirSync' -ScriptBlock {Start-ScheduledTask 'Azure AD Sync Scheduler'}
        Start-Sleep 30

        Connect-MsolService -Credential $Office365Credential
        [string]$Office365DeliveryDomain = Get-MsolDomain | Where Name -Like "*.mail.onmicrosoft.com" | Select -ExpandProperty Name
        [string]$License = Get-MsolAccountSku | Where {$_.ActiveUnits -LT 10000 -and $_.AccountSkuID -like "*ENTERPRISEPACK"} | Select -ExpandProperty AccountSkuId

        Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation 'US'
        Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $License

        Write-Verbose "Connect to Exchange Online"
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $Office365Credential
        Import-PSSession $Session -Prefix Cloud -DisableNameChecking

        [string]$InternalMailServerPublicDNS = Get-CloudOutboundConnector | Select -ExpandProperty SmartHosts
        New-CloudMoveRequest -Remote -RemoteHostName $InternalMailServerPublicDNS -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $Office365DeliveryDomain -identity $UserPrincipalName -SuspendWhenReadyToComplete:$false

        While (!((Get-CloudMoveRequest $DisplayName).Status -eq 'Completed')) {
            Start-Sleep 60
        }

        Set-cloudMailbox $UserPrincipalName -AuditOwner MailboxLogin,HardDelete,SoftDelete,Move,MoveToDeletedItems -AuditDelegate HardDelete,SendAs,Move,MoveToDeletedItems,SoftDelete -AuditEnabled $true -RetainDeletedItemsFor 30.00:00:00
        Enable-remoteMailbox $UserPrincipalName -Archive
        Get-CloudMailbox $UserPrincipalName -ResultSize Unlimited | Set-CloudClutter -Enable $false

        $Search = Get-CloudMailboxSearch | where InPlaceHoldEnabled -eq $true
        [string]$InPlaceHoldIdentity = $Search.InPlaceHoldIdentity
        $Mailboxes = Get-CloudMailbox –Resultsize Unlimited –IncludeInactiveMailbox | 
            where {$_.RecipientTypeDetails -eq 'UserMailbox' -and $_.InPlaceHolds -notcontains $InPlaceHoldIdentity -and $_.MailboxPlan -notlike "ExchangeOnlineDeskless*"} | 
            Select -ExpandProperty LegacyExchangeDN
        foreach ($Mailbox in $Mailboxes) {
            $Search.Sources.Add($Mailbox)
        }
        Set-CloudMailboxSearch "In-Place Hold" -SourceMailboxes $Search.Sources
    }
}