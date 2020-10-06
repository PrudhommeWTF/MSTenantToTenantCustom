#requires -modules ExchangeOnlineManagement, AzureAD, AzureRM.Insights, AzureRM.KeyVault, AzureRM.Resources
<#
    .SYNOPSIS
    This script can be used by a tenant that wishes to pull resources out of another tenant.
    For example contoso.com would run this script in order to pull mailboxes from fabrikam.com tenant.

    This script is intended for the target tenant and would setup the following using the SubscriptionId specified or the default subscription:
        1. Create a resource group or use the one specified as parameter
        2. Create a key vault in the above resource group specified as a parameter
        3. Setup above key vault's access policy to grant exchange access to secrets and certificates.
        4. Request a self-signed certificate to be put in the key vault.
        5. Retrieve the public part of certificate from key vault
        6. Create an AAD application and setup its permissions for MSGraph and exchange
        7. Set the secret for above application as the certificate in 4.
        8. Wait for the tenant admin to consent to the application permissions
        9. Once confirmed, send an email using initiation manager to the tenant admin of resource tenant.
        10. Create a migration endpoint in exchange with the ApplicationId, Pointer to application secret in KeyVault and RemoteTenant
        11. Create an organization relationship with resource tenant authorizing migration.

   .PARAMETER SubscriptionId
   The subscription to use for key vault.

   .PARAMETER ResourceTenantAdminEmail
   The resource tenant admin email.

   .PARAMETER ResourceGroup
   The resource group name.

   .PARAMETER KeyVaultName
   The key vault name.

   .PARAMETER KeyVaultLocation
   The location of the key vault

   .PARAMETER CertificateName
   The name of certificate to create in key vault

   .PARAMETER CertificateSubject
   The subject of certificate to create in key vault

   .PARAMETER AzureAppPermissions
   Fine grained control over the permissions to be given to the application.

   .PARAMETER UseAppAndCertGeneratedForSendingInvitation
   Download the private key of generated certificate from key vault to be used for sending invitation.

   .PARAMETER ResourceTenantDomain
   The resource tenant technical Domain Name (ie: fabrikam.onmicrosoft.com).

   .PARAMETER TargetTenantDomain
   The target tenant technical Domain Name (ie: contoso.onmicrosoft.com).

   .PARAMETER ResourceTenantId
   The resource tenant id.

   .PARAMETER ExistingApplicationId
   You can specify an existing Azure AD Application ID if you already have created previously.

   .EXAMPLE
   Set-CrossTenantRelationshipOnTargetTenant.ps1 -ResourceTenantDomain contoso.onmicrosoft.com -TargetTenantDomain fabrikam.onmicrosoft.com -ResourceTenantAdminEmail admin@contoso.onmicrosoft.com -ResourceGroup "TESTPSRG" -KeyVaultName "TestPSKV" -CertificateSubject "CN=TESTCERTSUBJ" -AzureAppPermissions Exchange, MSGraph -UseAppAndCertGeneratedForSendingInvitation -KeyVaultAuditStorageAccountName "KeyVaultLogsStorageAcnt" -ExistingApplicationId d7404497-1e2f-4b58-bdd5-93e82dad91a4

   .EXAMPLE
   Set-CrossTenantRelationshipOnTargetTenant.ps1 -ResourceTenantDomain contoso.onmicrosoft.com -TargetTenantDomain fabrikam.onmicrosoft.com -ResourceTenantId <ContosoTenantId>
#>
[CmdletBinding(
    SupportsShouldProcess = $true
)]
Param(
    [Parameter(
        Mandatory = $true,
        HelpMessage = 'SubscriptionId for key vault'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [String]$SubscriptionId,

    [Parameter(
        Mandatory = $true,
        HelpMessage = 'Resource tenant admin email'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [String]$ResourceTenantAdminEmail,

    [Parameter(
        Mandatory = $true,
        HelpMessage = 'Resource group for key vault'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [String]$ResourceGroup,

    [Parameter(
        Mandatory = $true,
        HelpMessage = 'KeyVault name'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [String]$KeyVaultName,

    [Parameter(
        HelpMessage = 'KeyVault location in Azure Regions'
    )]
    [String]$KeyVaultLocation = "North Europe",

    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Resource group for storage account used for key vault audit logs'
    )]
    [String]$KeyVaultAuditStorageResourceGroup,

    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Storage account name for storing key vault audit logs'
    )]
    [String]$KeyVaultAuditStorageAccountName,

    [Parameter(
        HelpMessage = 'Certificate name to use'
    )]
    [String]$CertificateName,

    [Parameter(
        HelpMessage = 'Certificate subject to use'
    )]
    [ValidateScript({$_.StartsWith("CN=") })]
    [String]$CertificateSubject,

    [Parameter(
        HelpMessage = 'Application permissions'
    )]
    $AzureAppPermissions = 'All',

    [Parameter(
        HelpMessage = 'Use the certificate generated for azure application when sending invitation'
    )]
    [Switch]$UseAppAndCertGeneratedForSendingInvitation,

    [Parameter(
        Mandatory = $true,
        HelpMessage='Resource tenant domain'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [String]$ResourceTenantDomain,

    [Parameter(
        Mandatory = $true,
        HelpMessage='Target tenant domain'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $TargetTenantDomain,

    [Parameter(
        Mandatory = $true,
        HelpMessage = 'Target tenant id. This is azure ad directory id or external directory object id in exchange online.'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $ResourceTenantId,

    [Parameter(
        HelpMessage = 'Existing Application Id. If existing application Id is present and can be found, new application will not be created.'
    )]
    [Guid]$ExistingApplicationId  = [guid]::Empty
)

#region Variables
#Verbose is meant to be cyan, not yellow...
$host.PrivateData.VerboseForegroundColor = 'cyan'

$ErrorActionPreference = 'Stop'

$MSGraphAppId = "00000003-0000-0000-c000-000000000000"
$MSGraphAppRole = 'Directory.ReadWrite.All','User.Invite.All'
$ExchangeOnlineAppId = "00000002-0000-0ff1-ce00-000000000000"
$ExchangeOnlineAppRole = "Mailbox.Migration"
$ReplyUrl = "https://office.com"
$MSGraphResourceUri = "https://graph.microsoft.com"
$PowerShellClientId = "a0c73c16-a7e3-4564-9a95-2bdf47383716"
$PowerShellClientRedirectUri = 'urn:ietf:wg:oauth:2.0:oob' -as [Uri]
$RandomId = '{0:D4}' -f [Random]::new().Next(0, 10000)
#endregion Variables

#region Main
$enumExists = $null
try {
    $enumExists = [ApplicationPermissions] | Get-Member
} catch { }

if (-not $enumExists) {
    Add-Type -TypeDefinition @"
       using System;

       [Flags]
       public enum ApplicationPermissions
       {
          Exchange = 1,
          MSGraph = 2,
          All = Exchange | MSGraph
       }
"@
}

#Check Exchange Online Powershell Connection
if ($null -eq (Get-Command -Name 'New-OrganizationRelationship' -ErrorAction SilentlyContinue)) {
    Write-Warning -Message 'The current PowerShell Host is not connected to Exchange Online. Trying to connect to Exchange Online Management. Please specify credential that have admin rights on the target tenant.'
    Connect-ExchangeOnline
} else {
    Write-Verbose -Message 'Already connected to Exchange Online.'
    if ((Get-AcceptedDomain).Name -notcontains $TargetTenantDomain) {
        #Not connected to the target tenant
        Write-Warning -Message 'Not connected to the target Exchange Online. Disconnecting and prompting to reconnect.'
        Disconnect-ExchangeOnline -Confirm:$false
        Connect-ExchangeOnline
    }
}

$AzureAppPermissions = ([ApplicationPermissions]$AzureAppPermissions)
#MSGraph App is missing permissions
if (-not $AzureAppPermissions.HasFlag([ApplicationPermissions]::MSGraph) -and $UseAppAndCertGeneratedForSendingInvitation) {
    throw "Cannot use application for sending invitation as it does not have permissions on MSGraph."
}

if ($null -eq (Get-Command -Name 'Get-AzureADUser' -ErrorAction SilentlyContinue)) {
    Write-Warning -Message 'The current PowerShell Host is not connected to Azure AD. Trying to connect to Azure AD. Please specify credential that have admin rights on the target tenant.'
    try {
        $AzureADConnection = Connect-AzureAD
        Write-Verbose -Message "Connected to AzureAD - $($AzureADConnection | Out-String)"
    }
    catch {throw $_}
} else {
    #Check that Azure AD is connected to the right Azure AD Subscription
    if ((Get-AzureADDomain).Name -notcontains $TargetTenantDomain) {
        Write-Warning -Message 'Not connected to the target Azure AD. Disconnecting and prompting to reconnect.'
        Disconnect-AzureAD -Confirm:$false
        Connect-AzureAD
    }
}
if ($null -eq (Get-Command -Name 'Get-AzureRmSubscription' -ErrorAction SilentlyContinue)) {
    try {
        $AzureRMAccount = Login-AzureRmAccount
        Write-Verbose -Message "Connected to Azure RM Account - $($AzureRMAccount | Out-String)"
    }
    catch {throw $_}
}

Write-Verbose -Message "Setting up key vault in $TargetTenantDomain tenant"
$AllSubscriptions = Get-AzureRmSubscription

Write-Verbose -Message "SubscriptionId - $SubscriptionId was provided. Searching for it in $($AllSubscriptions | Out-String)"
$Subscription = $AllSubscriptions | Where-Object -FilterScript { $_.SubscriptionId -eq $SubscriptionId}
if ($Subscription) {
    Write-Verbose -Message "Found subscription - $($SubscriptionId | Out-String)"
    $null = Set-AzureRmContext -SubscriptionId $SubscriptionId
} else {
    throw "Subscription with id $SubscriptionId was not found."
}

#Gather all Azure AD Service Principals
$AllAzureADSPNs = Get-AzureADServicePrincipal -All $true
$ExchangeOnlineAppSPN = $AllAzureADSPNs | Where-Object -FilterScript { $_.AppId -eq $ExchangeOnlineAppId }
Write-Verbose -Message "Found exchange service principal in $TargetTenantDomain - $($ExchangeOnlineAppSPN | Out-String)"

#region Manage Azure KeyVault and generate the Certificate
if ([string]::IsNullOrWhiteSpace($CertificateName)) {
    $CertificateName = 'TenantFriendingAppSecret{0}' -f $RandomId
}

#region Check and create Resource Group
$AzureADResourceGroup = $null
try {
    $AzureADResourceGroup = Get-AzureRmResourceGroup -Name $ResourceGroup
} catch {
    Write-Verbose -Message "Resource group $ResourceGroup not found, this will be created."
}

if ($AzureADResourceGroup) {
    Write-Verbose -Message "Resource group $ResourceGroup already exists."
} else {
    Write-Verbose -Message "Creating resource group - $ResourceGroup"
    try {
        $AzureADResourceGroup = New-AzureRmResourceGroup -Name $ResourceGroup -Location $KeyVaultLocation
        Write-Host -Object "Resource Group $ResourceGroup successfully created" -Foreground Green
    }
    catch {throw $_}
}
#endregion Check and create Resource Group

#region Check and create KeyVault
$AzureRmKeyVault = $null
try {
    $AzureRmKeyVault = Get-AzureRmKeyVault -Name $KeyVaultName -ResourceGroupName $ResourceGroup
} catch {
    Write-Verbose -Message "KeyVault $KeyVaultName not found, this will be created."
}

if ($AzureRmKeyVault) {
    Write-Verbose -Message "Keyvault $KeyVaultName already exists."
} else {
    Write-Verbose -Message "Creating KeyVault $KeyVaultName"
    try {
        $AzureRmKeyVault = New-AzureRmKeyVault -Name $KeyVaultName -Location $KeyVaultLocation -ResourceGroupName $ResourceGroup
        Write-Host -Object "KeyVault $KeyVaultName successfully created" -Foreground Green
    }
    catch {throw $_}
}
#endregion Check and create KeyVault

if ($KeyVaultAuditStorageResourceGroup -and $KeyVaultAuditStorageAccountName) {
    Write-Verbose -Message "Setting up auditing for key vault $KeyVaultName"
    $AzureRmStorageAccount = Get-AzureRmStorageAccount -ResourceGroupName $KeyVaultAuditStorageResourceGroup -Name $KeyVaultAuditStorageAccountName
    try {
        Set-AzureRmDiagnosticSetting -ResourceId $AzureRmKeyVault.ResourceId -StorageAccountId $AzureRmStorageAccount.Id -Enabled $true -Categories AuditEvent | Out-Null
        Write-Host -Object "Auditing setup successfully for $KeyVaultName" -Foreground Green
    }
    catch {throw $_}
}

Write-Verbose -Message "Setting up access for key vault $KeyVaultName"
try {
    Set-AzureRmKeyVaultAccessPolicy -ResourceId $AzureRmKeyVault.ResourceId -ObjectId $ExchangeOnlineAppSPN.ObjectId -PermissionsToSecrets get,list -PermissionsToCertificates get,list | Out-Null
    Write-Host -Object "Exchange app given access to KeyVault $KeyVaultName" -Foreground Green
}
catch {throw $_}

try {
    $certificatePublicKey = Get-AzureKeyVaultCertificate -VaultName $KeyVaultName -Name $CertificateName
    if ($certificatePublicKey.Certificate) {
        Write-Verbose -Message "Certificate $CertificateName already exists in $KeyVaultName"
        if ($UseAppAndCertGeneratedForSendingInvitation -eq $true) {
            Write-Verbose -Message "Retrieving certificate private key"
            $certificatePrivateKey = Get-AzureKeyVaultSecret -VaultName $KeyVaultName -Name $CertificateName
        }

        return $certificatePublicKey, $certificatePrivateKey
    }
} catch {
    Write-Verbose -Message "Certificate not found, a new request will be generated."
}

if ([string]::IsNullOrWhiteSpace($CertificateSubject)) {
    $CertificateSubject = 'CN={0}_{1}_{2}' -f $TargetTenantDomain, $ResourceTenantDomain, $RandomId
    Write-Verbose -Message "Cert subject not provided, generated subject - $CertificateSubject"
}

$AzureKeyVaultCertificatePolicy = New-AzureKeyVaultCertificatePolicy -SubjectName $CertificateSubject -IssuerName Self -ValidityInMonths 12
try {
    $null = Add-AzureKeyVaultCertificate -VaultName $KeyVaultName -Name $CertificateName -CertificatePolicy $AzureKeyVaultCertificatePolicy
    Write-Host -Object "Self signed certificate requested in key vault - $KeyVaultName. Certificate name - $CertificateName" -Foreground Green
}
catch {throw $_}

$tries = 5
while ($tries -gt 0) {
    try {
        Write-Verbose -Message "Looking for certificate $CertificateName. Attempt - $(6 - $tries)"
        $certificatePublicKey = Get-AzureKeyVaultCertificate -VaultName $KeyVaultName -Name $CertificateName
        if ($certificatePublicKey.Certificate) {
            Write-Verbose -Message "Certificate found - $($certificatePublicKey | Out-String)"
            if ($UseAppAndCertGeneratedForSendingInvitation -eq $true) {
                $certificatePrivateKey = Get-AzureKeyVaultSecret -VaultName $KeyVaultName -Name $CertificateName
                if ($certificatePrivateKey) {
                    Write-Verbose -Message "Certificate private key also found"
                    break
                } else {
                    if ($tries -lt 0) {
                        throw "Certificate private key not found after retries."
                    }
                    Write-Verbose -Message "Certificate public key is present, however, its private key is not available, waiting 5 secs and looking again."
                }
            }
        } else {
            if ($tries -lt 0) {
                throw "Certificate not found after retries."
            }

            Write-Verbose -Message "Certificate not found, waiting 5 secs and looking again."
            Start-Sleep 5
        }
    }
    catch {
        if ($tries -lt 0) {
            Write-Error "Certificate not found after retries."
        }

        Start-Sleep 5
    }

    $tries--
}

Write-Verbose -Message "Returning cert: $($certificatePublicKey.Certificate | Out-String)"
Write-Host -Object "Certificate $CertificateName successfully created" -Foreground Green
#endregion Manage Azure KeyVault and generate the Certificate

Write-Verbose -Message "Creating an application in $TargetTenantDomain"
if (-not $AzureAppPermissions.HasFlag([ApplicationPermissions]::MSGraph)) {
    Write-Warning -Message "MSGraph permission was not specified, however, an app needs at least one permission on ADGraph in order for admin to consent to it via the consent url. This app may only be consented from the azure portal."
}

#region Azure App creation
if ([guid]::Empty -eq $ExistingApplicationId) {
    $ExistingApp = Get-AzureADApplication -Filter "AppId eq '$ExistingApplicationId'"
    if ($ExistingApp) {
        Write-Warning -Message "Existing application '$ExistingApplicationId' found. Skipping new application creation."
        
        $AzureADAppCreated = $ExistingApp
    } else {
        Write-Warning -Message "No existing application found. Will create a new one."

        #Collect all the permissions first
        $AppPermissions = @()
        $MSGraphSPN = $null

        if ($AzureAppPermissions.HasFlag([ApplicationPermissions]::MSGraph)) {
            #Calculate permission on MSGraph
            $MSGraphSPN = $AllAzureADSPNs | Where-Object -FilterScript { $_.AppId -eq $MSGraphAppId }
            if (-not $MSGraphSPN) {
                Write-Error "Tenant does not have access to MSGraph"
            }

            $GraphRequiredAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
            $GraphRequiredAccess.ResourceAppId = $MSGraphSPN.AppId
            foreach ($role in $MSGraphAppRole) {
                $GraphRequiredAccess.ResourceAccess += New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList ($MSGraphSPN.AppRoles | Where-Object -FilterScript { $_.Value -eq $role }).Id, 'Role'
            }
        }
        $AppPermissions += $GraphRequiredAccess

        if ($AzureAppPermissions.HasFlag([ApplicationPermissions]::Exchange)) {
            #Calculate permission on EXO
            $ExchangeOnlineAppSPN = $AllAzureADSPNs | Where-Object -FilterScript { $_.AppId -eq $ExchangeOnlineAppId }
            if (-not $ExchangeOnlineAppSPN) {
                Write-Error "Tenant does not have Exchange enabled"
            }

            $ExchangeOnlineAppRole = $ExchangeOnlineAppSPN.AppRoles | Where-Object -FilterScript { $_.Value -eq $ExchangeOnlineAppRole }
            $ExchangeOnlineRequiredAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
            $ExchangeOnlineRequiredAccess.ResourceAppId = $ExchangeOnlineAppSPN.AppId
            $ExchangeOnlineRequiredAccess.ResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $ExchangeOnlineAppRole.Id, 'Role'
            $AppPermissions += $ExchangeOnlineRequiredAccess
        }

        #Create the app with all the permissions ####
        $AzureADAppName = '{0}_Friends_{1}_{2}' -f $TargetTenantDomain.Split('.')[0], $ResourceTenantDomain.Split('.')[0], $RandomId
        $appCreationParameters = @{
            "AvailableToOtherTenants" = $true;
            "DisplayName" = $AzureADAppName
            "Homepage" = $ReplyUrl
            "ReplyUrls" = $ReplyUrl
            "RequiredResourceAccess" = $AppPermissions
        }

        $AzureADAppCreated = New-AzureADApplication @appCreationParameters

        $base64CertHash = [System.Convert]::ToBase64String($certificatePublicKey.Certificate.GetCertHash())
        $base64CertVal = [System.Convert]::ToBase64String($certificatePublicKey.Certificate.GetRawCertData())
        $null = New-AzureADApplicationKeyCredential -ObjectId $AzureADAppCreated.ObjectId -CustomKeyIdentifier $base64CertHash -Value $base64CertVal -StartDate ([DateTime]::Now) -EndDate ([DateTime]::Now).AddDays(363) -Type AsymmetricX509Cert -Usage Verify
        $null = New-AzureADServicePrincipal -AppId $AzureADAppCreated.AppId -AccountEnabled $true -DisplayName $AzureADAppCreated.DisplayName
        $permissions = ""
        if ($AzureAppPermissions.HasFlag([ApplicationPermissions]::MSGraph)) {
            foreach ($role in $MSGraphAppRole) {
                $permissions += "MSGraph - $role. "
            }
        }

        if ($AzureAppPermissions.HasFlag([ApplicationPermissions]::Exchange)) {
            $permissions += "Exchange - $ExchangeOnlineAppRole"
        }

        Write-Host -Object "Application $AzureADAppName created successfully in $TargetTenantDomain tenant with following permissions. $permissions" -Foreground Green
        Write-Host -Object "Admin consent URI for $TargetTenantDomain tenant admin is:" -Foreground Yellow
        Write-Host -Object ("https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}" -f $TargetTenantDomain, $AzureADAppCreated.AppId, $AzureADAppCreated.ReplyUrls[0])

        Write-Host -Object "Admin consent URI for $ResourceTenantDomain tenant admin is:" -Foreground Yellow
        Write-Host -Object ("https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}" -f $ResourceTenantDomain, $AzureADAppCreated.AppId, $AzureADAppCreated.ReplyUrls[0])
    }
}
#endregion Azure App creation

Write-Verbose -Message "Sending the consent URI for this app to $ResourceTenantAdminEmail."

Read-Host "Please consent to the app for $TargetTenantDomain before sending invitation to $ResourceTenantAdminEmail"

#region Sending consent mail to the Resource Tenant Administrator email address specified
$AuthenticationResult = $null
Write-Verbose -Message "Preparing invitation. Waiting for 10 secs before requesting token for the consented application to give time for replication."
Start-Sleep 10
$AuthenticationContextObj = New-Object -TypeName 'Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext' -ArgumentList "https://login.microsoftonline.com/$TargetTenantDomain/oauth2/token"
if ($certificatePrivateKey) {
    $clientCreds = New-Object -TypeName 'Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate' -ArgumentList $AzureADAppCreated.AppId, ([System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String($certificatePrivateKey.SecretValueText)))
    Write-Verbose -Message "Acquiring token resourceAppIdURI $MSGraphResourceUri appSecret $certificatePrivateKey"
    $AuthenticationResult = $AuthenticationContextObj.AcquireTokenAsync($MSGraphResourceUri, $clientCreds).Result
} else {
    Write-Verbose -Message "Acquiring token resourceAppIdURI $MSGraphResourceUri"
    $AuthenticationResult = $AuthenticationContextObj.AcquireToken($MSGraphResourceUri, $PowerShellClientId, $PowerShellClientRedirectUri, [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always)
}

if (-not $AuthenticationResult) {
    throw "Could not retrieve a token for invitation manager api call"
}

$AADInviteBody = @{
    invitedUserEmailAddress = $ResourceTenantAdminEmail
    inviteRedirectUrl = ('https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}' -f $ResourceTenantDomain, $AzureADAppCreated.AppId, $AzureADAppCreated.ReplyUrls[0])
    sendInvitationMessage = $true
    invitedUserMessageInfo = @{
        customizedMessageBody = (@'
{0} wishes to pull mailboxes from your organization using {1}. `
If you recognize this application please click below to provide your consent. `
To authorize this application to be used for office 365 mailbox migration, please add its application id {2} to your organization relationship with {0} in the OAuthApplicationId property.
'@ -f $TargetTenantDomain, $AzureADAppCreated.DisplayName, $AzureADAppCreated.AppId)
    }
}

$AADInviteBodyJson = $AADInviteBody | ConvertTo-Json
$RequestHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$RequestHeaders.Add("Authorization", $AuthenticationResult.CreateAuthorizationHeader())
Write-Verbose -Message "Sending invitation"

$Response = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/invitations" -Body $AADInviteBodyJson -ContentType 'application/json' -Headers $RequestHeaders

if ($Response -and $Response.invitedUserEmailAddress) {
    Write-Host -Object "Successfully sent invitation to $($Response.invitedUserEmailAddress)" -Foreground Green
}
#endregion Sending consent mail to the Resource Tenant Administrator email address specified

#region Target Tenant Exchange Online configuration
Write-Host -Object "Setting up exchange components on target tenant: $TargetTenantDomain"

$ExistingOrgRelationship = Get-OrganizationRelationship | Where-Object -FilterScript { $_.DomainNames -contains $ResourceTenantId }
if ($null -eq $ExistingOrgRelationship) {
    $OrganizationalRelationshipName = '{0}_{1}_{2}' -f $TargetTenantDomain.Split('.')[0], $ResourceTenantDomain.Split('.')[0], $RandomId
    $OrganizationalRelationshipName = $OrganizationalRelationshipName.SubString(0, [System.Math]::Min(64, $OrganizationalRelationshipName.Length))

    Write-Verbose -Message ('Creating organization relationship: {0} in {1}. DomainName: {2}, OAuthApplicationId: {3}' -f $OrganizationalRelationshipName, $TargetTenantDomain, $ResourceTenantId, $AzureADAppCreated.AppId)
    try {
        $null = New-OrganizationRelationship  -DomainNames $ResourceTenantId -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability Inbound -Name $OrganizationalRelationshipName
        Write-Host -Object ('OrganizationRelationship created in {0} for source {1}' -f $TargetTenantDomain, $ResourceTenantDomain) -Foreground Green
    }
    catch {throw $_}
} else {
    Write-Verbose -Message "Organization relationship already exists with $ResourceTenantId. Updating it."
    $capabilities = @($ExistingOrgRelationship.MailboxMoveCapability.Split(",").Trim())
    if (-not $ExistingOrgRelationship.MailboxMoveCapability.Contains("Inbound")) {
        Write-Verbose -Message "Adding Inbound capability to the organization relationship. Existing capabilities: $capabilities"
        $capabilities += "Inbound"
    }

    try {
        $null = $ExistingOrgRelationship | Set-OrganizationRelationship -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability $capabilities
    }
    catch {throw $_}

    $OrganizationalRelationshipName = $ExistingOrgRelationship.Name
}

Write-Verbose -Message ('Creating migration endpoint {0} with remote tenant: {1}, appId: {2}, appSecret: {3}' -f $OrganizationalRelationshipName, $ResourceTenantDomain, $AzureADAppCreated.AppId, $certificatePublicKey.Id)
try {
    $global:MigrationEndpoint = New-MigrationEndpoint -Name $OrganizationalRelationshipName -RemoteTenant $ResourceTenantDomain  -RemoteServer "outlook.office.com" -ApplicationId $AzureADAppCreated.AppId -AppSecretKeyVaultUrl $certificatePublicKey.Id -ExchangeRemoteMove:$true
    Write-Host -Object ('MigrationEndpoint created in {0} for source {1}' -f $TargetTenantDomain, $ResourceTenantDomain) -Foreground Green
    Write-Verbose -Message $global:MigrationEndpoint
}
catch {throw $_}
#endregion Target Tenant Exchange Online configuration

Write-Host -Object 'Exchange setup complete. Migration endpoint details are available in $MigrationEndpoint variable' -Foreground Green

#Returning AppId and CertificateId as global variables to easy usage in Resource Tenant Configuration Script.
Write-Host -Object ('Application details to be registered in organization relationship: ApplicationId: [ {0} ]. KeyVault secret Id: [ {1} ]. These values are available in variables $global:ApplicationId and $global:CertificateId respectively' -f $AzureADAppCreated.AppId, $certificatePublicKey.Id) -Foreground Green
$global:ApplicationId = $AzureADAppCreated.AppId
$global:CertificateId = $certificatePublicKey.Id
#endregion Main