#requires -Modules ExchangeOnlineManagement, AzureAD, ActiveDirectory
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $true
    )]
    [String]$ResourceIdentity,

    [Parameter(
        Mandatory = $true
    )]
    [PSCredential]$ResourceCredential,

    [Parameter(
        Mandatory = $true
    )]
    [String]$ResourceScopingGroup,

    [Parameter(
        Mandatory = $true
    )]
    [PSCredential]$TargetCredential,

    [ValidateSet('ActiveDirectory','CloudOnly')]
    [String]$TargetAccountSource = 'CloudOnly',

    [String]$TargetLicenseSkuId = '189a915c-fe4f-4ffa-bde4-85b9628d07a0',

    [String]$TargetLicenseServiceMatchFilter = 'EXCHANGE'
)

#Connect resource Exchange Online and Azure AD
Connect-ExchangeOnline -Credential $ResourceCredential
Connect-AzureAD -Credential $ResourceCredential

#Get the Resource specified account
$ResourceMailbox = Get-Mailbox -Identity $ResourceIdentity
$ResourceAzureADAccount = Get-AzureADUser -Object $ResourceMailbox.ExternalDirectoryObjectId

#Add the Azure AD Accoun to the Migration scope group
Add-AzureADGroupMember -ObjectId $ResourceScopingGroup -RefObjectId $ResourceAzureADAccount.ObjectId

$CanAddLicenseToTargetAccount = $false
switch ($TargetAccountSource) {
    'CloudOnly' {
        #Disconnect from resource Azure AD and Exchange Online
        Disconnect-ExchangeOnline -Confirm:$false
        Disconnect-AzureAD -Confirm:$false

        #Connect to target Exchange Online and AzureAD
        Connect-ExchangeOnline -Credential $TargetCredential
        Connect-AzureAD -Credential $TargetCredential

        #Gather accepted domains
        $AcceptedDomains = Get-AcceptedDomain

        #check if the mailuser already exists
        $ExistingRecipient = Get-Recipient -Filter ('ExternalEmailAddress -eq "smtp:{0}"' -f $ResourceMailbox.PrimarySmtpAddress)
        if ($null -eq $ExistingRecipient) {
            #No recipient matched, creating a Mailuser
            [Hashtable]$NewMailUserParams = @{
                DisplayName          = $ResourceMailbox.DisplayName
                Name                 = $ResourceMailbox.Name
                LastName             = $ResourceAzureADAccount.Surname
                ExternalEmailAddress = $ResourceMailbox.PrimarySmtpAddress
                Alias                = $ResourceMailbox.Alias
                FirstName            = $ResourceAzureADAccount.GivenName
                PrimarySmtpAddress   = ('{0}@{1}' -f $ResourceMailbox.Alias, $($AcceptedDomains | Where-Object -FilterScript {$_.Default -eq $true} | Select-Object -ExpandPropert Name))
                MicrosoftOnlineServicesID = ('{0}@{1}' -f $ResourceMailbox.Alias, $($AcceptedDomains | Where-Object -FilterScript {$_.InitialDomain -eq $true} | Select-Object -ExpandPropert Name))
                Password             = (ConvertTo-SecureString -String 'P@ssw0rd' -AsPlainText -Force)
            }
            if ($null -ne $ResourceMailbox.Initials) {
                $NewMailUserParams.Add('Initials', $ResourceMailbox.Initials)
            }
            try {
                New-MailUser @NewMailUserParams
            }
            catch {throw}

            [Hashtable]$SetMailUserParams = @{
                Identity     = $ResourceMailbox.Alias
                ExchangeGuid = $ResourceMailbox.ExchangeGuid
                EmailAddresses = @{
                    Add = ('x500:{0}' -f $ResourceMailbox.LegacyExchangeDn)
                }
            }
            if ($null -ne $ResourceMailbox.ArchiveGuid) {
                $SetMailUserParams.Add('ArchiveGuid', $ResourceMailbox.ArchiveGuid)
            }
            try {
                Set-MailUser @SetMailUserParams
                $CanAddLicenseToTargetAccount = $true
            }
            catch {throw}
        } else {
            #Recipient matched

            #Need to work here
        }
    }
    'ActiveDirectory' {
        #Need to work here, also
    }
}

if ($CanAddLicenseToTargetAccount -eq $true) {
    $Sku = Get-AzureADSubscribedSku | Where-Object -FilterScript {$_.SkuId -eq $TargetLicenseSkuId}
    $FeaturesToDisable = $Sku.ServicePlans | ForEach-Object -Process {
        $_ | Where-Object -FilterScript {
            $_ -notin ($Sku.ServicePlans | Where-Object -FilterScript {
                $_.ServicePlanName -match $TargetLicenseServiceMatchFilter
            })
        }
    }

    #Create license object
    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $License.SkuId = $Sku
    $License.DisabledPlans = $FeaturesToDisable.ServicePlanId

    $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $LicensesToAssign.AddLicenses = $License

    $TargetAzureADAccount = Get-AzureADUser -Filter "UserPrincipalName eq '$($NewMailUserParams.PrimarySmtpAddress)'"
    Set-AzureADUserLicense -ObjectId $TargetAzureADAccount.ObjectId -AssignedLicenses $LicensesToAssign
}