#requires -modules ExchangeOnlineManagement
<#
    .SYNOPSIS
    This script can be used by a tenant that wishes to pull resources out of another tenant.
    For example contoso.com would run this script in order to pull mailboxes from fabrikam.com tenant.

    This script is intended for the resource tenant in above example fabrikam.com, and it sets up the organization relationship in exchange to authorize the migration.
    Following are key properties in organization relationship used here:
    - ApplicationId of the azure ad application that resource tenant consents to for mailbox migrations.
    - SourceMailboxMovePublishedScopes contains the groups of users that are in scope for migration. Without this no mailboxes can be migrated.


   .PARAMETER SourceMailboxMovePublishedScopes
   SourceMailboxMovePublishedScopes - Identity of the scope used by source tenant admin.

   .PARAMETER ResourceTenantDomain
   ResourceTenantDomain - the resource tenant.

   .PARAMETER TargetTenantDomain
   TargetTenantDomain - The target tenant.

   .PARAMETER TargetTenantId
   TargetTenantId - The target tenant id.

   .EXAMPLE
   SetupCrossTenantRelationshipForResourceTenant.ps1 -ResourceTenantDomain contoso.onmicrosoft.com -TargetTenantDomain fabrikam.onmicrosoft.com -TargetTenantId d925e0c6-d4db-40c6-a864-49db24af0460 -SourceMailboxMovePublishedScopes "SecurityGroupName"
#>
[CmdletBinding(
    SupportsShouldProcess = $true
)]
Param(
    [Parameter(
        Mandatory = $true,
        HelpMessage = 'Identity of the scope used by source tenant admin.'
    )]
    [String[]]$SourceMailboxMovePublishedScopes,

    [Parameter(
        Mandatory = $true,
        HelpMessage='Resource tenant domain'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [String]$ResourceTenantDomain,

    [Parameter(
        Mandatory = $true,
        HelpMessage = 'Target tenant domain'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [String]$TargetTenantDomain,

    [Parameter(
        HelpMessage = 'The application id for the azure ad application to be used for mailbox migrations'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $ApplicationId = $global:ApplicationId,

    [Parameter(
        Mandatory = $true,
        HelpMessage = 'Target tenant id. This is azure ad directory id or external directory object id in exchange online.'
    )]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $TargetTenantId
)

$ErrorActionPreference = 'Stop'

#Check Exchange Online Powershell Connection
if ($null -eq (Get-Command -Name 'New-OrganizationRelationship' -ErrorAction SilentlyContinue)) {
    Write-Warning -Message 'The current PowerShell Host is not connected to Exchange Online. Trying to connect to Exchange Online Management. Please specify credential that have admin rights on the target tenant.'
    Connect-ExchangeOnline
} else {
    Write-Verbose -Message 'Already connected to Exchange Online.'
    if ((Get-AcceptedDomain).Name -notcontains $ResourceTenantDomain) {
        #Not connected to the target tenant
        Write-Warning -Message 'Not connected to the target Exchange Online. Disconnecting and prompting to reconnect.'
        Disconnect-ExchangeOnline -Confirm:$false
        Connect-ExchangeOnline
    }
}

$ExistingOrgRelationship = Get-OrganizationRelationship | Where-Object -FilterScript { $_.DomainNames -contains $TargetTenantId }

if ($null -eq $ExistingOrgRelationship) {
    $RandomId = '{0:D4}' -f [Random]::new().Next(0, 10000)
    $OrganizationalRelationshipName = '{0}_{1}_{2}' -f $TargetTenantDomain.Split('.')[0], $ResourceTenantDomain.Split('.')[0], $RandomId
    $OrganizationalRelationshipName = $OrganizationalRelationshipName.SubString(0, [System.Math]::Min(64, $OrganizationalRelationshipName.Length))

    Write-Verbose -Message ('Creating organization relationship: {0} in ' -f $OrganizationalRelationshipName, $ResourceTenantDomain)
    try {
        $null = New-OrganizationRelationship -DomainNames $TargetTenantId -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability RemoteOutbound -Name $OrganizationalRelationshipName -OAuthApplicationId $ApplicationId -MailboxMovePublishedScopes $SourceMailboxMovePublishedScopes
        Write-Host -Object ('OrganizationRelationship created in {0} for source {1}' -f $TargetTenantDomain, $ResourceTenantDomain) -Foreground Green
    }
    catch {throw $_}
} else {
    Write-Verbose "Organization relationship already exists with $TargetTenantId. Updating it."
    $capabilities = @($ExistingOrgRelationship.MailboxMoveCapability.Split(",").Trim())
    if (-not $ExistingOrgRelationship.MailboxMoveCapability.Contains("RemoteOutbound")) {
        Write-Verbose "Adding RemoteOutbound capability to the organization relationship. Existing capabilities: $capabilities"
        $capabilities += "RemoteOutbound"
    }

    $ExistingOrgRelationship | Set-OrganizationRelationship -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability $capabilities -OAuthApplicationId $ApplicationId -MailboxMovePublishedScopes $SourceMailboxMovePublishedScopes
}