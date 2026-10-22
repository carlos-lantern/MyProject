<#
.SYNOPSIS
  Lantern employee offboarding runbook (Azure Automation, PowerShell 7.2).

.DESCRIPTION
  Order matters: licensing uses dynamic groups driven by Entra attributes, so
  attributes are only changed AFTER the mailbox is confirmed Shared.
    1. Disable account
    2. Exchange changes (shared, hide from GAL, forwarding, auto-reply)
    3. Verify mailbox is Shared (abort otherwise)
    4. Clear Entra attributes, manager, groups
    5. Intune device category
    6. Laptop return email (if label needed)

  Requires managed identity with Graph (User.ReadWrite.All, Group.ReadWrite.All,
  Directory.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All,
  DeviceManagementServiceConfig.Read.All, RoleManagement.Read.Directory) and
  Exchange.ManageAsApp, with a permanent (non-PIM) Exchange Recipient
  Administrator role assignment (or a least-privilege custom role).
  Automation variable 'FreshserviceApiKey' (encrypted) is required.
  Use -WhatIf for a dry run.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [string] $UserPrincipalName,
    [Parameter(Mandatory)] [datetime] $LastDay,
    [string] $ForwardToEmail,
    [int] $RetentionWeeks = 8,
    [bool] $LaptopLeftInOffice = $false,
    [string] $PersonalEmail,
    [datetime] $DevicePurchaseDate,
    [string] $TenantDomain = 'lanternstudios.com',
    [int] $TicketId,
    [string] $FreshserviceDomain = 'lanternstudios.freshservice.com'
)

$ErrorActionPreference = 'Stop'

function Write-Step($m) { Write-Output "[$(Get-Date -Format 'HH:mm:ss')] $m" }

Connect-MgGraph -Identity -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization $TenantDomain -ShowBanner:$false

$user = Get-MgUser -UserId $UserPrincipalName -Property Id, DisplayName, GivenName, Surname, AccountEnabled
Write-Step "Offboarding $($user.DisplayName) ($UserPrincipalName)"

# Guard: never offboard accounts holding directory roles
$roles = Get-MgUserTransitiveMemberOf -UserId $user.Id -All |
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' }
if ($roles) { throw "$UserPrincipalName holds directory roles; refusing to automate. Handle manually." }

# 1. Disable account
if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Disable Entra account')) {
    Update-MgUser -UserId $user.Id -AccountEnabled:$false
    Write-Step 'Account disabled'
}

# 2. Exchange
$mbx = Get-Mailbox -Identity $UserPrincipalName
if ($mbx.RecipientTypeDetails -ne 'SharedMailbox') {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Convert mailbox to Shared')) {
        Set-Mailbox -Identity $UserPrincipalName -Type Shared
        Write-Step 'Conversion to shared requested'
    }
}

if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Update mailbox settings')) {
    Set-Mailbox -Identity $UserPrincipalName -CustomAttribute1 $null -HiddenFromAddressListsEnabled $true

    if ($ForwardToEmail) {
        Set-Mailbox -Identity $UserPrincipalName -ForwardingSmtpAddress $ForwardToEmail -DeliverToMailboxAndForward $false
        $fwdName = (Get-MgUser -UserId $ForwardToEmail -Property DisplayName).DisplayName
        $reply = @"
Hi there,<br><br>
Thanks for reaching out! Just a heads-up, I'm Out of Office. But don't worry, your message will be forwarded to $fwdName ($ForwardToEmail), who will be happy to assist you.<br><br>
Cheers,
"@
        Set-MailboxAutoReplyConfiguration -Identity $UserPrincipalName -AutoReplyState Enabled `
            -InternalMessage $reply -ExternalMessage $reply
    }
    Write-Step 'Mailbox settings applied'
}

# 3. Verify shared before touching attributes (WhatIf skips the gate)
if (-not $WhatIfPreference) {
    $ok = $false
    foreach ($i in 1..30) {
        if ((Get-Mailbox -Identity $UserPrincipalName).RecipientTypeDetails -eq 'SharedMailbox') { $ok = $true; break }
        Start-Sleep -Seconds 20
    }
    if (-not $ok) { throw "Mailbox is not Shared after waiting. Aborting before attribute changes." }

    $stats = Get-MailboxStatistics -Identity $UserPrincipalName
    $sizeGB = [math]::Round(($stats.TotalItemSize.Value.ToBytes() / 1GB), 2)
    $hold = (Get-Mailbox -Identity $UserPrincipalName).LitigationHoldEnabled
    if ($sizeGB -ge 50 -or $hold) {
        Write-Warning "Mailbox is $sizeGB GB / LitigationHold=$hold - a license may be required. Review manually."
    }
    Write-Step "Mailbox verified Shared ($sizeGB GB)"
}

# 4. Entra attributes, manager, groups
if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Clear Entra attributes, manager and groups')) {
    $body = @{
        companyName    = 'Lantern-disabled'
        jobTitle       = $null
        department     = $null
        employeeType   = $null
        officeLocation = $null
        state          = $null
        country        = $null
    } | ConvertTo-Json
    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($user.Id)" `
        -Body $body -ContentType 'application/json'

    Remove-MgUserManagerByRef -UserId $user.Id -ErrorAction SilentlyContinue

    $groups = Get-MgUserMemberOf -UserId $user.Id -All |
        Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }
    foreach ($g in $groups) {
        $types = $g.AdditionalProperties['groupTypes']
        if ($types -contains 'DynamicMembership') { continue } # removed automatically by the attribute change
        try { Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $user.Id }
        catch { Write-Warning "Could not remove from $($g.AdditionalProperties['displayName']): $($_.Exception.Message)" }
    }
    Write-Step 'Entra attributes cleared and groups removed'
}

# 5. Intune device category
$category =
    if (-not $LaptopLeftInOffice) { 'Lantern-In Transit' }
    elseif ($DevicePurchaseDate -and $DevicePurchaseDate -le (Get-Date).AddYears(-3)) { 'Lantern-retired' }
    else { 'Lantern-Inventory' }

$devices = Get-MgDeviceManagementManagedDevice -Filter "userPrincipalName eq '$UserPrincipalName'" -All
$cat = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCategories').value |
    Where-Object displayName -eq $category
if (-not $cat) { Write-Warning "Device category '$category' not found" }
foreach ($d in $devices) {
    if ($cat -and $PSCmdlet.ShouldProcess($d.DeviceName, "Set device category $category")) {
        Invoke-MgGraphRequest -Method PUT `
            -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($d.Id)/deviceCategory/`$ref" `
            -Body (@{ '@odata.id' = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$($cat.id)" } | ConvertTo-Json) `
            -ContentType 'application/json'
        Write-Step "$($d.DeviceName) -> $category"
    }
}
if (-not $devices) { Write-Warning 'No Intune device found for user' }

# 6. Laptop return email
if (-not $LaptopLeftInOffice -and $PersonalEmail) {
    $html = @"
<p>Hi,</p>
<p>HR has confirmed that your employment with Lantern ended today on $($LastDay.ToString('MMMM d, yyyy')).<br>
In line with our asset-return policy, we need to collect the company-issued laptop and charger in your possession.</p>
<p><b>How the return works</b></p>
<ul>
<li>Reply with your preferred return address. We'll use this only as the "From" address on the prepaid UPS label.</li>
<li>You'll receive the label by email within one business day.</li>
<li>Place the laptop and charger in a sturdy box, affix the label, and drop the package at any UPS location.</li>
<li>Please ship within three business days of receiving the label.</li>
</ul>
<p>A few reminders:</p>
<ul>
<li>Be sure the laptop is fully powered off before packing.</li>
<li>Depending on the state of your laptop, it may be re-used or recycled. In either case, the drive will be wiped to destroy all the data on it.</li>
</ul>
<p>Thank you for helping us complete the hand-off quickly.</p>
<p>Best regards,</p>
"@
    if ($PSCmdlet.ShouldProcess("ticket $TicketId", 'Send laptop return reply via Freshservice')) {
        # Replying on the ticket keeps the employee's answer (return address) on the same ticket
        # and avoids mailing from helpdesk@, which would open a new ticket.
        $apiKey = Get-AutomationVariable -Name 'FreshserviceApiKey'
        $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${apiKey}:X"))
        $reply = @{ body = $html; cc_emails = @($PersonalEmail) } | ConvertTo-Json
        Invoke-RestMethod -Method POST -Uri "https://$FreshserviceDomain/api/v2/tickets/$TicketId/reply" `
            -Headers @{ Authorization = "Basic $auth" } -ContentType 'application/json' -Body $reply | Out-Null
        Write-Step 'Return email sent via Freshservice ticket reply'
    }
}

$deleteOn = $LastDay.AddDays(7 * $RetentionWeeks).ToString('yyyy-MM-dd')
Write-Step "DONE. Create reminder: delete shared mailbox $UserPrincipalName on $deleteOn"

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph | Out-Null
