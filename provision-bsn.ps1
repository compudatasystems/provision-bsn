<#
.SYNOPSIS
    Guided provisioning of Breach Secure Now "PII Protect" for a new client: automates the
    Microsoft Entra / Microsoft 365 setup and walks the operator through the PII Protect portal
    steps that have no API.

.DESCRIPTION
    Runs the BSN onboarding runbook in order. Each phase either:
      * automates an Entra / M365 action (security groups, dynamic membership, anti-spam allowed
        senders, the PII-Protect SSO enterprise app (non-gallery SAML)), or
      * pauses with numbered instructions + the exact URL for a step you must do by hand: the PII
        Protect portal (add tenant, enable Directory Sync, enable Direct Delivery) and the
        Microsoft 365 admin center (deploy the Catch Phish Outlook add-in), then continues.

    Entra work uses Microsoft Graph via Invoke-MgGraphRequest, so the only required module is
    Microsoft.Graph.Authentication. The anti-spam phase additionally uses ExchangeOnlineManagement
    and reuses your Graph sign-in via SSO. Any required module is installed automatically
    (CurrentUser scope) on first use. Everything honours -WhatIf and is safe to re-run (idempotent
    creates skip existing objects).

    Aimed at a non-technical operator: it attempts every automatable step by default (dynamic
    membership when the tenant has P1, the anti-spam allow-list), prompts for any required input
    that wasn't supplied, and if a phase fails it warns and keeps going rather than aborting.

    NOTE: portal.pii-protect.com has no public API, so tenant creation, Directory Sync consent,
    and Direct Delivery consent are inherently manual — the script prompts you through them. The
    Catch Phish add-in is manual for the same reason: Office centralized deployment is not on Graph,
    and its PowerShell module is Windows-only and documented as Basic-auth only (which mandatory
    admin MFA rules out), so the admin center is the supported path. Verify still checks the result.

.PARAMETER ClientName
    The client's business name, as it should appear in the PII Protect portal. Used in prompts and
    checklists. Keep it consistent with however you name clients elsewhere (your PSA, billing) so the
    BSN portal matches your records. If omitted you're prompted for it (required unless
    -NonInteractive).

.PARAMETER TenantId
    Client tenant id or domain to sign in to. Defaults to the tenant you authenticate against.

.PARAMETER IncludePartnerAdmins
    Also create the BSN-PartnerAdmins group. Per BSN, this is ONLY for your own internal BPP
    account — do not use it in a client tenant.

.PARAMETER TagGroup
    One or more organizational tag names. Creates BSN-TAG-<name> security groups (e.g.
    -TagGroup 'Outside Sales','Service' -> BSN-TAG-Outside Sales, BSN-TAG-Service) — the casing
    BSN's docs use predominantly. BSN-Tag- also appears in the wild, so -Verify matches either
    (Graph's name filter is case-insensitive) and reports the casing actually in the tenant.

.PARAMETER NoDynamicEmployees
    By default BSN-Employees is made a Dynamic User group (auto-enrolment) when the tenant has
    Entra ID P1+; without P1 it falls back to an assigned group and prints how to enrol users
    manually. Pass this switch to force a plain assigned group regardless of licensing.

.PARAMETER EnrollServicePlanId
    Service plan id(s) whose holders should be enrolled (used to build the dynamic rule). If
    omitted, the rule matches any user with an enabled service plan. Plan id reference:
    https://learn.microsoft.com/entra/identity/users/licensing-service-plan-reference

.PARAMETER ExcludeEmail
    Email address(es) to exclude from the dynamic BSN-Employees rule.

.PARAMETER SkipAllowedSenders
    By default the script adds the BSN sending addresses to the default anti-spam inbound policy
    (ExchangeOnlineManagement, reusing your Graph sign-in via SSO). Pass this to skip that phase.

.PARAMETER SkipCatchPhish
    Skip the guided Catch Phish add-in step, and its verification check. Catch Phish is deployed by
    hand in the Microsoft 365 admin center (Settings > Integrated apps) — Office centralized
    deployment has no Graph API, so the script links you straight to the page and the marketplace
    listing rather than automating it. Note it cannot be assigned to BSN-Employees/BSN-Managers:
    centralized deployment does not support mail-disabled security groups. Deploy to Everyone.

.PARAMETER CatchPhishOnly
    Run ONLY the Catch Phish step: the guided admin-center walk-through, then the check. Catch Phish
    is often deployed days or weeks after the initial onboarding, so this lets you come back and do
    just that visit — no ClientName needed, no other phases, and read-only Graph scopes (the step
    itself writes nothing; you do the deploying in the admin center).
    Add -Verify to only run the check, without the walk-through prompt.

.PARAMETER SsoRedirectUri
    The Redirect URL shown in the portal's Single Sign-On popup. If omitted, the SSO phase
    prompts for it interactively.

.PARAMETER SsoAppIdUri
    The Application ID URI shown in the portal's Single Sign-On popup. If omitted, prompted.

.PARAMETER SkipSso
    Skip the Single Sign-On phase entirely.

.PARAMETER SsoSyncType
    Which attribute BSN receives as the user's email in the SAML assertion. 'Email' (default)
    uses Azure's built-in emailaddress-from-user.mail claim (no claims policy). 'UPN' adds a
    claims-mapping policy remapping emailaddress -> user.userPrincipalName (this needs the per-app
    signing certificate the SSO phase creates). Match this to the "Portal Logon" choice you made
    when enabling Directory Sync.

.PARAMETER SsoCertYears
    Lifetime (1-3 years, default 3) of the per-app SAML signing certificate the SSO phase creates.
    3 is Azure's maximum for an auto-generated cert; the default minimises rotation, since rotating
    means re-pasting metadata in the BSN portal.

.PARAMETER RenewSsoCert
    Renew ONLY the SSO SAML signing certificate, regardless of how much life it has left, then stop.
    For a suspected compromise, a policy-driven rotation, or getting ahead of the expiry -Verify
    warns about. Adds a fresh certificate (-SsoCertYears, default 3), makes it the preferred signing
    key, and prints the metadata URL to paste back into the BSN portal.
    DISRUPTIVE: sign-in for that client breaks the moment the new certificate takes over and stays
    broken until the portal has the new metadata — so it asks before doing anything, and honours
    -WhatIf. Deliberately separate from -Fix, which only ever makes changes that cannot break a
    working setup. Only applies to the SAML enterprise app; the older app-registration shape signs
    with the tenant's Microsoft-managed default key and has nothing to renew.

.PARAMETER ReplaceSso
    If an app already holds this SSO Entity ID (Application ID URI) — e.g. a legacy PII-Protect app
    registration — delete it and recreate as the new SAML enterprise app WITHOUT prompting. A SAML
    Entity ID must be unique tenant-wide, so the old app has to go first. Interactively you are
    asked before deleting; in -NonInteractive nothing is deleted unless you pass this switch.

.PARAMETER NonInteractive
    Don't pause for the manual portal steps or prompt for values (automate what's possible, print
    the rest as a checklist). SSO is skipped unless -SsoRedirectUri and -SsoAppIdUri are supplied.

.PARAMETER UseDeviceCode
    Sign in with a device code instead of a browser (headless / SSH sessions).

.PARAMETER Verify
    Run the read-only verification pass ONLY (no provisioning), then finish with a plain-English
    summary of what it means and what to do next. Checks the security groups, the SSO app (including
    warning ~60 days BEFORE its SAML signing certificate expires — when that lapses, users can no
    longer sign in), the anti-spam allow-list, whether the Catch Phish add-in was ever deployed, and
    BSN's Directory-Sync / DMD enterprise apps. Connects with read-only Graph scopes (no write
    consent). Add -SkipAllowedSenders to avoid the extra Exchange Online sign-in (that also skips the
    Exchange half of the Catch Phish check). A normal run prints all of this automatically at the end.

.PARAMETER Fix
    Run the verification pass AND remediate the automatable gaps it finds (no manual portal steps,
    no ClientName needed). Only SAFE, additive fixes are offered: add missing BSN allowed senders,
    create genuinely-missing security groups, rename a misnamed SSO app to 'PII-Protect SSO', and
    set the SAML enterprise app's assignment-required to No. It never rotates a signing certificate
    or changes the email claim (those can break a working SSO — use a normal run / -ReplaceSso).
    Each fix is confirmed interactively (y/N); -NonInteractive applies them all; -WhatIf previews.
    Needs write scopes, so it connects like a provisioning run. The closing plain-English summary
    separates what was repaired from what still needs a human — e.g. deploying Catch Phish, which
    has no automatic fix.

.EXAMPLE
    # Full guided onboarding — attempts everything it can, prompts through the portal steps.
    ./provision-bsn.ps1 -ClientName "Acme Widgets"

.EXAMPLE
    # Entra-only bootstrap: no portal prompts, no Exchange, no SSO.
    ./provision-bsn.ps1 -ClientName "Acme Widgets" -NonInteractive -SkipAllowedSenders -SkipSso

.EXAMPLE
    # Rotate a client's SSO signing certificate on demand (asks first; -WhatIf to preview).
    ./provision-bsn.ps1 -RenewSsoCert

.EXAMPLE
    # Weeks later: just do the Catch Phish add-in for an already-onboarded client.
    ./provision-bsn.ps1 -CatchPhishOnly

.EXAMPLE
    # Just check whether Catch Phish is deployed — no prompts, read-only.
    ./provision-bsn.ps1 -CatchPhishOnly -Verify
#>
#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string]$ClientName,

    [string]$TenantId,
    [switch]$IncludePartnerAdmins,
    [string[]]$TagGroup,

    [switch]$NoDynamicEmployees,
    [string[]]$EnrollServicePlanId,
    [string[]]$ExcludeEmail,

    [switch]$SkipAllowedSenders,
    [switch]$SkipCatchPhish,
    [switch]$CatchPhishOnly,

    [string]$SsoRedirectUri,
    [string]$SsoAppIdUri,
    [switch]$SkipSso,
    [ValidateSet('Email', 'UPN')]
    [string]$SsoSyncType = 'Email',
    [ValidateRange(1, 3)]
    [int]$SsoCertYears = 3,
    [switch]$ReplaceSso,
    [switch]$RenewSsoCert,

    [switch]$NonInteractive,
    [switch]$UseDeviceCode,

    [switch]$Verify,
    [switch]$Fix
)

$ErrorActionPreference = 'Stop'

# Well-known ids.
$GraphAppId    = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
$GraphEmailScope = '64a6cdd6-aab1-4aac-85b6-d16ff39b2ee6' # delegated 'email'
$DefaultAppRole  = '00000000-0000-0000-0000-000000000000' # "default access" (no app roles)
$SamlAppTemplateId = '8adf8e6e-67b2-4cf2-a259-e3dc5476c621' # non-gallery "custom" application template

# BSN sending addresses to allow through anti-spam.
$BsnSenders = @('no-reply@pii-protect.com', 'no-reply@security-reminders.com', 'no-reply@breachsecurenow.com')

# Catch Phish (the BSN Outlook add-in). AssetIds are the Microsoft Marketplace id for an add-in and
# always start with 'WA'; the marketplace URL wants it lower-cased. Centralized deployment has no
# Graph API and its PowerShell module is Windows-only, so deployment is a guided admin-center step
# (Invoke-CatchPhishStep) and verify infers the result (Test-BsnCatchPhish).
$CatchPhishAssetId     = 'WA200008655'
$CatchPhishMarketUrl   = "https://marketplace.microsoft.com/en-us/product/office/$($CatchPhishAssetId.ToLower())"
$IntegratedAppsUrl     = 'https://admin.microsoft.com/Adminportal/Home#/Settings/IntegratedApps'

# --- Presentation helpers --------------------------------------------------------------
function Write-Section([string]$title) {
    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor DarkGray
    Write-Host "  $title" -ForegroundColor White
    Write-Host ('=' * 72) -ForegroundColor DarkGray
}

# Print a manual portal step and wait (unless -NonInteractive). Most steps are in the PII Protect
# portal (hence the default -Portal), but Catch Phish lives in the Microsoft 365 admin center.
# -Links takes an ordered label -> URL map for steps needing more than one address; most terminals
# render them clickable. -Notes are caveats printed after the steps, where they're still on screen
# while the operator works.
function Invoke-ManualStep {
    param(
        [string]$Title,
        [string]$Url,
        [string[]]$Steps,
        [string]$Portal = 'PII Protect portal',
        [System.Collections.Specialized.OrderedDictionary]$Links,
        [string[]]$Notes
    )
    Write-Section "MANUAL ($Portal): $Title"
    if ($Url) { Write-Host "  Open: $Url" -ForegroundColor Cyan }
    if ($Links) {
        # Pad the labels so the URLs line up and stay easy to click/copy.
        $width = ($Links.Keys | Measure-Object -Property Length -Maximum).Maximum
        foreach ($k in $Links.Keys) { Write-Host ("  {0}  {1}" -f "$($k):".PadRight($width + 1), $Links[$k]) -ForegroundColor Cyan }
    }
    # Same convention as -Notes: a caller-indented line is an aside on the step above, not a step of
    # its own, so it must not take a number — the numbers should match the source article's steps.
    $i = 1
    foreach ($s in $Steps) {
        if ($s -match '^\s') { Write-Host ("      {0}" -f $s.TrimStart()) -ForegroundColor DarkGray }
        else                 { Write-Host ("   {0}. {1}" -f $i++, $s) }
    }
    # '!' marks the start of a note; lines the caller indented are continuations of the one above.
    foreach ($n in $Notes) {
        if ($n -match '^\s') { Write-Host "    $($n.TrimStart())" -ForegroundColor Yellow }
        else                 { Write-Host "  ! $n" -ForegroundColor Yellow }
    }
    if ($NonInteractive) {
        Write-Host '  (-NonInteractive: do this manually later.)' -ForegroundColor Yellow
    } else {
        [void](Read-Host "`n  Press Enter when this step is complete")
    }
}

function Read-Required([string]$label, [string]$preset) {
    if ($preset) { return $preset }
    if ($NonInteractive) { return '' }
    return (Read-Host "  $label")
}

# Run an automated phase so a failure warns and the walk-through continues instead of aborting.
$script:phaseErrors = @()
function Invoke-Phase([string]$name, [scriptblock]$action) {
    try { & $action }
    catch {
        Write-Warning "Phase '$name' failed: $($_.Exception.Message)"
        Write-Host "  Skipping the rest of this phase; revisit '$name' afterwards." -ForegroundColor Yellow
        $script:phaseErrors += $name
    }
}

$script:fixesApplied = @()
# Did the SSO phase actually finish? The Done summary used to infer it from -SkipSso not being
# passed, so a run that bailed out ("Missing Redirect URL... skipping SSO") still reported SSO as
# automated. Report what happened, not what was asked for.
$script:ssoConfigured = $false
# Offer to remediate a gap the verification pass found. No-op unless -Fix. Interactive: prompt y/N;
# -NonInteractive: apply automatically; -WhatIf: preview only (via ShouldProcess). $action performs
# the write, so it runs only once the fix is approved; a failed fix warns and the pass continues.
# Only ever wired up for SAFE, additive fixes (see -Fix help) — nothing that can break a working SSO.
function Invoke-Fix([string]$what, [scriptblock]$action) {
    # Record that a fix EXISTS for the check just reported — before the -Fix gate, so a read-only
    # pass knows it too. That's what lets the summary offer "-Fix can repair some of this" only when
    # it's actually true, instead of sending the operator to a command that would do nothing.
    if ($script:checkResults.Count) { $script:checkResults[-1].Fixable = $true }
    if (-not $Fix) { return }
    if (-not $PSCmdlet.ShouldProcess($what, 'Fix')) { return }   # honours -WhatIf / -Confirm
    if ($NonInteractive) {
        Write-Host "       -> fixing: $what" -ForegroundColor Cyan
    } elseif ((Read-Host "       -> fix now: $what ? (y/N)") -notmatch '^\s*(y|yes)\s*$') {
        Write-Host '          skipped.' -ForegroundColor DarkGray
        return
    }
    try {
        & $action
        # Mark the check this fix answers, so the summary can separate "found and repaired" from
        # "still broken". Every call site runs Invoke-Fix directly after the failing Write-Check, so
        # the last recorded result is always the one being fixed.
        if ($script:checkResults.Count) { $script:checkResults[-1].Fixed = $true }
        $script:fixesApplied += $what
    }
    catch { Write-Warning "Fix failed ($what): $($_.Exception.Message)" }
}

# Make sure a required PowerShell module is present, installing it (CurrentUser) if not. Returns
# $true if usable. -Force trusts PSGallery so the install runs unattended.
function Ensure-Module([string]$name) {
    if (Get-Module -ListAvailable -Name $name) { return $true }
    Write-Host "Installing $name (one-time, current user)..." -ForegroundColor Cyan
    # Suppress the noisy 'PackageManagement/PowerShellGet in use' warnings; judge success by
    # whether the module is actually available afterwards, not by whether Install-Module threw.
    try { Install-Module $name -Scope CurrentUser -Force -AllowClobber -WarningAction SilentlyContinue -ErrorAction Stop } catch { }
    if (Get-Module -ListAvailable -Name $name) { return $true }
    Write-Warning "Could not install $name automatically. Install it manually, then re-run:  Install-Module $name -Scope CurrentUser"
    return $false
}

# --- Graph REST convenience ------------------------------------------------------------
$GraphBase = 'https://graph.microsoft.com/v1.0'
function Graph-Get([string]$path) {
    return Invoke-MgGraphRequest -Method GET -OutputType PSObject -Uri "$GraphBase$path"
}
function Graph-Post([string]$path, [hashtable]$body) {
    return Invoke-MgGraphRequest -Method POST -OutputType PSObject -Uri "$GraphBase$path" -Body $body
}
function Graph-Patch([string]$path, [hashtable]$body) {
    Invoke-MgGraphRequest -Method PATCH -Uri "$GraphBase$path" -Body $body | Out-Null
}
function Graph-Delete([string]$path) {
    Invoke-MgGraphRequest -Method DELETE -Uri "$GraphBase$path" | Out-Null
}

# $true if the tenant owns Entra ID P1/P2 (required for dynamic groups); $false if not; $null if
# it can't be determined (treated as "attempt, but fall back to assigned on failure").
function Test-EntraP1 {
    try {
        $skus = (Graph-Get '/subscribedSkus?$select=skuPartNumber,servicePlans').value
        foreach ($sku in $skus) {
            foreach ($p in $sku.servicePlans) {
                if ($p.servicePlanName -in @('AAD_PREMIUM', 'AAD_PREMIUM_P2') -and $p.provisioningStatus -eq 'Success') {
                    return $true
                }
            }
        }
        return $false
    } catch {
        Write-Warning "Could not check Entra ID P1 availability: $($_.Exception.Message)"
        return $null
    }
}

# --- Phase: security groups ------------------------------------------------------------
function New-MailNickname([string]$name) { return ($name -replace '[^A-Za-z0-9]', '') }

function Ensure-BsnGroup {
    param([string]$Name, [string]$Description, [switch]$Dynamic, [string]$Rule)
    try {
        $safe = $Name -replace "'", "''"
        $found = @((Graph-Get "/groups?`$filter=displayName eq '$safe'&`$select=id,displayName,groupTypes,membershipRule").value)
        if ($found.Count -gt 0) {
            $g = $found[0]
            Write-Host "  '$Name' already exists (id $($g.id))." -ForegroundColor DarkGray
            if ($Dynamic -and ($g.groupTypes -notcontains 'DynamicMembership')) {
                if ($PSCmdlet.ShouldProcess($Name, 'Convert to Dynamic User membership')) {
                    Write-Warning "Converting '$Name' to dynamic removes any manually-assigned members."
                    Graph-Patch "/groups/$($g.id)" @{
                        groupTypes = @('DynamicMembership'); membershipRule = $Rule; membershipRuleProcessingState = 'On'
                    }
                    Write-Host "  converted '$Name' to dynamic membership." -ForegroundColor Green
                    # Reflect the conversion in what we hand back. $g was read BEFORE the PATCH, so
                    # without this the caller sees a stale 'assigned' shape and wrongly reports the
                    # group still needs hand-enrolment.
                    $g.groupTypes = @('DynamicMembership')
                    $g.membershipRule = $Rule
                }
            }
            return $g
        }
        if (-not $PSCmdlet.ShouldProcess($Name, 'Create security group')) { return $null }
        $body = @{
            displayName = $Name; description = $Description
            mailEnabled = $false; mailNickname = (New-MailNickname $Name); securityEnabled = $true
        }
        if ($Dynamic) {
            $body.groupTypes = @('DynamicMembership')
            $body.membershipRule = $Rule
            $body.membershipRuleProcessingState = 'On'
        }
        try {
            $g = Graph-Post '/groups' $body
        } catch {
            if (-not $Dynamic) { throw }
            # Fall back to an assigned group (e.g. P1 check was inconclusive but the tenant lacks it).
            Write-Warning "Dynamic create failed for '$Name' ($($_.Exception.Message)); creating it as an ASSIGNED group instead."
            'groupTypes', 'membershipRule', 'membershipRuleProcessingState' | ForEach-Object { $body.Remove($_) }
            $g = Graph-Post '/groups' $body
            $Dynamic = $false
        }
        Write-Host "  created '$Name' (id $($g.id))$(if ($Dynamic) { ' [dynamic]' })." -ForegroundColor Green
        return $g
    } catch {
        Write-Warning "Could not ensure group '$Name': $($_.Exception.Message)"
        return $null
    }
}

# Build the BSN-Employees dynamic membership rule.
function Build-EmployeeRule([string[]]$planIds, [string[]]$excludeEmails) {
    if ($planIds) {
        $clauses = $planIds | ForEach-Object {
            "(user.assignedPlans -any (assignedPlan.servicePlanId -eq `"$_`" -and assignedPlan.capabilityStatus -eq `"Enabled`"))"
        }
        $licensed = '(' + ($clauses -join ' or ') + ')'
    } else {
        $licensed = '(user.assignedPlans -any (assignedPlan.capabilityStatus -eq "Enabled"))'
    }
    $rule = "$licensed and (user.surname -ne null) and (user.givenName -ne null)"
    if ($excludeEmails) {
        $list = ($excludeEmails | ForEach-Object { "`"$_`"" }) -join ','
        $rule = "($rule) and (user.mail -notIn [$list])"
    }
    return $rule
}

function Invoke-GroupsPhase {
    Write-Section 'Microsoft Entra: BSN security groups'

    # Dynamic membership is attempted by default (needs Entra ID P1/P2). Preflight, then degrade
    # gracefully if P1 is missing. -NoDynamicEmployees forces a plain assigned group.
    $useDynamic = -not $NoDynamicEmployees
    if ($useDynamic) {
        $p1 = Test-EntraP1
        if ($p1 -eq $false) {
            Write-Host '  Entra ID P1/P2 not found; BSN-Employees will be a standard (assigned) group.' -ForegroundColor Yellow
            $useDynamic = $false
        } elseif ($null -eq $p1) {
            Write-Host '  Could not confirm P1; attempting dynamic membership (falls back to assigned on failure).' -ForegroundColor Yellow
        } else {
            Write-Host '  Entra ID P1/P2 detected; BSN-Employees will auto-enrol via dynamic membership.' -ForegroundColor DarkGray
        }
    }

    $rule = if ($useDynamic) { Build-EmployeeRule $EnrollServicePlanId $ExcludeEmail }
    if ($useDynamic) { Write-Host "  Dynamic rule for BSN-Employees:`n    $rule" -ForegroundColor DarkGray }

    $emp = Ensure-BsnGroup -Name 'BSN-Employees'   -Description 'PII/PHI Protect Standard Users'             -Dynamic:$useDynamic -Rule $rule
    Ensure-BsnGroup -Name 'BSN-Managers'       -Description 'PII/PHI Protect Manager Role'                 | Out-Null
    Ensure-BsnGroup -Name 'BSN-ManagerAdmins'  -Description 'PII/PHI Protect Manager Administrator Role'   | Out-Null
    if ($IncludePartnerAdmins) {
        Ensure-BsnGroup -Name 'BSN-PartnerAdmins' -Description 'PII/PHI Protect Partner Administrator Role' | Out-Null
    }
    foreach ($tag in $TagGroup) {
        Ensure-BsnGroup -Name "BSN-TAG-$tag" -Description "PII/PHI Protect Tag: $tag" | Out-Null
    }

    # If BSN-Employees didn't end up dynamic, say what that means and how to enrol by hand — but
    # only when it's actually true. Under -WhatIf nothing was converted, so a preview that just
    # showed the conversion must not then announce the group needs hand-enrolment. And name the real
    # reason: telling an operator to buy P1 they already own is worse than saying nothing.
    $empDynamic = [bool]($emp -and (@($emp.groupTypes) -contains 'DynamicMembership'))
    if (-not $empDynamic -and -not ($useDynamic -and $WhatIfPreference)) {
        Write-Host ''
        Write-Host '  ACTION NEEDED — BSN-Employees is a standard (assigned) group:' -ForegroundColor Yellow
        Write-Host '    Employees are NOT enrolled automatically. Add them by hand:'
        Write-Host '      1. Go to https://entra.microsoft.com  >  Groups  >  All groups  >  BSN-Employees'
        Write-Host '      2. Members  >  + Add members  >  select each employee  >  Select'
        if ($NoDynamicEmployees) {
            Write-Host '    You passed -NoDynamicEmployees. Re-run without it to switch to automatic'
            Write-Host '    enrolment (conversion drops any manually-added members).'
        } elseif (-not $useDynamic) {
            Write-Host '    To switch to automatic enrolment later, add Microsoft Entra ID P1 licensing and'
            Write-Host '    re-run this script — it will convert BSN-Employees to dynamic membership.'
        } else {
            Write-Host '    Automatic enrolment was attempted but did not take effect — re-run with -Verify'
            Write-Host '    to see the current state.'
        }
    }
}

# --- Phase: anti-spam allowed senders (Exchange Online) --------------------------------
function Invoke-AllowedSendersPhase {
    Write-Section 'Microsoft 365 Defender: anti-spam allowed senders'
    if (-not (Ensure-Module 'ExchangeOnlineManagement')) {
        Write-Host "  Add these to the default anti-spam inbound policy manually: $($BsnSenders -join ', ')" -ForegroundColor Yellow
        return
    }
    try {
        Import-Module ExchangeOnlineManagement
        # Reuse the Graph sign-in: passing the same account uses the cached token (SSO), so this
        # connect should be silent or a one-click account pick rather than a fresh credential entry.
        $connect = @{ ShowBanner = $false }
        $upn = (Get-MgContext).Account
        if ($upn)      { $connect.UserPrincipalName = $upn }
        if ($TenantId) { $connect.DelegatedOrganization = $TenantId }
        Write-Host "  Connecting to Exchange Online$(if ($upn) { " as $upn (single sign-on)" })..." -ForegroundColor Cyan
        Connect-ExchangeOnline @connect | Out-Null
        if ($PSCmdlet.ShouldProcess('Default anti-spam inbound policy', "Allow senders: $($BsnSenders -join ', ')")) {
            Set-HostedContentFilterPolicy -Identity Default -AllowedSenders @{ Add = $BsnSenders } | Out-Null
            Write-Host "  Allowed: $($BsnSenders -join ', ')" -ForegroundColor Green
        }
    } catch {
        Write-Warning "Anti-spam configuration failed: $($_.Exception.Message)"
        Write-Host "  Add manually at https://security.microsoft.com/antispam : $($BsnSenders -join ', ')" -ForegroundColor Yellow
    } finally {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    }
}

# --- Phase: SSO enterprise app (non-gallery SAML) --------------------------------------
# SOURCE: BSN's official article "How to Set Up Microsoft Office 365 Single Sign-On (SSO) in the
# Breach Secure Now Portal" (BSN partner portal, retrieved 2026-07-15) — CONFIRMED 2026-07-15. It
# specifies exactly what this phase builds: Enterprise applications > New application > Create your
# own > non-gallery; Single sign-on > SAML; identifier = the portal's Application ID URL; reply URL =
# the portal's Redirect URL; Assignment required = No; then paste the App Federation Metadata URL
# back into the portal. Email sync claims user.mail; UPN sync claims user.userprincipalname.
# (That article is marked BSN Confidential — do not reproduce it here; this comment records only
# which end state it requires, which is what the code has to build.)
#
# TWO DEVIATIONS from that article, both deliberate, both worth re-checking on a pilot:
#  1. Per-app signing certificate. The article never creates one — it just copies the metadata URL,
#     using whatever cert Entra generates when SAML is configured in the PORTAL. This phase builds
#     the app over Graph, where that auto-generation can't be relied on, so it adds an explicit cert
#     (-SsoCertYears). Extra, not contrary.
#  2. Claims. The article uses the portal's Attributes & Claims UI; this uses a claimsMappingPolicy
#     (UPN only). Same intended assertion, different mechanism — verify the assertion on a pilot
#     before trusting it for UPN clients.
#
# Note the article also has a SEPARATE procedure for GCC High / GCC Low tenants.

# Page app registrations and return those that conflict with the SSO we're about to create: any
# app already holding this Entity ID (identifierUris), or one carrying BSN's Cognito SAML
# fingerprint (reply URL …amazoncognito.com/saml2/idpresponse, or Entity ID urn:amazon:cognito).
# Name / owner-tenant agnostic — display names and appIds can be anything. Queries /applications,
# so it returns only LOCAL app registrations (a legacy PII-Protect SSO app), not BSN's own
# multi-tenant enterprise apps (Directory Sync / DMD live in BSN's tenant, so they never match).
function Find-ExistingSsoApps([string]$EntityId) {
    $matched = @()
    $uri = "/applications?`$select=id,appId,displayName,identifierUris,web&`$top=999"
    while ($uri) {
        $resp = Graph-Get $uri
        foreach ($a in $resp.value) {
            $ids  = @($a.identifierUris)
            $urls = ($ids + @($a.web.redirectUris)) -join ' '
            $holdsEntityId = [bool]($ids | Where-Object { $_ -ieq $EntityId })
            $fingerprint = ($urls -match '(?i)amazoncognito\.com/saml2/idpresponse') -or ($urls -match '(?i)urn:amazon:cognito')
            if ($holdsEntityId -or $fingerprint) { $matched += $a }
        }
        $next = $resp.'@odata.nextLink'
        $uri = if ($next) { $next -replace '^https://graph\.microsoft\.com/v1\.0', '' } else { $null }
    }
    return , $matched
}

# New service principals / applications replicate lazily and can 404 for a few seconds after
# instantiate. Poll a GET until it succeeds (or give up). Returns $true if the object became
# readable in time.
function Wait-ForGraphObject([string]$path, [int]$tries = 8, [int]$delaySeconds = 3) {
    for ($i = 1; $i -le $tries; $i++) {
        try { Graph-Get $path | Out-Null; return $true } catch { }
        Start-Sleep -Seconds $delaySeconds
    }
    return $false
}

# UPN sync: remap the SAML emailaddress claim to user.userPrincipalName via a claims-mapping policy,
# and assign it to the enterprise app. Restricted-claim remapping like this requires the app's own
# signing key — which is why the SSO phase creates the cert BEFORE calling this.
function Set-UpnEmailClaim([string]$SpId) {
    $definition = '{"ClaimsMappingPolicy":{"Version":1,"IncludeBasicClaimSet":"true","ClaimsSchema":[{"Source":"user","ID":"userprincipalname","SamlClaimType":"http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress"}]}}'
    $policy = Graph-Post '/policies/claimsMappingPolicies' @{
        definition            = @($definition)
        displayName           = 'PII-Protect SSO email=UPN'
        isOrganizationDefault = $false
    }
    Graph-Post "/servicePrincipals/$SpId/claimsMappingPolicies/`$ref" @{
        '@odata.id' = "$GraphBase/policies/claimsMappingPolicies/$($policy.id)"
    } | Out-Null
    Write-Host '  email claim: mapped emailaddress -> user.userPrincipalName (UPN claims policy).' -ForegroundColor Green
}

function Invoke-SsoPhase {
    Write-Section 'Single Sign-On: PII-Protect SSO (non-gallery SAML enterprise app)'
    Write-Host '  In the portal: User Management > Single Sign On > click "Microsoft".' -ForegroundColor Cyan
    Write-Host '  A popup shows a Redirect URL and an Application ID URI. Leave it open.' -ForegroundColor Cyan

    $redirect = Read-Required 'Paste the Redirect URL from the portal popup' $SsoRedirectUri
    $appIdUri = Read-Required 'Paste the Application ID URI from the portal popup' $SsoAppIdUri
    if (-not $redirect -or -not $appIdUri) {
        Write-Warning 'Missing Redirect URL / Application ID URI; skipping SSO. Re-run with -SsoRedirectUri and -SsoAppIdUri, or interactively.'
        return
    }

    # 1. Existing-app check. A SAML Entity ID must be unique tenant-wide, so any app already holding
    #    this Entity ID — or one carrying BSN's Cognito fingerprint (e.g. a legacy PII-Protect app
    #    registration) — blocks a fresh create. Offer to replace it (guarded delete-then-recreate).
    $conflicts = @(Find-ExistingSsoApps $appIdUri)
    if ($conflicts.Count -gt 0) {
        Write-Host ''
        Write-Host "  Found $($conflicts.Count) existing app(s) that conflict with this SSO Entity ID:" -ForegroundColor Yellow
        foreach ($c in $conflicts) {
            Write-Host "    - $($c.displayName) (appId $($c.appId)); App ID URI $((@($c.identifierUris) -join ', '))" -ForegroundColor DarkGray
        }
        if ($WhatIfPreference) {
            Write-Host '  (-WhatIf: would offer to delete-then-recreate the above.)' -ForegroundColor DarkGray
        } else {
            $replace = [bool]$ReplaceSso
            if (-not $replace) {
                if ($NonInteractive) {
                    Write-Warning 'An app already holds this SSO Entity ID. Not deleting it in -NonInteractive without -ReplaceSso; skipping SSO.'
                    Write-Host '  Re-run with -ReplaceSso to delete-then-recreate, or remove the old app by hand.' -ForegroundColor Yellow
                    return
                }
                $ans = Read-Host '  Delete the above and recreate as a SAML enterprise app? (y/N)'
                $replace = ($ans -match '^\s*(y|yes)\s*$')
            }
            if (-not $replace) {
                Write-Warning 'Left the existing SSO app in place; skipping SSO (the Entity ID cannot be reused while it exists).'
                return
            }
            foreach ($c in $conflicts) {
                if ($PSCmdlet.ShouldProcess("$($c.displayName) (appId $($c.appId))", 'Delete existing SSO app')) {
                    Graph-Delete "/applications/$($c.id)"
                    Write-Host "  deleted old app '$($c.displayName)'." -ForegroundColor Green
                }
            }
        }
    }

    if (-not $PSCmdlet.ShouldProcess('PII-Protect SSO', 'Create + configure non-gallery SAML enterprise app')) { return }

    # 2. Instantiate the non-gallery ("custom") application template — creates the application
    #    object AND its enterprise app (service principal) together, in the local tenant.
    $inst = Graph-Post "/applicationTemplates/$SamlAppTemplateId/instantiate" @{ displayName = 'PII-Protect SSO' }
    $appObjId = $inst.application.id
    $appId    = $inst.application.appId
    $spId     = $inst.servicePrincipal.id
    Write-Host "  created enterprise app 'PII-Protect SSO' (client id $appId)." -ForegroundColor Green

    # 3. The new SP/app replicate lazily — wait until each is readable before patching it.
    if (-not (Wait-ForGraphObject "/servicePrincipals/$spId")) {
        Write-Warning "Service principal $spId not readable yet; the steps below may need a re-run (-Verify to confirm)."
    }

    # 4. Switch the enterprise app to SAML SSO and turn off the user-assignment requirement.
    Graph-Patch "/servicePrincipals/$spId" @{ preferredSingleSignOnMode = 'saml'; appRoleAssignmentRequired = $false }
    Write-Host '  set SSO mode = SAML; assignment required = No.' -ForegroundColor Green

    # 5. Entity ID (Application ID URI) + Reply URL from the portal popup.
    [void](Wait-ForGraphObject "/applications/$appObjId")
    Graph-Patch "/applications/$appObjId" @{ identifierUris = @($appIdUri); web = @{ redirectUris = @($redirect) } }
    Write-Host '  set Entity ID (Application ID URI) and Reply URL.' -ForegroundColor Green

    # 6. Per-app SAML signing certificate (independent rotation; also unlocks the restricted-claim
    #    remap for UPN sync). 3y is Azure's max for an auto-generated cert. Do this BEFORE claims.
    $end = (Get-Date).ToUniversalTime().AddYears($SsoCertYears).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $cert = Graph-Post "/servicePrincipals/$spId/addTokenSigningCertificate" @{ displayName = 'CN=PII-Protect SSO'; endDateTime = $end }
    Graph-Patch "/servicePrincipals/$spId" @{ preferredTokenSigningKeyThumbprint = $cert.thumbprint }
    Write-Host "  added SAML signing certificate (thumbprint $($cert.thumbprint); expires in ${SsoCertYears}y)." -ForegroundColor Green

    # 7. Email claim. Email (default): Azure already emits emailaddress from user.mail — nothing to
    #    do. UPN: add the claims-mapping policy that remaps emailaddress -> user.userPrincipalName.
    if ($SsoSyncType -eq 'UPN') { Set-UpnEmailClaim $spId }
    else { Write-Host '  email claim: using Azure default (emailaddress from user.mail).' -ForegroundColor DarkGray }

    # 8. App-specific federation metadata URL — deterministic; guide the operator to paste it.
    $tenant = (Get-MgContext).TenantId
    $fedMeta = "https://login.microsoftonline.com/$tenant/federationmetadata/2007-06/federationmetadata.xml?appid=$appId"
    Write-Host ''
    Write-Host '  Back in the PII Protect portal SSO popup:' -ForegroundColor Cyan
    Write-Host "    Metadata URL: $fedMeta" -ForegroundColor White
    Write-Host '    Paste it into the "Metadata URL" field and click Connect.' -ForegroundColor Cyan
    Write-Host '    If the portal shows a "Skip Identity Provider Logout" toggle, set it per BSN guidance.' -ForegroundColor Cyan
    if (-not $NonInteractive) { [void](Read-Host "`n  Press Enter once you've clicked Connect in the portal") }
    $script:ssoConfigured = $true
}

# --- Phase: renew the SSO SAML signing certificate (-RenewSsoCert) ----------------------
# On-demand rotation, regardless of how much life the current cert has left — for a compromise, a
# policy-driven rotation, or just getting ahead of the expiry -Verify warns about.
#
# Deliberately NOT part of -Fix. Rotation is disruptive, not additive: BSN caches the signing cert
# from the metadata you paste into their portal, so the moment Entra starts signing with a new one,
# sign-in breaks until the metadata is re-pasted. -Fix only ever does things that cannot break a
# working setup; this can, so it's an explicit, confirmed, one-job mode.
function Invoke-RenewSsoCertPhase {
    Write-Section 'Renew the SSO SAML signing certificate'

    $apps = @((Find-BsnSsoApps).Apps)
    if (-not $apps.Count) {
        Write-Warning 'No BSN SSO app found in this tenant (nothing points at the PII Protect portal or carries the Cognito SAML fingerprint).'
        Write-Host '  Run -Verify to see what is there, or a normal run to set SSO up.' -ForegroundColor Yellow
        return
    }
    if ($apps.Count -gt 1) {
        Write-Warning "$($apps.Count) apps match the BSN SSO fingerprint — renewing only '$($apps[0].displayName)'. Run -Verify and clean up the duplicates first if that's the wrong one."
    }
    $app = $apps[0]
    Write-Host "  SSO app: $($app.displayName) (client id $($app.appId))" -ForegroundColor Cyan

    $sp = @((Graph-Get "/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$select=id,displayName,preferredSingleSignOnMode,preferredTokenSigningKeyThumbprint,keyCredentials").value) | Select-Object -First 1
    if (-not $sp) {
        Write-Warning 'That app has no service principal (enterprise app) in this tenant, so there is no SAML signing certificate to renew.'
        return
    }

    # Gate on whether a per-app certificate ACTUALLY exists, not on the app's shape. Renewing means
    # replacing a certificate this app already owns; if it owns none, Entra is signing with the
    # tenant's default key (Microsoft's, not ours) and there is nothing here to renew. Adding a first
    # certificate to such an app would not be a renewal at all — it would switch how the app signs,
    # and break sign-in until the portal is re-pasted. That's a migration decision, not this switch's
    # job. (Shape is reported for context, but it does not decide.)
    # Wrap the PIPELINE, not just the input: with exactly one match Where-Object emits a bare
    # object, and Graph hands these back as hashtables — whose .Count is the number of KEYS.
    # Without the outer @(), one certificate reports as "8 signing certificate(s)".
    $signCerts = @(@($sp.keyCredentials) | Where-Object { $_.usage -eq 'Sign' -and $_.type -eq 'AsymmetricX509Cert' })
    $shape = if ($sp.preferredSingleSignOnMode -eq 'saml') { 'SAML enterprise app' } else { "app registration (preferredSingleSignOnMode = '$(if ($sp.preferredSingleSignOnMode) { $sp.preferredSingleSignOnMode } else { 'not set' })')" }
    Write-Host "  SSO shape: $shape" -ForegroundColor Cyan

    if (-not $signCerts.Count -and -not $sp.preferredTokenSigningKeyThumbprint) {
        Write-Host ''
        Write-Wrapped 'Nothing to renew: this app has no signing certificate of its own, so Entra signs it with your tenant''s default key — which Microsoft manages, not this script.' '  ' 'Yellow'
        Write-Host ''
        Write-Wrapped 'Adding a certificate here would not be a renewal — it would change how the app signs and break sign-in until you re-paste metadata into the BSN portal. If that is what you want, it belongs to a migration (a normal run with -ReplaceSso), not to this switch.' '  ' 'DarkGray'
        return
    }

    # Show what's there now, so the operator sees what they're replacing.
    $current = if ($sp.preferredTokenSigningKeyThumbprint) {
        @($signCerts | Where-Object { (Get-KeyThumbprint $_) -eq "$($sp.preferredTokenSigningKeyThumbprint)".ToLower() }) | Select-Object -First 1
    }
    if ($current -and $current.endDateTime) {
        $end  = ([datetime]$current.endDateTime).ToUniversalTime()
        $days = [int][Math]::Floor(($end - (Get-Date).ToUniversalTime()).TotalDays)
        $state = if ($days -lt 0) { "EXPIRED $([Math]::Abs($days)) days ago" } else { "$days days left" }
        Write-Host "  Current certificate: expires $($end.ToString('yyyy-MM-dd')) ($state)" -ForegroundColor Cyan
    } else {
        Write-Host '  Current certificate: none set as preferred.' -ForegroundColor Yellow
    }
    Write-Host "  New certificate:     $SsoCertYears year$(if ($SsoCertYears -ne 1) { 's' }) from today (-SsoCertYears to change)" -ForegroundColor Cyan

    Write-Host ''
    Write-Wrapped 'HEADS UP: single sign-on for this client will STOP working the moment the new certificate takes over, and stay broken until you paste the metadata URL below back into the BSN portal. Do this when you can finish the portal step straight away — not at 5pm on a Friday.' '  ' 'Yellow'
    Write-Host ''

    if (-not $NonInteractive) {
        if ((Read-Host "  Renew the signing certificate for '$($app.displayName)'? (y/N)") -notmatch '^\s*(y|yes)\s*$') {
            Write-Host '  Skipped — nothing changed.' -ForegroundColor DarkGray
            return
        }
    }
    if (-not $PSCmdlet.ShouldProcess($app.displayName, "Add a new SAML signing certificate ($SsoCertYears years) and make it the preferred signing key")) { return }

    $endDate = (Get-Date).ToUniversalTime().AddYears($SsoCertYears).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $cert = Graph-Post "/servicePrincipals/$($sp.id)/addTokenSigningCertificate" @{ displayName = 'CN=PII-Protect SSO'; endDateTime = $endDate }
    Graph-Patch "/servicePrincipals/$($sp.id)" @{ preferredTokenSigningKeyThumbprint = $cert.thumbprint }
    Write-Host ''
    Write-Host "  Added and selected a new signing certificate (thumbprint $($cert.thumbprint))." -ForegroundColor Green
    Write-Host "  It expires $((Get-Date).ToUniversalTime().AddYears($SsoCertYears).ToString('yyyy-MM-dd'))." -ForegroundColor Green
    Write-Host '  The old certificate is left on the app; Entra now signs with the new one.' -ForegroundColor DarkGray

    # Same URL as always (it keys on appid) — but its CONTENTS now carry the new certificate, which
    # is exactly why the portal has to be pointed at it again.
    $tenant = (Get-MgContext).TenantId
    $fedMeta = "https://login.microsoftonline.com/$tenant/federationmetadata/2007-06/federationmetadata.xml?appid=$($app.appId)"
    Write-Section 'MANUAL (PII Protect portal): finish the renewal'
    Write-Host '  Sign-in is broken for this client until you do this.' -ForegroundColor Yellow
    Write-Host "  Metadata URL: $fedMeta" -ForegroundColor White
    Write-Host '   1. Log in to the BSN portal as a Partner Administrator.' -ForegroundColor Gray
    Write-Host '   2. Manage Clients > the client > User Management > Single Sign-On.' -ForegroundColor Gray
    Write-Host '   3. Paste the Metadata URL above and save/Connect, so BSN picks up the new certificate.' -ForegroundColor Gray
    Write-Host '   4. Test: log out of BSN and log back in with a client user account.' -ForegroundColor Gray
    if (-not $NonInteractive) { [void](Read-Host "`n  Press Enter once the portal has the new metadata") }
    Write-Host '  Re-run -Verify to confirm the new certificate is the preferred key.' -ForegroundColor Cyan
}

# --- Phase: Catch Phish Outlook add-in (guided; Microsoft 365 admin center) -------------
# Uses the "Get apps" flow on the Integrated apps page — Microsoft's recommended path for deploying
# a marketplace add-in. Note this is NOT the "Add-ins" link on the same page, which is the older
# add-in page: a separate surface with its own store. The distinction matters: an Integrated-apps
# deployment is invisible to Exchange's Get-App (see Test-BsnCatchPhish), and only the
# Integrated-apps flow ends in the admin-consent step that leaves an enterprise app behind for
# verify to find.
#
# SOURCING: steps below mirror BSN's official article "How to Deploy the Catch Phish Outlook Add-In
# via Microsoft 365 Admin Center" (BSN partner portal, retrieved 2026-07-15). That article is the
# authority here — NOT Microsoft's generic add-in docs, and NOT a reseller's rebranded copy. Keep
# these in step with it. The audience advice in -Notes is OURS: BSN's article only says "select the
# users to deploy to".
#
# It's manual because Office centralized deployment has no Graph API, and the one PowerShell module
# (O365CentralizedAddInDeployment) is Windows-only and documented as Basic-auth/no-MFA, which
# mandatory admin MFA rules out. The integrated apps portal is Microsoft's recommended path anyway.
function Invoke-CatchPhishStep {
    Invoke-ManualStep -Portal 'Microsoft 365 admin center' -Title 'Deploy the Catch Phish Outlook add-in' `
        -Links ([ordered]@{
            'Integrated apps' = $IntegratedAppsUrl
            'Catch Phish'     = $CatchPhishMarketUrl
        }) `
        -Steps @(
            'Log in to the Microsoft 365 admin center as a Global Admin of the CLIENT tenant.',
            'Settings > Integrated apps, then click "Get apps".',
            '   (NOT the "Add-ins" link on that page — a different, older surface.)',
            'Search for "Catch Phish" and click "Get it now".',
            "   (Confirm it is the Breach Secure Now listing — AssetId $CatchPhishAssetId.)",
            'Review the Microsoft permission and click "Get it now" again.',
            'Select the users to deploy to (see the audience note below) and click Next.',
            'Review the App permissions and click "Accept permissions".',
            'Log in as a Global Administrator, review the App Permissions, and click Accept.',
            'Click Next, then "Finish deployment".'
        ) `
        -Notes @(
            'Use "Get apps", not the "Add-ins" link. Both live on this page and both deploy Catch',
            '  Phish, but they are separate surfaces with separate stores — "Get apps" is the flow',
            '  BSN''s article gives, and the one -Verify can confirm afterwards.',
            'Audience — OUR guidance, not BSN''s (their article just says "select the users"):',
            '  deploy to the entire organization. Centralized deployment does NOT support',
            '  mail-disabled security groups, and this script creates BSN-Employees / BSN-Managers',
            '  mail-disabled, so you cannot pick them here. Everyone is the right audience anyway —',
            '  Catch Phish is a report-phishing button for all staff. Narrower scoping needs a',
            '  mail-enabled group or distribution list (top-level only; nested groups are NOT',
            '  assigned).',
            'GCC High / GCC Low tenants: BSN publishes a SEPARATE article — these steps do not apply.',
            'BSN notes it can take up to 72 hours for the add-in to appear in Outlook. -Verify can',
            '  confirm the deployment well before that, but not instantly.'
        )
}

# --- Verification (read-only) ----------------------------------------------------------
# Every check funnels through here, so this is also where results are collected for the plain-English
# summary at the end of the run. $plain says what a FAIL/WARN MEANS for the client in one sentence
# ("BSN's emails may land in junk"), as opposed to $detail, which says what the tool observed
# ("missing"). Authored at the check itself, because that's the only place the meaning is known.
# Omit $plain and the summary falls back to the technical text — degraded, not broken.
$script:checkResults = @()
function Write-Check([string]$status, [string]$item, [string]$detail, [string]$plain) {
    $label = @{ OK = '  [ OK ]'; FAIL = '  [FAIL]'; WARN = '  [WARN]'; INFO = '  [ -- ]' }[$status]
    $color = @{ OK = 'Green'; FAIL = 'Red'; WARN = 'Yellow'; INFO = 'Cyan' }[$status]
    Write-Host ("{0} {1}{2}" -f $label, $item, $(if ($detail) { " — $detail" })) -ForegroundColor $color
    $script:checkResults += [pscustomobject]@{
        Status = $status; Item = $item; Detail = $detail; Plain = $plain
        Fixed = $false; Fixable = $false
    }
}

# How many members a group actually has. Uses $count rather than paging the whole membership, which
# needs the eventual-consistency header that Graph-Get doesn't send — hence the direct call. Returns
# $null on failure so callers just omit the number instead of failing the check.
function Get-GroupMemberCount([string]$id) {
    try {
        $r = Invoke-MgGraphRequest -Method GET -OutputType PSObject `
                -Uri "$GraphBase/groups/$id/members?`$top=1&`$count=true" `
                -Headers @{ ConsistencyLevel = 'eventual' }
        return [int]$r.'@odata.count'
    } catch { return $null }
}

# A deep link straight to where you'd edit this group in the Entra admin center — the rule editor for
# a dynamic group, the members list otherwise. Saves hunting through Groups > All groups > search >
# click > left-nav every time.
#
# Microsoft publishes NO deep-link reference for these blades, so the format is verified by clicking,
# not by documentation:
#  * 'DynamicGroupMembershipRule' is CONFIRMED working (2026-07-15) — note it is NOT
#    'DynamicMembershipRule', which is the obvious guess and is wrong.
#  * 'Members' for assigned groups is NOT yet confirmed.
#  * No #@<tenant> prefix: the confirmed URL has none, and an unverified tenant hint that breaks the
#    link would be worse than no hint. That means the portal opens in whichever directory the browser
#    used last — so callers must tell the operator to check they're in the right tenant, which for an
#    MSP is the difference between editing the client's group and their own.
# If a link stops landing, the printed navigation fallback is the fix.
function Get-EntraGroupUrl($group) {
    $blade = if (@($group.groupTypes) -contains 'DynamicMembership') { 'DynamicGroupMembershipRule' } else { 'Members' }
    return "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/$blade/groupId/$($group.id)"
}

# What a group actually contains, and — if dynamic — the rule that decides it. "The group exists"
# says nothing about whether it works: an empty BSN-Employees enrols nobody, and a Paused rule
# freezes membership silently. Both look identical to a name check.
function Write-GroupDetail($group, [switch]$Core) {
    $dynamic = (@($group.groupTypes) -contains 'DynamicMembership')
    $count = Get-GroupMemberCount $group.id

    if ($null -eq $count) {
        Write-Host '         members: could not read' -ForegroundColor DarkGray
    } elseif ($count -eq 0 -and $Core) {
        # An empty required group is the quiet version of "not set up at all".
        Write-Check WARN "  $($group.displayName) membership" 'the group exists but has NO members' "Nobody is actually enrolled through $($group.displayName) — the group is there but empty, so it is doing nothing."
    } else {
        Write-Host "         members: $count" -ForegroundColor DarkGray
    }

    if (-not $dynamic) { return }

    # Paused means Entra has stopped evaluating the rule: membership is frozen wherever it was.
    if ($group.membershipRuleProcessingState -and $group.membershipRuleProcessingState -ne 'On') {
        Write-Check WARN "  $($group.displayName) rule" "processing is $($group.membershipRuleProcessingState) — membership is frozen and will not update" "$($group.displayName) has stopped updating itself: new staff will not be added and leavers will not be removed until it is switched back on."
    }
    if ($group.membershipRule) {
        Write-Host '         rule:' -ForegroundColor DarkGray
        Write-Wrapped $group.membershipRule '           ' 'DarkGray'
    }
    Write-EntraGroupLink $group $(if ($dynamic) { 'edit rule' } else { 'edit in Entra' })
}

# The Entra link for a group, printed where the operator is already looking at that group. Separate
# from Write-GroupDetail so callers can show it only where editing is plausible, rather than hanging
# a URL off every line.
function Write-EntraGroupLink($group, [string]$what) {
    Write-Host "         $($what): $(Get-EntraGroupUrl $group)" -ForegroundColor DarkCyan
}

function Test-BsnGroups {
    function Find-Group([string]$name) {
        $safe = $name -replace "'", "''"
        return @((Graph-Get "/groups?`$filter=displayName eq '$safe'&`$select=id,displayName,groupTypes,membershipRule,membershipRuleProcessingState").value) | Select-Object -First 1
    }

    # Required — BSN can't enrol without these, so missing is a real failure. -Fix only CREATES a
    # missing group (safe/additive); it never converts an existing assigned BSN-Employees to
    # dynamic, since that would drop manually-added members.
    $emp = Find-Group 'BSN-Employees'
    if ($emp) {
        $dyn = (@($emp.groupTypes) -contains 'DynamicMembership')
        if ($dyn) {
            Write-Check OK 'Group BSN-Employees' 'dynamic (auto-enrol)'
            Write-GroupDetail $emp -Core
        } else {
            # Assigned. Whether that's the best available depends on licensing — dynamic membership
            # (auto-enrol) needs Entra ID P1/P2 — so say which case this is: a missed opportunity
            # (tenant HAS P1) vs. genuinely the best it can do (no P1).
            $p1 = Test-EntraP1
            if ($p1 -eq $true) {
                Write-Check WARN 'Group BSN-Employees' 'assigned, but this tenant HAS Entra ID P1 — it could auto-enrol via dynamic membership. Re-run without -NoDynamicEmployees to convert (conversion drops any manually-added members).' 'New staff will NOT be enrolled in training automatically — someone has to add each person by hand. This tenant is licensed for automatic enrolment, so that is avoidable.'
            } elseif ($p1 -eq $false) {
                Write-Check OK 'Group BSN-Employees' 'assigned (Entra ID P1 not present, so dynamic auto-enrol is unavailable; add users by hand)'
            } else {
                Write-Check OK 'Group BSN-Employees' 'assigned (add users by hand; could not confirm Entra ID P1)'
            }
        }
    } else {
        Write-Check FAIL 'Group BSN-Employees' 'not found' 'Nobody at this client will be enrolled in security training. The BSN-Employees group is missing, and that is the group BSN pulls staff from.'
        Invoke-Fix 'create BSN-Employees' {
            $useDynamic = -not $NoDynamicEmployees
            if ($useDynamic -and (Test-EntraP1) -eq $false) { $useDynamic = $false }
            $rule = if ($useDynamic) { Build-EmployeeRule $EnrollServicePlanId $ExcludeEmail }
            Ensure-BsnGroup -Name 'BSN-Employees' -Description 'PII/PHI Protect Standard Users' -Dynamic:$useDynamic -Rule $rule | Out-Null
        }
    }

    $mgr = Find-Group 'BSN-Managers'
    if ($mgr) {
        Write-Check OK 'Group BSN-Managers' "exists$(if (@($mgr.groupTypes) -contains 'DynamicMembership') { ' (dynamic)' })"
        Write-GroupDetail $mgr
    }
    else {
        Write-Check FAIL 'Group BSN-Managers' 'not found' 'Managers will not be able to see their team''s training progress. The BSN-Managers group is missing.'
        Invoke-Fix 'create BSN-Managers' { Ensure-BsnGroup -Name 'BSN-Managers' -Description 'PII/PHI Protect Manager Role' | Out-Null }
    }

    # Optional admin groups — set up only when wanted, so absence is informational, not a failure
    # (and not something -Fix creates on its own).
    foreach ($name in @('BSN-ManagerAdmins', 'BSN-PartnerAdmins')) {
        $g = Find-Group $name
        if ($g) {
            Write-Check OK "Group $name" "exists (optional)$(if (@($g.groupTypes) -contains 'DynamicMembership') { '; dynamic' })"
            Write-GroupDetail $g
        }
        else { Write-Check INFO "Group $name" 'not set up (optional)' }
    }

    # Tag groups you explicitly asked for — missing is a failure.
    # Say this once, not per group: the links above are undocumented portal URLs, so give the manual
    # route too. They carry a #@<tenant> hint, which for an MSP is the difference between editing the
    # client's group and editing your own.
    Write-Host '         (Entra links do NOT pin the tenant — check the portal opened the client you' -ForegroundColor DarkGray
    Write-Host '          just verified, not your own. If a link does not land: entra.microsoft.com >' -ForegroundColor DarkGray
    Write-Host '          Groups > All groups > the group > Dynamic membership rules.)' -ForegroundColor DarkGray

    # Tag groups are BSN-TAG-<tagname> — the casing BSN's docs use predominantly, though BSN-Tag-
    # appears occasionally in both their docs and real tenants. Graph's displayName filter is
    # case-insensitive, so either casing is found; report the casing actually in the tenant rather
    # than the form we asked for, so drift stays visible instead of being silently normalised.
    foreach ($tag in $TagGroup) {
        $name = "BSN-TAG-$tag"
        $tg = Find-Group $name
        if ($tg) {
            $actual = if ($tg.displayName -cne $name) { " — actually named '$($tg.displayName)'" }
            Write-Check OK "Group $name" "exists$(if (@($tg.groupTypes) -contains 'DynamicMembership') { '; dynamic' })$actual"
            Write-GroupDetail $tg
        }
        else {
            Write-Check FAIL "Group $name" 'not found'
            Invoke-Fix "create $name" { Ensure-BsnGroup -Name $name -Description "PII/PHI Protect Tag: $tag" | Out-Null }
        }
    }
}

# Identify the SSO app by its BSN portal footprint (reply URL / Entity ID), NOT by name or appId — a
# manual setup may have named it anything, and appIds can change. BSN SSO federates to AWS Cognito,
# so the reply URL (…amazoncognito.com/saml2/idpresponse) and Entity ID (urn:amazon:cognito:…) are
# the reliable, name/tenant/region-agnostic fingerprint, and they hold for BOTH the new SAML
# enterprise app and the legacy app registration.
#
# Shared by -Verify and -RenewSsoCert: both must agree on which app is "the SSO app", so this lives
# in one place rather than being reimplemented per caller.
function Find-BsnSsoApps {
    $select = 'id,appId,displayName,web,api,identifierUris,optionalClaims,requiredResourceAccess'
    $apps = @()            # confident: portal domain or Cognito SAML fingerprint
    $samlCandidates = @()  # possible: legacy SAML 'email' optional claim but no portal match
    $uri = "/applications?`$select=$select&`$top=999"
    while ($uri) {
        $resp = Graph-Get $uri
        foreach ($a in $resp.value) {
            $urls = (@($a.web.redirectUris) + @($a.identifierUris)) -join ' '
            $hasSamlClaim = [bool](@($a.optionalClaims.saml2Token) | Where-Object { $_.name -eq 'email' })
            $cognito = ($urls -match '(?i)amazoncognito\.com/saml2/idpresponse') -or ($urls -match '(?i)urn:amazon:cognito')
            $isBsnSso = ($urls -match '(?i)pii-protect|breachsecurenow') -or $cognito
            if ($isBsnSso) { $apps += $a }
            elseif ($hasSamlClaim) { $samlCandidates += $a }
        }
        $next = $resp.'@odata.nextLink'
        $uri = if ($next) { $next -replace '^https://graph\.microsoft\.com/v1\.0', '' } else { $null }
    }
    return [pscustomobject]@{ Apps = $apps; SamlCandidates = $samlCandidates }
}

# The SAML signing certificate: which key Entra actually signs with, and how long it has left.
#
# Runs for BOTH SSO shapes on purpose. It used to run only for the SAML enterprise app, on the
# assumption that the older app-registration shape always signs with the tenant's Microsoft-managed
# default key and so has nothing worth reporting. That assumption was never verified — and gating the
# check on it meant a legacy app WITH its own certificate would sail past silently until the day it
# expired and sign-in broke. So: look, then report what is actually there.
function Test-SsoSigningCert($sp, [string]$Shape) {
    $thumb = $sp.preferredTokenSigningKeyThumbprint
    # Wrap the PIPELINE, not just the input: with exactly one match Where-Object emits a bare
    # object, and Graph hands these back as hashtables — whose .Count is the number of KEYS.
    # Without the outer @(), one certificate reports as "8 signing certificate(s)".
    $signCerts = @(@($sp.keyCredentials) | Where-Object { $_.usage -eq 'Sign' -and $_.type -eq 'AsymmetricX509Cert' })
    $preferred = if ($thumb) { @($signCerts | Where-Object { (Get-KeyThumbprint $_) -eq "$thumb".ToLower() }) | Select-Object -First 1 }

    # No certificate of its own: Entra signs with the tenant default, which Microsoft owns and rolls
    # over. Nothing for us to renew — but say so plainly rather than silently reporting nothing.
    if (-not $signCerts.Count -and -not $thumb) {
        if ($Shape -eq 'enterprise') {
            Write-Check FAIL '  SAML signing cert' 'no per-app signing certificate configured' 'Single sign-on has no certificate of its own to sign with, so signing in to BSN with a Microsoft account may not work.'
        } else {
            Write-Check INFO '  SAML signing cert' 'none on this app — signs with the tenant default key (Microsoft-managed; nothing here expires for you to renew)'
        }
        return
    }

    if ($preferred -and $preferred.endDateTime) {
        $end  = ([datetime]$preferred.endDateTime).ToUniversalTime()
        $exp  = $end.ToString('yyyy-MM-dd')
        $days = [int][Math]::Floor(($end - (Get-Date).ToUniversalTime()).TotalDays)
        if ($days -lt 0) {
            Write-Check FAIL '  SAML signing cert' "preferred key's certificate EXPIRED on $exp — rotate it" 'Single sign-on is broken: the certificate that signs the sign-in has expired, so users cannot log in to BSN with their Microsoft account. Fixing it also means pasting a new metadata URL into the BSN portal.'
        } elseif ($days -le $script:CertWarnDays) {
            Write-Check WARN '  SAML signing cert' "expires in $days day$(if ($days -ne 1) { 's' }) ($exp) — renew with -RenewSsoCert" "Single sign-on for this client will STOP working on $exp — users will not be able to log in to BSN with their Microsoft account. Renewing it also means pasting a new metadata URL into the BSN portal, so book the time before then."
        } else {
            Write-Check OK '  SAML signing cert' "preferred key set; valid to $exp ($days days)"
        }
        return
    }

    if ($thumb) {
        Write-Check FAIL '  SAML signing cert' "preferred key $thumb does not match any signing certificate on the app" 'Single sign-on is misconfigured: the app is set to sign with a certificate that is not there. Users may not be able to log in to BSN with their Microsoft account.'
        return
    }

    # Certificates present but none marked preferred. On a legacy app this may be normal (Entra can
    # still use the tenant default), so report the facts and let the operator judge.
    $ends = @($signCerts | Where-Object { $_.endDateTime } | ForEach-Object { ([datetime]$_.endDateTime).ToUniversalTime() } | Sort-Object -Descending)
    $latest = if ($ends.Count) { $ends[0].ToString('yyyy-MM-dd') } else { 'unknown' }
    Write-Check WARN '  SAML signing cert' "$($signCerts.Count) signing certificate(s) on the app but none set as preferred (newest expires $latest)"
}

function Test-BsnSso {
    if ($SkipSso) { Write-Check WARN 'SSO app' 'skipped (-SkipSso)'; return }
    $found = Find-BsnSsoApps
    $apps = $found.Apps
    $samlCandidates = $found.SamlCandidates
    if ($apps.Count -eq 0) {
        if ($samlCandidates.Count) {
            Write-Check WARN 'SSO app' "no portal-URL match; $($samlCandidates.Count) app(s) carry a SAML email claim — possible BSN SSO, confirm the reply URL:"
            foreach ($c in $samlCandidates) {
                Write-Host "         - $($c.displayName): reply $((@($c.web.redirectUris) -join ', '))  |  App ID URI $((@($c.identifierUris) -join ', '))" -ForegroundColor DarkGray
            }
            Write-Host '         If one is your BSN SSO, tell me its reply-URL domain and I can teach the check to match it.' -ForegroundColor DarkGray
        } else {
            Write-Check FAIL 'SSO app' 'no app registration points at the PII Protect portal or carries the BSN SAML fingerprint' 'Users cannot sign in to the BSN portal with their Microsoft account — single sign-on is not set up for this client.'
        }
        return
    }
    $app = $apps[0]
    $rename = if ($app.displayName -ne 'PII-Protect SSO') { " — named '$($app.displayName)'; consider renaming to 'PII-Protect SSO'" }
    Write-Check OK 'SSO app' "$($app.displayName) (client id $($app.appId))$rename"
    if ($apps.Count -gt 1) { Write-Check WARN '  duplicate SSO apps' "$($apps.Count) registrations point at the portal — review for duplicates" }

    # Reply URL + Entity ID live on the application object (both SSO shapes need them).
    $redir = @($app.web.redirectUris).Count -gt 0
    Write-Check $(if ($redir) { 'OK' } else { 'FAIL' }) '  reply URL' $(if ($redir) { $app.web.redirectUris[0] } else { 'missing' })
    $idu = @($app.identifierUris).Count -gt 0
    Write-Check $(if ($idu) { 'OK' } else { 'FAIL' }) '  Application ID URI (Entity ID)' $(if ($idu) { $app.identifierUris[0] } else { 'missing' })

    # BSN's current SAML config lives on the enterprise app (service principal): the SSO mode, the
    # assignment requirement, and the per-app signing certificate.
    $sp = @((Graph-Get "/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$select=id,preferredSingleSignOnMode,appRoleAssignmentRequired,preferredTokenSigningKeyThumbprint,keyCredentials").value) | Select-Object -First 1
    Write-Check $(if ($sp) { 'OK' } else { 'FAIL' }) '  enterprise app (SP)' $(if ($sp) { 'present' } else { 'missing — SAML SSO needs an enterprise app' })

    # Safe -Fix: rename a misnamed app to the canonical name. Cosmetic — SAML keys on Entity ID /
    # reply URL / signing cert, not display name — so this can't affect a working federation.
    if ($app.displayName -ne 'PII-Protect SSO') {
        Invoke-Fix "rename SSO app '$($app.displayName)' -> 'PII-Protect SSO'" {
            Graph-Patch "/applications/$($app.id)" @{ displayName = 'PII-Protect SSO' }
            if ($sp) { Graph-Patch "/servicePrincipals/$($sp.id)" @{ displayName = 'PII-Protect SSO' } }
            Write-Host "          renamed to 'PII-Protect SSO'." -ForegroundColor Green
        }
    }

    # Federation shape drives which checks matter. NEW method = SAML enterprise app
    # (SP.preferredSingleSignOnMode=saml); LEGACY method = an app registration carrying a SAML
    # 'email' optional claim; anything else is treated as OIDC (checking OIDC config on a SAML app
    # would only produce false failures, and vice-versa).
    $spSaml = [bool]($sp -and $sp.preferredSingleSignOnMode -eq 'saml')
    $legacySaml = ([bool](@($app.optionalClaims.saml2Token) | Where-Object { $_.name -eq 'email' })) -or
                  ([bool](@($app.identifierUris) -match '(?i)urn:amazon:cognito')) -or
                  ([bool](@($app.web.redirectUris) -match '(?i)saml2/idpresponse'))
    # Labels confirmed against BSN's own docs (2026-07-15): their current setup article specifies a
    # non-gallery SAML enterprise app, while their 2-year-old SSO troubleshooting article still talks
    # about "the app registration" — so BSN did move, and "older" is accurate, not editorialising.
    $shape = if ($spSaml) { 'SAML enterprise app (BSN current method)' } elseif ($legacySaml) { 'SAML app registration (BSN older method)' } else { 'OIDC / OAuth' }
    Write-Check INFO '  type' $shape

    if ($spSaml) {
        Write-Check OK '  SSO mode' 'SAML'
        $noAssign = ($sp.appRoleAssignmentRequired -eq $false)
        Write-Check $(if ($noAssign) { 'OK' } else { 'WARN' }) '  assignment required' $(if ($noAssign) { 'No (all users can sign in)' } else { 'Yes — users must be assigned first' })
        if (-not $noAssign) {
            # Safe -Fix: only LOOSENS access (lets all tenant users sign in via the app); it cannot
            # break an existing federation. "No" is what BSN's SSO setup article requires (Step 6),
            # confirmed against that article 2026-07-15.
            Invoke-Fix 'set assignment required = No (all users can sign in)' {
                Graph-Patch "/servicePrincipals/$($sp.id)" @{ appRoleAssignmentRequired = $false }
                Write-Host '          set appRoleAssignmentRequired = No.' -ForegroundColor Green
            }
        }

        Test-SsoSigningCert $sp -Shape 'enterprise'

        # Email claim. Email sync = Azure default (emailaddress from user.mail, no policy). UPN sync
        # = a claims-mapping policy remapping emailaddress -> user.userPrincipalName.
        $policies = $null
        try { $policies = @((Graph-Get "/servicePrincipals/$($sp.id)/claimsMappingPolicies").value) }
        catch { Write-Check WARN '  email claim' "could not read claims-mapping policy (needs Policy.Read.All): $($_.Exception.Message)" }
        if ($null -ne $policies) {
            if ($SsoSyncType -eq 'UPN') {
                if ($policies.Count) { Write-Check OK '  email claim' "UPN: claims policy assigned ('$($policies[0].displayName)')" }
                else { Write-Check FAIL '  email claim' 'UPN requested but no claims-mapping policy assigned — email will be user.mail, not UPN' 'Users may get stuck at "Redirecting" when signing in: BSN will be sent the wrong email address for them, so it cannot match them to their account.' }
            } else {
                if ($policies.Count) { Write-Check WARN '  email claim' "Email sync but a claims policy is assigned ('$($policies[0].displayName)') — email may be remapped" }
                else { Write-Check OK '  email claim' 'Email: Azure default (emailaddress from user.mail)' }
            }
        }

        # App-specific federation metadata must be reachable for the portal's Connect step.
        $tenant = (Get-MgContext).TenantId
        $fedMeta = "https://login.microsoftonline.com/$tenant/federationmetadata/2007-06/federationmetadata.xml?appid=$($app.appId)"
        try {
            $r = Invoke-WebRequest -Uri $fedMeta -Method GET -TimeoutSec 20 -ErrorAction Stop
            $reachable = ($r.StatusCode -eq 200 -and $r.Content -match 'EntityDescriptor')
            Write-Check $(if ($reachable) { 'OK' } else { 'WARN' }) '  federation metadata' $(if ($reachable) { 'reachable (app-specific)' } else { "unexpected response (HTTP $($r.StatusCode))" })
        } catch {
            Write-Check WARN '  federation metadata' "not reachable yet: $($_.Exception.Message)"
        }
    }
    elseif ($legacySaml) {
        $samlEmail = [bool](@($app.optionalClaims.saml2Token) | Where-Object { $_.name -eq 'email' })
        Write-Check $(if ($samlEmail) { 'OK' } else { 'WARN' }) '  SAML email claim' $(if ($samlEmail) { 'present (legacy optional claim)' } else { 'not an optional claim — email may come from a claims policy' })
        if ($sp) { Test-SsoSigningCert $sp -Shape 'legacy' }
        # BSN's current article does specify the enterprise app (confirmed 2026-07-15), so naming
        # this the older shape is fair. Still don't push the migration: nothing in BSN's docs asks
        # partners to rebuild working SSO, and -ReplaceSso DELETES the app to do it. State it; let a
        # human weigh it.
        Write-Host '         App-registration SSO — working. This is BSN''s older method; their current' -ForegroundColor DarkGray
        Write-Host '         article builds a non-gallery SAML enterprise app instead. Migrating is optional' -ForegroundColor DarkGray
        Write-Host '         and NOT risk-free — -ReplaceSso deletes and recreates the app, and the portal' -ForegroundColor DarkGray
        Write-Host '         needs the new metadata URL pasted back before users can sign in again.' -ForegroundColor DarkGray
    }
    else {
        # OIDC / OAuth (legacy OIDC runbook shape): token-v2, Graph email permission, admin consent.
        $v2 = ($app.api.requestedAccessTokenVersion -eq 2)
        Write-Check $(if ($v2) { 'OK' } else { 'FAIL' }) '  access token v2' $(if ($v2) { 'set' } else { 'NOT set' })
        $emailPerm = $false
        foreach ($rra in @($app.requiredResourceAccess)) {
            if ($rra.resourceAppId -eq $GraphAppId -and (@($rra.resourceAccess) | Where-Object { $_.id -eq $GraphEmailScope })) { $emailPerm = $true }
        }
        Write-Check $(if ($emailPerm) { 'OK' } else { 'FAIL' }) '  Graph email permission' $(if ($emailPerm) { 'present' } else { 'missing' })
        if ($sp) {
            $grants = @((Graph-Get "/servicePrincipals/$($sp.id)/oauth2PermissionGrants").value) | Where-Object { $_.scope -match 'email' }
            Write-Check $(if ($grants) { 'OK' } else { 'WARN' }) '  admin consent' $(if ($grants) { "granted ($($grants[0].scope.Trim()))" } else { 'not detected — grant in Entra' })
        }
    }
}

function Test-BsnAllowedSenders {
    if ($SkipAllowedSenders) { Write-Check WARN 'Anti-spam allow-list' 'skipped (-SkipAllowedSenders)'; return }
    if (-not (Ensure-Module 'ExchangeOnlineManagement')) { Write-Check WARN 'Anti-spam allow-list' 'ExchangeOnlineManagement unavailable' ; return }
    $opened = $false
    try {
        Import-Module ExchangeOnlineManagement
        if (-not (@(Get-ConnectionInformation -ErrorAction SilentlyContinue).Count)) {
            $connect = @{ ShowBanner = $false }
            $upn = (Get-MgContext).Account; if ($upn) { $connect.UserPrincipalName = $upn }
            if ($TenantId) { $connect.DelegatedOrganization = $TenantId }
            Connect-ExchangeOnline @connect | Out-Null; $opened = $true
        }
        $policy = Get-HostedContentFilterPolicy -Identity Default
        $allowed = @()
        foreach ($a in @($policy.AllowedSenders)) {
            if ($a.Sender.Address) { $allowed += $a.Sender.Address } elseif ($a.Address) { $allowed += $a.Address } else { $allowed += [string]$a }
        }
        # A manual setup may allow-list whole DOMAINS instead of individual addresses — that still
        # covers the BSN senders, so check both.
        $allowedDomains = @()
        foreach ($d in @($policy.AllowedSenderDomains)) {
            if ($d.Domain) { $allowedDomains += $d.Domain } elseif ($d.Name) { $allowedDomains += $d.Name } else { $allowedDomains += [string]$d }
        }
        foreach ($s in $BsnSenders) {
            $domain = ($s -split '@')[-1]
            if     ($allowed -contains $s)        { Write-Check OK   "Allowed sender $s" 'present (address)' }
            elseif ($allowedDomains -contains $domain) { Write-Check OK "Allowed sender $s" "covered by allowed domain '$domain'" }
            else {
                Write-Check FAIL "Allowed sender $s" 'missing' "Training and phishing-test emails from BSN may land in users' junk folders, because $s is not on the anti-spam allow-list."
                # Safe -Fix: additive Add to the Default anti-spam policy's allowed-senders list.
                Invoke-Fix "add allowed sender $s" {
                    Set-HostedContentFilterPolicy -Identity Default -AllowedSenders @{ Add = @($s) } | Out-Null
                    Write-Host "          added $s to the anti-spam allow-list." -ForegroundColor Green
                }
            }
        }
        # Echo what the policy actually contains so the operator can confirm the check sees it.
        Write-Host "         allow-list senders: $(if ($allowed) { $allowed -join ', ' } else { '(none)' })" -ForegroundColor DarkGray
        Write-Host "         allow-list domains: $(if ($allowedDomains) { $allowedDomains -join ', ' } else { '(none)' })" -ForegroundColor DarkGray
    } catch {
        Write-Check WARN 'Anti-spam allow-list' "could not check: $($_.Exception.Message)"
    } finally {
        if ($opened) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} }
    }
}

# Directory Sync and Direct Delivery leave a BSN enterprise app behind. Match by name/publisher
# (appIds can change), and report each app's OWNER TENANT so BSN-owned multi-tenant apps are
# distinguishable from locally-created ones (e.g. a manually-made, possibly misnamed SSO app).
function Test-BsnConsentApps {
    try {
        $localTenant = (Get-MgContext).TenantId
        $found = @()
        $uri = "/servicePrincipals?`$select=id,displayName,appId,appOwnerOrganizationId,verifiedPublisher,createdDateTime&`$top=999"
        while ($uri) {
            $resp = Graph-Get $uri
            foreach ($sp in $resp.value) {
                # BSN ships apps under two publishers: "Breach Secure Now" and "Entegration Inc"
                # (Entegration is the company that develops the BSN platform). 'catch phish' is in
                # the name list because BSN ships more than one Catch Phish identity and at least one
                # carries NO verified publisher — so neither the name nor the publisher pattern alone
                # catches them all.
                $pub = $sp.verifiedPublisher.displayName
                if ($sp.displayName -match '(?i)pii|breach|\bbsn\b|\bdmd\b|secure now|catch\s*phish' -or $pub -match '(?i)breach|secure now|entegration') { $found += $sp }
            }
            $next = $resp.'@odata.nextLink'
            $uri = if ($next) { $next -replace '^https://graph\.microsoft\.com/v1\.0', '' } else { $null }
        }
        if ($found.Count -eq 0) {
            Write-Check WARN 'BSN enterprise apps (Directory Sync / DMD)' 'none found — finish the portal consents, then -Verify again' 'This client''s users are probably not syncing into BSN at all: none of BSN''s own apps are connected to the tenant, which usually means the portal steps (Directory Sync / Direct Delivery) were never completed.'
            return
        }
        # Sorting by creation date turns this list into a rough timeline of the BSN deployment: when
        # Directory Sync was consented, when SSO was registered, when Catch Phish arrived. That dates
        # the PII-Protect rollout itself, which is independent of the tenant's own age.
        foreach ($m in ($found | Sort-Object { $_.createdDateTime })) {
            $owner = if ($m.appOwnerOrganizationId -eq $localTenant) { 'owned by THIS tenant (local app registration)' }
                     elseif ($m.appOwnerOrganizationId)              { "owned by BSN tenant $($m.appOwnerOrganizationId)" }
                     else                                            { 'owner unknown' }
            $pub = if ($m.verifiedPublisher.displayName) { "; publisher '$($m.verifiedPublisher.displayName)'" }
            $added = Format-CreatedDate $m.createdDateTime
            Write-Check OK "App '$($m.displayName)'" "$(if ($added) { "added $added; " })appId $($m.appId); $owner$pub"
        }
        Write-Host '         (Apps owned by a BSN tenant are the real Directory Sync / DMD apps; apps owned by' -ForegroundColor DarkGray
        Write-Host '          THIS tenant are locally created — e.g. a manually-made, maybe misnamed SSO app.' -ForegroundColor DarkGray
        Write-Host '          Listed oldest-first: the dates are when THIS BSN deployment was built.)' -ForegroundColor DarkGray
    } catch {
        Write-Check WARN 'BSN enterprise apps' "could not enumerate: $($_.Exception.Message)"
    }
}

# Was Catch Phish ever deployed? Nothing can answer that authoritatively from PowerShell, so this
# check is deliberately built to never claim more than it knows.
#
# Office centralized deployment has no Graph API, and the two admin-center surfaces write to two
# different stores. Add-ins deployed through Integrated apps are INVISIBLE to Exchange's
# Get-App -OrganizationApp — confirmed against two tenants running the add-in, where Get-App
# returned its other org add-ins but no Catch Phish. (The Get-App docs hint at this but file the
# caveat under -Mailbox.) So:
#
#   * Entra service principal — the primary signal. The "Get apps" flow ends in Global Admin consent,
#     which leaves an enterprise app behind. Strong evidence, though it proves CONSENT rather than
#     deployment, and can't show the audience.
#   * Get-App -OrganizationApp — a positive-only supplement. It only sees the older "Add-ins" page
#     deployments, so a hit is meaningful but a miss means nothing at all.
#
# Absence therefore never reports FAIL: the admin center is the only real authority, so a no-trace
# result points there instead of asserting the add-in is missing. No -Fix: deployment is manual.
# How long before a SAML signing cert expires we start warning. 60 days is chosen to comfortably
# clear a monthly -Verify cadence, so a cert can't slip from "fine" to "expired" between two runs,
# and to leave room to book the portal metadata re-paste that a rotation forces.
$script:CertWarnDays = 60

# A keyCredential's thumbprint, in the same form as preferredTokenSigningKeyThumbprint (lower-case
# hex). Graph reports the two differently — the preferred key is hex, while customKeyIdentifier is
# the same bytes base64-encoded (and Invoke-MgGraphRequest may hand it back as a byte[] already) —
# which is why they can't just be string-compared. Returns $null if it can't be read, so callers
# treat it as "no match" rather than throwing.
function Get-KeyThumbprint($keyCredential) {
    $id = $keyCredential.customKeyIdentifier
    if (-not $id) { return $null }
    try {
        $bytes = if ($id -is [byte[]]) { $id } else { [Convert]::FromBase64String("$id") }
        return (-join ($bytes | ForEach-Object { $_.ToString('x2') })).ToLower()
    } catch { return $null }
}

# Render a Graph createdDateTime as a plain date. Graph hands this back as a string or a DateTime
# depending on the call, and it's absent on older objects — so parse defensively and say nothing
# rather than printing a broken date. Parse via DateTimeOffset and print UTC: casting straight to
# [datetime] converts to local time, which slides anything consented near midnight UTC onto the
# wrong day.
function Format-CreatedDate($value) {
    if (-not $value) { return $null }
    try { return ([datetimeoffset]$value).UtcDateTime.ToString('yyyy-MM-dd') } catch { return "$value" }
}

function Test-BsnCatchPhish {
    if ($SkipCatchPhish) { Write-Check WARN 'Catch Phish add-in' 'skipped (-SkipCatchPhish)'; return }

    $sawIt = $false
    # Tracks whether the Exchange half actually ran. It decides how loudly we can call an absence:
    # the two signals cover the two deployment surfaces, so "nothing in either" is real evidence —
    # but only if both were genuinely checked.
    $exoChecked = $false

    # 1. Primary: the Catch Phish app identity left by the Integrated apps consent step. Keyed on
    # name/publisher, not appId (BSN's appIds change — see the app inventory in the README).
    try {
        $sp = @((Graph-Get "/servicePrincipals?`$filter=startswith(displayName,'Catch Phish')&`$select=id,displayName,appId,appOwnerOrganizationId,verifiedPublisher,createdDateTime").value)
        $localTenant = (Get-MgContext).TenantId
        foreach ($s in $sp) {
            $sawIt = $true
            $pub = if ($s.verifiedPublisher.displayName) { "publisher '$($s.verifiedPublisher.displayName)'" } else { 'publisher unverified' }
            # Report provenance; do NOT grade it. Observed 2026-07-15: a tenant may hold more than
            # one app named "Catch Phish", the appId differs per tenant (unlike BSN's multi-tenant
            # DMD/BSN_Azure_Sync, which share one appId everywhere), and some carry no verified
            # publisher. A PII-Protect deployment stood up years ago may predate the current deploy
            # method or app version, so no shape here is reliably "wrong" — grading it would just cry
            # wolf on whichever vintage the operator is looking at. Existence is the signal.
            #
            # The creation date is the useful context: it dates THIS BSN deployment (when Catch Phish
            # was consented here), independent of how old the tenant is — which is what tells you
            # which era's method a client was set up with.
            $owner = if ($s.appOwnerOrganizationId -eq $localTenant) { 'registered in this tenant' }
                     elseif ($s.appOwnerOrganizationId)              { "owned by tenant $($s.appOwnerOrganizationId)" }
                     else                                            { 'owner unknown' }
            $added = Format-CreatedDate $s.createdDateTime
            Write-Check OK "Catch Phish app '$($s.displayName)'" "present$(if ($added) { "; added $added" }) (appId $($s.appId); $owner; $pub)"
        }
        # State the observation only. The verdict comes at the end, once BOTH signals have reported —
        # concluding here would contradict a legacy-page deployment that Get-App is about to find.
        if (-not $sp.Count) { Write-Check INFO 'Catch Phish app identity' 'not found in Entra (the Integrated apps flow normally leaves one behind)' }
    } catch {
        Write-Check WARN 'Catch Phish app identity' "could not check: $($_.Exception.Message)"
    }

    # 2. Supplement: catch the older "Add-ins"-page deployments, which DO land in Exchange. Only
    # reported when it finds something — a miss here is expected for the documented method and would
    # be pure noise.
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        $opened = $false
        try {
            Import-Module ExchangeOnlineManagement
            if (-not (@(Get-ConnectionInformation -ErrorAction SilentlyContinue).Count)) {
                $connect = @{ ShowBanner = $false }
                $upn = (Get-MgContext).Account; if ($upn) { $connect.UserPrincipalName = $upn }
                if ($TenantId) { $connect.DelegatedOrganization = $TenantId }
                Connect-ExchangeOnline @connect | Out-Null; $opened = $true
            }
            # Match on the rendered object, not one property name: Get-App's shape varies.
            $orgApps = @(Get-App -OrganizationApp -ErrorAction Stop)
            $exoChecked = $true
            $found = @($orgApps | Where-Object { ($_ | Out-String) -match '(?i)catch\s*phish|breach\s*secure' })
            foreach ($a in $found) {
                $sawIt = $true
                $name = if ($a.DisplayName) { $a.DisplayName } else { 'Catch Phish' }
                # Deployed-but-disabled is a silent failure, so it's worth calling out.
                if ($null -ne $a.Enabled -and -not $a.Enabled) {
                    Write-Check FAIL "Catch Phish add-in '$name'" 'deployed via the older Add-ins page, but DISABLED — users cannot see it' 'The Catch Phish button is installed but switched off, so staff cannot see it in Outlook. Turn it back on in the Microsoft 365 admin center.'
                } else {
                    Write-Check OK "Catch Phish add-in '$name'" "deployed via the older Add-ins page$(if ($a.AppVersion) { "; version $($a.AppVersion)" })"
                }
            }
        } catch {
            # Expected when the sign-in lacks an Exchange role. The Entra signal above is the one
            # that matters, so this is a footnote, not a failure.
            Write-Verbose "Get-App -OrganizationApp unavailable: $($_.Exception.Message)"
        } finally {
            if ($opened) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} }
        }
    }

    # 3. Verdict. The two signals cover the two deployment surfaces — an Integrated-apps deploy
    # leaves the Entra app, an older Add-ins-page deploy lands in Get-App — so silence from BOTH is
    # good evidence of a genuine gap. Calibrated 2026-07-15 on four live tenants: the three running
    # the add-in each had the Entra app, and the one without it (confirmed empty in both
    # admin-center lists) had neither. If Exchange couldn't be reached, only half the ground is
    # covered, so stay hedged.
    if ($sawIt) {
        Write-Host '         Neither signal reports the audience, and consent is not proof of deployment —' -ForegroundColor DarkGray
        Write-Host "         confirm it is listed and assigned: $IntegratedAppsUrl" -ForegroundColor DarkGray
    } elseif ($exoChecked) {
        Write-Check FAIL 'Catch Phish add-in' 'not deployed — nothing in Entra and nothing in Exchange (both surfaces checked)' 'Staff have no "report phishing" button in Outlook. Deploy it with: ./provision-bsn.ps1 -CatchPhishOnly'
        Write-Host "         Deploy it: ./provision-bsn.ps1 -CatchPhishOnly   (or $IntegratedAppsUrl)" -ForegroundColor DarkGray
    } else {
        Write-Check WARN 'Catch Phish add-in' 'no app in Entra, and Exchange could not be checked — probably not deployed, but confirm by hand before redeploying' 'Staff probably have no "report phishing" button in Outlook, but this could not be confirmed — check the Microsoft 365 admin center before deploying it again.'
        Write-Host "         Integrated apps > Get apps: $IntegratedAppsUrl (AssetId $CatchPhishAssetId)" -ForegroundColor DarkGray
    }
}

# Plain-English wrap-up. The check lines above are a technical audit trail; this answers the only
# questions a non-technical operator actually has: is this client set up, what's broken, what does it
# mean for them, and what do I do next. Wraps prose to the terminal so it reads as sentences.
function Write-Wrapped([string]$text, [string]$indent, [string]$colour) {
    # Bulleted text gets a hanging indent, so wrapped lines sit under the sentence rather than under
    # the marker — otherwise a two-line bullet reads as two separate points.
    $hang = if ($text -match '^[*-] ') { "$indent  " } else { $indent }
    # Fall back to 80 when there's no real console (piped/redirected output reports width 0 or null).
    $console = $Host.UI.RawUI.WindowSize.Width
    if (-not $console -or $console -lt 20) { $console = 80 }
    $width = [Math]::Max(40, [Math]::Min(96, $console - 2) - $hang.Length)
    $line = ''
    $prefix = $indent
    foreach ($word in ($text -split '\s+' | Where-Object { $_ })) {
        if ($line -and ($line.Length + 1 + $word.Length) -gt $width) {
            Write-Host "$prefix$line" -ForegroundColor $colour
            $line = $word; $prefix = $hang
        } else { $line = if ($line) { "$line $word" } else { $word } }
    }
    if ($line) { Write-Host "$prefix$line" -ForegroundColor $colour }
}

function Write-PlainSummary {
    # A check with no $plain is still counted, just described in its own technical words.
    $say = { param($r) if ($r.Plain) { $r.Plain } else { "$($r.Item): $($r.Detail)" } }

    $problems = @($script:checkResults | Where-Object { $_.Status -eq 'FAIL' -and -not $_.Fixed })
    $watch    = @($script:checkResults | Where-Object { $_.Status -eq 'WARN' -and -not $_.Fixed -and $_.Plain })
    $repaired = @($script:checkResults | Where-Object { $_.Fixed })

    Write-Section 'In plain English'

    if ($repaired.Count) {
        Write-Host "  Fixed $($repaired.Count) thing$(if ($repaired.Count -ne 1) { 's' }):" -ForegroundColor Green
        foreach ($f in $script:fixesApplied) { Write-Wrapped "- $f" '    ' 'Green' }
        Write-Host ''
    }

    if (-not $problems.Count -and -not $watch.Count) {
        Write-Wrapped 'This client looks fully set up. Nothing needs your attention.' '  ' 'Green'
    } elseif (-not $problems.Count) {
        Write-Wrapped 'No real problems. A couple of things are worth a look when you have a minute:' '  ' 'Green'
    } else {
        Write-Host "  $($problems.Count) thing$(if ($problems.Count -ne 1) { 's' }) need$(if ($problems.Count -eq 1) { 's' }) attention:" -ForegroundColor Red
    }

    foreach ($p in $problems) { Write-Host ''; Write-Wrapped "* $(& $say $p)" '  ' 'Red' }
    foreach ($w in $watch)    { Write-Host ''; Write-Wrapped "- $(& $say $w)" '  ' 'Yellow' }

    # What to do next, in the operator's terms — only offer -Fix if something here is actually
    # fixable, so we never send someone to a command that will do nothing.
    $fixable = @($script:checkResults | Where-Object { $_.Status -in @('FAIL', 'WARN') -and -not $_.Fixed -and $_.Fixable })
    if ($problems.Count -or $watch.Count) {
        Write-Host ''
        if ($repaired.Count) {
            Write-Wrapped 'Re-run with -Verify to confirm the fixes took effect (some changes take a minute to apply).' '  ' 'Cyan'
        } elseif ($fixable.Count) {
            Write-Wrapped 'Some of these can be repaired automatically: re-run this script with -Fix.' '  ' 'Cyan'
        }
    }
    Write-Host ''
}

function Invoke-VerifyPhase {
    $script:checkResults = @()   # a fresh pass reports on itself, not on an earlier run
    Write-Section $(if ($Fix) { 'Verify + fix (remediate detected gaps)' } else { 'Verification (read-only)' })
    Test-BsnGroups
    Test-BsnSso
    Test-BsnAllowedSenders
    Test-BsnCatchPhish
    Test-BsnConsentApps
    Write-PlainSummary
}

# =======================================================================================
if ($CatchPhishOnly -and $SkipCatchPhish) { throw '-CatchPhishOnly and -SkipCatchPhish are contradictory: pick one.' }
if ($RenewSsoCert -and $SkipSso) { throw '-RenewSsoCert and -SkipSso are contradictory: pick one.' }
if ($RenewSsoCert -and $CatchPhishOnly) { throw '-RenewSsoCert and -CatchPhishOnly each do one job: run them one at a time.' }

# Prompt for required input that wasn't supplied (keeps a bare, no-argument run friendly).
# ClientName is only needed for the provisioning run, not for -Verify / -Fix / -CatchPhishOnly.
if (-not $ClientName -and -not $Verify -and -not $Fix -and -not $CatchPhishOnly -and -not $RenewSsoCert) {
    if ($NonInteractive) { throw 'ClientName is required: pass -ClientName, or omit -NonInteractive to be prompted.' }
    $ClientName = Read-Host 'Client business name (as it should appear in the PII Protect portal)'
    if (-not $ClientName) { throw 'A client name is required.' }
}

Write-Section $(
    if     ($RenewSsoCert)   { 'Renew SSO signing certificate' }
    elseif ($CatchPhishOnly) { "Catch Phish add-in $(if ($Verify) { '(check only)' } else { '(guided)' })" }
    elseif ($Fix)            { 'Fix BSN PII Protect' }
    elseif ($Verify)         { 'Verify BSN PII Protect' }
    else                     { "Provision BSN PII Protect  —  client: $ClientName" }
)

# --- Prerequisites: install EVERY module this run needs up front, before the main body, so
# nothing installs mid-run (which interrupts output and conflicts with in-use PowerShellGet). ----
$needed = @('Microsoft.Graph.Authentication')
# Exchange Online serves two callers: the anti-spam phase (provision + verify) and the Catch Phish
# check (Get-App -OrganizationApp), so it's needed unless BOTH are skipped.
if (-not $SkipAllowedSenders -or -not $SkipCatchPhish) { $needed += 'ExchangeOnlineManagement' }
Write-Host 'Checking required PowerShell modules...' -ForegroundColor Cyan
foreach ($m in $needed) {
    if (Ensure-Module $m) { Write-Host "  $m — ready." -ForegroundColor Green }
    elseif ($m -eq 'Microsoft.Graph.Authentication') { throw 'Microsoft.Graph.Authentication is required and could not be installed.' }
}
Import-Module Microsoft.Graph.Authentication

# Connect to Graph with the scopes the run needs. A look-only -Verify requests read-only scopes
# (no write consent); -Fix needs write access to remediate, so it connects like a provisioning run.
# -CatchPhishOnly writes nothing via Graph (the deploying happens in the admin center) and the only
# Graph read is Test-BsnCatchPhish's service-principal fallback — so it stays read-only even with
# -Fix, and asks for the narrowest consent of any mode.
$readOnly = $CatchPhishOnly -or ($Verify -and -not $Fix)
$scopes = if ($CatchPhishOnly) {
    @('Application.Read.All', 'Directory.Read.All')
} elseif ($RenewSsoCert) {
    # Writes one certificate onto one service principal — nothing else, so don't ask for group or
    # policy write consent it will never use.
    @('Application.ReadWrite.All', 'Directory.Read.All')
} elseif ($readOnly) {
    # Read-only. Policy.Read.All lets verify read the SSO app's assigned claims-mapping policy (UPN
    # sync); the check degrades to a WARN if that scope is refused.
    @('Group.Read.All', 'Application.Read.All', 'Directory.Read.All', 'Policy.Read.All')
} else {
    $s = @('Group.ReadWrite.All', 'Application.ReadWrite.All', 'Directory.Read.All', 'Policy.Read.All')
    # Only a provisioning run's UPN sync writes a claims-mapping policy (-Fix never touches claims),
    # so request that write scope only then.
    if (-not $Fix -and -not $SkipSso -and $SsoSyncType -eq 'UPN') { $s += 'Policy.ReadWrite.ApplicationConfiguration' }
    $s
}
$connectArgs = @{ Scopes = $scopes }
if ($TenantId)      { $connectArgs.TenantId = $TenantId }
if ($UseDeviceCode) { $connectArgs.UseDeviceCode = $true }
Write-Host 'Connecting to Microsoft Graph (sign in to the CLIENT tenant)...' -ForegroundColor Cyan
Connect-MgGraph @connectArgs | Out-Null
$me = Graph-Get '/me?$select=userPrincipalName'
Write-Host "Signed in as $($me.userPrincipalName) on tenant $((Get-MgContext).TenantId)." -ForegroundColor Green

try {
    # -CatchPhishOnly: the add-in is usually deployed days or weeks after the onboarding visit, so
    # it runs on its own — guided step, then the check. With -Verify, just the check.
    if ($CatchPhishOnly) {
        if (-not $Verify) { Invoke-CatchPhishStep }
        Invoke-Phase 'Catch Phish' { Test-BsnCatchPhish }
        if ($script:phaseErrors.Count) { Write-Warning 'The Catch Phish check hit a problem.' }
        return
    }

    # -RenewSsoCert: rotate the SAML signing cert and stop. One job, because it is disruptive and
    # ends in a portal step the operator must finish immediately.
    if ($RenewSsoCert) {
        Invoke-Phase 'Renew SSO certificate' { Invoke-RenewSsoCertPhase }
        if ($script:phaseErrors.Count) { Write-Warning 'The certificate renewal hit a problem — re-run -Verify to see the current state.' }
        return
    }

    # -Fix: verify + remediate the safe/automatable gaps, no manual portal steps, then done.
    if ($Fix) {
        Invoke-Phase 'Fix' { Invoke-VerifyPhase }
        Write-Section 'Done'
        Write-Host 'Applied the safe automatable fixes. Re-run with -Verify to confirm they stuck' -ForegroundColor Cyan
        Write-Host '(some writes replicate with a short delay). Cert / UPN / legacy-SSO migration are not' -ForegroundColor Cyan
        Write-Host 'auto-fixed — use a normal run (add -ReplaceSso to migrate a legacy SSO app).' -ForegroundColor Cyan
        if ($script:phaseErrors.Count) { Write-Warning "The fix pass hit a problem: $($script:phaseErrors -join ', ')." }
        return
    }

    # -Verify: read-only checks only, then done.
    if ($Verify) {
        Invoke-Phase 'Verification' { Invoke-VerifyPhase }
        return
    }

    # 1. Entra security groups (+ dynamic BSN-Employees by default).
    Invoke-Phase 'Security groups' { Invoke-GroupsPhase }

    # 2. Anti-spam allowed senders (default on; separate Exchange Online connection via SSO).
    if (-not $SkipAllowedSenders) { Invoke-Phase 'Anti-spam allowed senders' { Invoke-AllowedSendersPhase } }
    else { Write-Host "`n(Anti-spam allowed senders skipped by -SkipAllowedSenders.)" -ForegroundColor DarkGray }

    # 3. Manual: add the client as a tenant in the PII Protect portal.
    Invoke-ManualStep -Title 'Add the client tenant' -Url 'https://portal.pii-protect.com' -Steps @(
        'Log in as a Partner Administrator.',
        'Manage Clients > "+ New Client".',
        "Business name: $ClientName (match your own client records exactly).",
        'Pick the closest Industry vertical, then Create.',
        'Open the new client > User Management tab.',
        'Click "Welcome Message" to toggle it OFF (red) so welcome emails are not sent prematurely.'
    )

    # 4. Manual: enable Directory Sync (Azure AD) — portal-driven Global Admin consent.
    Invoke-ManualStep -Title 'Enable Directory Sync (Azure AD)' -Url 'https://portal.pii-protect.com' -Steps @(
        'Client > User Management > "Directory Sync".',
        'Choose Your Sync Type > "Azure Active Directory".',
        'Click Enable. Portal Logon: select "Email".',
        'Click "Authorize Directory Access" and consent with a Global Admin of the client tenant.',
        'Review permissions > Accept. Wait for "Verified Successfully!".'
    )

    # 5. SSO (hybrid portal + Entra automation).
    if (-not $SkipSso) { Invoke-Phase 'Single Sign-On app' { Invoke-SsoPhase } }

    # 6. Manual: Direct Delivery / phishing whitelisting — portal-driven consent.
    Invoke-ManualStep -Title 'Enable Direct Delivery (phishing)' -Url 'https://portal.pii-protect.com' -Steps @(
        'Client > Phishing tab > Whitelisting.',
        'Click Enable; authenticate with a Global Admin of the client tenant.',
        'Review permissions > Accept.',
        'If it does not turn green, re-check later; if you disable/re-enable you may need to',
        '   re-grant permissions under Entra: App registrations/Enterprise applications > DMD.'
    )

    # 7. Manual: Catch Phish Outlook add-in (Microsoft 365 admin center, not the BSN portal).
    if (-not $SkipCatchPhish) { Invoke-CatchPhishStep }
    else { Write-Host "`n(Catch Phish add-in skipped by -SkipCatchPhish.)" -ForegroundColor DarkGray }

    # 8. Verify what left a footprint in the tenant.
    Invoke-Phase 'Verification' { Invoke-VerifyPhase }

    Write-Section 'Done'
    Write-Host "Automated: Entra groups$(if (-not $NoDynamicEmployees) { ' (+ dynamic BSN-Employees where P1)' })$(if (-not $SkipAllowedSenders) { ', anti-spam allowed senders' })$(if ($script:ssoConfigured) { ', PII-Protect SSO (SAML enterprise app)' })." -ForegroundColor Green
    if (-not $SkipSso -and -not $script:ssoConfigured) {
        Write-Warning 'SSO was NOT configured (see above). Re-run with -SsoRedirectUri / -SsoAppIdUri, or interactively with the portal popup open.'
    }
    if ($script:phaseErrors.Count) {
        Write-Warning "These automated phases hit problems and need a second look: $($script:phaseErrors -join ', ')."
    }
    Write-Host 'Verify in the portal: tenant added, Directory Sync "Verified", SSO Connected, Direct Delivery green.' -ForegroundColor Cyan
    if (-not $SkipCatchPhish) {
        Write-Host 'Catch Phish is deployed by hand and can take 24-72h to reach ribbons — re-run -Verify later.' -ForegroundColor Cyan
    }
}
finally {
    Disconnect-MgGraph | Out-Null
}
