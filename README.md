# provision-bsn

A guided PowerShell script that onboards a new client onto **Breach Secure Now "PII Protect"**.
It automates the Microsoft Entra / Microsoft 365 work and **walks the operator through the PII
Protect portal steps** that have no API, in the correct runbook order.

Everything lives in one cross-platform script — Windows, macOS, and Linux, on PowerShell 7.

## What it does

| Phase | Type | Notes |
|---|---|---|
| Create BSN security groups | **Automated** (Graph) | `BSN-Employees`, `BSN-Managers`, `BSN-ManagerAdmins`, optional `BSN-PartnerAdmins`, and `BSN-TAG-*` tag groups |
| Dynamic `BSN-Employees` membership | **Automated** (Graph) | on by default when the tenant has Entra ID **P1+** (`-NoDynamicEmployees` forces a plain assigned group); rule from service-plan ids + surname/givenName-not-null (+ `-ExcludeEmail`) |
| Anti-spam allowed senders | **Automated** (Exchange Online) | on by default (`-SkipAllowedSenders` to skip); adds the 3 BSN addresses to the default inbound policy |
| Add client tenant to portal | **Manual, prompted** | portal.pii-protect.com has no API |
| Enable Directory Sync (Azure AD) | **Manual, prompted** | portal-driven Global Admin consent |
| Single Sign-On app | **Hybrid** | you paste the portal's Redirect URL + Application ID URI; the script builds the `PII-Protect SSO` **non-gallery SAML enterprise app** (SAML SSO mode, per-app signing certificate, Entity ID + Reply URL, email claim) and prints the federation-metadata URL to paste back. `-SkipSso` to skip |
| Direct Delivery (phishing) | **Manual, prompted** | portal-driven consent |
| Catch Phish Outlook add-in | **Manual, prompted** | Microsoft 365 admin center (Settings → Integrated apps); links you to the page and the marketplace listing. Runs standalone via `-CatchPhishOnly` (it's usually deployed weeks later); `-SkipCatchPhish` to skip |

Manual steps pause with numbered instructions and the exact URL; the script resumes when you
press Enter (or are skipped with `-NonInteractive`).

## Requirements

- **PowerShell 7+**
- **Microsoft.Graph.Authentication** — the only module needed for the Entra phases; all Graph
  calls go through `Invoke-MgGraphRequest`.
- **ExchangeOnlineManagement** — used by the anti-spam phase and the Catch Phish check (both on by
  default); it reuses your Graph sign-in via SSO, so it should connect without a second credential
  prompt. Only skipped if you pass **both** `-SkipAllowedSenders` and `-SkipCatchPhish`.
- A sign-in with rights in the **client tenant** to manage groups, register apps, grant consent
  (Global Admin is typical, since the portal consent steps need it anyway).

The modules are **installed automatically** (`-Scope CurrentUser`) on first use — no manual setup.

## Usage

It's built for a **non-technical operator**: run it with just the client name and it attempts
every automatable step, prompts for anything required that you didn't pass, and — if a phase
fails — warns and keeps going instead of aborting.

```powershell
# Full guided walk-through (recommended). Prompts for the client name if you omit it.
./provision-bsn.ps1 -ClientName "Acme Widgets"

# Entra-only bootstrap: no portal prompts, no Exchange, no SSO.
./provision-bsn.ps1 -ClientName "Acme Widgets" -NonInteractive -SkipAllowedSenders -SkipSso
```

### Run it straight from GitHub

No clone, no download — paste this into a **PowerShell 7** prompt. Parameters work exactly as they
do locally:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/compudatasystems/provision-bsn/master/provision-bsn.ps1))) -ClientName "Acme Widgets"
```

```powershell
# Same thing, read-only verification pass:
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/compudatasystems/provision-bsn/master/provision-bsn.ps1))) -Verify
```

> Run it from **pwsh 7**, not Windows PowerShell 5.1. The script's `#Requires -Version 7.0` guard is
> **not enforced** when it's run this way (`#Requires` only applies to script *files*), so on 5.1 it
> will start and then fail somewhere less obvious. If you're unsure, check `$PSVersionTable.PSVersion`.

Read the script before you pipe it into your shell — that goes for this one as much as any other.

By default it makes `BSN-Employees` **dynamic** when the tenant has Entra ID P1 (falling back to a
standard assigned group, with printed enrol-by-hand instructions, if not) and configures the
anti-spam allow-list. Turn those off with `-NoDynamicEmployees` / `-SkipAllowedSenders`.

Everything honours `-WhatIf`, and re-running is safe (existing groups/objects are detected and
skipped). See `Get-Help ./provision-bsn.ps1 -Full` for all parameters.

### Other options

- `-TenantId <id|domain>` — the client tenant to sign in to (defaults to whatever you authenticate
  against).
- `-IncludePartnerAdmins` — also create the `BSN-PartnerAdmins` group. Per BSN this is **only** for
  your own internal BPP account — don't use it in a client tenant.
- `-TagGroup 'Outside Sales','Service'` — create `BSN-TAG-<name>` organizational tag groups.
- `-UseDeviceCode` — sign in with a device code instead of a browser (headless / SSH sessions).

### Single Sign-On (SAML enterprise app)

The SSO phase builds a **non-gallery SAML enterprise app**
named `PII-Protect SSO`, configured entirely through Graph. You paste the portal's Redirect URL and
Application ID URI; the script instantiates the app, sets SAML SSO mode + assignment-not-required,
adds a per-app signing certificate, sets the email claim, and prints the app-specific federation
metadata URL to paste back into the portal. You can pre-supply the pasted values with
`-SsoRedirectUri` / `-SsoAppIdUri` (required to run SSO under `-NonInteractive`).

- `-SsoSyncType Email|UPN` (default **Email**) — Email uses Azure's built-in `emailaddress`-from-
  `user.mail` claim; UPN adds a claims-mapping policy remapping it to `user.userPrincipalName`.
  Match this to the portal's Directory Sync → "Portal Logon" choice.
- `-SsoCertYears 1..3` (default **3**) — lifetime of the signing certificate (3 is Azure's max for
  an auto-generated cert; rotating it means re-pasting metadata in the portal, so the default
  minimises churn).
- `-ReplaceSso` — a SAML Entity ID must be unique tenant-wide, so if an app already holds it (e.g.
  an older PII-Protect app registration) that app has to go first. `-ReplaceSso` deletes-then-
  recreates without prompting; interactively you're asked before anything is deleted; in
  `-NonInteractive` nothing is deleted unless you pass it.
- `-RenewSsoCert` — rotate the SAML signing certificate on demand, whatever life it has left (a
  suspected compromise, a policy rotation, or getting ahead of the expiry warning). Runs on its own,
  needs no `-ClientName`, asks before touching anything, and honours `-WhatIf`.

#### Renewing the signing certificate

```powershell
./provision-bsn.ps1 -RenewSsoCert            # asks first, then prints the metadata URL to paste back
./provision-bsn.ps1 -RenewSsoCert -WhatIf    # preview
```

> **This breaks sign-in until you finish the portal step.** BSN caches the certificate from the
> metadata you paste into their portal, so the moment Entra signs with the new one, users can't log
> in until the portal has the new metadata. Do it when you can finish immediately.

It renews only where a certificate actually exists to renew. If the app has none of its own, Entra
signs with your tenant's default key — Microsoft's, not yours — and it says so rather than adding
one, because that would change how the app signs rather than renew anything.

This is deliberately **not** part of `-Fix`, which only ever makes changes that cannot break a
working setup.

#### Editing membership rules

`-Verify` prints a link to each group's rule editor. These are undocumented portal URLs — verified by
clicking, not by any Microsoft reference — so if one stops working, navigate **entra.microsoft.com →
Groups → All groups → the group → Dynamic membership rules** (the script prints that fallback too).

> The links **don't pin the tenant**, so the portal opens in whichever directory your browser used
> last. If you work across clients, check the portal landed on the one you just verified before you
> start editing.

Rules are authored in Entra, not here: the script creates the groups and applies one opinionated
default for `BSN-Employees`, and **never touches the rule on a group that's already dynamic** — so a
rule you craft by hand survives every re-run, including `-Fix`. `-Verify` then shows it back to you.

> Thinking of building `BSN-TAG-Sales` from a Sales *group*? That needs Entra's `memberOf` operator,
> which is [preview and explicitly not for production](https://learn.microsoft.com/en-us/entra/identity/users/groups-dynamic-rule-member-of):
> it can't be combined with other clauses (so it can't keep the `surname`/`givenName` guard BSN needs)
> and it doesn't remove people when they leave the source group. An attribute rule like
> `user.department -eq "Sales"` does the same job, is supported, and self-heals.

### The dynamic membership rule

Dynamic `BSN-Employees` (the default on P1) sets **Dynamic User**. Supply the license(s) that mean
"should be enrolled" with `-EnrollServicePlanId <ids>` (service-plan id reference:
<https://learn.microsoft.com/entra/identity/users/licensing-service-plan-reference>). Omit it to
match any user with an enabled service plan. The rule always requires `surname` and `givenName`
(BSN won't sync users missing a first/last name), and `-ExcludeEmail` drops specific addresses.

> Converting an existing **Assigned** group to Dynamic removes manually-added members — the script
> warns before doing so, and `-Fix` never does it automatically.

### Verify and fix

A normal run ends with a read-only **verification pass**, and you can run it on its own anytime.

```powershell
# Read-only checks (safe to point at any tenant; read-only Graph scopes, no write consent).
./provision-bsn.ps1 -Verify -SkipAllowedSenders

# Verify AND remediate the safe, automatable gaps (write scopes; no manual portal steps).
./provision-bsn.ps1 -Fix
```

Both end with a **plain-English summary**, so you don't have to read the check lines to know where a
client stands. A healthy tenant says so in one line; otherwise it explains what each problem *means*
for the client and what to do next:

```
========================================================================
  In plain English
========================================================================
  3 things need attention:

  * Managers will not be able to see their team's training progress. The
    BSN-Managers group is missing.

  * Training and phishing-test emails from BSN may land in users' junk
    folders, because no-reply@breachsecurenow.com is not on the anti-spam
    allow-list.

  * Staff have no "report phishing" button in Outlook. Deploy it with:
    ./provision-bsn.ps1 -CatchPhishOnly

  Some of these can be repaired automatically: re-run this script with -Fix.
```

It only suggests `-Fix` when `-Fix` can actually repair something. Under `-Fix`, the summary
separates what it repaired from what still needs a human.

Above the summary, `-Verify` reports a ✓ / ✗ / ⚠ line per step:

- **Security groups** — plus whether `BSN-Employees` is dynamic; if it's assigned, it says whether
  the tenant *has* Entra ID P1, so you know if auto-enrol is even on the table. For every BSN group
  it also reports the **member count** and, when dynamic, the **membership rule** — because "the
  group exists" says nothing about whether it works. An empty required group enrols nobody, and a
  **paused** rule freezes membership silently; both look identical to a name check, so both are
  called out. Each group also gets a **deep link straight into Entra** — the rule editor for a
  dynamic group, the members list otherwise — so you don't have to go hunting for the blade.
- **The `PII-Protect SSO` app** — identified by its BSN/Cognito footprint (reply URL / Entity ID),
  not by name. It recognises all three shapes (new **SAML enterprise app** / legacy app
  registration / OIDC) and checks what matters for each: SAML SSO mode, assignment-not-required, a
  valid per-app signing certificate, the UPN claims policy (when applicable), reply URL, Entity ID,
  and that the federation metadata is reachable. The signing certificate is resolved by
  **thumbprint** — the one Entra actually signs with, not merely "some valid cert on the app" — and
  it **warns 60 days before expiry**, since when that certificate dies users simply cannot sign in,
  and renewing it also means re-pasting the metadata URL into the BSN portal.
- **Anti-spam allowed senders** — matched by address or covering domain.
- **Catch Phish add-in** — best-effort only, and it says so: no API can confirm this. See
  [Catch Phish](#catch-phish-outlook-add-in).
- **Best-effort**: Directory Sync and Direct Delivery leave a BSN **enterprise app** in the tenant;
  the script lists matching service principals with their **owner tenant + verified publisher**, so
  BSN-owned multi-tenant apps are distinguishable from locally-created ones.

Both `-Verify` and `-Fix` work without a `-ClientName`. `-Fix` connects with **write** scopes (like
a provisioning run) but runs **no manual portal steps**; it runs the same checks as `-Verify` and
offers to remediate the **safe, additive** gaps — asking per gap (y/N; `-NonInteractive` applies
all; `-WhatIf` previews):

- add missing BSN allowed senders,
- create genuinely-missing security groups,
- rename a misnamed SSO app to `PII-Protect SSO`,
- set the SAML enterprise app's assignment-required to No.

It deliberately does **not** rotate signing certificates, change the email claim, or migrate a
legacy SSO app — those can break a working SSO, so they're left to a normal run / `-ReplaceSso`.

### Catch Phish Outlook add-in

The **Catch Phish** add-in is a prompted step like the portal ones: the script links you straight to
**Settings → Integrated apps** and to the [marketplace listing](https://marketplace.microsoft.com/en-us/product/office/wa200008655)
(AssetId `WA200008655`), walks you through the **"Get apps"** flow, then waits.
`-SkipCatchPhish` skips the step and its check.

Catch Phish is often deployed **days or weeks after** the initial onboarding, so it also runs on its
own — no `-ClientName`, no other phases:

```powershell
# That later visit: guided walk-through, then check the result.
./provision-bsn.ps1 -CatchPhishOnly

# Just check whether it's deployed — no prompts, read-only.
./provision-bsn.ps1 -CatchPhishOnly -Verify
```

`-CatchPhishOnly` asks for the narrowest consent of any mode (`Application.Read.All`,
`Directory.Read.All`) — it can't write to the tenant at all, because the deploying happens in the
admin center, not here.

It isn't automated because Office **centralized deployment** has no Graph API. The only programmatic
surface is the `O365CentralizedAddInDeployment` module, which is Windows-only *and* [documented as
Basic-auth only](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins)
("Multifactor authentication isn't supported") — a non-starter under mandatory admin MFA. The
[integrated apps portal is Microsoft's recommended path](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins)
regardless, so a prompt beats a script that couldn't sign in.

> **Assign it to "Everyone".** Centralized deployment doesn't support **mail-disabled security
> groups**, and this script creates `BSN-Employees` / `BSN-Managers` mail-disabled — so you *cannot*
> pick them here. Everyone is the right audience anyway (Catch Phish is a report-phishing button for
> all staff). To scope it narrower you need a mail-enabled group or distribution list, and only
> **top-level** members are assigned — nested groups are not.

#### Use "Get apps", not the "Add-ins" link

The Integrated apps page offers **two** ways to deploy, and they are different surfaces writing to
**different stores**:

| Surface | Method | Visible to `Get-App -OrganizationApp`? |
|---|---|---|
| **Get apps** (Integrated apps) | Microsoft's recommended path — ends in Global Admin consent | **No** |
| **Add-ins** link (older add-in page) | Legacy centralized deployment | Yes |

Use **Get apps** — it's the flow BSN's own article specifies, and its consent step leaves an Entra
enterprise app behind, which is the only thing `-Verify` can reliably see.

> **Source:** BSN's official article *"How to Deploy the Catch Phish Outlook Add-In via Microsoft 365
> Admin Center"* (BSN partner portal). The prompt's steps mirror it. Two caveats it adds: **GCC High
> / GCC Low** tenants need BSN's separate article, and the add-in can take up to **72 hours** to
> appear in Outlook. The "deploy to everyone" advice below is **ours**, not BSN's — their article
> only says "select the users to deploy to".

#### What `-Verify` can and can't tell you

No Graph API exists for centralized deployment, and the one PowerShell module is Windows-only and
can't do MFA — so the check combines two signals, one per surface:

- **Entra service principal** — left behind by the Get apps consent step, so it catches the
  **Integrated apps** path — the one that matters for the Get apps flow above.
- **`Get-App -OrganizationApp`** — only sees **legacy add-in page** deployments, so it covers the
  other surface. Reported only on a hit; a miss here is expected and would be noise.

Because the two signals cover both surfaces, **silence from both is real evidence**, and `-Verify`
reports a genuine gap as a FAIL — but only when Exchange was actually reachable. If it wasn't, only
half the ground was covered and it stays a hedged WARN rather than sending you to redeploy something
that may already be there.

Neither signal reports the **audience**, and consent isn't strictly proof of deployment, so a
positive result still points you at the admin center to confirm who it's assigned to.

> Calibrated across four tenants (2026-07-15): three running the add-in each had the Entra app and
> were invisible to `Get-App`; one with no trace was confirmed empty in both admin-center lists.
> Each deployment also registers its **own** Catch Phish app — the appId differs per tenant, and an
> unverified publisher is normal, so the check reports provenance without grading it.

> Allow up to **72 hours** for the button to reach users' ribbons after deployment (they may need to
> relaunch Outlook). If it's deployed but not appearing, check `AppsForOfficeEnabled` and **EWS** —
> add-in delivery currently depends on EWS being enabled, and it silently fails if it's off.

## Not automated (by design)

`portal.pii-protect.com` exposes no public API, so **tenant creation, Directory Sync, and Direct
Delivery** are consent/UI flows that must be done by a human — the script guides you through each.
The **Catch Phish** add-in is manual for a different reason: centralized deployment has no Graph API
and its module can't do MFA (see above).

## Status

The read-only `-Verify` pass has been run against live tenants. The **provisioning** paths (groups,
anti-spam, and the SSO SAML rewrite) should still be validated with `-WhatIf` and a pilot tenant
before you rely on them. Pure logic (dynamic-rule construction, name sanitization, `-Fix` gating)
is unit-tested.

The **Catch Phish** check has been calibrated against four live tenants (three deployed, one not,
including a known-recent deployment as the anchor) and detects both presence and absence correctly.
It still can't report the assigned audience — nothing but the admin center can.

## License

[MIT](LICENSE) — © 2026 Compudata Systems.

Not affiliated with or endorsed by Breach Secure Now. "PII Protect", "Catch Phish" and "Breach
Secure Now" are their marks; this is an independent tool for partners who deploy their product.
