# provision-bsn — working conventions for contributors and coding agents

## Sourcing (this bit matters most)

**Only Breach Secure Now's own official articles are a source of truth for what BSN requires.** A
reseller's or distributor's rebranded copy is not, even when it's the only thing publicly indexed
and even when it turns out to match. Microsoft Learn is authoritative for how Entra / M365 behave —
never for what BSN wants.

BSN's technical articles are behind the `portal.pii-protect.com` partner login and are not publicly
reachable, so **ask the maintainer to export the current article** rather than substituting something
that looks close. Exported PDFs are **deliberately not committed**: they stale, and a stale copy in
the repo is worse than no copy. Always work from a freshly exported article.

Never write "BSN documents X" unless BSN's own article was actually read. Keep our recommendations
visibly separate from BSN's instructions (e.g. Catch Phish "deploy to everyone" is ours — BSN's
article only says "select the users to deploy to").

Some BSN articles are marked **Confidential — not to be redistributed**. This repo is public: record
*which end state* an article requires (a configuration fact the code must build), never reproduce the
article's walkthrough.

## Don't publish client data

This repo is public. No client names, tenant ids, Cognito pool names, or Entity IDs in code, docs, or
comments — describe shapes ("four live tenants", "the tenant without the add-in") instead.

## Stay vendor-agnostic outside of BSN

The tool automates **Breach Secure Now on Microsoft 365** — that's the scope. Everything else about
the maintainer's stack is incidental and must not leak into prompts, help, or docs: no PSA
(Halo/Autotask/ConnectWise), no RMM, no billing system, no internal process names. Anchor wording to
something the operator can see on screen (e.g. "as it should appear in the PII Protect portal"), not
to a system only we happen to run. Users of this tool have different stacks.

## Code

- The comment-based help block must come **before** `#Requires`, or `Get-Help` falls back to
  auto-generated syntax.
- All Graph goes through `Invoke-MgGraphRequest` (helpers `Graph-Get/Post/Patch`), so
  `Microsoft.Graph.Authentication` stays the only hard dependency. Modules auto-install via
  `Ensure-Module` in the Prerequisites step up front — never mid-run.
- `-Verify` uses read-only scopes; `-CatchPhishOnly` uses the narrowest scopes of any mode and must
  stay read-only. Only ever wire `Invoke-Fix` to **safe, additive** remediations.
- Everything honours `-WhatIf` and must be safe to re-run.

## Claims and evidence

- **Don't assume — surface gaps and ask.** (Maintainer's explicit preference.)
- Prefer empirical evidence over documentation inference: `Get-App -OrganizationApp` is documented as
  listing org add-ins, but Integrated-apps deployments genuinely aren't there. Four live tenants beat
  a confident reading of a doc.
- Detect BSN apps by **traits, not appIds** — and for Catch Phish, by display name only: its appId
  differs per tenant and its publisher is often unverified.
- Graph can't be live-tested from the maintainer's box. Parse-check (`[Parser]::ParseFile`),
  unit-test pure logic, and drive functions with stubbed Graph/Exchange; the maintainer validates
  with `-WhatIf` and a pilot tenant.

## Commits

**No `Co-Authored-By` trailer in this repo** (maintainer's rule). Commit as
`Stan Clemance <stan@compudata.ca>`.
