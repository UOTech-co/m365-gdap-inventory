# m365-multi-tenant-inventory

A PowerShell 7 toolkit that runs a Microsoft 365 posture inventory across every customer a Cloud Solution Provider has GDAP access to, in one pass. Produces per-tenant Excel workbooks plus a cross-tenant rollup keyed on tenant id. The wrapper enumerates GDAP customers from your partner tenant and drives the bundled per-tenant collector against each one as a child process. Built exclusively for GDAP-driven multi-tenant inventory; not intended for standalone single-tenant use.

## What's in the box

- `Get-O365MailGroupInventory-Multi.ps1` — the wrapper. Authenticates against your partner tenant, enumerates GDAP customers, drives the per-tenant collector per customer, builds the rollup.
- `Get-O365MailGroupInventory.ps1` — the per-tenant collector. Invoked once per customer by the wrapper as a child process. Connects to the target customer tenant via GDAP-delegated Microsoft Graph + Exchange Online (`-DelegatedOrganization`) + Microsoft Teams and writes a 16-tab `.xlsx` of posture data (mailboxes, groups, Teams, Conditional Access, directory roles, admin users, break-glass candidates). Not designed to be run directly — fails fast if `-TenantId` and `-DelegatedOrganization` aren't supplied.
- `scripts/Register-PartnerCenterApp.ps1` — one-time partner-app registration with the right permissions and admin consent applied programmatically.
- `scripts/Setup-LocalConfig.ps1` — interactive config setup. Run once per operator machine.
- `scripts/Test-TenantPreflight.ps1` — pre-flight validation. Silently tests every GDAP customer tenant and prints a punchlist of the ones that need admin consent or are blocked by Conditional Access, with the consent URL inline.

## Status

Working end-to-end. Process-per-tenant orchestration: the wrapper handles partner-tenant auth + GDAP enumeration, then spawns the collector as a child process per customer with `-TenantId`, `-ClientId`, and `-DelegatedOrganization` set per customer. The collector does its own Connect-* calls inside the child process; MSAL token caching means only the first customer prompts for sign-in.

## Requirements

- **PowerShell 7+** on Windows, macOS, or Linux. Verify with `pwsh -v`. The bundled scripts auto-install the `Microsoft.Graph.*`, `ExchangeOnlineManagement`, `MicrosoftTeams`, and `ImportExcel` modules into your `CurrentUser` scope on first run.
- **A Microsoft Cloud Solution Provider relationship** with at least one customer that has granted GDAP. Without that, the GDAP enumeration step returns zero customers and there's nothing to inventory.
- **An Entra app registration** in your partner tenant. The included `scripts/Register-PartnerCenterApp.ps1` handles this end-to-end.
- **For the operator running the registration**: Global Administrator (or Privileged Role Admin / Cloud App Admin) in your partner tenant. Required once, when the app is registered. Day-to-day operators can be standard users with GDAP-delegated rights.

## Quick start

```powershell
# 1. Clone
git clone <your-fork-url> m365-multi-tenant-inventory
cd m365-multi-tenant-inventory

# 2. One-time partner-app registration (run by a Global Admin in your home tenant)
./scripts/Register-PartnerCenterApp.ps1

# 3. One-time per-machine config setup (interactive prompts; writes to tenants.config.local.json)
./scripts/Setup-LocalConfig.ps1

# 4. Day-to-day run
./Get-O365MailGroupInventory-Multi.ps1 -ConfigPath ./tenants.config.local.json
```

First run prompts for interactive sign-in once. Subsequent runs use cached refresh tokens; no popup per tenant.

## One-time setup

### Register the partner app

Run `scripts/Register-PartnerCenterApp.ps1` once, signed in as a Global Administrator in your home tenant. The script:

- Creates the Entra app with `signInAudience = AzureADMultipleOrgs` (multi-tenant; required so customer tenants accept tokens for this clientId).
- Adds 14 Microsoft Graph delegated permissions for the staff workflow + 14 Microsoft Graph application permissions and `Exchange.ManageAsApp` for the unattended `-AppOnly` workflow.
- Generates a self-signed cert (2048-bit RSA, SHA-256, 2-year validity) and writes a password-protected `.pfx` to `~/.m365-multi-tenant-inventory/partner-app.pfx` by default. You set the password when prompted; never written to a config file.
- Attaches the cert as a key credential, creates the service principal, and grants admin consent programmatically tenant-wide (`appRoleAssignment` for application permissions, `oauth2PermissionGrant` with `consentType=AllPrincipals` for delegated). No browser admin-consent step.
- Prints `clientId`, `tenantId`, cert thumbprint, and PFX path. Optional `-ConfigPathToUpdate` writes them straight into your local config.

The script is idempotent and re-entrant. Re-running it patches an existing app rather than creating a duplicate; running it on a second operator machine appends a new cert (or reuses an existing one if the PFX is shared) without overwriting key credentials.

### Set up per-machine configuration

Each operator runs `scripts/Setup-LocalConfig.ps1` once. The script auto-discovers as much as it can before prompting, so most operators just press Enter through the prompts.

```powershell
# Default (delegated mode; what most operators need)
./scripts/Setup-LocalConfig.ps1

# With cert handling for -AppOnly use (extra prompts: cert thumbprint, PFX path, env var name)
./scripts/Setup-LocalConfig.ps1 -AppOnly

# Skip the auto-discovery sign-in (faster if you know your tenant won't allow it)
./scripts/Setup-LocalConfig.ps1 -NoAutoDiscover
```

What happens when you run it:

1. **Auto-discovery sign-in.** A browser opens to `login.microsoftonline.com`. Sign in with your tenant account; the script reads your home tenant id from the resulting Microsoft Graph context and queries app registrations to find the partner app.
2. **Prompts with auto-filled defaults.** Each prompt shows the auto-discovered (or existing-config) value in square brackets, e.g. `Partner-app clientId (GUID) [00000000-...]:`. **Press Enter to accept the bracketed value, or type a different value to override.** Same convention for every subsequent prompt.
3. **Config written.** `tenants.config.local.json` lands at the repo root (gitignored). The cert fields are populated only if you ran with `-AppOnly`.

Re-running is safe and idempotent. Existing config values take precedence over auto-discovered ones, so fixing one field doesn't require re-typing the rest. Placeholder-looking values (`<PARTNER-APP-CLIENT-ID-GUID>`, etc.) are rejected outright.

If auto-discovery fails (sign-in declined, `Application.Read.All` not consented in your tenant, or zero/many app candidates) the script falls back to manual prompts cleanly. The first time the script runs in a tenant where the partner app hasn't been registered yet, expect this fallback; run `Register-PartnerCenterApp.ps1` first or have someone else hand you the values.

## Usage

### Default (delegated)

The day-to-day operator workflow. Standard users with GDAP-delegated rights to customer tenants run the wrapper directly:

```powershell
./Get-O365MailGroupInventory-Multi.ps1 -ConfigPath ./tenants.config.local.json
```

First run prompts for interactive sign-in once. Customer tenants are authenticated using cached refresh tokens that GDAP/Lighthouse propagates through; subsequent customer tenants in the same run usually go silent.

### Unattended / scheduled (`-AppOnly`)

For cron / launchd / Task Scheduler runs without a human at the keyboard. Cert-based, app authenticates as itself.

```powershell
$env:M365_MULTI_PFX_PASSWORD = '<your-pfx-password>'
./Get-O365MailGroupInventory-Multi.ps1 -ConfigPath ./tenants.config.local.json -AppOnly
```

`-AppOnly` requires `certificateThumbprint` and `certificatePfxPath` populated in the local config. The PFX password is resolved from the env var named in `certificatePfxPasswordEnvVar` (default `M365_MULTI_PFX_PASSWORD`); if the env var is not set, falls back to an interactive prompt.

**Caveat.** App-only auth against customer tenants requires the partner-app's service principal to be authorized in each customer tenant. With standard GDAP role grants (which assign Entra roles to user security groups, not to the partner-app SP), `-AppOnly` will fail at the per-customer connect step until you either (a) extend GDAP role templates to grant app-management to the partner-app SP, or (b) each customer's admin manually consents the partner-app SP. Plan delegated mode as the working path; treat `-AppOnly` as future capability.

### Pre-flight validation

Before a long multi-tenant run, you can validate that every GDAP customer tenant is reachable and properly consented. The pre-flight script does a single device-code sign-in to your partner tenant, then silently tests each customer tenant's `/token` endpoint and prints a punchlist of the tenants that need attention:

```powershell
./scripts/Test-TenantPreflight.ps1
```

Each problem tenant comes back with a category and a fix. The most common ones:

- **`NEEDS_CONSENT (no SP)`** / **`NEEDS_CONSENT (scopes)`** — the customer tenant has no service principal for Microsoft Graph PowerShell, or the SP exists but the requested scopes aren't consented. Send the customer admin the printed `/adminconsent` URL — one click and they're done. (Heads-up: Microsoft drops them on a "This is not the right page" / phishing-warning page after they accept. That's the normal endpoint for first-party admin-consent flows; consent did land.)
- **`BLOCKED_BY_CA (device-code only)`** — the customer's "Authentication Flows Policy" CA control blocks device code flow. Pre-flight uses device code flow so it can't validate them, but the actual wrapper uses interactive browser auth and may still work. Try a `-OnlyTenant` run against the tenant; only escalate to the customer admin if that also fails.
- **`NO_ACCESS (GDAP)`** — your account isn't in the customer tenant. The GDAP relationship has lapsed.

Re-test a single tenant after consent with `./scripts/Test-TenantPreflight.ps1 -OnlyTenant '<name-or-guid>'`.

A clean pre-flight run prints a one-line success message and writes nothing to disk.

### Per-tenant confirmation prompt

By default, the wrapper prompts before each customer tenant with a 5-second auto-Y countdown:

```
Process 'Acme Corp' (acme.onmicrosoft.com)? [Y/n/q]  (auto-Y in  4s)...
```

The `delegatedAdminCustomers` API includes customer-of-record relationships from offboarded customers; this prompt lets you skip stale ones at runtime instead of polluting the rollup. Press `Y` or Enter to process, `N` to skip (the tenant lands in the run summary with `Status='skipped'`), or `Q` to stop the run entirely (the rollup still gets built from tenants processed up to that point). Any other key accepts the default. Wait 5 seconds and the prompt auto-Y's.

For unattended runs, pass `-NoConfirm`. The prompt is also auto-skipped when `-OnlyTenant` targets a single tenant and when stdin is redirected.

### Other useful flags

| Flag | Effect |
|---|---|
| `-OnlyTenant <name-or-guid>` | Run a single tenant. Matches case-insensitively against `tenantId`, `shortName`, or `displayName`. Useful for retrying a failed tenant. |
| `-SkipGdapEnumeration` | Skip the partner-tenant enumeration; use only tenants explicitly listed in your config. Useful for testing or fallback. |
| `-SkipRollup` | Produce per-tenant workbooks but skip the rollup. |
| `-NoConfirm` | Skip the per-tenant Y/n/q prompt entirely. Required for unattended runs. |
| `-ConfirmTimeoutSec <n>` | Per-tenant prompt timeout. Default 5. Set to 0 for no wait (same effect as `-NoConfirm`). |
| `-OutputRoot <path>` | Override where workbooks land. Default: `./output/`. |

## Configuration

`tenants.config.json` is the schema example checked into the repo with placeholders. `tenants.config.local.json` (gitignored) is the real config the wrapper reads. Don't edit the schema example directly; copy via `Setup-LocalConfig.ps1`.

Top-level structure:

```json
{
  "collectorScriptPath": "<optional override path; defaults to ./Get-O365MailGroupInventory.ps1>",
  "partner": {
    "homeTenantId":          "<HOME-TENANT-ID-GUID>",
    "clientId":              "<PARTNER-APP-CLIENT-ID-GUID>",
    "certificateThumbprint": null,
    "certificatePfxPath":    null,
    "certificatePfxPasswordEnvVar": "M365_MULTI_PFX_PASSWORD"
  },
  "defaults": { ... per-tenant skip flags ... },
  "exclude":  [ ... tenant ids to skip ... ],
  "tenants":  [ ... per-tenant overrides ... ]
}
```

`partner.homeTenantId` and `partner.clientId` are always required. The cert fields are required only with `-AppOnly`. The per-tenant overrides are optional; the default behavior enumerates every GDAP customer with no overrides needed.

## Auth modes

The wrapper supports two auth modes against the same Entra app registration. Both are added side-by-side at registration time so flipping between them needs no app-side changes.

| Mode | When | What it does |
|---|---|---|
| **Delegated** (default) | Day-to-day operator runs | User signs in interactively against the partner-app's `clientId`. GDAP/Lighthouse propagates the user's role assignments per customer tenant. No cert involved on the operator machine. |
| **App-only** (`-AppOnly`) | Cron / launchd / Task Scheduler | Cert-based. App authenticates as itself using the PFX. Subject to the SP-level GDAP grant caveat above. |

## Collector skip flags

The collector accepts skip flags when sections of the inventory are too expensive or aren't applicable. Configure them in your local config under `defaults` (see Configuration above) and the wrapper threads them through to every per-tenant invocation.

| Flag | What it skips |
|---|---|
| `-SkipMailboxStats` | Per-mailbox size / item-count / last-logon. Speeds the run on tenants with many mailboxes. |
| `-SkipPermissions` | FullAccess / SendAs / SendOnBehalf delegate collection. |
| `-SkipUserMailboxes` | The User Mailboxes sheet. |
| `-SkipSharePointStats` | SharePoint URL + drive-quota lookups on M365 Groups and Teams. |
| `-SkipConditionalAccess` | All four CA-related sheets. |
| `-SkipAdminPosture` | Directory Roles + Admin Users + Admin Service Principals + Break-Glass Candidates sheets. |

## How it works

Process-per-tenant orchestration. The multi-tenant wrapper:

1. Authenticates against your partner tenant via `Connect-MgGraph` (delegated by default; cert-based with `-AppOnly`).
2. Enumerates GDAP customers via `GET /v1.0/tenantRelationships/delegatedAdminCustomers` (paginated).
3. Merges in any statically-configured tenants from `tenants.config.local.json`.
4. For each customer, spawns the collector as a child process with `-TenantId <customer>`, `-ClientId <partner-app>`, and `-DelegatedOrganization <customer-domain>`. The child process does its own auth using GDAP/Lighthouse delegation; MSAL token caching means only the first customer prompts for sign-in.
5. Captures the per-tenant xlsx, status, duration, and warnings.
6. After all customers finish, builds a rollup workbook by reading each per-tenant Summary sheet and projecting the headline counts into one row per tenant. Writes a per-run summary CSV alongside.

Per-tenant runtime is 3–4 minutes; an end-to-end run across ~30 GDAP customers finishes in 1.5–2 hours. Sequential by default; the wrapper does not parallelise across tenants (rate-limit pressure on Graph + EXO would be the bottleneck if it did).

## Repository structure

```
.
├── .gitignore
├── LICENSE
├── README.md                                  ← you are here
├── Get-O365MailGroupInventory.ps1             ← per-tenant collector
├── Get-O365MailGroupInventory-Multi.ps1       ← multi-tenant wrapper
├── tenants.config.json                        ← schema example (with placeholders)
└── scripts/
    ├── Register-PartnerCenterApp.ps1          ← one-time partner-app registration
    └── Setup-LocalConfig.ps1                  ← per-machine interactive config
```

Per-run output (gitignored):

```
output/
├── <tenant-shortname>/
│   └── O365-MailGroupInventory_<UTC-timestamp>.xlsx
├── _rollup/
│   └── Multi-Tenant-Rollup_<UTC-timestamp>.xlsx
└── run-summary_<UTC-timestamp>.csv
```

## Contributing

Issues and pull requests welcome. A few conventions:

- Match the existing PowerShell style: verbose comments, defensive error handling, idempotent re-runs.
- Don't commit real tenant ids, client ids, or cert thumbprints; the schema example uses `<…>` placeholders for a reason.
- The two scripts share a name prefix on purpose. Changes to the collector that break its `-TenantId / -ClientId / -DelegatedOrganization` contract also need a wrapper update; please bundle them in one PR.

## License

Apache-2.0. See [LICENSE](./LICENSE).
