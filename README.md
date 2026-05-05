# m365-multi-tenant-inventory

A PowerShell 7 wrapper that runs a single-tenant Microsoft 365 posture inventory across every customer a Cloud Solution Provider has GDAP access to, in one pass. Produces per-tenant Excel workbooks plus a cross-tenant rollup keyed on tenant id.

The wrapper is auth + orchestration only. The actual collection script (the per-tenant inventory tool) is **not bundled** with this repo. You install it separately and point the wrapper at it via configuration.

## Status

**Scaffold.** Wrapper structure, registration helper, configuration schema, and interactive setup are in place. The per-tenant collector call-site is intentionally a TODO pending the architectural decision on how the wrapper invokes the upstream single-tenant script (three workable paths described in [Architecture](#architecture-collector-wiring)). Until that's wired, the wrapper enumerates tenants and writes a rollup, but per-tenant runs are no-ops with explicit `[WARN] STUB:` log lines.

## Requirements

- **PowerShell 7+** on Windows, macOS, or Linux. Verify with `pwsh -v`. The bundled scripts auto-install the `Microsoft.Graph.*`, `ExchangeOnlineManagement`, `MicrosoftTeams`, and `ImportExcel` modules into your `CurrentUser` scope on first run.
- **A single-tenant M365 inventory script** (PowerShell 7+) that accepts `-OutputPath` and produces a 16-tab `.xlsx`. The wrapper drives it; you choose how it's wired in. See [v1 dependency](#v1-dependency).
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
- Generates a self-signed cert (2048-bit RSA, SHA-256, 2-year validity) and writes a password-protected `.pfx` to `~/.uotech/m365-multi-tenant-inventory/partner-app.pfx` by default. You set the password when prompted; never written to a config file.
- Attaches the cert as a key credential, creates the service principal, and grants admin consent programmatically tenant-wide (`appRoleAssignment` for application permissions, `oauth2PermissionGrant` with `consentType=AllPrincipals` for delegated). No browser admin-consent step.
- Prints `clientId`, `tenantId`, cert thumbprint, and PFX path. Optional `-ConfigPathToUpdate` writes them straight into your local config.

The script is idempotent and re-entrant. Re-running it patches an existing app rather than creating a duplicate; running it on a second operator machine appends a new cert (or reuses an existing one if the PFX is shared) without overwriting key credentials.

### Set up per-machine configuration

Each operator runs `scripts/Setup-LocalConfig.ps1` once. It prompts for `clientId` and `homeTenantId` (always required), validates the GUIDs, and writes `tenants.config.local.json`. Pass `-AppOnly` to also collect the cert thumbprint, PFX path, and PFX-password env-var name needed for unattended runs.

```powershell
# Default (delegated mode; what most operators need)
./scripts/Setup-LocalConfig.ps1

# With cert handling for -AppOnly use
./scripts/Setup-LocalConfig.ps1 -AppOnly

# Non-default config destination (kept entirely outside the repo)
./scripts/Setup-LocalConfig.ps1 -ConfigDestination "$HOME/your-org/m365-multi/tenants.config.local.json"
```

Existing values are shown as defaults on re-run, so fixing one field doesn't require re-typing the rest. Placeholder-looking values (`<PARTNER-APP-CLIENT-ID-GUID>`, etc.) are rejected.

### Wire up your single-tenant inventory script

Edit `tenants.config.local.json` and set `v1ScriptPath` to the path of your single-tenant inventory script. The wrapper expects:

- A PowerShell 7+ script, executable as `pwsh -File <path>`.
- Accepts `-OutputPath <xlsx>` and writes a 16-tab `.xlsx` of M365 posture data.
- Does its own auth (the wrapper handles per-tenant authentication separately and drops the script into a pre-authenticated context, OR spawns it as a child process; depends on which collector-wiring path you pick).

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

### Other useful flags

| Flag | Effect |
|---|---|
| `-OnlyTenant <name-or-guid>` | Run a single tenant. Matches case-insensitively against `tenantId`, `shortName`, or `displayName`. Useful for retrying a failed tenant. |
| `-SkipGdapEnumeration` | Skip the partner-tenant enumeration; use only tenants explicitly listed in your config. Useful for testing or fallback. |
| `-SkipRollup` | Produce per-tenant workbooks but skip the rollup. |
| `-OutputRoot <path>` | Override where workbooks land. Default: `./output/`. |

## Configuration

`tenants.config.json` is the schema example checked into the repo with placeholders. `tenants.config.local.json` (gitignored) is the real config the wrapper reads. Don't edit the schema example directly; copy via `Setup-LocalConfig.ps1`.

Top-level structure:

```json
{
  "v1ScriptPath": "<path to your single-tenant inventory script>",
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

## v1 dependency

The single-tenant inventory script is intentionally external. This repo ships only the wrapper. Reasons:

- The single-tenant script may be your own, an internally-vendored one, or one you've licensed from elsewhere.
- Decoupling the wrapper from any specific single-tenant implementation means upgrading either independently is a path-config change, not a re-vendor.
- Keeping the wrapper repo small and free of someone else's licensed code keeps the legal posture clean.

What the wrapper expects of the v1 script:

- PowerShell 7+ entry point (`.ps1`).
- Accepts `-OutputPath <xlsx>` and writes a single workbook with 16 tabs (Summary, Shared Mailboxes, User Mailboxes, Resource Mailboxes, Distribution Groups, Mail-Enabled Sec Groups, M365 Groups, Teams, Security Groups (cloud), CA Policies, CA Named Locations, CA Auth Strengths, Auth Methods Policy, Directory Roles, Admin Users, Admin Service Principals, Break-Glass Candidates, Run Log).
- Either runs end-to-end as a child process (process-per-tenant orchestration), exposes its collectors as importable functions, or supports a library mode that produces the same row arrays the wrapper writes itself.

If you don't have a single-tenant script yet, open-source M365 posture-inventory scripts exist on GitHub; pick one that produces the 16-tab schema and point `v1ScriptPath` at it.

## Architecture: collector wiring

How the wrapper invokes the single-tenant script per customer tenant is an open architectural decision. Three workable paths:

| Path | What changes in v1 | Wrapper behavior | Trade-off |
|---|---|---|---|
| 1. **Refactor v1 into named collector functions** | v1's per-section logic becomes `Invoke-O365InventoryCollect-<Section>` functions returning row arrays | Dot-source v1, call each collector, build workbook in-process | Cleanest long-term. Largest v1-side change. |
| 2. **Library-mode env-var guard** | One-line guard at the top of v1: `if ($env:M365INV_LIBRARY_MODE) { return }` after function defs | Wrapper sets the env var, dot-sources v1 to load helpers, then duplicates the inline collection logic in v2 (parameterised on tenant context) | Smallest v1 diff. v2 carries a copy of collection blocks → drift risk between versions. |
| 3. **Process-per-tenant orchestration** | Nothing in v1 | Wrapper spawns `pwsh -File $config.v1ScriptPath -OutputPath ...` per tenant, reads the resulting xlsx files for the rollup | Zero v1 risk. ~few-second process-spawn cost per tenant. Rollup parses xlsx instead of using in-memory rows. |

Recommendation: path 3 as immediate stopgap; path 1 long-term once the v1 source can absorb the refactor.

## Repository structure

```
.
├── .gitignore
├── LICENSE
├── README.md                                  ← you are here
├── Get-O365MailGroupInventory-Multi.ps1       ← the wrapper
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
- v1 (the single-tenant collector) is intentionally not vendored here. PRs adding a vendor copy will be declined unless they bring their own license + attribution and stay scoped under `vendor/<source>/`.

## License

Apache-2.0. See [LICENSE](./LICENSE).
