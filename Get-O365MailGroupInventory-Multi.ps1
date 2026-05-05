<#
.SYNOPSIS
    Multi-tenant wrapper for Get-O365MailGroupInventory.ps1 — runs the v1
    single-tenant inventory across every tenant UOTech has GDAP access to,
    plus any statically-configured non-GDAP tenants, in one pass. Produces
    a per-tenant workbook (identical to v1 output) and a roll-up workbook
    keyed on tenant.

.DESCRIPTION
    STUB. This file lays down the wrapper structure — partner-app auth,
    GDAP customer enumeration, static-config tenant load, the per-tenant
    loop, and the roll-up writer. The auth body, the per-tenant collector
    invocation, and the roll-up writer are all marked TODO. See README.md
    for the architectural decision still open on how v2 calls v1's
    collectors (v1's collection logic is inline procedural code, not
    function-scoped).

    Do NOT run this script as-is. The TODO markers are load-bearing.

.PARAMETER ConfigPath
    Path to tenants.config.json. Defaults to ./tenants.config.json. The
    config supplies partner-app credentials, default skip flags, an
    exclude list, and per-tenant auth overrides for non-GDAP tenants.

.PARAMETER OutputRoot
    Root folder for run artefacts. Defaults to ./output. Per-tenant
    workbooks land in <OutputRoot>/<tenant-shortname>/, and the roll-up
    in <OutputRoot>/_rollup/.

.PARAMETER OnlyTenant
    Optional. If supplied, runs only the matching tenant — matches against
    tenantId, shortName, or displayName (case-insensitive). Useful for
    re-running a single tenant after a transient failure without rerunning
    the whole partner population.

.PARAMETER SkipGdapEnumeration
    If set, only tenants explicitly listed in tenants.config.json are run.
    Useful when partner-app enumeration is failing and you want to fall
    back to the static config.

.PARAMETER SkipRollup
    If set, per-tenant workbooks are produced but the roll-up is not built.

.NOTES
    Author : Mike Maser, UOTech
    Created: 2026-05-05
    Status : Scaffold / stub — not runnable.
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [string] $ConfigPath           = (Join-Path $PSScriptRoot 'tenants.config.json'),
    [string] $OutputRoot           = (Join-Path $PSScriptRoot 'output'),
    [string] $OnlyTenant,
    [switch] $SkipGdapEnumeration,
    [switch] $SkipRollup,
    # Auth mode. Default = delegated: the running user signs in interactively
    # and relies on GDAP/Lighthouse for customer-tenant authorization. Set
    # -AppOnly for unattended/scheduled runs that authenticate via the
    # partner-app cert (loaded from $config.partner.certificatePfxPath).
    [switch] $AppOnly
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Run-state
# ============================================================================

$script:RunStarted   = Get-Date
$script:RunStamp     = $script:RunStarted.ToString('yyyyMMdd_HHmmss')
$script:RunResults   = [System.Collections.Generic.List[object]]::new()
$script:RunWarnings  = [System.Collections.Generic.List[object]]::new()

function Write-MultiLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO')
    $stamp = (Get-Date).ToString('HH:mm:ss')
    $line  = '[{0}] {1,-5} {2}' -f $stamp, $Level, $Message
    if     ($Level -eq 'ERROR') { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq 'WARN')  { Write-Host $line -ForegroundColor Yellow }
    else                        { Write-Host $line }
}

# ============================================================================
# Module / library bootstrap
# ============================================================================

Write-MultiLog 'Bootstrapping modules...'

# v1 (the single-tenant inventory script this wrapper drives) ships its own
# Ensure-Module helper. v1 is NOT bundled with this repo — operators install
# it separately and point at it via $config.v1ScriptPath. Once the v1-collector
# wiring path is picked (see README "Status / open items"), this section
# either dot-sources v1 in library mode, imports a refactored collectors
# module, or does nothing (process-per-tenant orchestration).
#
# TODO: pick wiring path, then either:
#       . $config.v1ScriptPath
#         (with $env:O365INV_LIBRARY_MODE = '1' set first if v1 supports it)
#   OR  Import-Module $config.v1CollectorsModulePath
#   OR  pass — process-per-tenant doesn't need an in-process load.

# ============================================================================
# Config load
# ============================================================================

if (-not (Test-Path $ConfigPath)) {
    throw "Config file not found at $ConfigPath. Copy tenants.config.json (the example) to a real config and populate it."
}

Write-MultiLog "Loading config from $ConfigPath..."
$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json

# Validate the always-required partner fields. cert fields are optional in
# delegated mode (the default) and required only when -AppOnly is set.
if (-not $config.partner -or
    -not $config.partner.homeTenantId -or
    -not $config.partner.clientId) {
    throw "config.partner is missing required fields (homeTenantId, clientId)."
}

# Refuse to run against the schema-example file in version control. Detect
# placeholder values instead of silently iterating <EXAMPLE-…-TENANT-GUID>
# entries as if they were real customers.
$placeholderPattern = '<.*-GUID.*>|<PARTNER-APP-CLIENT-ID-GUID>|<UOTECH-HOME-TENANT-GUID>'
if ($config.partner.clientId -match $placeholderPattern -or
    $config.partner.homeTenantId -match $placeholderPattern) {
    throw "config.partner contains placeholder values (looks like the schema-example tenants.config.json). Copy it to tenants.config.local.json (kept outside git), populate with real values from scripts/Register-PartnerCenterApp.ps1 output, then re-run with -ConfigPath ./tenants.config.local.json."
}

$partnerCert = $null

if ($AppOnly) {
    Write-MultiLog 'Auth mode: APP-ONLY (cert-based). Loading partner cert from PFX...'

    if (-not $config.partner.certificateThumbprint -or -not $config.partner.certificatePfxPath) {
        throw "-AppOnly requires config.partner.certificateThumbprint and certificatePfxPath. Either populate them (re-run scripts/Register-PartnerCenterApp.ps1 with -ConfigPathToUpdate) or omit -AppOnly to use delegated auth."
    }

    $pfxPathExpanded = [System.Environment]::ExpandEnvironmentVariables($config.partner.certificatePfxPath)
    $pfxPathExpanded = $pfxPathExpanded -replace '^~', $HOME
    if (-not (Test-Path $pfxPathExpanded)) {
        throw "Partner cert PFX not found at $pfxPathExpanded. Re-run scripts/Register-PartnerCenterApp.ps1 on this machine, or copy the .pfx into place per the README's pattern A."
    }

    $pfxPwd = $null
    if ($config.partner.certificatePfxPasswordEnvVar) {
        $envName = $config.partner.certificatePfxPasswordEnvVar
        $envVal  = [System.Environment]::GetEnvironmentVariable($envName)
        if ($envVal) { $pfxPwd = ConvertTo-SecureString $envVal -AsPlainText -Force }
    }
    if (-not $pfxPwd) {
        $pfxPwd = Read-Host "Enter PFX password for $pfxPathExpanded" -AsSecureString
    }

    $partnerCert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $pfxPathExpanded,
        $pfxPwd,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    )
    if ($partnerCert.Thumbprint -ne $config.partner.certificateThumbprint) {
        throw "Loaded cert thumbprint ($($partnerCert.Thumbprint)) does not match config.partner.certificateThumbprint ($($config.partner.certificateThumbprint)). Wrong PFX, wrong password, or config drift."
    }
    Write-MultiLog ("Partner cert loaded ({0}, expires {1:yyyy-MM-dd})" -f $partnerCert.Thumbprint, $partnerCert.NotAfter)
} else {
    Write-MultiLog 'Auth mode: DELEGATED. Staff user signs in interactively; GDAP/Lighthouse propagates roles into customer tenants.'
}

$excludeIds = @($config.exclude | Where-Object { $_.tenantId } | ForEach-Object { $_.tenantId.ToLower() })
Write-MultiLog ("Excluding {0} tenant(s) from config.exclude." -f $excludeIds.Count)

# ============================================================================
# Partner-tenant auth (Step 1 of design sketch)
# ============================================================================

Write-MultiLog "Connecting to partner home tenant ($($config.partner.homeTenantId))..."
Write-MultiLog 'STUB: partner-tenant Connect-MgGraph not yet implemented — auth call skipped.' 'WARN'

# TODO: real implementation. Two branches by auth mode:
#
# DELEGATED (default; staff workflow):
#   Connect-MgGraph -TenantId $config.partner.homeTenantId `
#                   -ClientId $config.partner.clientId `
#                   -Scopes   'DelegatedAdminRelationship.Read.All','Directory.Read.All' `
#                   -NoWelcome
#   The user is prompted (browser/device-code) on first run; subsequent
#   runs in the same session reuse cached tokens.
#
# APP-ONLY (-AppOnly switch; unattended/scheduled):
#   Connect-MgGraph -TenantId    $config.partner.homeTenantId `
#                   -ClientId    $config.partner.clientId `
#                   -Certificate $partnerCert `
#                   -NoWelcome
#
# Confirm Get-MgContext returns the right tenant before proceeding.

# ============================================================================
# GDAP customer enumeration (Step 2)
# ============================================================================

$gdapCustomers = @()

if (-not $SkipGdapEnumeration) {
    Write-MultiLog 'Enumerating GDAP customers via /v1.0/tenantRelationships/delegatedAdminCustomers...'
    Write-MultiLog 'STUB: GDAP enumeration not yet implemented — returning empty list.' 'WARN'

    # TODO: real implementation —
    #   $next = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers?$top=100'
    #   while ($next) {
    #       $page = Invoke-MgGraphRequest -Method GET -Uri $next
    #       $gdapCustomers += $page.value
    #       $next = $page.'@odata.nextLink'
    #   }
    #
    # Each customer object has: id (customer tenant id), displayName, defaultDomainName,
    # tenantId, customerId — schema:
    # https://learn.microsoft.com/en-us/graph/api/resources/delegatedadmincustomer

    Write-MultiLog ("GDAP enumeration returned {0} customer(s)." -f $gdapCustomers.Count)
} else {
    Write-MultiLog '-SkipGdapEnumeration set; using only statically-configured tenants.'
}

# ============================================================================
# Tenant target list — merge GDAP results with static config, apply exclusions
# ============================================================================

$tenantTargets = [System.Collections.Generic.List[object]]::new()

# GDAP-discovered tenants. Static-config entries override defaults if they
# match by tenantId — enabling per-tenant skip flags, shortName overrides,
# and auth-mode overrides (e.g. 'gdap' switched to 'app-only-cert' for a
# tenant where the GDAP relationship is unreliable).
foreach ($gd in $gdapCustomers) {
    $tid = ($gd.tenantId, $gd.id, $gd.customerId | Where-Object { $_ } | Select-Object -First 1)
    if (-not $tid) { continue }
    if ($excludeIds -contains $tid.ToLower()) {
        Write-MultiLog "Skipping $tid ($($gd.displayName)) — in config.exclude." 'WARN'
        continue
    }
    $override = $config.tenants | Where-Object { $_.tenantId -and $_.tenantId.ToLower() -eq $tid.ToLower() } | Select-Object -First 1
    $tenantTargets.Add([pscustomobject]@{
        TenantId       = $tid
        DisplayName    = if ($override.displayName) { $override.displayName } else { $gd.displayName }
        ShortName      = if ($override.shortName)   { $override.shortName }   else { ($gd.displayName -replace '[^A-Za-z0-9]+','-').Trim('-').ToLower() }
        PrimaryDomain  = if ($override.primaryDomain) { $override.primaryDomain } else { $gd.defaultDomainName }
        Source         = 'gdap'
        Auth           = if ($override.auth) { $override.auth } else { @{ mode = 'gdap' } }
        Overrides      = $override.overrides
    })
}

# Static-only tenants (not in GDAP enumeration).
foreach ($t in @($config.tenants)) {
    if (-not $t.tenantId) { continue }
    if ($excludeIds -contains $t.tenantId.ToLower()) { continue }
    $alreadyAdded = $tenantTargets | Where-Object { $_.TenantId.ToLower() -eq $t.tenantId.ToLower() } | Select-Object -First 1
    if ($alreadyAdded) { continue }
    $tenantTargets.Add([pscustomobject]@{
        TenantId       = $t.tenantId
        DisplayName    = $t.displayName
        ShortName      = if ($t.shortName) { $t.shortName } else { ($t.displayName -replace '[^A-Za-z0-9]+','-').Trim('-').ToLower() }
        PrimaryDomain  = $t.primaryDomain
        Source         = 'static'
        Auth           = $t.auth
        Overrides      = $t.overrides
    })
}

if ($OnlyTenant) {
    $needle = $OnlyTenant.ToLower()
    $tenantTargets = $tenantTargets | Where-Object {
        $_.TenantId.ToLower()    -eq $needle -or
        $_.ShortName.ToLower()   -eq $needle -or
        $_.DisplayName.ToLower() -eq $needle
    }
    Write-MultiLog ("Filtered to single tenant: {0}" -f $OnlyTenant)
}

Write-MultiLog ("Tenant target count: {0}" -f @($tenantTargets).Count)

# ============================================================================
# Per-tenant loop (Step 3)
# ============================================================================

if (-not (Test-Path $OutputRoot)) { New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null }

foreach ($tenant in $tenantTargets) {
    $tStart = Get-Date
    Write-MultiLog ("─── {0} ({1}) — {2} ───" -f $tenant.DisplayName, $tenant.ShortName, $tenant.TenantId)
    Write-MultiLog 'STUB: per-tenant connect + collect not yet implemented — no inventory data captured for this tenant.' 'WARN'

    $tenantOutDir = Join-Path $OutputRoot $tenant.ShortName
    if (-not (Test-Path $tenantOutDir)) { New-Item -ItemType Directory -Path $tenantOutDir -Force | Out-Null }
    $tenantXlsx   = Join-Path $tenantOutDir ("O365-MailGroupInventory_{0}.xlsx" -f $script:RunStamp)

    $tenantStatus = 'pending'
    $tenantWarnings = @()

    try {
        # ----- Connect (mode-dispatched per tenants.config.json auth block) -----
        switch ($tenant.Auth.mode) {
            'gdap' {
                # TODO — auth-mode-aware dispatch:
                #
                # DELEGATED (default; $AppOnly = $false):
                #   The running user has GDAP delegated rights in the
                #   customer tenant via security-group membership. Acquire a
                #   delegated token for the customer tenant under our
                #   partner-app's clientId; GDAP propagates the user's role.
                #
                #     Connect-MgGraph -TenantId $tenant.TenantId `
                #                     -ClientId $config.partner.clientId `
                #                     -Scopes   $script:DelegatedScopesForCustomer `
                #                     -NoWelcome
                #
                #     # EXO uses first-party GDAP delegated auth — no app
                #     # involvement, just the customer's domain.
                #     Connect-ExchangeOnline -DelegatedOrganization $tenant.PrimaryDomain `
                #                            -ShowBanner:$false
                #
                # APP-ONLY ($AppOnly = $true):
                #   Cert-based, app acts on its own behalf. Requires the
                #   customer tenant to authorize this app's SP either by
                #   manual consent or via a GDAP role template that
                #   includes app-management propagation.
                #
                #     Connect-MgGraph -TenantId    $tenant.TenantId `
                #                     -ClientId    $config.partner.clientId `
                #                     -Certificate $partnerCert `
                #                     -NoWelcome
                #     Connect-ExchangeOnline -DelegatedOrganization $tenant.PrimaryDomain `
                #                            -AppId                 $config.partner.clientId `
                #                            -Certificate           $partnerCert `
                #                            -ShowBanner:$false
            }
            'app-only-cert' {
                # TODO: Connect-MgGraph / Connect-ExchangeOnline using the in-tenant
                #       app reg's clientId + certificateThumbprint from $tenant.Auth.
            }
            'stored-credential' {
                # TODO: resolve secret via Microsoft.PowerShell.SecretManagement
                #   $sec  = Get-Secret -Vault $tenant.Auth.secretManagement.vaultName `
                #                      -Name  $tenant.Auth.secretManagement.secretName
                #   $cred = [pscredential]::new($tenant.Auth.userPrincipalName, $sec)
                # then Connect-MgGraph + Connect-ExchangeOnline with -Credential.
                # Caveat: requires CA + MFA exemption on the admin account, or use app-only-cert.
            }
            'interactive' {
                # TODO: Connect-MgGraph -TenantId $tenant.TenantId -Scopes ... -NoWelcome
                # plus Connect-ExchangeOnline -UserPrincipalName $tenant.Auth.userPrincipalName.
            }
            default {
                throw "Unknown auth.mode '$($tenant.Auth.mode)' for tenant $($tenant.DisplayName)"
            }
        }

        # ----- Run v1 collectors against this tenant -----
        # TODO (architectural decision — see README):
        #   Path 1: dot-sourced & refactored v1 → call each Invoke-O365InventoryCollect-*
        #           function and capture row arrays into $collected.
        #   Path 2: library-mode v1 → call $collected = Invoke-O365InventoryCollect.
        #   Path 3: process-per-tenant → spawn `pwsh -File $config.v1ScriptPath
        #           -OutputPath $tenantXlsx` with appropriate env vars; parse $tenantXlsx after.
        #
        # In Paths 1 and 2, after collection: write $collected to $tenantXlsx using v1's
        # exact Export-Excel block (factor it out into a Write-O365InventoryWorkbook helper).

        $tenantStatus = 'ok'
    }
    catch {
        $tenantStatus = 'error'
        $tenantWarnings += $_.Exception.Message
        Write-MultiLog ("Tenant '{0}' failed: {1}" -f $tenant.DisplayName, $_.Exception.Message) 'ERROR'
        $script:RunWarnings.Add([pscustomobject]@{
            TenantId    = $tenant.TenantId
            DisplayName = $tenant.DisplayName
            Severity    = 'ERROR'
            Message     = $_.Exception.Message
        })
    }
    finally {
        # ----- Disconnect (best-effort; never let teardown abort the loop) -----
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        try { Disconnect-MgGraph                        -ErrorAction SilentlyContinue } catch { }
        # MicrosoftTeams disconnect — only if v1 connected it; the wiring path will decide.
    }

    $tEnd = Get-Date
    $script:RunResults.Add([pscustomobject]@{
        TenantId      = $tenant.TenantId
        DisplayName   = $tenant.DisplayName
        ShortName     = $tenant.ShortName
        PrimaryDomain = $tenant.PrimaryDomain
        Source        = $tenant.Source
        AuthMode      = $tenant.Auth.mode
        Status        = $tenantStatus
        DurationSec   = [int]($tEnd - $tStart).TotalSeconds
        OutputPath    = if ($tenantStatus -eq 'ok') { $tenantXlsx } else { $null }
        WarningCount  = @($tenantWarnings).Count
    })
}

# ============================================================================
# Roll-up workbook (Step 4)
# ============================================================================

if (-not $SkipRollup) {
    $rollupDir  = Join-Path $OutputRoot '_rollup'
    if (-not (Test-Path $rollupDir)) { New-Item -ItemType Directory -Path $rollupDir -Force | Out-Null }
    $rollupXlsx = Join-Path $rollupDir ("Multi-Tenant-Rollup_{0}.xlsx" -f $script:RunStamp)

    Write-MultiLog "Building roll-up workbook at $rollupXlsx..."

    # TODO: build the roll-up tabs. One row per tenant on each headline tab.
    # Source data depends on the wiring path:
    #   Paths 1/2: roll-up is built from the in-memory $collected hashtables we
    #              kept around per tenant — flatten the headline counts into rows.
    #   Path 3:    roll-up is built by reading each per-tenant xlsx with
    #              Import-Excel, projecting the Summary-tab counts into a row.
    #
    # Headline tabs to produce:
    #   - Run        : one row per tenant — TenantId, DisplayName, ShortName, Source, AuthMode, Status, DurationSec, OutputPath
    #   - MFA        : per tenant — admin count, admins MFA-registered, non-MFA admins,
    #                  admin-MFA coverage %, distinct human admins, GA count, GA non-MFA count.
    #   - Admin      : per tenant — total admins, eligible-PIM admins, role-holding service principals,
    #                  role-holding groups, break-glass candidates flagged.
    #   - CA         : per tenant — total policies, enabled, report-only, disabled,
    #                  named locations, trusted named locations, auth-strength policies.
    #   - Capacity   : per tenant — user-mailbox count, total mailbox size GB,
    #                  M365 group count, Teams count, SharePoint storage GB.
    #   - License    : per tenant — top SKUs by assignment count.
    #   - External-Sender : per tenant — accept-from-external CA presence,
    #                       external-tag enabled, anti-impersonation policies count.
    #   - Run Issues : flattened from $script:RunWarnings.
    #
    # Use the same xlsx style helpers v1 uses (FreezeTopRow, AutoFilter, BoldTopRow, AutoSize).
}

# ============================================================================
# Wrap-up
# ============================================================================

$elapsed = (Get-Date) - $script:RunStarted
$ok      = @($script:RunResults | Where-Object Status -eq 'ok').Count
$err     = @($script:RunResults | Where-Object Status -eq 'error').Count
Write-MultiLog ("Done. Tenants attempted: {0}; ok: {1}; error: {2}; elapsed: {3:hh\:mm\:ss}" -f
    @($script:RunResults).Count, $ok, $err, $elapsed)

# Per-run summary CSV alongside the rollup, useful for diffing runs.
$summaryCsv = Join-Path $OutputRoot ("run-summary_{0}.csv" -f $script:RunStamp)
$script:RunResults | Export-Csv -NoTypeInformation -Path $summaryCsv
Write-MultiLog "Per-run summary at $summaryCsv"
