<#
.SYNOPSIS
    Multi-tenant wrapper for Get-O365MailGroupInventory.ps1 — runs the
    bundled single-tenant inventory script across every tenant a Cloud
    Solution Provider has GDAP access to, plus any statically-configured
    non-GDAP tenants, in one pass. Produces a per-tenant workbook
    (identical to the single-tenant output) and a roll-up workbook keyed
    on tenant.

.DESCRIPTION
    Process-per-tenant orchestration: the wrapper handles partner-app auth
    + GDAP enumeration, then spawns the bundled single-tenant script as a
    child process per customer. v1 does its own connect (delegated GDAP-
    aware Connect-MgGraph + Connect-ExchangeOnline -DelegatedOrganization)
    using the multi-tenant params the wrapper passes in. After all tenants
    finish, the wrapper builds a roll-up workbook by reading each per-tenant
    Summary tab and projecting the headline counts into one row per tenant.

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
    License: Apache-2.0 (see LICENSE in repo root)
    Created: 2026-05-05
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
    [switch] $AppOnly,

    # Path to the single-tenant inventory script invoked per customer. Default
    # is the bundled Get-O365MailGroupInventory.ps1 next to this script. Can be
    # overridden via param OR config.v1ScriptPath.
    [string] $V1ScriptPath = (Join-Path $PSScriptRoot 'Get-O365MailGroupInventory.ps1')
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

# v1 (the bundled single-tenant inventory script this wrapper drives) handles
# its own module bootstrap inside its child process. The parent process only
# needs Microsoft.Graph.Authentication for the partner-tenant connect + GDAP
# enumeration, plus ImportExcel for the rollup writer (loaded near where
# they're used).

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
$placeholderPattern = '<.*-GUID.*>|<PARTNER-APP-CLIENT-ID-GUID>|<HOME-TENANT-ID-GUID>'
if ($config.partner.clientId -match $placeholderPattern -or
    $config.partner.homeTenantId -match $placeholderPattern) {
    throw "config.partner contains placeholder values (looks like the schema-example tenants.config.json). Copy it to tenants.config.local.json (kept outside git), populate with real values from scripts/Register-PartnerCenterApp.ps1 output, then re-run with -ConfigPath ./tenants.config.local.json."
}

# Resolve v1 script path. Param overrides config; config overrides default.
if ($config.v1ScriptPath -and -not $PSBoundParameters.ContainsKey('V1ScriptPath')) {
    $expanded = $config.v1ScriptPath -replace '^~', $HOME
    $expanded = [System.Environment]::ExpandEnvironmentVariables($expanded)
    if ($expanded -and $expanded -notmatch '^<.*>$') { $V1ScriptPath = $expanded }
}
if (-not (Test-Path $V1ScriptPath)) {
    throw "v1 script not found at $V1ScriptPath. Either it's missing from the bundled location, or config.v1ScriptPath / -V1ScriptPath points at the wrong file."
}
Write-MultiLog ("v1 script: {0}" -f $V1ScriptPath)

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

# Make sure Microsoft.Graph.Authentication is available (auto-install on first run).
if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
    Write-MultiLog 'Installing Microsoft.Graph.Authentication (CurrentUser scope)...'
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

if ($AppOnly) {
    Connect-MgGraph -TenantId    $config.partner.homeTenantId `
                    -ClientId    $config.partner.clientId `
                    -Certificate $partnerCert `
                    -NoWelcome
} else {
    # Delegated: user signs in interactively under the partner-app clientId.
    # First run prompts (browser/device-code); subsequent runs in the same
    # session reuse cached refresh tokens.
    Connect-MgGraph -TenantId $config.partner.homeTenantId `
                    -ClientId $config.partner.clientId `
                    -Scopes   'DelegatedAdminRelationship.Read.All','Directory.Read.All' `
                    -NoWelcome
}

$ctx = Get-MgContext
if (-not $ctx) {
    throw 'Connect-MgGraph did not produce a context.'
}
if ($ctx.TenantId -ne $config.partner.homeTenantId) {
    throw "Connected to the wrong tenant. Expected $($config.partner.homeTenantId), got $($ctx.TenantId)."
}
Write-MultiLog ("Connected. Tenant: {0}  Account: {1}" -f $ctx.TenantId, $ctx.Account)

# ============================================================================
# GDAP customer enumeration (Step 2)
# ============================================================================

$gdapCustomers = @()

if (-not $SkipGdapEnumeration) {
    Write-MultiLog 'Enumerating GDAP customers via /v1.0/tenantRelationships/delegatedAdminCustomers...'

    $next = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers?$top=100'
    while ($next) {
        try {
            $page = Invoke-MgGraphRequest -Method GET -Uri $next
        } catch {
            Write-MultiLog ("GDAP enumeration failed: {0}" -f $_.Exception.Message) 'ERROR'
            $script:RunWarnings.Add([pscustomobject]@{
                TenantId = '(partner)'; DisplayName = '(partner)'
                Severity = 'ERROR'; Message = "GDAP enumeration: $($_.Exception.Message)"
            })
            break
        }
        if ($page.value) { $gdapCustomers += $page.value }
        $next = $page.'@odata.nextLink'
    }

    Write-MultiLog ("GDAP enumeration returned {0} customer(s)." -f @($gdapCustomers).Count)
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

    $tenantOutDir = Join-Path $OutputRoot $tenant.ShortName
    if (-not (Test-Path $tenantOutDir)) { New-Item -ItemType Directory -Path $tenantOutDir -Force | Out-Null }
    $tenantXlsx   = Join-Path $tenantOutDir ("O365-MailGroupInventory_{0}.xlsx" -f $script:RunStamp)

    $tenantStatus = 'pending'
    $tenantWarnings = @()

    try {
        # ----- Spawn v1 as a child process for this customer tenant -----
        # Process-per-tenant orchestration: v1 does its own Connect-MgGraph,
        # Connect-ExchangeOnline, and Connect-MicrosoftTeams using the multi-
        # tenant params we pass in. MSAL's token cache is per-user, so the
        # first customer prompts for sign-in (delegated mode); subsequent
        # customers in the same run silently use cached refresh tokens.
        # AppOnly mode is currently delegated-equivalent for the per-customer
        # call until SP-level GDAP grants are in place — see README.

        $v1Args = @(
            '-NoProfile',
            '-NoLogo',
            '-File', $V1ScriptPath,
            '-OutputPath', $tenantXlsx,
            '-TenantId',   $tenant.TenantId,
            '-ClientId',   $config.partner.clientId
        )
        if ($tenant.PrimaryDomain) {
            $v1Args += @('-DelegatedOrganization', $tenant.PrimaryDomain)
        }

        # Pass through skip flags from the config defaults block.
        if ($config.defaults) {
            if ($config.defaults.skipMailboxStats)      { $v1Args += '-SkipMailboxStats' }
            if ($config.defaults.skipPermissions)       { $v1Args += '-SkipPermissions' }
            if ($config.defaults.skipUserMailboxes)     { $v1Args += '-SkipUserMailboxes' }
            if ($config.defaults.skipSharePointStats)   { $v1Args += '-SkipSharePointStats' }
            if ($config.defaults.skipConditionalAccess) { $v1Args += '-SkipConditionalAccess' }
            if ($config.defaults.skipAdminPosture)      { $v1Args += '-SkipAdminPosture' }
        }

        Write-MultiLog ("Invoking v1: pwsh {0}" -f ($v1Args -join ' '))
        & pwsh @v1Args 2>&1 | ForEach-Object { Write-Host ("    [v1] {0}" -f $_) }
        $v1Exit = $LASTEXITCODE
        if ($v1Exit -ne 0) {
            throw "v1 child process exited with code $v1Exit"
        }
        if (-not (Test-Path $tenantXlsx)) {
            throw "v1 reported success but no workbook at $tenantXlsx"
        }
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
    # No per-tenant disconnect needed: v1 ran in its own child process, which
    # terminated when the script finished. The parent process keeps its
    # Microsoft.Graph context (for the GDAP enumeration) until script end.

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

    # ImportExcel must be available — v1 uses it too, so it's typically present.
    if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
        Write-MultiLog 'Installing ImportExcel (CurrentUser scope)...' 'INFO'
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ImportExcel -ErrorAction Stop

    # ----- Run sheet: one row per tenant attempted -----
    $runRows = @($script:RunResults | Select-Object TenantId, DisplayName, ShortName, PrimaryDomain, Source, AuthMode, Status, DurationSec, OutputPath, WarningCount)

    # ----- Counts sheet: read each per-tenant Summary tab, project to columns -----
    # v1's Summary tab has metadata at the top (Field/Value rows) followed by
    # the counts table starting at $meta.Count + 3, with columns Category +
    # Count. We read every category cell as a column header and put one row
    # per tenant.
    $countsRows = @()
    foreach ($r in ($script:RunResults | Where-Object Status -eq 'ok')) {
        if (-not $r.OutputPath -or -not (Test-Path $r.OutputPath)) { continue }
        try {
            # Pull the full Summary sheet; v1 exports it with no header row,
            # so first-row-is-data. Filter to rows that look like Category/Count.
            $summary = Import-Excel -Path $r.OutputPath -WorksheetName 'Summary' -NoHeader -ErrorAction Stop
            $row = [ordered]@{
                TenantId    = $r.TenantId
                DisplayName = $r.DisplayName
                ShortName   = $r.ShortName
            }
            foreach ($cell in $summary) {
                # cell.P1 is column A, cell.P2 is column B
                $cat   = $cell.P1
                $count = $cell.P2
                if ($cat -and ($count -is [int] -or $count -is [double] -or $count -is [long])) {
                    # Skip section dividers (rows where Category contains "—")
                    if ($cat -match '^—.*—$') { continue }
                    $row[[string]$cat] = $count
                }
            }
            $countsRows += [pscustomobject]$row
        } catch {
            Write-MultiLog ("Rollup: failed to read Summary from {0}: {1}" -f $r.OutputPath, $_.Exception.Message) 'WARN'
        }
    }

    # ----- Run Issues sheet: every captured warning -----
    $issueRows = @($script:RunWarnings)

    # ----- Write the workbook -----
    $xl = @{ AutoSize = $true; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true }
    if (Test-Path $rollupXlsx) { Remove-Item $rollupXlsx -Force }
    $runRows | Export-Excel -Path $rollupXlsx -WorksheetName 'Run' @xl
    if (@($countsRows).Count -gt 0) { $countsRows | Export-Excel -Path $rollupXlsx -WorksheetName 'Counts' @xl }
    if (@($issueRows).Count -gt 0)  { $issueRows  | Export-Excel -Path $rollupXlsx -WorksheetName 'Run Issues' @xl }

    Write-MultiLog ("Rollup written: {0} tenant rows, {1} counts rows, {2} issues" -f
        @($runRows).Count, @($countsRows).Count, @($issueRows).Count)
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

# Best-effort disconnect of the parent's Graph context.
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
