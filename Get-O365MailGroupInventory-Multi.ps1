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
    + GDAP enumeration, then spawns the bundled per-tenant collector as a
    child process per customer. The collector connects to the customer
    tenant via delegated GDAP (Connect-MgGraph -TenantId, Connect-Exchange-
    Online -DelegatedOrganization, Connect-MicrosoftTeams -TenantId) using
    the parameters the wrapper passes in. After all tenants finish, the
    wrapper builds a roll-up workbook by reading each per-tenant Summary
    tab and projecting the headline counts into one row per tenant.

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
    the whole partner population. Mutually exclusive with -TenantListCsv.

.PARAMETER TenantListCsv
    Optional. Path to a CSV that defines a subset of customers to process.
    Header row must include at least one of: TenantId, ShortName,
    DisplayName, Tenant (case-insensitive; other columns are ignored).
    Match is case-insensitive against any of the three identifier columns
    on each tenant row; first match wins. CSV rows that don't match any
    GDAP tenant are warned and skipped — they don't fail the run.
    Composes with the preflight CSV verbatim:
        Import-Csv preflight-results_*.csv |
            Where-Object Status -eq 'OK' |
            Export-Csv ok-tenants.csv -NoTypeInformation
        ./Get-O365MailGroupInventory-Multi.ps1 -TenantListCsv ./ok-tenants.csv -NoConfirm
    Mutually exclusive with -OnlyTenant.

.PARAMETER SkipGdapEnumeration
    If set, only tenants explicitly listed in tenants.config.json are run.
    Useful when partner-app enumeration is failing and you want to fall
    back to the static config.

.PARAMETER SkipRollup
    If set, per-tenant workbooks are produced but the roll-up is not built.

.PARAMETER NoConfirm
    Skip the per-tenant Y/n/q confirmation prompt. Default is to prompt
    before each tenant with a 5-second auto-Y countdown so the operator
    can skip stale GDAP relationships from offboarded customers at runtime.
    Use -NoConfirm for unattended runs (cron / launchd / Task Scheduler).
    The prompt is also auto-skipped when -OnlyTenant targets a single
    tenant, and when stdin is redirected.

.PARAMETER ConfirmTimeoutSec
    Per-tenant prompt timeout in seconds. Auto-Y when the countdown
    elapses. Set to 0 to skip the wait entirely (effectively the same as
    -NoConfirm). Default 30.

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

    # Run against a CSV-defined subset of GDAP customers. The CSV must have
    # a header row including at least one of: TenantId, ShortName,
    # DisplayName (case-insensitive). Other columns are ignored — the
    # preflight CSV (Test-TenantPreflight.ps1 -Csv) plugs in directly.
    # Mutually exclusive with -OnlyTenant. The wrapper still does GDAP
    # enumeration (and /v1.0/contracts lookup) so domain resolution works;
    # the filter applies after the tenant target list is built.
    [string] $TenantListCsv,

    [switch] $SkipGdapEnumeration,
    [switch] $SkipRollup,
    # Auth mode. Default = delegated: the running user signs in interactively
    # and relies on GDAP/Lighthouse for customer-tenant authorization. Set
    # -AppOnly for unattended/scheduled runs that authenticate via the
    # partner-app cert (loaded from $config.partner.certificatePfxPath).
    [switch] $AppOnly,

    # Path to the per-tenant collector invoked per customer. Default is the
    # bundled Get-O365MailGroupInventory.ps1 next to this script. Can be
    # overridden via param OR config.collectorScriptPath. The historical
    # config key (v1ScriptPath) and parameter alias (-V1ScriptPath) still
    # work for backward compatibility with older local configs.
    [Alias('V1ScriptPath')]
    [string] $CollectorScriptPath = (Join-Path $PSScriptRoot 'Get-O365MailGroupInventory.ps1'),

    # Skip the per-tenant Y/n/q confirmation prompt. Use for unattended runs
    # (cron / launchd / Task Scheduler) where there's no human at the keyboard.
    # Default behaviour prompts before each tenant with a 5-second auto-Y
    # countdown so stale GDAP relationships from offboarded customers can be
    # skipped at runtime.
    [switch] $NoConfirm,

    # Per-tenant prompt timeout in seconds. Auto-Y when the countdown elapses.
    # Set to 0 to skip the wait entirely (effectively the same as -NoConfirm).
    [int] $ConfirmTimeoutSec = 5
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

# Per-tenant confirmation with a countdown that auto-Y's at timeout.
# Returns 'Y' (process), 'N' (skip), or 'Q' (stop the run after this point).
# Auto-returns 'Y' when stdin is redirected (non-interactive context like
# launchd/Task Scheduler), when the timeout is 0, or when the countdown elapses.
function Read-TenantConfirmation {
    param(
        [Parameter(Mandatory)] [string] $DisplayName,
        [string] $Domain,
        [int]    $TimeoutSec = 30
    )

    if ($TimeoutSec -le 0) { return 'Y' }
    if ([Console]::IsInputRedirected) { return 'Y' }

    # Drain any keystrokes left over from a prior prompt
    while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null }

    $label = if ($Domain) { "'$DisplayName' ($Domain)" } else { "'$DisplayName'" }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $ansiClearLine = [char]27 + '[K'

    while ((Get-Date) -lt $deadline) {
        $remaining = [Math]::Max(0, [int][Math]::Ceiling(($deadline - (Get-Date)).TotalSeconds))
        $line = "Process {0}? [Y/n/q]  (auto-Y in {1,2}s)..." -f $label, $remaining
        Write-Host -NoNewline ("`r{0}{1}" -f $line, $ansiClearLine)
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            Write-Host ''   # advance past the in-place line
            if ($key.Key -eq [ConsoleKey]::Enter) { return 'Y' }
            $char = $key.KeyChar.ToString().ToUpperInvariant()
            if ($char -in 'Y','N','Q') { return $char }
            return 'Y'   # any other key = accept default
        }
        Start-Sleep -Milliseconds 200
    }

    Write-Host ''   # advance past the in-place line
    return 'Y'   # timeout = auto-Y
}

# ============================================================================
# Module / library bootstrap
# ============================================================================

Write-MultiLog 'Bootstrapping modules...'

# The collector (the bundled per-tenant inventory script this wrapper drives) handles
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

# Resolve collector script path. Param overrides config; config overrides
# default. Both the new key (collectorScriptPath) and the legacy key
# (v1ScriptPath) are accepted; the new one wins if both are present.
$collectorPathFromConfig = $null
if ($config.collectorScriptPath) { $collectorPathFromConfig = $config.collectorScriptPath }
elseif ($config.v1ScriptPath)    { $collectorPathFromConfig = $config.v1ScriptPath }

if ($collectorPathFromConfig -and
    -not $PSBoundParameters.ContainsKey('CollectorScriptPath') -and
    -not $PSBoundParameters.ContainsKey('V1ScriptPath')) {
    $expanded = $collectorPathFromConfig -replace '^~', $HOME
    $expanded = [System.Environment]::ExpandEnvironmentVariables($expanded)
    if ($expanded -and $expanded -notmatch '^<.*>$') { $CollectorScriptPath = $expanded }
}
if (-not (Test-Path $CollectorScriptPath)) {
    throw "Collector script not found at $CollectorScriptPath. Either it's missing from the bundled location, or config.collectorScriptPath / -CollectorScriptPath points at the wrong file."
}
Write-MultiLog ("Collector script: {0}" -f $CollectorScriptPath)

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

    # /tenantRelationships/delegatedAdminCustomers returns id + tenantId +
    # displayName but NOT the customer's defaultDomainName. /v1.0/contracts
    # (the CSP-relationship endpoint) does, keyed on customerId. Pull both
    # and join on tenantId so each customer row carries a domain we can
    # pass as -DelegatedOrganization to Connect-ExchangeOnline.
    Write-MultiLog 'Resolving customer default domains via /v1.0/contracts...'
    $script:DomainByTenantId = @{}
    $contractsNext = 'https://graph.microsoft.com/v1.0/contracts?$top=100'
    while ($contractsNext) {
        try {
            $page = Invoke-MgGraphRequest -Method GET -Uri $contractsNext
        } catch {
            Write-MultiLog ("/v1.0/contracts lookup failed: {0}. PrimaryDomain may be missing on some tenants — populate manually in tenants.config.local.json under config.tenants[].primaryDomain to override." -f $_.Exception.Message) 'WARN'
            break
        }
        foreach ($c in @($page.value)) {
            $cid = ($c.customerId, $c.tenantId, $c.id | Where-Object { $_ } | Select-Object -First 1)
            if ($cid -and $c.defaultDomainName) {
                $script:DomainByTenantId[$cid.ToLower()] = $c.defaultDomainName
            }
        }
        $contractsNext = $page.'@odata.nextLink'
    }
    Write-MultiLog ("Resolved domain for {0} customer(s) from /v1.0/contracts." -f $script:DomainByTenantId.Count)
} else {
    Write-MultiLog '-SkipGdapEnumeration set; using only statically-configured tenants.'
    $script:DomainByTenantId = @{}
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

    # Resolve PrimaryDomain in priority order:
    #   1. Static config override (config.tenants[].primaryDomain)
    #   2. /v1.0/contracts hashtable from the lookup above
    #   3. defaultDomainName on the GDAP row (rarely populated, but cheap to try)
    $resolvedDomain = $null
    if ($override.primaryDomain) {
        $resolvedDomain = $override.primaryDomain
    } elseif ($script:DomainByTenantId.ContainsKey($tid.ToLower())) {
        $resolvedDomain = $script:DomainByTenantId[$tid.ToLower()]
    } elseif ($gd.defaultDomainName) {
        $resolvedDomain = $gd.defaultDomainName
    }

    $tenantTargets.Add([pscustomobject]@{
        TenantId       = $tid
        DisplayName    = if ($override.displayName) { $override.displayName } else { $gd.displayName }
        ShortName      = if ($override.shortName)   { $override.shortName }   else { ($gd.displayName -replace '[^A-Za-z0-9]+','-').Trim('-').ToLower() }
        PrimaryDomain  = $resolvedDomain
        Source         = 'gdap'
        Auth           = if ($override.auth) { $override.auth } else { @{ mode = 'gdap' } }
        Overrides      = $override.overrides
    })
}

# Static-only tenants (not in GDAP enumeration).
# Filter out placeholder rows from the schema example. The schema's
# tenants[] array carries example entries with placeholder GUIDs like
# <CUSTOMER-A-TENANT-GUID> as documentation; if an operator copied
# tenants.config.json to tenants.config.local.json without stripping
# them, those placeholder rows would be processed as real tenants and
# blow up on the per-tenant child invocation. Defensively drop anything
# whose tenantId or primaryDomain looks like a <…> placeholder.
$placeholderPattern = '<.*>'
foreach ($t in @($config.tenants)) {
    if (-not $t.tenantId) { continue }
    if ($t.tenantId -match $placeholderPattern -or
        ($t.primaryDomain -and $t.primaryDomain -match $placeholderPattern) -or
        ($t.displayName   -and $t.displayName   -match $placeholderPattern)) {
        Write-MultiLog ("Skipping placeholder tenant row from config.tenants[]: '{0}' ({1}). Strip example rows from tenants.config.local.json — see tenants.config.json for the schema." -f $t.displayName, $t.tenantId) 'WARN'
        continue
    }
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

# Mutually-exclusive filters: -OnlyTenant picks one customer; -TenantListCsv
# picks a CSV-defined subset. They're alternative selectors; allowing both
# would either be redundant (CSV is a superset of one) or contradictory
# (CSV doesn't include the OnlyTenant). Bail out cleanly if both supplied.
if ($OnlyTenant -and $TenantListCsv) {
    throw "-OnlyTenant and -TenantListCsv are mutually exclusive; pick one. -OnlyTenant for a single customer; -TenantListCsv for a multi-customer subset."
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

if ($TenantListCsv) {
    if (-not (Test-Path $TenantListCsv)) {
        throw "-TenantListCsv path not found: $TenantListCsv"
    }
    Write-MultiLog ("Loading tenant subset from CSV: {0}" -f $TenantListCsv)

    # Read CSV. Tolerate column casing (TenantId / tenantId / TENANTID all work)
    # by normalizing each row's keys to lowercase.
    $csvRows = @()
    try {
        $csvRows = @(Import-Csv -Path $TenantListCsv)
    } catch {
        throw "Failed to read CSV at ${TenantListCsv}: $($_.Exception.Message)"
    }
    if (@($csvRows).Count -eq 0) {
        throw "-TenantListCsv at $TenantListCsv has no rows. Expected at least one row with TenantId / ShortName / DisplayName."
    }

    # Header sanity check — at least one of our identifier columns must exist.
    $cols = @($csvRows[0].PSObject.Properties.Name | ForEach-Object { $_.ToLower() })
    $hasIdentifierCol = ($cols -contains 'tenantid') -or
                       ($cols -contains 'shortname') -or
                       ($cols -contains 'displayname') -or
                       ($cols -contains 'tenant')   # preflight CSV uses 'Tenant'
    if (-not $hasIdentifierCol) {
        throw "-TenantListCsv at $TenantListCsv has no recognizable identifier column. Need at least one of: TenantId, ShortName, DisplayName, Tenant. Other columns are ignored."
    }

    # Build lowercased lookup set from CSV rows.
    $csvNeedles = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $csvRows) {
        # Resolve each identifier with case-insensitive property lookup.
        function Get-CsvCol { param($r, [string]$name)
            $prop = $r.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
            if ($prop) { return [string]$prop.Value } else { return $null }
        }
        foreach ($colName in 'TenantId','ShortName','DisplayName','Tenant') {
            $val = Get-CsvCol $row $colName
            if ($val -and $val.Trim()) { [void]$csvNeedles.Add($val.Trim()) }
        }
    }

    if ($csvNeedles.Count -eq 0) {
        throw "-TenantListCsv at $TenantListCsv has rows but no non-empty TenantId/ShortName/DisplayName values to match against."
    }

    # Filter targets — match on any of the three identifiers.
    $beforeCount = @($tenantTargets).Count
    $matched = @()
    $matchedNeedles = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $tenantTargets) {
        if ($csvNeedles.Contains($t.TenantId) -or
            ($t.ShortName   -and $csvNeedles.Contains($t.ShortName)) -or
            ($t.DisplayName -and $csvNeedles.Contains($t.DisplayName))) {
            $matched += $t
            # Record which needles matched, so we can warn about the rest.
            foreach ($n in @($t.TenantId, $t.ShortName, $t.DisplayName)) {
                if ($n -and $csvNeedles.Contains($n)) { [void]$matchedNeedles.Add($n) }
            }
        }
    }

    # Warn about CSV rows that didn't match any GDAP tenant — typo, offboarded,
    # whatever. Don't fail the whole run; just surface them so the operator
    # knows their CSV had stale entries.
    foreach ($needle in $csvNeedles) {
        if (-not $matchedNeedles.Contains($needle)) {
            Write-MultiLog ("CSV row '{0}' didn't match any GDAP tenant — skipped." -f $needle) 'WARN'
        }
    }

    $tenantTargets = $matched
    Write-MultiLog ("Filtered {0} tenants → {1} matching CSV rows." -f $beforeCount, @($tenantTargets).Count)
}

Write-MultiLog ("Tenant target count: {0}" -f @($tenantTargets).Count)

# ============================================================================
# Per-tenant loop (Step 3)
# ============================================================================

if (-not (Test-Path $OutputRoot)) { New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null }

$userQuit = $false
foreach ($tenant in $tenantTargets) {
    $tStart = Get-Date
    Write-MultiLog ("─── {0} ({1}) — {2} ───" -f $tenant.DisplayName, $tenant.ShortName, $tenant.TenantId)

    # Per-tenant Y/n/q confirmation. Skip when -NoConfirm, when -OnlyTenant
    # is targeting a single tenant (no ambiguity to resolve), or when stdin
    # is redirected (the helper auto-Y's in that case).
    if (-not $NoConfirm -and -not $OnlyTenant) {
        $answer = Read-TenantConfirmation `
            -DisplayName $tenant.DisplayName `
            -Domain      $tenant.PrimaryDomain `
            -TimeoutSec  $ConfirmTimeoutSec
        if ($answer -eq 'N') {
            Write-MultiLog ("Skipped {0} by user choice." -f $tenant.DisplayName) 'WARN'
            $script:RunResults.Add([pscustomobject]@{
                TenantId      = $tenant.TenantId
                DisplayName   = $tenant.DisplayName
                ShortName     = $tenant.ShortName
                PrimaryDomain = $tenant.PrimaryDomain
                Source        = $tenant.Source
                AuthMode      = $tenant.Auth.mode
                Status        = 'skipped'
                DurationSec   = 0
                OutputPath    = $null
                WarningCount  = 0
            })
            continue
        }
        if ($answer -eq 'Q') {
            Write-MultiLog 'User chose to quit. Stopping the per-tenant loop; rollup will reflect tenants processed up to this point.' 'WARN'
            $userQuit = $true
            break
        }
    }

    $tenantOutDir = Join-Path $OutputRoot $tenant.ShortName
    if (-not (Test-Path $tenantOutDir)) { New-Item -ItemType Directory -Path $tenantOutDir -Force | Out-Null }
    $tenantXlsx   = Join-Path $tenantOutDir ("O365-MailGroupInventory_{0}.xlsx" -f $script:RunStamp)

    $tenantStatus = 'pending'
    $tenantWarnings = @()

    try {
        # ----- Spawn the collector as a child process for this customer ----
        # Process-per-tenant orchestration: the collector does its own
        # Connect-MgGraph / Connect-ExchangeOnline / Connect-MicrosoftTeams
        # using the GDAP delegated params we pass in. MSAL's token cache is
        # per-user, so the first customer prompts for sign-in; subsequent
        # customers in the same run silently use cached refresh tokens.
        # -DelegatedOrganization is required by the collector — without a
        # PrimaryDomain on the tenant row, EXO can't connect.

        if (-not $tenant.PrimaryDomain) {
            # Skip this tenant cleanly rather than throwing — other customers
            # in this run shouldn't be punished for one missing domain. The
            # operator can recover by adding `primaryDomain` under
            # config.tenants[] for this tenantId in tenants.config.local.json.
            Write-MultiLog ("Tenant '{0}' has no PrimaryDomain (not in /v1.0/contracts and no static override). Skipping. To recover, add this tenant under config.tenants[] with primaryDomain populated, e.g. {{ tenantId='{1}', primaryDomain='customer.onmicrosoft.com' }}." -f $tenant.DisplayName, $tenant.TenantId) 'WARN'
            $tenantStatus = 'skipped (no PrimaryDomain)'
            $script:RunWarnings.Add([pscustomobject]@{
                TenantId    = $tenant.TenantId
                DisplayName = $tenant.DisplayName
                Severity    = 'WARN'
                Message     = 'Skipped: no PrimaryDomain available from /v1.0/contracts or static config.'
            })
            continue
        }

        $collectorArgs = @(
            '-NoProfile',
            '-NoLogo',
            '-File', $CollectorScriptPath,
            '-OutputPath', $tenantXlsx,
            '-TenantId',   $tenant.TenantId,
            '-ClientId',   $config.partner.clientId,
            '-DelegatedOrganization', $tenant.PrimaryDomain
        )

        # Pass through skip flags from the config defaults block.
        if ($config.defaults) {
            if ($config.defaults.skipMailboxStats)      { $collectorArgs += '-SkipMailboxStats' }
            if ($config.defaults.skipPermissions)       { $collectorArgs += '-SkipPermissions' }
            if ($config.defaults.skipUserMailboxes)     { $collectorArgs += '-SkipUserMailboxes' }
            if ($config.defaults.skipSharePointStats)   { $collectorArgs += '-SkipSharePointStats' }
            if ($config.defaults.skipConditionalAccess) { $collectorArgs += '-SkipConditionalAccess' }
            if ($config.defaults.skipAdminPosture)      { $collectorArgs += '-SkipAdminPosture' }
        }

        Write-MultiLog ("Invoking collector: pwsh {0}" -f ($collectorArgs -join ' '))
        & pwsh @collectorArgs 2>&1 | ForEach-Object { Write-Host ("    [collector] {0}" -f $_) }
        $collectorExit = $LASTEXITCODE
        if ($collectorExit -ne 0) {
            throw "Collector child process exited with code $collectorExit"
        }
        if (-not (Test-Path $tenantXlsx)) {
            throw "Collector reported success but no workbook at $tenantXlsx"
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
    # No per-tenant disconnect needed: the collector ran in its own child
    # process, which terminated when the script finished. The parent process
    # keeps its Microsoft.Graph context (for the GDAP enumeration) until the
    # wrapper exits.

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

    # ImportExcel must be available — the collector uses it too, so it's typically present.
    if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
        Write-MultiLog 'Installing ImportExcel (CurrentUser scope)...' 'INFO'
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ImportExcel -ErrorAction Stop

    # ----- Run sheet: one row per tenant attempted -----
    $runRows = @($script:RunResults | Select-Object TenantId, DisplayName, ShortName, PrimaryDomain, Source, AuthMode, Status, DurationSec, OutputPath, WarningCount)

    # ----- Counts sheet: read each per-tenant Summary tab, project to columns -----
    # The collector's Summary tab has metadata at the top (Field/Value rows) followed by
    # the counts table starting at $meta.Count + 3, with columns Category +
    # Count. We read every category cell as a column header and put one row
    # per tenant.
    $countsRows = @()
    foreach ($r in ($script:RunResults | Where-Object Status -eq 'ok')) {
        if (-not $r.OutputPath -or -not (Test-Path $r.OutputPath)) { continue }
        try {
            # Pull the full Summary sheet; the collector exports it with no header row,
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
$skipped = @($script:RunResults | Where-Object Status -eq 'skipped').Count
$quitNote = if ($userQuit) { ' (run stopped early by user)' } else { '' }
Write-MultiLog ("Done. Tenants in summary: {0}; ok: {1}; error: {2}; skipped: {3}; elapsed: {4:hh\:mm\:ss}{5}" -f
    @($script:RunResults).Count, $ok, $err, $skipped, $elapsed, $quitNote)

# Per-run summary CSV alongside the rollup, useful for diffing runs.
$summaryCsv = Join-Path $OutputRoot ("run-summary_{0}.csv" -f $script:RunStamp)
$script:RunResults | Export-Csv -NoTypeInformation -Path $summaryCsv
Write-MultiLog "Per-run summary at $summaryCsv"

# Best-effort disconnect of the parent's Graph context.
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
