<#
.SYNOPSIS
    Pre-flight validation. Tests whether each GDAP customer tenant has a
    service principal for the OAuth client the inventory tool uses
    (Microsoft Graph PowerShell, 14d82eec-204b-4c2f-b7e8-296a70dab67e).
    Outputs a punchlist of tenants needing admin consent with the URL to
    send each customer admin.

.DESCRIPTION
    Run this once before a full inventory run to identify which customer
    tenants will work cleanly and which need per-tenant admin consent.
    Re-run with -OnlyTenant after a customer admin has consented to
    confirm their tenant is now ready.

    How it works:
      1. Single device-code sign-in to your partner tenant. You enter a
         short code at microsoft.com/devicelogin once. No per-tenant
         browser pop-ups for any of the customer tenants — even the
         ones that will fail.
      2. The script enumerates every GDAP customer.
      3. For each customer, it sends an OAuth refresh-token grant to
         that tenant's /token endpoint. Success = the OAuth client has
         a service principal in the customer tenant; failure with
         AADSTS90099 = the SP is missing and admin consent is needed.
      4. Output: clean run (every tenant passes) → one-line success
         message. Partial run → punchlist with one row per failing
         tenant + the admin-consent URL.

    The script does NOT attempt to acquire tokens for Microsoft Teams
    PowerShell separately — that would need a second device-code flow.
    But customer tenants that reject Microsoft Graph PowerShell almost
    always reject the Teams client too, so the punchlist includes a
    second consent URL for Teams alongside each Graph URL. The customer
    admin can click both and be done.

    No file output. Pipe to a file if you want persistence.

.PARAMETER ConfigPath
    Path to your tenants.config.local.json. Used for the partner home
    tenant id. Default: ../tenants.config.local.json.

.PARAMETER OnlyTenant
    Test a single tenant by tenant id or display name (case-insensitive).
    Useful for verifying a customer admin's consent has actually landed.

.PARAMETER Csv
    Also write the results to a CSV file alongside the on-screen output.
    With no path argument, writes to ./preflight-results_<timestamp>.csv in
    the current working directory. The CSV has one row per tenant (passing
    tenants included with Status=OK so the file is a complete snapshot)
    and includes the consent URL inline for the NEEDS_CONSENT cases.
    Useful for MSPs with larger customer bases who want to mail-merge or
    bulk-process the punchlist.

.PARAMETER CsvPath
    Explicit path for the CSV output. Implies -Csv. Use when you want a
    specific filename (e.g. for a scheduled run that overwrites a known
    location, or for stable diffs between runs).

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1
    # Test every GDAP customer; print punchlist or success message.

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1 -OnlyTenant 'Acme Corp'
    # Re-test a single customer after their admin consented.

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1 -OnlyTenant 00000000-0000-0000-0000-000000000000
    # Same, by tenant id.

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1 -Csv
    # Run against everyone and dump every tenant + status + consent URL
    # to a timestamped CSV in the cwd alongside the on-screen punchlist.

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1 -CsvPath ./reports/preflight-2026-05-06.csv
    # Same, with an explicit output path.

.NOTES
    License: Apache-2.0 (see LICENSE in repo root)
    Created: 2026-05-05
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [string] $ConfigPath = (Join-Path $PSScriptRoot '..' 'tenants.config.local.json'),
    [string] $OnlyTenant,

    # CSV output. -Csv writes to a timestamped path in cwd; -CsvPath overrides
    # to a specific location and implies -Csv.
    [switch] $Csv,
    [string] $CsvPath
)

$ErrorActionPreference = 'Stop'

# Well-known clientIds. We acquire a refresh token via device-code flow for
# Microsoft Graph PowerShell; that refresh token is then used to test each
# customer tenant's /token endpoint. The Teams clientId is included only
# in the punchlist consent URLs — we don't authenticate against it here.
$GraphPsClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$TeamsPsClientId = '12128f48-ec9e-42f0-b203-ea49fb6af367'

# v1 collector's full Graph scope set. We mirror this list when building the
# admin-consent URL so the customer admin pre-grants every scope v1 needs.
# Out of date if v1's $graphScopes drifts; check Get-O365MailGroupInventory.ps1.
$V1GraphScopes = @(
    'Group.Read.All','GroupMember.Read.All','User.Read.All',
    'Team.ReadBasic.All','Channel.ReadBasic.All','Directory.Read.All',
    'Sites.Read.All','Reports.Read.All','Policy.Read.All',
    'RoleManagement.Read.Directory','AuditLog.Read.All',
    'UserAuthenticationMethod.Read.All','Application.Read.All'
)

function Get-AdminConsentUrl {
    param([string] $TenantId)
    # v2.0 /adminconsent with an explicit scope list. Microsoft Graph PowerShell
    # is a dynamic-scope client, so /adminconsent without /v2.0 (which consents
    # the app's static requiredResourceAccess) leaves elevated scopes un-granted.
    # /v2.0/adminconsent accepts a scope= parameter and forces the admin to
    # consent the listed scopes specifically.
    #
    # Note on prompt=admin_consent: that's only valid on the v2.0 /authorize
    # endpoint paired with response_type=code (and even then it's been
    # deprecated). On /v2.0/adminconsent the consent itself IS the prompt;
    # no prompt= parameter needed.
    $scopeString = ($V1GraphScopes | ForEach-Object { "https://graph.microsoft.com/$_" }) -join ' '
    $params = @{
        client_id    = $GraphPsClientId
        redirect_uri = 'https://login.microsoftonline.com/common/oauth2/nativeclient'
        scope        = $scopeString
        state        = [guid]::NewGuid().Guid
    }
    $query = ($params.GetEnumerator() | ForEach-Object {
        '{0}={1}' -f $_.Key, [System.Net.WebUtility]::UrlEncode($_.Value)
    }) -join '&'
    "https://login.microsoftonline.com/$TenantId/v2.0/adminconsent?$query"
}

# ============================================================================
# Load config
# ============================================================================

if (-not (Test-Path $ConfigPath)) {
    throw "Config not found at $ConfigPath. Run scripts/Setup-LocalConfig.ps1 first to create it."
}
$config = Get-Content -Raw $ConfigPath | ConvertFrom-Json

if (-not $config.partner.homeTenantId) {
    throw "config.partner.homeTenantId is empty in $ConfigPath. Re-run Setup-LocalConfig.ps1 to populate it."
}
$partnerTenantId = $config.partner.homeTenantId

# ============================================================================
# Step 1: device-code sign-in. Get a refresh token we can use silently
# against each customer tenant's /token endpoint.
# ============================================================================

# Request a representative admin-consent scope (Directory.Read.All) at sign-in
# time so we can use it as the per-customer probe. Earlier versions tested
# only User.Read, which can pass for a tenant where the SP exists but admin
# scopes haven't been consented — exactly the case the wrapper actually
# fails on (AADSTS90099). Directory.Read.All is requested by every v1 run,
# so if we can acquire it for a customer tenant, the wrapper will too.
$scopes = 'offline_access https://graph.microsoft.com/Directory.Read.All https://graph.microsoft.com/DelegatedAdminRelationship.Read.All'

Write-Host '==> Initiating device-code sign-in...' -ForegroundColor Cyan

$dc = Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/$partnerTenantId/oauth2/v2.0/devicecode" `
    -ContentType 'application/x-www-form-urlencoded' `
    -Body @{
        client_id = $GraphPsClientId
        scope     = $scopes
    }

Write-Host ''
Write-Host '    To sign in:' -ForegroundColor Yellow
Write-Host "      1. Open $($dc.verification_uri) in any browser" -ForegroundColor Yellow
Write-Host "      2. Enter code: $($dc.user_code)" -ForegroundColor Yellow
Write-Host "      3. Sign in with your partner-tenant account" -ForegroundColor Yellow
Write-Host ''
Write-Host '    Waiting for authentication' -NoNewline -ForegroundColor Cyan

$accessToken  = $null
$refreshToken = $null
$deadline = (Get-Date).AddSeconds($dc.expires_in)
$pollInterval = [Math]::Max(1, $dc.interval)

while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds $pollInterval
    Write-Host -NoNewline '.' -ForegroundColor Cyan
    try {
        $token = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$partnerTenantId/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $GraphPsClientId
                device_code = $dc.device_code
            }
        $accessToken  = $token.access_token
        $refreshToken = $token.refresh_token
        break
    }
    catch {
        $err = $null
        try { $err = $_.ErrorDetails.Message | ConvertFrom-Json } catch {}
        if ($err -and $err.error -in @('authorization_pending','slow_down')) {
            if ($err.error -eq 'slow_down') { $pollInterval++ }
            continue
        }
        $msg = if ($err) { $err.error_description } else { $_.Exception.Message }
        throw "Authentication failed: $msg"
    }
}

if (-not $refreshToken) { throw 'Authentication timed out before user signed in.' }
Write-Host ' done.' -ForegroundColor Green

# ============================================================================
# Step 2: enumerate GDAP customers
# ============================================================================

Write-Host '==> Enumerating GDAP customers...' -ForegroundColor Cyan

$customers = @()
$next = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers?$top=100'
while ($next) {
    $page = Invoke-RestMethod -Method GET -Uri $next `
        -Headers @{ Authorization = "Bearer $accessToken" }
    if ($page.value) { $customers += $page.value }
    $next = $page.'@odata.nextLink'
}
Write-Host ('    Found {0} GDAP customer(s).' -f @($customers).Count) -ForegroundColor Cyan

if ($OnlyTenant) {
    $needle = $OnlyTenant.ToLower()
    $customers = @($customers | Where-Object {
        $_.tenantId.ToLower()    -eq $needle -or
        $_.displayName.ToLower() -eq $needle
    })
    Write-Host ('    Filtered to {0} matching customer(s).' -f @($customers).Count) -ForegroundColor Cyan
    if (@($customers).Count -eq 0) {
        throw "OnlyTenant '$OnlyTenant' didn't match any GDAP customer. Try a tenant id or exact display name."
    }
}

# ============================================================================
# Step 3: silent test per customer
# ============================================================================

Write-Host '==> Testing each customer tenant...' -ForegroundColor Cyan

$punchlist = @()
# allResults: every tenant + outcome (including OK), used for CSV output.
# Punchlist still drives the on-screen output and remains the human-eyes view.
$allResults = @()
$total = @($customers).Count
$idx = 0
foreach ($c in $customers) {
    $idx++
    $tenantId    = $c.tenantId
    $displayName = $c.displayName

    # Pad / truncate the status line so the carriage-return overwrite is clean
    $line = '    [{0,3}/{1}] {2,-50}' -f $idx, $total, $displayName.Substring(0, [Math]::Min(50, $displayName.Length))
    $ansiClearLine = [char]27 + '[K'
    Write-Host -NoNewline ("`r{0}{1}" -f $line, $ansiClearLine)

    try {
        $null = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body @{
                grant_type    = 'refresh_token'
                client_id     = $GraphPsClientId
                refresh_token = $refreshToken
                scope         = 'https://graph.microsoft.com/Directory.Read.All'
            }
        # Success — the customer tenant has the SP and an admin-consent scope
        # representative of the collector's actual scope set is consented.
        $allResults += [pscustomobject]@{
            Tenant     = $displayName
            TenantId   = $tenantId
            Status     = 'OK'
            ConsentUrl = ''
            ErrorDesc  = ''
        }
    }
    catch {
        $err = $null
        try { $err = $_.ErrorDetails.Message | ConvertFrom-Json } catch {}
        $errCode = if ($err) { $err.error } else { '' }
        $errDesc = if ($err) { $err.error_description } else { $_.Exception.Message }

        $status = $null
        $consentUrls = @()

        # Both AADSTS90099 (no SP) and AADSTS65001 (SP exists but the
        # requested scopes aren't admin-consented) need the same fix: a
        # Cloud Application Administrator in the customer tenant clicks an
        # /authorize URL with prompt=admin_consent and the full scope list.
        # The simpler /adminconsent?client_id=X URL only grants the app's
        # static-declared permissions, which for Microsoft Graph PowerShell
        # is a small set — the elevated scopes v1 actually requests
        # (Directory.Read.All, Application.Read.All, etc.) stay un-consented
        # and the wrapper still fails. The /authorize URL with explicit
        # scope= and prompt=admin_consent grants every scope listed.
        if ($errDesc -match 'AADSTS90099') {
            $status = 'NEEDS_CONSENT (no SP)'
            $consentUrls = @(Get-AdminConsentUrl -TenantId $tenantId)
        }
        elseif ($errDesc -match 'AADSTS65001') {
            $status = 'NEEDS_CONSENT (scopes)'
            $consentUrls = @(Get-AdminConsentUrl -TenantId $tenantId)
        }
        elseif ($errDesc -match 'AADSTS50020|AADSTS50034|AADSTS50158') {
            # User account doesn't exist in the tenant — usually means GDAP isn't active for this customer
            $status = 'NO_ACCESS (GDAP)'
        }
        elseif ($errDesc -match 'authentication flow checks|authentication flows policy') {
            # Customer's "Authentication Flows Policy" CA control blocks device code
            # flow specifically (phishing mitigation). The partner-app interactive
            # browser flow used by the actual inventory wrapper is a different grant
            # type and may still work — preflight just can't validate it.
            $status = 'BLOCKED_BY_CA (device-code only)'
        }
        elseif ($errDesc -match 'AADSTS530036|AADSTS500131|AADSTS530003') {
            # Generic Conditional Access block — applies to all auth, not just device code
            $status = 'BLOCKED_BY_CA'
        }
        else {
            $status = "ERROR: $errCode"
        }

        $punchlist += [pscustomobject]@{
            Tenant      = $displayName
            TenantId    = $tenantId
            Status      = $status
            ConsentUrls = $consentUrls
            ErrorDesc   = $errDesc
        }
        # CSV-friendly mirror — flat row, primary URL only, single-line error.
        # Newlines in error_description would break Excel CSV import; collapse.
        $allResults += [pscustomobject]@{
            Tenant     = $displayName
            TenantId   = $tenantId
            Status     = $status
            ConsentUrl = if (@($consentUrls).Count -gt 0) { $consentUrls[0] } else { '' }
            ErrorDesc  = ($errDesc -replace '\s+', ' ').Trim()
        }
    }
}

Write-Host -NoNewline ("`r{0}" -f $ansiClearLine)
Write-Host ''

# ============================================================================
# Step 4: output
# ============================================================================

# CSV emit — runs before the early-return on a clean run AND at the end of
# the punchlist path. -CsvPath implies -Csv. Default path lands a timestamped
# file in the cwd so two runs in a row don't clobber each other.
function Write-PreflightCsv {
    param([object[]] $Rows, [string] $Path)
    if (@($Rows).Count -eq 0) { return }
    $resolvedPath = $Path
    if (-not $resolvedPath) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmm'
        $resolvedPath = Join-Path (Get-Location) "preflight-results_$stamp.csv"
    }
    $dir = Split-Path -Parent $resolvedPath
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $Rows | Export-Csv -Path $resolvedPath -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host ('CSV written: {0} ({1} row(s))' -f $resolvedPath, @($Rows).Count) -ForegroundColor Cyan
}

$wantCsv = $Csv.IsPresent -or $CsvPath

if (@($punchlist).Count -eq 0) {
    Write-Host ''
    Write-Host ('OK — all {0} customer tenant(s) passed the Directory.Read.All refresh-token probe.' -f $total) -ForegroundColor Green
    Write-Host '   Caveat: preflight uses a refresh-token grant; the wrapper uses interactive browser auth. They mostly agree, but a tenant can pass preflight and still fail the wrapper if the customer''s tenant requires admin consent on first interactive use of the elevated scope set. If that happens, the collector logs the consent URL inline — same fix as a NEEDS_CONSENT punchlist entry.' -ForegroundColor DarkGray
    Write-Host '   Preflight also doesn''t separately verify the Microsoft Teams PowerShell client. Teams data flows through Graph for our inventory, so this rarely matters; if Teams data is empty for a tenant after a real run, send the customer admin the Teams consent URL too.' -ForegroundColor DarkGray
    if ($wantCsv) { Write-PreflightCsv -Rows $allResults -Path $CsvPath }
    return
}

# Summary header
Write-Host ''
Write-Host ('Punchlist — {0} tenant(s) need attention out of {1}:' -f @($punchlist).Count, $total) -ForegroundColor Yellow
Write-Host ''

# Group by status so the operator sees actionable items first
$ordered = @(
    $punchlist | Where-Object Status -like 'NEEDS_CONSENT*'
    $punchlist | Where-Object Status -like 'BLOCKED_BY_CA*'
    $punchlist | Where-Object Status -eq  'NO_ACCESS (GDAP)'
    $punchlist | Where-Object Status -like 'ERROR:*'
)

foreach ($p in $ordered) {
    Write-Host ('  {0}' -f $p.Tenant) -ForegroundColor Yellow
    Write-Host ('  {0}  [{1}]' -f $p.TenantId, $p.Status) -ForegroundColor DarkGray
    if ($p.Status -like 'NEEDS_CONSENT*') {
        if ($p.Status -eq 'NEEDS_CONSENT (scopes)') {
            Write-Host '    Service principal exists but the required scopes aren''t consented. Send the customer admin this URL (one click):'
        } else {
            Write-Host '    Service principal not present in tenant. Send the customer admin this URL (one click):'
        }
        Write-Host ('      {0}' -f $p.ConsentUrls[0]) -ForegroundColor White
    }
    elseif ($p.Status -eq 'BLOCKED_BY_CA (device-code only)') {
        Write-Host '    Customer tenant''s "Authentication Flows Policy" CA control blocks device code'
        Write-Host '    flow (phishing mitigation, applies to all apps in their tenant). Preflight uses'
        Write-Host '    device code flow, so it can''t validate this tenant — but the actual inventory'
        Write-Host '    wrapper uses an interactive browser auth flow (different grant type) and may'
        Write-Host '    still work. Try a real wrapper run with -OnlyTenant against this tenant.'
        Write-Host '    If THAT fails too, the customer admin needs to add an exclusion to their CA'
        Write-Host '    Authentication Flows Policy for partner / GDAP delegated admin accounts.'
        Write-Host ('    Diagnostic: {0}' -f $p.ErrorDesc) -ForegroundColor DarkGray
    }
    elseif ($p.Status -eq 'BLOCKED_BY_CA') {
        Write-Host '    Customer tenant''s Conditional Access policy is blocking the partner sign-in.'
        Write-Host '    The customer admin needs to either:'
        Write-Host '      - exempt partner / GDAP delegated admins from the blocking CA policy, or'
        Write-Host '      - allow the Microsoft Graph PowerShell client (14d82eec-204b-4c2f-b7e8-296a70dab67e) explicitly.'
        Write-Host ('    Diagnostic: {0}' -f $p.ErrorDesc) -ForegroundColor DarkGray
    }
    elseif ($p.Status -eq 'NO_ACCESS (GDAP)') {
        Write-Host '    Your account isn''t in this customer tenant — likely GDAP is no longer active or the relationship lapsed. Check Partner Center → Customers → this customer → GDAP relationships.'
    }
    else {
        Write-Host ('    Error: {0}' -f $p.ErrorDesc)
    }
    Write-Host ''
}

# Closing summary — only relevant when NEEDS_CONSENT items are present
$needsConsent = @($punchlist | Where-Object Status -like 'NEEDS_CONSENT*')
if (@($needsConsent).Count -gt 0) {
    Write-Host '----' -ForegroundColor DarkGray
    Write-Host 'Heads-up to share with each customer admin:' -ForegroundColor Cyan
    Write-Host '  After they sign in and click Accept on the consent URL, Microsoft redirects them'
    Write-Host '  to login.microsoftonline.com/common/oauth2/nativeclient (or /common/wrongplace).'
    Write-Host '  The page literally says "This is not the right page" or warns about phishing.'
    Write-Host '  That is the expected, normal endpoint for first-party admin-consent flows — the'
    Write-Host '  consent landed correctly. They can close the tab.' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'If a customer admin can''t use the URL (e.g. admin-consent-only policies block the' -ForegroundColor Cyan
    Write-Host 'Microsoft Graph PowerShell first-party client), they can run this once in PowerShell 7' -ForegroundColor Cyan
    Write-Host 'from inside their own tenant:' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '    Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","Directory.Read.All", `' -ForegroundColor White
    Write-Host '      "Organization.Read.All","Domain.Read.All","Policy.Read.All", `' -ForegroundColor White
    Write-Host '      "RoleManagement.Read.Directory","AuditLog.Read.All","Reports.Read.All"' -ForegroundColor White
    Write-Host ''
    Write-Host '  …signing in as a Cloud Application Administrator and answering Yes to "Consent on' -ForegroundColor Cyan
    Write-Host '  behalf of your organization." That records the same SP + grants in their tenant.' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Re-test a single tenant after consent with:' -ForegroundColor Cyan
    Write-Host '  ./scripts/Test-TenantPreflight.ps1 -OnlyTenant ''<display name or tenant id>''' -ForegroundColor White
    Write-Host ''
}

if ($wantCsv) { Write-PreflightCsv -Rows $allResults -Path $CsvPath }
