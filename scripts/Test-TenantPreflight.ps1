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

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1
    # Test every GDAP customer; print punchlist or success message.

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1 -OnlyTenant 'Acme Corp'
    # Re-test a single customer after their admin consented.

.EXAMPLE
    ./scripts/Test-TenantPreflight.ps1 -OnlyTenant 00000000-0000-0000-0000-000000000000
    # Same, by tenant id.

.NOTES
    License: Apache-2.0 (see LICENSE in repo root)
    Created: 2026-05-05
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [string] $ConfigPath = (Join-Path $PSScriptRoot '..' 'tenants.config.local.json'),
    [string] $OnlyTenant
)

$ErrorActionPreference = 'Stop'

# Well-known clientIds. We acquire a refresh token via device-code flow for
# Microsoft Graph PowerShell; that refresh token is then used to test each
# customer tenant's /token endpoint. The Teams clientId is included only
# in the punchlist consent URLs — we don't authenticate against it here.
$GraphPsClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$TeamsPsClientId = '12128f48-ec9e-42f0-b203-ea49fb6af367'

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
        # representative of v1's actual scope set is consented.
    }
    catch {
        $err = $null
        try { $err = $_.ErrorDetails.Message | ConvertFrom-Json } catch {}
        $errCode = if ($err) { $err.error } else { '' }
        $errDesc = if ($err) { $err.error_description } else { $_.Exception.Message }

        $status = $null
        $consentUrls = @()

        # Both AADSTS90099 (no SP at all) and AADSTS65001 (SP exists but the
        # requested scopes aren't consented in that customer tenant) are
        # fixable the same way: a Cloud Application Administrator in the
        # customer tenant has to consent the app + scopes once.
        #
        # Empirically: consenting just the Microsoft Graph PowerShell client
        # is enough to clear preflight for both Graph and Teams. The Teams
        # PS client uses Graph for inventory-relevant data anyway. So we
        # only emit the Graph URL.
        if ($errDesc -match 'AADSTS90099') {
            $status = 'NEEDS_CONSENT (no SP)'
            $consentUrls = @(
                "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$GraphPsClientId"
            )
        }
        elseif ($errDesc -match 'AADSTS65001') {
            $status = 'NEEDS_CONSENT (scopes)'
            $consentUrls = @(
                "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$GraphPsClientId"
            )
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
    }
}

Write-Host -NoNewline ("`r{0}" -f $ansiClearLine)
Write-Host ''

# ============================================================================
# Step 4: output
# ============================================================================

if (@($punchlist).Count -eq 0) {
    Write-Host ''
    Write-Host ('OK — all {0} customer tenant(s) passed the Directory.Read.All refresh-token probe.' -f $total) -ForegroundColor Green
    Write-Host '   Caveat: preflight uses a refresh-token grant; the wrapper uses interactive browser auth. They mostly agree, but a tenant can pass preflight and still fail the wrapper if the customer''s tenant requires admin consent on first interactive use of the elevated scope set. If that happens, the v1 collector logs the consent URL inline — same fix as a NEEDS_CONSENT punchlist entry.' -ForegroundColor DarkGray
    Write-Host '   Preflight also doesn''t separately verify the Microsoft Teams PowerShell client. Teams data flows through Graph for our inventory, so this rarely matters; if Teams data is empty for a tenant after a real run, send the customer admin the Teams consent URL too.' -ForegroundColor DarkGray
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
