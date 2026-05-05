<#
.SYNOPSIS
    Register the partner-center app used by
    Get-O365MailGroupInventory-Multi.ps1 — creates the Entra app
    registration, generates a self-signed cert, attaches it as a key
    credential, creates the corresponding service principal, and prints
    the values needed for tenants.config.json plus the admin-consent URL.

.DESCRIPTION
    Run this in your partner-tenant Entra ID as a Global Administrator.
    Cross-platform: works on Windows, macOS, and Linux PowerShell 7+. The
    cert is persisted as a password-protected PFX file on disk (not in any
    platform keystore), so the same script and the same cert work
    everywhere. Re-entrant across machines and re-runs.

    Multi-machine usage — two supported patterns:

    A) SHARED CERT (recommended).
       Generate the cert on Machine 1, store the .pfx where both machines
       can read it (e.g. shared encrypted storage your operators already
       trust), point both machines at the same -CertPfxPath. The first
       run registers the app + attaches the cert; every subsequent run
       on either machine detects "app exists, this cert is already
       attached" and exits cleanly with the values printed.

    B) PER-MACHINE CERT.
       Each machine has its own .pfx (different path on each machine, or
       the same default path with no shared storage). The first run on
       Machine 1 registers the app and attaches cert C1. The first run on
       Machine 2 generates cert C2, sees the app already exists, and after
       confirming with you, APPENDS C2 to the app's key credentials
       (alongside C1, never overwriting). Each machine connects with its
       own private key. Decommissioning a machine = removing only that
       cert from the app in Entra ID.

    What this script does, in order:
      1. Connect-MgGraph with Application.ReadWrite.All + Directory.Read.All
         (interactive, MFA-aware).
      2. Look up Microsoft Graph and Office 365 Exchange Online service
         principals to resolve the application-permission GUIDs at runtime
         (no hardcoded magic strings).
      3. Generate a self-signed cert (2048-bit RSA, SHA-256, configurable
         validity) using .NET CertificateRequest APIs and persist it as a
         password-protected .pfx — or reuse an existing .pfx at the same
         path. The cert is the only secret produced; protect the .pfx
         accordingly.
      4. Look up the app by displayName.
         - If app exists AND this cert is already attached:
             verify, patch missing pieces (signInAudience, missing perms,
             missing http://localhost redirect URI), print values, exit.
         - If app exists AND this cert is NOT attached:
             show currently-attached certs, prompt for confirmation, then
             APPEND this cert to the app's KeyCredentials (multi-machine
             pattern B).
         - If app does not exist:
             create it with the right requiredResourceAccess block,
             publicClient.redirectUris = http://localhost (required for
             delegated interactive sign-in), attach the cert as the first
             key credential, create the service principal.
      5. Grant admin consent programmatically by creating an
         appRoleAssignment for each requested permission against the
         relevant resource service principal (Microsoft Graph, Office 365
         Exchange Online). Idempotent — already-granted assignments are
         skipped. Equivalent to clicking "Grant admin consent for <tenant>"
         in the portal API-permissions blade, but without the browser.
      6. Print ClientId, TenantId, CertificateThumbprint, PFX path, and a
         verification deep-link to the portal.

    The app supports two auth modes used by the multi-tenant wrapper:

      DELEGATED (default for staff):
        Multi-tenant audience (signInAudience = AzureADMultipleOrgs).
        Staff users sign in interactively against the partner-app clientId
        and rely on GDAP/Lighthouse role propagation in customer tenants.
        No cert needed for staff machines. Admin-consented tenant-wide
        once by the operator running this script so staff don't see
        consent prompts. Permissions consented as oauth2PermissionGrant
        with consentType=AllPrincipals.

      APP-ONLY (fallback for unattended runs):
        Cert-based, app acts on its own behalf. Used by the wrapper's
        -AppOnly switch for scheduled/background runs you trigger
        yourself. Permissions consented as appRoleAssignment.

    Both flavours are added to the same app registration in a single run.
    Delegated EXO has no app permission (Connect-ExchangeOnline
    -DelegatedOrganization uses first-party EXO PowerShell auth and
    bypasses our app), so Office 365 Exchange Online has only the
    application Exchange.ManageAsApp.

    Required application permissions:
       Microsoft Graph (00000003-0000-0000-c000-000000000000):
         Policy.Read.All, RoleManagement.Read.Directory, AuditLog.Read.All,
         UserAuthenticationMethod.Read.All, Application.Read.All,
         Directory.Read.All, Group.Read.All, GroupMember.Read.All,
         User.Read.All, Reports.Read.All, Sites.Read.All,
         Team.ReadBasic.All, Channel.ReadBasic.All,
         DelegatedAdminRelationship.Read.All
       Office 365 Exchange Online (00000002-0000-0ff1-ce00-000000000000):
         Exchange.ManageAsApp

    Required delegated permissions:
       Microsoft Graph: same 14 permission names as Application above.

    Required Connect-MgGraph scopes (you, the operator, must consent):
       Application.ReadWrite.All
       Directory.Read.All
       AppRoleAssignment.ReadWrite.All       (Application consent)
       DelegatedPermissionGrant.ReadWrite.All (Delegated consent)

    What this script does NOT do:
      - Configure GDAP role templates. That's a Partner Center step done
        via the relationship-request flow, not by this app's manifest.
      - Move .pfx files between machines for you. Either pre-position the
        .pfx on shared storage (pattern A), or generate per-machine and
        let the script append (pattern B).

.PARAMETER AppDisplayName
    Display name for the app registration in Entra ID.
    Default: 'M365 Multi-Tenant Inventory'.

.PARAMETER CertSubject
    CN= subject for the self-signed cert.
    Default: 'CN=M365-Multi-Tenant-Inventory'.

.PARAMETER CertPfxPath
    Path to the .pfx file holding the cert + private key. If the file
    exists, it is reused (you'll be prompted for its password). Otherwise
    a new cert is generated and written here. Default:
    ~/.m365-multi-tenant-inventory/partner-app.pfx (cross-platform;
    $HOME on macOS/Linux, $env:USERPROFILE on Windows).

.PARAMETER CertValidYears
    How many years the self-signed cert is valid for. Default 2.

.PARAMETER ForceNewCert
    If set, generate a new cert even if a .pfx exists at -CertPfxPath.
    The existing file is overwritten.

.PARAMETER ConfigPathToUpdate
    Optional path to a tenants.config.json file. If set, the script will
    overwrite the partner.{homeTenantId,clientId,certificateThumbprint}
    fields in that file with the values it just produced. Useful for the
    real-config copy outside the repo. Default: $null (don't touch any
    file — print to console only).

.PARAMETER AppendCertSilently
    When the app already exists and this cert is not attached, append it
    without prompting. Useful for non-interactive multi-machine setup.
    Default: $false (prompt for confirmation).

.EXAMPLE
    # Standard run: register the app, generate a new cert at the default
    # path, print values, do not touch any config file.
    ./scripts/Register-PartnerCenterApp.ps1

.EXAMPLE
    # Multi-machine, shared-cert pattern (pattern A). Run identically on
    # every machine, pointing at the same OneDrive-synced .pfx path.
    # First run registers + attaches; later runs verify and exit.
    ./scripts/Register-PartnerCenterApp.ps1 `
        -CertPfxPath '<path to shared encrypted storage>/partner-app.pfx'

.EXAMPLE
    # Multi-machine, per-machine-cert pattern (pattern B). Generates a
    # cert at the local default path; on the second machine, the script
    # detects the app exists, prompts you, and appends the new cert to
    # its key credentials.
    ./scripts/Register-PartnerCenterApp.ps1
    # ...and on subsequent machines, with no prompt:
    ./scripts/Register-PartnerCenterApp.ps1 -AppendCertSilently

.NOTES
    License: Apache-2.0 (see LICENSE in repo root)
    Created: 2026-05-05

    Required PowerShell modules (auto-installed if missing):
      Microsoft.Graph.Authentication      2.0+
      Microsoft.Graph.Applications        2.0+
      Microsoft.Graph.Identity.SignIns    2.0+   (oauth2PermissionGrant cmdlets)
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [string] $AppDisplayName     = 'M365 Multi-Tenant Inventory',
    [string] $CertSubject        = 'CN=M365-Multi-Tenant-Inventory',
    [string] $CertPfxPath        = (Join-Path $HOME '.m365-multi-tenant-inventory/partner-app.pfx'),
    [int]    $CertValidYears     = 2,
    [switch] $ForceNewCert,
    [string] $ConfigPathToUpdate,
    [switch] $AppendCertSilently
)

$ErrorActionPreference = 'Stop'

# Permissions we need, by resource and flavour. Resolved at runtime against
# live service principals — no hardcoded permission GUIDs; we only know the
# well-known resource AppIds.
#
#   Application = appRole (Role) — used in -AppOnly mode (cert auth)
#   Delegated   = oauth2PermissionScope (Scope) — used in default delegated
#                 mode where staff users sign in interactively and rely on
#                 GDAP/Lighthouse for customer-tenant authorization.
#
# Office 365 Exchange Online has no Delegated entries because staff
# Connect-ExchangeOnline uses -DelegatedOrganization, which authenticates
# the running user via Microsoft's first-party EXO PowerShell auth and
# does NOT consume our app's permissions. App-only EXO (-AppOnly mode)
# does need Exchange.ManageAsApp.
$RequiredPermissions = @{
    # Microsoft Graph
    '00000003-0000-0000-c000-000000000000' = @{
        Application = @(
            'Policy.Read.All',
            'RoleManagement.Read.Directory',
            'AuditLog.Read.All',
            'UserAuthenticationMethod.Read.All',
            'Application.Read.All',
            'Directory.Read.All',
            'Group.Read.All',
            'GroupMember.Read.All',
            'User.Read.All',
            'Reports.Read.All',
            'Sites.Read.All',
            'Team.ReadBasic.All',
            'Channel.ReadBasic.All',
            'DelegatedAdminRelationship.Read.All'
        )
        Delegated = @(
            'Policy.Read.All',
            'RoleManagement.Read.Directory',
            'AuditLog.Read.All',
            'UserAuthenticationMethod.Read.All',
            'Application.Read.All',
            'Directory.Read.All',
            'Group.Read.All',
            'GroupMember.Read.All',
            'User.Read.All',
            'Reports.Read.All',
            'Sites.Read.All',
            'Team.ReadBasic.All',
            'Channel.ReadBasic.All',
            'DelegatedAdminRelationship.Read.All'
        )
    }
    # Office 365 Exchange Online — Application only
    '00000002-0000-0ff1-ce00-000000000000' = @{
        Application = @('Exchange.ManageAsApp')
        Delegated   = @()
    }
}

function Write-Step { param([string]$Msg) Write-Host "==> $Msg" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Msg) Write-Host "    $Msg" -ForegroundColor Green }
function Write-Warn { param([string]$Msg) Write-Host "    $Msg" -ForegroundColor Yellow }

# ============================================================================
# Module bootstrap
# ============================================================================

Write-Step 'Bootstrapping modules...'
foreach ($mod in @('Microsoft.Graph.Authentication','Microsoft.Graph.Applications','Microsoft.Graph.Identity.SignIns')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Warn "Installing $mod (CurrentUser scope)..."
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -ErrorAction Stop
}
Write-Ok 'Modules ready.'

# ============================================================================
# Connect to Graph as a Global Admin (interactive)
# ============================================================================

Write-Step 'Connecting to Microsoft Graph (interactive, MFA-aware)...'
# AppRoleAssignment.ReadWrite.All  — to grant Application permissions (appRoleAssignment)
# DelegatedPermissionGrant.ReadWrite.All — to admin-consent Delegated permissions tenant-wide (oauth2PermissionGrant)
Connect-MgGraph -Scopes 'Application.ReadWrite.All','Directory.Read.All','AppRoleAssignment.ReadWrite.All','DelegatedPermissionGrant.ReadWrite.All' -NoWelcome
$ctx = Get-MgContext
if (-not $ctx) { throw 'Connect-MgGraph did not produce a context.' }
Write-Ok ("Tenant: {0}  ({1})" -f $ctx.TenantId, $ctx.Account)

# ============================================================================
# Resolve permission GUIDs at runtime
# ============================================================================

Write-Step 'Resolving requiredResourceAccess block from live service principals...'
$resourceAccessList = @()
# Cache resolved resource SPs + appRole/scope objects so the consent step
# below can reuse them without re-fetching.
$script:ResolvedResources = @{}

foreach ($resourceAppId in $RequiredPermissions.Keys) {
    $resourceSp = Get-MgServicePrincipal -Filter "appId eq '$resourceAppId'" -ErrorAction Stop | Select-Object -First 1
    if (-not $resourceSp) { throw "Could not resolve service principal for resource appId $resourceAppId in this tenant." }

    $resourceAccess = @()
    $resolvedAppRoles = @()
    $resolvedScopes   = @()

    # Application permissions → type='Role'
    foreach ($permName in $RequiredPermissions[$resourceAppId].Application) {
        $appRole = $resourceSp.AppRoles | Where-Object { $_.Value -eq $permName -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $appRole) { throw "Application permission '$permName' not found on resource $($resourceSp.DisplayName) ($resourceAppId)." }
        $resourceAccess  += @{ id = $appRole.Id; type = 'Role' }
        $resolvedAppRoles += $appRole
    }

    # Delegated permissions → type='Scope'
    foreach ($permName in $RequiredPermissions[$resourceAppId].Delegated) {
        $scope = $resourceSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $permName } | Select-Object -First 1
        if (-not $scope) { throw "Delegated permission '$permName' not found on resource $($resourceSp.DisplayName) ($resourceAppId)." }
        $resourceAccess += @{ id = $scope.Id; type = 'Scope' }
        $resolvedScopes += $scope
    }

    $resourceAccessList += @{
        resourceAppId  = $resourceAppId
        resourceAccess = $resourceAccess
    }
    $script:ResolvedResources[$resourceAppId] = @{
        ServicePrincipal = $resourceSp
        AppRoles         = $resolvedAppRoles
        Scopes           = $resolvedScopes
    }
    Write-Ok ("{0}: {1} app perm(s), {2} delegated perm(s) resolved." -f $resourceSp.DisplayName, $resolvedAppRoles.Count, $resolvedScopes.Count)
}

# ============================================================================
# Generate (or reuse) the self-signed cert — cross-platform via .NET
# ============================================================================
# We deliberately do NOT use New-SelfSignedCertificate / Cert:\CurrentUser\My /
# Export-PfxCertificate — those are part of the Windows-only PKI module. The
# .NET CertificateRequest + X509Certificate2 APIs are present on .NET 6+ and
# work identically on Windows, macOS, and Linux. The cert lives as a
# password-protected .pfx file at $CertPfxPath; downstream Connect-MgGraph and
# Connect-ExchangeOnline calls accept either a thumbprint (when the cert is
# imported into a keystore) or an X509Certificate2 object loaded from a file.

# Ensure parent directory exists
$pfxDir = Split-Path -Parent $CertPfxPath
if ($pfxDir -and -not (Test-Path $pfxDir)) {
    New-Item -ItemType Directory -Path $pfxDir -Force | Out-Null
}

# Pin the X509KeyStorageFlags enum locally for readability.
$exportable = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable -bor `
              [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet

if ((Test-Path $CertPfxPath) -and -not $ForceNewCert) {
    Write-Step "Loading existing cert from $CertPfxPath..."
    $pfxPwd = Read-Host '    PFX password (will not echo)' -AsSecureString
    $cert   = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $CertPfxPath, $pfxPwd, $exportable
    )
} else {
    Write-Step "Generating new self-signed cert and writing to $CertPfxPath..."
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        $CertSubject,
        $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    $notBefore = [DateTimeOffset]::Now
    $notAfter  = $notBefore.AddYears($CertValidYears)
    $cert      = $req.CreateSelfSigned($notBefore, $notAfter)

    Write-Host '    Set a password for the new .pfx (will not echo). You will need this'
    Write-Host '    password every time the wrapper script loads the cert.'
    $pfxPwd = Read-Host '    PFX password' -AsSecureString
    $bytes  = $cert.Export(
        [System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx,
        $pfxPwd
    )
    [System.IO.File]::WriteAllBytes($CertPfxPath, $bytes)

    # Tighten file permissions on POSIX (best-effort; ignored on Windows).
    if ($IsMacOS -or $IsLinux) {
        try { & chmod 600 $CertPfxPath } catch { Write-Warn "Could not chmod 600 on $CertPfxPath — set it manually." }
    }
}
Write-Ok ("Subject:    {0}" -f $cert.Subject)
Write-Ok ("Thumbprint: {0}" -f $cert.Thumbprint)
Write-Ok ("NotAfter:   {0}" -f $cert.NotAfter.ToString('yyyy-MM-dd'))
Write-Ok ("PFX path:   {0}" -f $CertPfxPath)

# ============================================================================
# Locate or create the app — multi-machine re-entrant logic
# ============================================================================
# The app exists once across all machines. On each machine, this script
# decides between three branches:
#
#   1. App doesn't exist        → create it, attach this cert as the first
#                                 key credential, create SP. (First-ever run.)
#   2. App exists, this cert is
#      already in its KeyCreds  → idempotent re-run; nothing to mutate.
#   3. App exists, this cert is
#      NOT in its KeyCreds      → APPEND this cert (don't overwrite). Used
#                                 in the per-machine-cert pattern.
#
# Match is by SHA-1 thumbprint. Graph stores the cert thumbprint in each
# key credential's CustomKeyIdentifier as raw bytes; we hex it and compare
# to $cert.Thumbprint (also uppercase hex).

Write-Step "Looking up app '$AppDisplayName'..."
$existing = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ErrorAction SilentlyContinue | Select-Object -First 1

$ourThumbHex = $cert.Thumbprint   # uppercase hex, no separators

function Test-CertAttached {
    param($KeyCredentials, [string]$ThumbHex)
    foreach ($k in @($KeyCredentials)) {
        if (-not $k.CustomKeyIdentifier) { continue }
        $existingHex = [Convert]::ToHexString([byte[]]$k.CustomKeyIdentifier)
        if ($existingHex -eq $ThumbHex) { return $true }
    }
    return $false
}

if ($existing) {
    Write-Ok ("App already exists. AppId (ClientId): {0}" -f $existing.AppId)
    Write-Ok ("                   ObjectId:         {0}" -f $existing.Id)

    # Patch signInAudience to multi-tenant if it's still single-tenant. Required
    # for staff to sign in via the partner-app clientId against customer tenants.
    if ($existing.SignInAudience -ne 'AzureADMultipleOrgs') {
        Write-Step "Patching signInAudience: $($existing.SignInAudience) -> AzureADMultipleOrgs..."
        Update-MgApplication -ApplicationId $existing.Id -SignInAudience 'AzureADMultipleOrgs'
        Write-Ok 'signInAudience updated.'
    }

    # Patch requiredResourceAccess. We always SET to the full desired list —
    # safe because no other process modifies this app's manifest. The list
    # includes both Application (Role) and Delegated (Scope) entries built
    # in the resolution pass above.
    Write-Step 'Ensuring requiredResourceAccess includes all Application + Delegated permissions...'
    Update-MgApplication -ApplicationId $existing.Id -RequiredResourceAccess $resourceAccessList
    Write-Ok 'requiredResourceAccess set.'

    # Patch publicClient.redirectUris. Required for delegated interactive
    # sign-in via Connect-MgGraph -ClientId <partner-app>: MSAL spins up a
    # local loopback server to receive the OAuth callback, and Entra refuses
    # to redirect back to a URI that's not registered on the app
    # (AADSTS500113: No reply address is registered for the application).
    $existingRedirects = @($existing.PublicClient.RedirectUris)
    if ($existingRedirects -notcontains 'http://localhost') {
        Write-Step 'Adding http://localhost to publicClient.redirectUris...'
        $combined = @($existingRedirects + 'http://localhost' | Where-Object { $_ } | Select-Object -Unique)
        Update-MgApplication -ApplicationId $existing.Id -PublicClient @{ RedirectUris = $combined }
        Write-Ok 'publicClient.redirectUris updated.'
    }

    # Refresh the app object to pick up the updates above.
    $existing = Get-MgApplication -ApplicationId $existing.Id

    if (Test-CertAttached -KeyCredentials $existing.KeyCredentials -ThumbHex $ourThumbHex) {
        # Branch 2 — idempotent
        Write-Ok "This cert (thumbprint $ourThumbHex) is already attached. Nothing to mutate."
        $app = $existing
    } else {
        # Branch 3 — append
        Write-Warn "App exists, but this machine's cert ($ourThumbHex) is NOT attached."
        Write-Host '    Currently attached cert(s):' -ForegroundColor DarkGray
        if (-not $existing.KeyCredentials -or @($existing.KeyCredentials).Count -eq 0) {
            Write-Host '      (none)' -ForegroundColor DarkGray
        } else {
            foreach ($k in @($existing.KeyCredentials)) {
                $hex = if ($k.CustomKeyIdentifier) { [Convert]::ToHexString([byte[]]$k.CustomKeyIdentifier) } else { '(no CustomKeyIdentifier)' }
                $end = if ($k.EndDateTime) { ([DateTime]$k.EndDateTime).ToString('yyyy-MM-dd') } else { '?' }
                Write-Host ("      - {0}  DisplayName: {1}  NotAfter: {2}" -f $hex, $k.DisplayName, $end) -ForegroundColor DarkGray
            }
        }

        if (-not $AppendCertSilently) {
            $ans = Read-Host '    Append this machine''s cert as an additional key credential? (yes/no)'
            if ($ans.Trim().ToLower() -ne 'yes') { throw 'Aborted by user — no changes made.' }
        }

        Write-Step 'Appending cert to existing app key credentials...'
        $newKey = @{
            Type        = 'AsymmetricX509Cert'
            Usage       = 'Verify'
            Key         = $cert.RawData
            DisplayName = $CertSubject
        }
        # KeyCredentials must be set as a complete list — copy the existing
        # ones forward so we don't drop them.
        $combined = @()
        foreach ($k in @($existing.KeyCredentials)) {
            $combined += @{
                Type                = $k.Type
                Usage               = $k.Usage
                Key                 = $k.Key
                DisplayName         = $k.DisplayName
                StartDateTime       = $k.StartDateTime
                EndDateTime         = $k.EndDateTime
                KeyId               = $k.KeyId
                CustomKeyIdentifier = $k.CustomKeyIdentifier
            }
        }
        $combined += $newKey

        Update-MgApplication -ApplicationId $existing.Id -KeyCredentials $combined
        Write-Ok ("Appended. App now has {0} key credential(s)." -f $combined.Count)
        $app = $existing
    }

    # Service principal lookup — should always exist for an existing app.
    $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $sp) {
        Write-Step 'No service principal found for existing app — creating one...'
        $sp = New-MgServicePrincipal -AppId $app.AppId
    }
    Write-Ok ("ServicePrincipalId: {0}" -f $sp.Id)
}
else {
    # Branch 1 — first run anywhere, create the app
    Write-Step 'App does not exist; creating it...'
    # Multi-tenant audience: required so customer tenants will mint tokens for
    # this clientId when staff sign in via GDAP-mediated delegated flows.
    # publicClient.redirectUris: required for delegated interactive sign-in
    # via Connect-MgGraph -ClientId <partner-app>; MSAL spins up a local
    # loopback callback. Without this, sign-in fails with AADSTS500113.
    $app = New-MgApplication `
        -DisplayName            $AppDisplayName `
        -SignInAudience         'AzureADMultipleOrgs' `
        -RequiredResourceAccess $resourceAccessList `
        -PublicClient           @{ RedirectUris = @('http://localhost') }
    Write-Ok ("AppId (ClientId): {0}" -f $app.AppId)
    Write-Ok ("ObjectId:         {0}" -f $app.Id)

    Write-Step 'Attaching cert as the first key credential...'
    $keyCred = @{
        Type        = 'AsymmetricX509Cert'
        Usage       = 'Verify'
        Key         = $cert.RawData
        DisplayName = $CertSubject
    }
    Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCred)
    Write-Ok 'Key credential attached.'

    Write-Step 'Creating service principal in this tenant...'
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Ok ("ServicePrincipalId: {0}" -f $sp.Id)
}

# ============================================================================
# Grant admin consent programmatically — replaces the broken /adminconsent URL
# ============================================================================
# Two flavours of consent to handle:
#   1. Application permissions  → appRoleAssignment on our SP
#                                  (principalId = appSP, resourceId = resourceSP,
#                                   appRoleId = the appRole)
#   2. Delegated permissions    → oauth2PermissionGrant on our SP with
#                                  consentType='AllPrincipals' (admin-consent
#                                  tenant-wide so staff don't see prompts)
#                                  scope = space-separated permission names
#
# Both are idempotent: existing grants are detected and skipped. Either or
# both branches may produce zero new grants on a re-run; that's fine.

Write-Step 'Granting admin consent (Application + Delegated)...'

$existingAppRoleAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue)
$existingOauth2Grants = @(
    Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)'" -All -ErrorAction SilentlyContinue
)
$grantedCount      = 0
$skippedCount      = 0
$failedAssignments = @()

foreach ($resourceAppId in $RequiredPermissions.Keys) {
    $resolved = $script:ResolvedResources[$resourceAppId]
    if (-not $resolved) { continue }
    $resourceSp = $resolved.ServicePrincipal

    # ---- Application permissions: appRoleAssignment per perm ----
    foreach ($appRole in $resolved.AppRoles) {
        $alreadyAssigned = $existingAppRoleAssignments | Where-Object {
            $_.ResourceId -eq $resourceSp.Id -and $_.AppRoleId -eq $appRole.Id
        } | Select-Object -First 1

        if ($alreadyAssigned) { $skippedCount++; continue }

        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id `
                -PrincipalId        $sp.Id `
                -ResourceId         $resourceSp.Id `
                -AppRoleId          $appRole.Id | Out-Null
            $grantedCount++
        }
        catch {
            $failedAssignments += "$($resourceSp.DisplayName)/App/$($appRole.Value) ($($_.Exception.Message))"
        }
    }

    # ---- Delegated permissions: one oauth2PermissionGrant per resource ----
    if ($resolved.Scopes.Count -gt 0) {
        $wantedScopeNames = @($resolved.Scopes | ForEach-Object { $_.Value })
        $wantedScopeText  = $wantedScopeNames -join ' '

        $existingGrant = $existingOauth2Grants | Where-Object {
            $_.ResourceId -eq $resourceSp.Id -and $_.ConsentType -eq 'AllPrincipals'
        } | Select-Object -First 1

        if ($existingGrant) {
            # Compute the union: existing scopes + any wanted scopes not yet in the grant.
            $existingScopeNames = @(($existingGrant.Scope -split ' ') | Where-Object { $_ })
            $missingScopeNames  = @($wantedScopeNames | Where-Object { $existingScopeNames -notcontains $_ })

            if ($missingScopeNames.Count -eq 0) {
                $skippedCount += $wantedScopeNames.Count
            } else {
                try {
                    $unionScopeText = (($existingScopeNames + $missingScopeNames) | Select-Object -Unique) -join ' '
                    Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $existingGrant.Id -Scope $unionScopeText | Out-Null
                    $grantedCount += $missingScopeNames.Count
                    $skippedCount += ($wantedScopeNames.Count - $missingScopeNames.Count)
                }
                catch {
                    foreach ($n in $missingScopeNames) {
                        $failedAssignments += "$($resourceSp.DisplayName)/Del/$n ($($_.Exception.Message))"
                    }
                }
            }
        } else {
            try {
                New-MgOauth2PermissionGrant `
                    -ClientId    $sp.Id `
                    -ConsentType 'AllPrincipals' `
                    -ResourceId  $resourceSp.Id `
                    -Scope       $wantedScopeText | Out-Null
                $grantedCount += $wantedScopeNames.Count
            }
            catch {
                foreach ($n in $wantedScopeNames) {
                    $failedAssignments += "$($resourceSp.DisplayName)/Del/$n ($($_.Exception.Message))"
                }
            }
        }
    }
}

Write-Ok ("Consent grants — newly granted: {0}; already-granted (skipped): {1}; failed: {2}" -f $grantedCount, $skippedCount, $failedAssignments.Count)
if ($failedAssignments.Count -gt 0) {
    Write-Warn 'Some grants failed. You can verify and grant the rest in the portal:'
    foreach ($f in $failedAssignments) { Write-Warn "  - $f" }
}

# ============================================================================
# Optionally update a tenants.config.json
# ============================================================================

if ($ConfigPathToUpdate) {
    Write-Step "Patching $ConfigPathToUpdate with partner block values..."
    if (-not (Test-Path $ConfigPathToUpdate)) {
        throw "Config path '$ConfigPathToUpdate' does not exist. Create the file first (copy from tenants.config.json in the repo)."
    }
    $cfgRaw = Get-Content -Raw -Path $ConfigPathToUpdate
    $cfg    = $cfgRaw | ConvertFrom-Json
    $cfg.partner.homeTenantId          = $ctx.TenantId
    $cfg.partner.clientId              = $app.AppId
    $cfg.partner.certificateThumbprint = $cert.Thumbprint
    $cfg.partner.certificatePfxPath    = $CertPfxPath
    $cfg | ConvertTo-Json -Depth 12 | Set-Content -Path $ConfigPathToUpdate -NoNewline
    Write-Ok 'Config updated. Confirm the file is excluded from any git index.'
}

# ============================================================================
# Print results
# ============================================================================

# Verification deep-link to the API permissions blade. NOT used to grant
# consent — consent was already granted programmatically above. This is for
# eyes-on confirmation: every row should show "Granted for <tenant>".
$verifyUrl = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)"

Write-Host ''
Write-Host '────────────────────────────────────────────────────────────────────────────'
Write-Host 'Partner app registered. Values below go into tenants.config.json -> partner:'
Write-Host '────────────────────────────────────────────────────────────────────────────'
Write-Host ('  homeTenantId          = "{0}"' -f $ctx.TenantId)
Write-Host ('  clientId              = "{0}"' -f $app.AppId)
Write-Host ('  certificateThumbprint = "{0}"' -f $cert.Thumbprint)
Write-Host ('  certificatePfxPath    = "{0}"' -f $CertPfxPath)
Write-Host ''
Write-Host 'The wrapper resolves the PFX password from the env var named in'
Write-Host '  partner.certificatePfxPasswordEnvVar  (default: M365_MULTI_PFX_PASSWORD)'
Write-Host 'or prompts interactively if the env var is unset. Set it in your shell:'
Write-Host '  $env:M365_MULTI_PFX_PASSWORD = ''<your-pfx-password>''  # PowerShell'
Write-Host ''
if ($failedAssignments.Count -eq 0) {
    Write-Host 'Admin consent: GRANTED programmatically. No manual portal step required.' -ForegroundColor Green
} else {
    Write-Host 'Admin consent: PARTIAL. See warnings above and grant the missing permissions in the portal.' -ForegroundColor Yellow
}
Write-Host 'Verify in the portal (API permissions should all show "Granted for <tenant>"):' -ForegroundColor DarkGray
Write-Host "  $verifyUrl" -ForegroundColor DarkGray
Write-Host ''
Write-Host 'GDAP REMINDER: this app uses customer-side roles granted via your GDAP'
Write-Host 'relationship templates, NOT by consenting in each customer tenant. Confirm'
Write-Host 'every active GDAP relationship grants at minimum:'
Write-Host '  - Global Reader'
Write-Host '  - Exchange Recipient Administrator (or Exchange Administrator)'
Write-Host '  - Privileged Role Reader'
Write-Host '────────────────────────────────────────────────────────────────────────────'

Disconnect-MgGraph | Out-Null
