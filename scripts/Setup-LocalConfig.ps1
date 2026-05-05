<#
.SYNOPSIS
    Interactive setup for tenants.config.local.json — prompts for the values
    the wrapper needs and writes them to a gitignored config file.

.DESCRIPTION
    Run this once on each operator machine before the first run of
    Get-O365MailGroupInventory-Multi.ps1. The script reads the schema-example
    `tenants.config.json` (kept in version control with placeholders),
    prompts you for the values that need to be filled in, validates each
    input, and writes the populated result to `tenants.config.local.json`
    (gitignored by default, lives outside the index).

    By default, the script begins with an interactive sign-in to Microsoft
    Graph that auto-discovers your home tenant id and the partner-app
    clientId from your tenant's app registrations. The discovered values
    pre-fill the prompts, so most operators just press Enter to confirm. If
    auto-discovery fails (consent not granted, app not registered yet, or
    multiple ambiguous candidates) the script falls back to manual entry
    cleanly.

    Idempotent. If `tenants.config.local.json` already exists, current
    values are shown as defaults and you can press Enter to keep them —
    useful for fixing one field without re-typing the rest. Existing config
    values take precedence over auto-discovered ones.

    What it asks for:

      Always required (delegated mode is the default staff workflow):
        - Partner-app clientId          (GUID; auto-discovered from your
                                         tenant's app registrations when
                                         Application.Read.All is consented,
                                         else prompt-only)
        - Home tenant id                (GUID; auto-discovered from your
                                         signed-in Microsoft Graph context)

      Required only with -AppOnly:
        - Certificate thumbprint        (40 hex characters, from registration output)
        - PFX file path                 (path to the cert; default: ~/.m365-multi-tenant-inventory/partner-app.pfx)
        - PFX password env-var name     (default: M365_MULTI_PFX_PASSWORD)

    What it does NOT ask for:
        - The PFX password itself. That's a runtime env var, never written
          to a config file. Set it in your shell before running -AppOnly:
              $env:M365_MULTI_PFX_PASSWORD = '<your-pfx-password>'

    What it does NOT do:
        - Register the Entra app. That's scripts/Register-PartnerCenterApp.ps1
          (run once by your tenant's Global Admin for all operators).
        - Generate certs. That's also Register-PartnerCenterApp.ps1.
        - Install PowerShell modules. The wrapper handles that on first run.

.PARAMETER ConfigSource
    Schema example used as the structural template. Default:
    ../tenants.config.json (relative to this script's directory).

.PARAMETER ConfigDestination
    Where to write the populated config. Default:
    ../tenants.config.local.json (gitignored).

.PARAMETER AppOnly
    Also prompt for cert + PFX values used by the wrapper's -AppOnly path.
    Default: skip cert prompts (delegated-only setup, sufficient for staff).

.PARAMETER NonInteractive
    Skip all prompts; require every value via parameter. Useful for
    automated provisioning. Pair with -ClientId, -HomeTenantId, etc.

.PARAMETER NoAutoDiscover
    Skip the interactive Microsoft Graph sign-in that auto-discovers
    home tenant id + clientId. Use when you know your tenant doesn't
    have Application.Read.All consented for the Microsoft Graph
    PowerShell client and you'd rather not waste 5–10 seconds on the
    failed attempt.

.PARAMETER ClientId
    Partner-app client id (when -NonInteractive).

.PARAMETER HomeTenantId
    Home tenant id (when -NonInteractive).

.PARAMETER CertificateThumbprint
    Cert thumbprint (when -NonInteractive -AppOnly).

.PARAMETER CertificatePfxPath
    PFX path (when -NonInteractive -AppOnly).

.PARAMETER CertificatePfxPasswordEnvVar
    PFX password env-var name (when -NonInteractive -AppOnly). Default:
    M365_MULTI_PFX_PASSWORD.

.EXAMPLE
    ./scripts/Setup-LocalConfig.ps1
    # Default: signs you in once, auto-discovers home tenant id and the
    # partner-app clientId, prompts you to confirm each (press Enter to
    # accept the auto-discovered values), writes tenants.config.local.json.

.EXAMPLE
    ./scripts/Setup-LocalConfig.ps1 -NoAutoDiscover
    # Skip the auto-discovery sign-in. Two manual prompts only.

.EXAMPLE
    ./scripts/Setup-LocalConfig.ps1 -AppOnly
    # Full setup including cert path, thumbprint, PFX-password env-var name.

.EXAMPLE
    ./scripts/Setup-LocalConfig.ps1 -ConfigDestination "$HOME/.m365-multi-tenant-inventory/tenants.config.local.json"
    # Write the config outside the repo entirely. Belt-and-braces if you
    # don't trust .gitignore alone.

.EXAMPLE
    ./scripts/Setup-LocalConfig.ps1 -NonInteractive `
        -ClientId 00000000-0000-0000-0000-000000000000 `
        -HomeTenantId 11111111-1111-1111-1111-111111111111
    # Automated provisioning — no prompts.

.NOTES
    License: Apache-2.0 (see LICENSE in repo root)
    Created: 2026-05-05
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [string] $ConfigSource                  = (Join-Path $PSScriptRoot '..' 'tenants.config.json'),
    [string] $ConfigDestination             = (Join-Path $PSScriptRoot '..' 'tenants.config.local.json'),
    [switch] $AppOnly,
    [switch] $NonInteractive,
    [switch] $NoAutoDiscover,
    [string] $ClientId,
    [string] $HomeTenantId,
    [string] $CertificateThumbprint,
    [string] $CertificatePfxPath,
    [string] $CertificatePfxPasswordEnvVar  = 'M365_MULTI_PFX_PASSWORD'
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Msg) Write-Host "==> $Msg" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Msg) Write-Host "    $Msg" -ForegroundColor Green }
function Write-Warn { param([string]$Msg) Write-Host "    $Msg" -ForegroundColor Yellow }

# ============================================================================
# Validators
# ============================================================================

$GuidRegex  = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
$ThumbRegex = '^[0-9a-fA-F]{40}$'
$PlaceholderMarker = '^<.*>$'

function Test-Guid { param([string]$v) if ($v -notmatch $GuidRegex)  { return 'must be a GUID, e.g. 12345678-1234-1234-1234-123456789012' } }
function Test-Thumb { param([string]$v) if ($v -notmatch $ThumbRegex) { return 'must be 40 hex characters (the cert SHA-1 thumbprint, no colons or spaces)' } }
function Test-NonPlaceholder { param([string]$v) if ($v -match $PlaceholderMarker) { return 'looks like a placeholder; supply the real value' } }

function Read-Required {
    param(
        [Parameter(Mandatory)] [string]$Prompt,
        [string]$Default,
        [scriptblock]$Validator,
        [switch]$AllowEmpty
    )
    while ($true) {
        $shown = if ($Default) { " [$Default]" } else { '' }
        $entered = Read-Host "    $Prompt$shown"
        if (-not $entered -and $Default) { $entered = $Default }
        if (-not $entered) {
            if ($AllowEmpty) { return '' }
            Write-Warn '(required — please enter a value)'
            continue
        }
        # Always reject placeholder-looking inputs
        $err = Test-NonPlaceholder $entered
        if ($err) { Write-Warn $err; continue }
        if ($Validator) {
            $err = & $Validator $entered
            if ($err) { Write-Warn $err; continue }
        }
        return $entered
    }
}

# ============================================================================
# Load source + existing destination
# ============================================================================

if (-not (Test-Path $ConfigSource)) {
    throw "Source config not found at $ConfigSource. Expected the schema example tenants.config.json at the repo root."
}

Write-Step "Loading schema example from $ConfigSource..."
$source = Get-Content -Raw $ConfigSource | ConvertFrom-Json

$existing = $null
if (Test-Path $ConfigDestination) {
    Write-Step "Existing config found at $ConfigDestination — values will be shown as defaults."
    $existing = Get-Content -Raw $ConfigDestination | ConvertFrom-Json
}

function Default-From {
    param($obj, [string]$path)
    if (-not $obj) { return $null }
    $val = $obj
    foreach ($seg in ($path -split '\.')) {
        if (-not $val) { return $null }
        $val = $val.$seg
    }
    if (-not $val) { return $null }
    if ($val -match $PlaceholderMarker) { return $null }   # never use a placeholder as a default
    return $val
}

# ============================================================================
# Auto-discovery: sign in once, grab homeTenantId from the auth context, query
# app registrations to suggest clientId. Best-effort — falls back to manual
# prompts if the user lacks Application.Read.All consent or if the partner app
# can't be found unambiguously.
# ============================================================================

function Try-AutoDiscover {
    param([switch]$Skip)
    $result = @{ TenantId = $null; ClientId = $null }
    if ($Skip) { return $result }

    Write-Host ''
    Write-Host '=== Auto-discovery (sign-in to your tenant) ===' -ForegroundColor Cyan
    Write-Host '    Sign in with your tenant account; the script will read your'
    Write-Host '    tenant id and look for the partner-app registration.'
    Write-Host '    Press Ctrl+C to skip and enter values manually.'
    Write-Host ''

    # Bootstrap modules. Microsoft.Graph.Authentication is enough for the
    # tenant-id read; Microsoft.Graph.Applications is what gives us
    # Get-MgApplication for the clientId discovery.
    foreach ($mod in @('Microsoft.Graph.Authentication','Microsoft.Graph.Applications')) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            Write-Step "Installing $mod (CurrentUser scope)..."
            try {
                Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            } catch {
                Write-Warn "Could not install $mod automatically: $($_.Exception.Message)"
                return $result
            }
        }
        Import-Module $mod -ErrorAction SilentlyContinue
    }

    # Try sign-in. Application.Read.All is treated as optional — a
    # tenant-wide admin-consent on the Microsoft Graph PowerShell client lets
    # regular users hold this scope, but if it's not consented the call fails
    # and we fall back gracefully.
    try {
        Connect-MgGraph -Scopes 'User.Read','Application.Read.All' -NoWelcome -ErrorAction Stop | Out-Null
    } catch {
        Write-Warn "Sign-in failed or Application.Read.All not consented: $($_.Exception.Message)"
        Write-Warn 'Falling back to manual entry.'
        return $result
    }

    $ctx = Get-MgContext
    if (-not $ctx) {
        Write-Warn 'No Graph context after sign-in; falling back to manual entry.'
        return $result
    }

    $result.TenantId = $ctx.TenantId
    Write-Ok ("Detected home tenant: {0}  (signed in as {1})" -f $ctx.TenantId, $ctx.Account)

    # Look for partner-app candidates: multi-tenant audience, has Microsoft
    # Graph in requiredResourceAccess. Filter by name as a tie-breaker.
    try {
        $candidates = Get-MgApplication -All -ErrorAction Stop |
            Where-Object {
                $_.SignInAudience -eq 'AzureADMultipleOrgs' -and
                @($_.RequiredResourceAccess.ResourceAppId) -contains '00000003-0000-0000-c000-000000000000'
            }
    } catch {
        Write-Warn "Could not list app registrations: $($_.Exception.Message)"
        Write-Warn '(This usually means Application.Read.All needs admin consent in your tenant.)'
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        return $result
    }

    $candidates = @($candidates)
    if ($candidates.Count -eq 0) {
        Write-Warn "No multi-tenant Entra apps with Microsoft Graph permissions found in this tenant."
        Write-Warn "Either the partner app hasn't been registered yet (run scripts/Register-PartnerCenterApp.ps1) or Application.Read.All didn't return enough scope. Falling back to manual entry."
    } elseif ($candidates.Count -eq 1) {
        $result.ClientId = $candidates[0].AppId
        Write-Ok ("Auto-discovered partner app: {0}  (clientId {1})" -f $candidates[0].DisplayName, $candidates[0].AppId)
    } else {
        Write-Host "    Found $($candidates.Count) candidate apps. Pick one:"
        for ($i = 0; $i -lt $candidates.Count; $i++) {
            Write-Host ("      [{0}] {1}  (clientId {2})" -f $i, $candidates[$i].DisplayName, $candidates[$i].AppId)
        }
        $picked = Read-Host '    Enter number (or press Enter to skip and enter manually)'
        if ($picked -match '^\d+$' -and [int]$picked -lt $candidates.Count) {
            $result.ClientId = $candidates[[int]$picked].AppId
            Write-Ok ("Selected: {0}  (clientId {1})" -f $candidates[[int]$picked].DisplayName, $candidates[[int]$picked].AppId)
        }
    }

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    return $result
}

# ============================================================================
# Collect values
# ============================================================================

if (-not $NonInteractive) {
    $discovered = Try-AutoDiscover -Skip:$NoAutoDiscover

    Write-Host ''
    Write-Host '    Tip: square-bracket values shown next to each prompt are defaults' -ForegroundColor DarkGray
    Write-Host '         (auto-discovered or from your existing config). Press Enter to' -ForegroundColor DarkGray
    Write-Host '         accept, or type a different value to override.' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '=== Partner-app values (always required) ===' -ForegroundColor Cyan

    # Default precedence: auto-discovered (live truth from current tenant) > existing config > nothing.
    # Existing-config-first was wrong: when an app gets re-registered, the
    # stale config wins and the prompt offers a clientId that no longer
    # exists in the tenant, leading to AADSTS700016 on the next sign-in.
    # Auto-discovery is the live state of the tenant and should win when
    # present; warn the operator if existing config disagrees so they know
    # something changed.
    $existingClientId   = (Default-From $existing 'partner.clientId')
    $existingTenantId   = (Default-From $existing 'partner.homeTenantId')

    if ($discovered.ClientId) {
        $clientIdDefault = $discovered.ClientId
        if ($existingClientId -and $existingClientId -ne $discovered.ClientId) {
            Write-Warn "existing config clientId ($existingClientId) differs from the live auto-discovered value."
            Write-Warn "  Using auto-discovered ($($discovered.ClientId)) as the prompt default. Press Enter to accept; type the existing value to override."
        }
    } else {
        $clientIdDefault = $existingClientId
    }

    if ($discovered.TenantId) {
        $homeTenantIdDefault = $discovered.TenantId
        if ($existingTenantId -and $existingTenantId -ne $discovered.TenantId) {
            Write-Warn "existing config homeTenantId ($existingTenantId) differs from the live auto-discovered value."
            Write-Warn "  Using auto-discovered ($($discovered.TenantId)) as the prompt default. Press Enter to accept; type the existing value to override."
        }
    } else {
        $homeTenantIdDefault = $existingTenantId
    }

    $ClientId = Read-Required `
        -Prompt    'Partner-app clientId (GUID)' `
        -Default   $clientIdDefault `
        -Validator { param($v) Test-Guid $v }

    $HomeTenantId = Read-Required `
        -Prompt    'Home tenant id (GUID)' `
        -Default   $homeTenantIdDefault `
        -Validator { param($v) Test-Guid $v }

    if ($AppOnly) {
        Write-Host ''
        Write-Host '=== Cert values (-AppOnly only) ===' -ForegroundColor Cyan

        $CertificateThumbprint = Read-Required `
            -Prompt    'Certificate thumbprint (40 hex chars, no separators)' `
            -Default   (Default-From $existing 'partner.certificateThumbprint') `
            -Validator { param($v) Test-Thumb $v }

        $defaultPfx = Default-From $existing 'partner.certificatePfxPath'
        if (-not $defaultPfx) { $defaultPfx = (Join-Path $HOME '.m365-multi-tenant-inventory/partner-app.pfx') }
        $CertificatePfxPath = Read-Required `
            -Prompt   'PFX file path' `
            -Default  $defaultPfx

        $defaultEnvVar = Default-From $existing 'partner.certificatePfxPasswordEnvVar'
        if (-not $defaultEnvVar) { $defaultEnvVar = $CertificatePfxPasswordEnvVar }
        $CertificatePfxPasswordEnvVar = Read-Required `
            -Prompt   'Env-var name for the PFX password' `
            -Default  $defaultEnvVar
    }
} else {
    # Non-interactive: validate the provided values
    if (-not $ClientId) { throw '-NonInteractive requires -ClientId' }
    $err = Test-Guid $ClientId; if ($err) { throw "ClientId $err" }
    if (-not $HomeTenantId) { throw '-NonInteractive requires -HomeTenantId' }
    $err = Test-Guid $HomeTenantId; if ($err) { throw "HomeTenantId $err" }
    if ($AppOnly) {
        if (-not $CertificateThumbprint) { throw '-NonInteractive -AppOnly requires -CertificateThumbprint' }
        $err = Test-Thumb $CertificateThumbprint; if ($err) { throw "CertificateThumbprint $err" }
        if (-not $CertificatePfxPath) { throw '-NonInteractive -AppOnly requires -CertificatePfxPath' }
    }
}

# ============================================================================
# Compose final config
# ============================================================================

# Start from the source structure (so we get defaults, exclude, tenants
# blocks, and all _comment guidance) and overwrite the partner block.
$out = $source.PSObject.Copy()

$out.partner.clientId      = $ClientId
$out.partner.homeTenantId  = $HomeTenantId

if ($AppOnly) {
    $out.partner.certificateThumbprint        = $CertificateThumbprint
    $out.partner.certificatePfxPath           = $CertificatePfxPath
    $out.partner.certificatePfxPasswordEnvVar = $CertificatePfxPasswordEnvVar
} else {
    # Delegated mode — explicitly null the cert fields so the wrapper's
    # AppOnly-only validation skips them cleanly.
    $out.partner.certificateThumbprint = $null
    $out.partner.certificatePfxPath    = $null
    # Leave the env-var name as-is (harmless when null cert path means it's never read)
}

# ============================================================================
# Write
# ============================================================================

# Ensure parent dir exists (for non-default destinations).
$destDir = Split-Path -Parent $ConfigDestination
if ($destDir -and -not (Test-Path $destDir)) {
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
}

$json = $out | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText($ConfigDestination, $json)
Write-Step "Wrote $ConfigDestination ($([System.IO.File]::ReadAllBytes($ConfigDestination).Length) bytes)"

# Tighten permissions on POSIX (best-effort; ignored on Windows).
if ($IsMacOS -or $IsLinux) {
    try { & chmod 600 $ConfigDestination } catch { Write-Warn "Could not chmod 600 on $ConfigDestination — set manually." }
}

# ============================================================================
# Confirm gitignore coverage
# ============================================================================

$gitignorePath = Join-Path $PSScriptRoot '..' '.gitignore'
$destBasename  = Split-Path $ConfigDestination -Leaf
$relativeToRepo = $false
try {
    $repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
    $resolved = Resolve-Path $ConfigDestination -ErrorAction SilentlyContinue
    if ($resolved -and $resolved.Path.StartsWith($repoRoot.Path)) { $relativeToRepo = $true }
} catch { }

Write-Host ''
if (-not $relativeToRepo) {
    Write-Ok "Config is OUTSIDE the repo at $ConfigDestination — git can't see it. No .gitignore concern."
} elseif (Test-Path $gitignorePath) {
    $patterns = @(Get-Content $gitignorePath | Where-Object { $_ -and -not $_.TrimStart().StartsWith('#') })
    $covered = $false
    foreach ($p in $patterns) {
        $p = $p.Trim()
        if ($p -eq $destBasename -or $p -eq "*.local.json" -or $p -like "tenants.config.*") {
            $covered = $true; break
        }
    }
    if ($covered) {
        Write-Ok ".gitignore covers $destBasename — safe to commit the rest of the repo."
    } else {
        Write-Warn ".gitignore may NOT cover $destBasename — verify, or move the config outside the repo."
    }
} else {
    Write-Warn 'No .gitignore found in repo root — ensure your config destination is excluded from version control.'
}

# ============================================================================
# Print next steps
# ============================================================================

Write-Host ''
Write-Host '────────────────────────────────────────────────────────────────────────────'
Write-Host 'Setup complete. Next steps:'
Write-Host '────────────────────────────────────────────────────────────────────────────'
Write-Host "  1. Verify the values in $ConfigDestination."
if ($AppOnly) {
    Write-Host '  2. Set the PFX password env var in your shell:'
    Write-Host ('       $env:{0} = ''<your-pfx-password>''' -f $CertificatePfxPasswordEnvVar) -ForegroundColor DarkGray
    Write-Host '     (or fetch from Keychain / Credential Manager — see the App-Only KB article).'
    Write-Host '  3. Run with -AppOnly:'
    Write-Host ('       ./Get-O365MailGroupInventory-Multi.ps1 -ConfigPath "{0}" -AppOnly' -f $ConfigDestination) -ForegroundColor DarkGray
} else {
    Write-Host '  2. Run delegated (default; no env-var needed):'
    Write-Host ('       ./Get-O365MailGroupInventory-Multi.ps1 -ConfigPath "{0}"' -f $ConfigDestination) -ForegroundColor DarkGray
    Write-Host '     First run will prompt for sign-in once; subsequent runs use cached tokens.'
}
Write-Host '────────────────────────────────────────────────────────────────────────────'
