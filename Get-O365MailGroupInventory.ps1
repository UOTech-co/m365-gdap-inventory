<#
.SYNOPSIS
    Microsoft 365 mail, group, identity, and Conditional Access inventory —
    shared / user / resource mailboxes, distribution groups, mail-enabled
    security groups, Microsoft 365 (Unified) Groups, Microsoft Teams, cloud
    security groups, Conditional Access policies and named locations,
    authentication-methods policy, directory-role assignments (active +
    PIM-eligible), per-admin MFA registration and sign-in activity, and
    service principals with directory roles — exported to a single
    multi-tab .xlsx workbook.

.DESCRIPTION
    Per-customer-tenant collector for the GDAP partner inventory tool.
    Driven by Get-O365MailGroupInventory-Multi.ps1 (the wrapper), one child
    process per customer. Authenticates as the partner-tenant operator via
    GDAP-delegated rights and pulls structural + usage metadata from
    Exchange Online, Microsoft Graph, and Microsoft Teams against the
    target customer tenant.

    Not intended to be invoked directly. The wrapper passes -TenantId,
    -ClientId, and -DelegatedOrganization for every customer in scope; the
    collector fails fast if those are missing.

    Workbook tabs produced (in workbook order):
      - Summary             : object counts + run metadata
      - Shared Mailboxes    : delegates (FullAccess / SendAs / SendOnBehalf),
                              forwarding, retention, mailbox stats
      - User Mailboxes      : every user mailbox in the tenant; size, last
                              logon, forwarding (skip with -SkipUserMailboxes)
      - Resource Mailboxes  : rooms / equipment, capacity, calendar processing
      - Distribution Groups
      - Mail-Enabled Sec Groups
      - M365 Groups         : Unified Groups (incl. Teams-provisioned flag),
                              SharePoint site URL, doc-library file count,
                              storage used (skip SP enrichment with
                              -SkipSharePointStats)
      - Teams               : channel counts (standard / private / shared),
                              owner / member / guest counts, plus the same
                              SharePoint site / file count / storage data
                              cross-referenced from the underlying group
      - Security Groups (cloud) : non-mail security groups from Graph
      - CA Policies         : every Conditional Access policy in the tenant
                              with state, included/excluded users + groups +
                              roles + apps, client-app types, risk levels,
                              grant controls, session controls (skip with
                              -SkipConditionalAccess)
      - CA Named Locations  : IP and country named locations with trusted flag
      - CA Auth Strengths   : authentication-strength policies referenced by
                              CA grant controls
      - Auth Methods Policy : tenant-level Authentication Methods Policy
                              (per-method state + included/excluded targets)
      - Directory Roles     : every directory role with active and eligible
                              member counts (skip with -SkipAdminPosture)
      - Admin Users         : one row per (user, role) tuple — display name,
                              UPN, role, active/eligible, MFA-registered,
                              MFA methods, last-interactive sign-in, last
                              non-interactive sign-in, account-enabled,
                              license summary, dirsync flag
      - Admin Service Principals : service principals holding any directory
                              role (apps acting as admins) with appId,
                              owner, role list, role-assignment source
      - Break-Glass Candidates : surfaced from name-pattern matches AND
                              from Global Admins with no recent sign-in;
                              flags for documentation hygiene
      - Run Log             : timestamped log of the run

    Every sheet has a frozen header row, autosized columns, and an autofilter.
    Every mail-enabled sheet includes an SmtpDomain column so you can filter
    by primary domain post hoc.

.PARAMETER OutputPath
    Destination .xlsx path. Defaults to current working directory with a
    timestamped filename like O365-MailGroupInventory_20260430_1530.xlsx.

.PARAMETER UserPrincipalName
    Optional UPN to pre-fill the modern auth prompt.

.PARAMETER SkipMailboxStats
    Skip Get-MailboxStatistics calls (size, item count, last logon). Useful
    for a fast structural-only pass against a tenant with many mailboxes.

.PARAMETER SkipPermissions
    Skip FullAccess / SendAs / SendOnBehalf delegate collection on shared
    mailboxes. The other sheets are unaffected.

.PARAMETER SkipUserMailboxes
    Skip the user-mailbox sheet. By default every user mailbox in the tenant
    is collected (size, last logon, forwarding, retention). Use this switch
    on large tenants when you only need the shared / group / team view.

.PARAMETER SkipSharePointStats
    Skip the SharePoint enrichment on M365 Groups and Teams (site URL,
    document-library file count, storage used). Useful when running without
    the Sites.Read.All scope or when you just need a structural pass.

.PARAMETER SkipConditionalAccess
    Skip the Conditional Access collection (CA Policies, CA Named Locations,
    CA Auth Strengths, Auth Methods Policy tabs). Useful for dry runs or
    when Policy.Read.All consent is not yet granted.

.PARAMETER SkipAdminPosture
    Skip the admin-posture collection (Directory Roles, Admin Users, Admin
    Service Principals, Break-Glass Candidates tabs). Useful when
    RoleManagement.Read.Directory consent is not yet granted.

.PARAMETER BreakGlassNamePattern
    Regex matched (case-insensitive) against UPN and DisplayName to flag
    likely break-glass / emergency-access accounts. Default catches the
    common conventions: BREAKGLASS, BREAK-GLASS, BREAK_GLASS, EMERGENCY,
    EMERG-, BG-, GLASS-. Override if your tenant uses a different convention.

.PARAMETER InactiveAdminDays
    Threshold (in days) above which an active Global Admin with no recent
    sign-in is surfaced on the Break-Glass Candidates tab. Default 60.
    Real break-glass accounts should rarely sign in; long-idle Global Admins
    with no break-glass naming are documentation-hygiene candidates.

.PARAMETER GroupMemberPreviewCount
    How many member display names to inline as a preview column on each
    group sheet. Default 25 — keeps the workbook readable while full counts
    are always captured in the MemberCount column.

.EXAMPLE
    .\Get-O365MailGroupInventory.ps1

.EXAMPLE
    .\Get-O365MailGroupInventory.ps1 -OutputPath 'C:\Reports\inv.xlsx' -SkipUserMailboxes

.EXAMPLE
    .\Get-O365MailGroupInventory.ps1 -SkipMailboxStats -SkipPermissions

.NOTES
    Required PowerShell modules (auto-installed if missing, Scope=CurrentUser):
      ExchangeOnlineManagement       3.0.0+
      Microsoft.Graph.Authentication 2.0.0+
      Microsoft.Graph.Groups         2.0.0+
      Microsoft.Graph.Users          2.0.0+
      MicrosoftTeams                 5.0.0+
      ImportExcel                    7.0.0+

    All Conditional Access and admin-posture queries are issued via
    Invoke-MgGraphRequest against the v1.0/beta endpoints — no extra
    Microsoft.Graph submodules are required for those passes.

    Required Microsoft Graph delegated scopes:
      Core (mail / group / team inventory):
        Group.Read.All, GroupMember.Read.All, User.Read.All,
        Team.ReadBasic.All, Channel.ReadBasic.All, Directory.Read.All,
        Sites.Read.All     (SharePoint site URL / drive quota lookups)
        Reports.Read.All   (per-site file count via SP usage report —
                            admin-consent scope; also covers the MFA
                            user-registration-details report below)
      Conditional Access (skip with -SkipConditionalAccess):
        Policy.Read.All    (CA policies, named locations,
                            authentication-strength policies,
                            authentication-methods policy)
      Admin posture (skip with -SkipAdminPosture):
        RoleManagement.Read.Directory  (role definitions, role
                            assignments, role eligibility schedules / PIM)
        AuditLog.Read.All  (signInActivity property on /users; also
                            backs the MFA user-registration-details report)
        UserAuthenticationMethod.Read.All  (registered MFA methods per
                            privileged user — strong-auth method type
                            enumeration)
        Application.Read.All  (service-principal display name + appId
                            enrichment when SPs hold directory roles)

    Required Exchange Online role (minimum):
      View-Only Recipients   — for mailbox / group reads
      View-Only Configuration — for organization-level reads
      (Global Reader covers both.)

    Required Entra (Azure AD) role for the Conditional Access /
    admin-posture passes (minimum):
      Global Reader  — covers Policy.Read.All, RoleManagement.Read.Directory,
                       and the role/sign-in/auth-method reads.
      Security Reader is sufficient for everything *except*
      UserAuthenticationMethod.Read.All; if running as Security Reader,
      pass -SkipAdminPosture or accept blank MFA-method columns.

    License: Apache-2.0 (see LICENSE in repo root)
    Created: 2026-04-30
    Updated: 2026-05-05 — collector for the GDAP partner inventory tool.
                          Always driven by Get-O365MailGroupInventory-Multi.ps1.
                          Standalone single-tenant mode removed.
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [string] $OutputPath = (Join-Path (Get-Location) ("O365-MailGroupInventory_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmm'))),

    # GDAP delegated auth (set by Get-O365MailGroupInventory-Multi.ps1):
    #   -TenantId               target customer tenant id (GUID, REQUIRED)
    #   -DelegatedOrganization  customer's primary domain
    #                           (e.g. customer.onmicrosoft.com, REQUIRED).
    #                           Used for Connect-ExchangeOnline -DelegatedOrganization.
    #   -ClientId               partner-app clientId. Accepted for future
    #                           cert-based app-only flows; intentionally NOT
    #                           used for delegated Connect-MgGraph (the script
    #                           falls back to the default Microsoft Graph
    #                           PowerShell client to dodge AADSTS90099).
    [Parameter(Mandatory = $true)]
    [string] $TenantId,

    [Parameter(Mandatory = $true)]
    [string] $DelegatedOrganization,

    [string] $ClientId,

    [switch] $SkipMailboxStats,
    [switch] $SkipPermissions,
    [switch] $SkipUserMailboxes,
    [switch] $SkipSharePointStats,
    [switch] $SkipConditionalAccess,
    [switch] $SkipAdminPosture,
    [string] $BreakGlassNamePattern = '(?i)(break.?glass|emerg(ency)?|^bg-|-bg$|glass-)',
    [int]    $InactiveAdminDays = 60,
    [int]    $GroupMemberPreviewCount = 25
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ fail-fast
# This script is the per-customer-tenant collector for the wrapper at
# Get-O365MailGroupInventory-Multi.ps1. It is not designed to run standalone.
# CmdletBinding's Mandatory enforcement covers the explicit-call case, but
# we double-check here to give a clearer message if someone source-loads it
# or invokes it from a hand-rolled script that bypasses parameter binding.
if ([string]::IsNullOrWhiteSpace($TenantId) -or
    [string]::IsNullOrWhiteSpace($DelegatedOrganization)) {
    throw "Get-O365MailGroupInventory.ps1 is the per-customer collector for the multi-tenant wrapper. Run Get-O365MailGroupInventory-Multi.ps1 instead — it drives this script with the right parameters per GDAP customer. See README for the wrapper invocation."
}

# ---------------------------------------------------------------------- helpers
$script:RunLog = [System.Collections.Generic.List[object]]::new()

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO')
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "[{0}] [{1}] {2}" -f $stamp, $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
    $script:RunLog.Add([pscustomobject]@{ Timestamp = $stamp; Level = $Level; Message = $Message })
}

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name, [string]$MinVersion)
    $installed = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending | Select-Object -First 1
    if (-not $installed) {
        Write-Log "Installing missing module: $Name" 'WARN'
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    } elseif ($MinVersion -and [version]$installed.Version -lt [version]$MinVersion) {
        Write-Log "Updating module $Name (have $($installed.Version), need $MinVersion)" 'WARN'
        Update-Module -Name $Name -Force
    }
    Import-Module $Name -DisableNameChecking | Out-Null
}

function Get-SmtpDomain {
    param([string]$Smtp)
    if ([string]::IsNullOrWhiteSpace($Smtp)) { return $null }
    return ($Smtp -split '@')[-1]
}

function Join-Strings {
    param([object]$Items, [int]$Max = 0, [string]$Separator = '; ')
    if (-not $Items) { return $null }
    $arr = @($Items | Where-Object { $_ -and "$_".Trim() })
    if ($arr.Count -eq 0) { return $null }
    if ($Max -gt 0 -and $arr.Count -gt $Max) {
        return ($arr[0..($Max - 1)] -join $Separator) + " ... (+$($arr.Count - $Max) more)"
    }
    return ($arr -join $Separator)
}

function Format-Size {
    param($SizeObj)
    if (-not $SizeObj) { return $null }
    try { return $SizeObj.ToString() } catch { return "$SizeObj" }
}

function Try-Block {
    # Run a scriptblock, log + swallow any error, return $null on failure.
    # Forces ErrorActionPreference=Stop inside the block so non-terminating
    # cmdlet errors (common in EXO / Graph) are converted and caught.
    param([scriptblock]$ScriptBlock, [string]$Context = 'operation')
    try {
        $oldEAP = $ErrorActionPreference
        $ErrorActionPreference = 'Stop'
        return & $ScriptBlock
    } catch {
        Write-Log "$Context failed: $($_.Exception.Message)" 'WARN'
        return $null
    } finally {
        $ErrorActionPreference = $oldEAP
    }
}

function Resolve-GroupIdentity {
    # Pick the most-unique identifier available on a recipient/group object.
    # Avoids 'matches multiple entries' errors when DisplayName collides.
    param($Obj)
    if ($Obj.ExternalDirectoryObjectId) { return [string]$Obj.ExternalDirectoryObjectId }
    if ($Obj.Guid)                      { return [string]$Obj.Guid }
    if ($Obj.ExchangeObjectId)          { return [string]$Obj.ExchangeObjectId }
    if ($Obj.PrimarySmtpAddress)        { return [string]$Obj.PrimarySmtpAddress }
    return [string]$Obj.Identity
}

function Get-GroupSharePointStats {
    # Looks up SharePoint site stats for a unified group. Two-layer lookup:
    #
    #   Layer 1 — site usage report (bulk call, Reports.Read.All):
    #     Provides FileCount, ActiveFileCount, StorageUsedMB,
    #     StorageQuotaBytes, LastSPActivityDate. Affected by the M365
    #     privacy setting "Display concealed user, group, and site names
    #     in all reports" — when ON, Site URLs in the report are hashed
    #     and the URL match fails for every site. File count cannot be
    #     recovered programmatically when concealment is ON.
    #
    #   Layer 2 — live drive quota fallback (per-group, Sites.Read.All):
    #     When the report didn't produce storage, resolve the site by URL
    #     (/sites/{host}:{path}, NOT /groups/{id}/sites/root which is
    #     membership-gated) and read /sites/{siteId}/drive quota directly.
    #     Fills StorageUsedMB / StorageUsedBytes / StorageQuotaBytes.
    [CmdletBinding()]
    param(
        [string]$SharePointSiteUrl,
        [string]$SharePointDocumentsUrl
    )

    $r = [pscustomobject]@{
        SharePointSiteUrl      = $SharePointSiteUrl
        SharePointDocumentsUrl = $SharePointDocumentsUrl
        SharePointSiteId       = $null
        FileCount              = $null
        ActiveFileCount        = $null
        StorageUsedBytes       = $null
        StorageUsedMB          = $null
        StorageQuotaBytes      = $null
        LastSPActivityDate     = $null
        SPStatsSource          = $null   # 'report', 'drive', 'report+drive', or $null
    }

    if (-not $SharePointSiteUrl) { return $r }

    # ---- Layer 1: report cache lookup
    $usage = $null
    if ($script:spUsageBySiteUrl -and $script:spUsageBySiteUrl.ContainsKey($SharePointSiteUrl)) {
        $usage = $script:spUsageBySiteUrl[$SharePointSiteUrl]
    }
    if ($usage) {
        if ($usage.'File Count')               { $r.FileCount         = [int]$usage.'File Count' }
        if ($usage.'Active File Count')        { $r.ActiveFileCount   = [int]$usage.'Active File Count' }
        if ($usage.'Storage Used (Byte)') {
            $r.StorageUsedBytes = [long]$usage.'Storage Used (Byte)'
            $r.StorageUsedMB    = [math]::Round([long]$usage.'Storage Used (Byte)' / 1MB, 2)
        }
        if ($usage.'Storage Allocated (Byte)') { $r.StorageQuotaBytes  = [long]$usage.'Storage Allocated (Byte)' }
        if ($usage.'Last Activity Date')       { $r.LastSPActivityDate = $usage.'Last Activity Date' }
        $r.SPStatsSource = 'report'
    }

    # ---- Layer 2: live drive quota fallback (only if storage didn't come from the report)
    if ($null -eq $r.StorageUsedBytes -and $graphConnected) {
        try {
            $u = [System.Uri]$SharePointSiteUrl
            # /sites/{hostname}:{server-relative-path}
            $reqPath = "/v1.0/sites/$($u.Host):$($u.AbsolutePath)"
            $site = Invoke-MgGraphRequest -Method GET -Uri $reqPath -ErrorAction Stop
            if ($site -and $site.id) {
                $r.SharePointSiteId = $site.id
                $drive = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$($site.id)/drive" -ErrorAction Stop
                if ($drive.quota) {
                    if ($drive.quota.used) {
                        $r.StorageUsedBytes = [long]$drive.quota.used
                        $r.StorageUsedMB    = [math]::Round($drive.quota.used / 1MB, 2)
                    }
                    if ($drive.quota.total -and -not $r.StorageQuotaBytes) {
                        $r.StorageQuotaBytes = [long]$drive.quota.total
                    }
                }
                $r.SPStatsSource = if ($r.SPStatsSource) { 'report+drive' } else { 'drive' }
            }
        } catch {
            Write-Log "Drive-quota fallback failed for $SharePointSiteUrl : $($_.Exception.Message)" 'WARN'
        }
    }

    return $r
}

function Initialize-SharePointUsageCache {
    # One bulk call to /reports/getSharePointSiteUsageDetail; populates
    # the script-scoped hashtables used by Get-GroupSharePointStats.
    # Period D7 is the shortest available window — picks up sites that
    # were active in the last week. Sites with no recent activity still
    # appear in the report with their last-activity timestamp.
    $script:spUsageBySiteUrl = @{}
    $script:spUsageBySiteId  = @{}

    if ($SkipSharePointStats -or -not $graphConnected) { return }

    Write-Log "Fetching SharePoint site usage report (period=D7)..."
    $tmpCsv = [IO.Path]::Combine([IO.Path]::GetTempPath(), "sp_usage_$([guid]::NewGuid().Guid).csv")
    try {
        Invoke-MgGraphRequest -Method GET `
            -Uri "/v1.0/reports/getSharePointSiteUsageDetail(period='D7')" `
            -OutputFilePath $tmpCsv -ErrorAction Stop

        $usage = Import-Csv -Path $tmpCsv
        foreach ($row in $usage) {
            if ($row.'Site Id')  { $script:spUsageBySiteId[$row.'Site Id']   = $row }
            if ($row.'Site URL') { $script:spUsageBySiteUrl[$row.'Site URL'] = $row }
        }
        Write-Log "SharePoint usage report indexed: $(@($usage).Count) sites."

        # If every Site URL looks hashed, concealment is on in the M365
        # admin center → real URLs are stripped from reports. Flag it so
        # the operator knows why the SP columns came back partial.
        $sample = @($usage | Select-Object -First 5)
        if ($sample.Count -gt 0 -and -not ($sample.'Site URL' -match '^https?://')) {
            Write-Log "SharePoint report has concealed Site URLs — file counts cannot be matched. To fix: M365 Admin Center → Settings → Org settings → Services → Reports → uncheck 'Display concealed user, group, and site names in all reports' → Save, then re-run. Storage columns will be filled by the live-drive fallback in the meantime; FileCount stays blank." 'WARN'
        }
    } catch {
        Write-Log "SharePoint usage report fetch failed: $($_.Exception.Message). FileCount / storage columns will be blank; SharePointSiteUrl will still populate from Get-UnifiedGroup." 'WARN'
    } finally {
        if (Test-Path $tmpCsv) { Remove-Item $tmpCsv -Force -ErrorAction SilentlyContinue }
    }
}

# ----------------------------------------------------- directory-object resolution
# Cache for directory-object lookups (users, groups, service principals, roles).
# Conditional Access policies and role assignments come back as bare GUIDs;
# resolving them per-policy without a cache is N*M Graph calls. We populate
# this lazily and reuse across CA + admin posture.
$script:DirObjectCache = @{}

function Get-DirObjectName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Id,
        [string]$ExpectedKind  # 'user', 'group', 'servicePrincipal', 'role' — best effort
    )
    if (-not $Id -or $Id -eq 'All' -or $Id -eq 'None' -or $Id -eq 'GuestsOrExternalUsers') {
        return $Id  # CA shorthand values pass through
    }
    if ($script:DirObjectCache.ContainsKey($Id)) {
        return $script:DirObjectCache[$Id].Display
    }

    # Try the lightweight directoryObjects/{id} endpoint — works for users,
    # groups, and service principals; gives us @odata.type back so we can
    # tag the cache entry with the actual object kind. Stale references
    # (deleted users still listed in CA policies) come back 404 — silently
    # cache the GUID so we don't keep retrying or filling the run log.
    $obj = $null
    try {
        $obj = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/directoryObjects/$Id" -ErrorAction Stop
    } catch {
        # 404 / NotFound is the common case (deleted object). Cache & move on.
    }
    if ($obj) {
        $kind = ($obj.'@odata.type' -replace '#microsoft\.graph\.','')
        $name = if ($obj.displayName) { $obj.displayName }
                elseif ($obj.userPrincipalName) { $obj.userPrincipalName }
                else { $Id }
        $script:DirObjectCache[$Id] = [pscustomobject]@{ Id = $Id; Kind = $kind; Display = $name; Upn = $obj.userPrincipalName; AppId = $obj.appId }
        return $name
    }

    # Some endpoints (role definitions, applications) aren't reachable through
    # /directoryObjects — fall through to specific endpoints when we have a hint.
    if ($ExpectedKind -eq 'role') {
        try {
            $r = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/roleManagement/directory/roleDefinitions/$Id" -ErrorAction Stop
            if ($r) {
                $script:DirObjectCache[$Id] = [pscustomobject]@{ Id = $Id; Kind = 'roleDefinition'; Display = $r.displayName }
                return $r.displayName
            }
        } catch { }
    }
    if ($ExpectedKind -eq 'servicePrincipal') {
        try {
            $sp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/servicePrincipals/$Id" -ErrorAction Stop
            if ($sp) {
                $script:DirObjectCache[$Id] = [pscustomobject]@{ Id = $Id; Kind = 'servicePrincipal'; Display = $sp.displayName; AppId = $sp.appId }
                return $sp.displayName
            }
        } catch { }
    }

    # Last resort — return the GUID so the cell isn't empty, and cache the
    # tombstone so we don't refetch the same dead reference.
    $script:DirObjectCache[$Id] = [pscustomobject]@{ Id = $Id; Kind = 'unknown (likely deleted)'; Display = $Id }
    return $Id
}

function Resolve-DirObjectIds {
    # Bulk resolver — takes an array of object IDs, returns "Name [kind]; Name [kind]; ..."
    param(
        [object]$Ids,
        [int]$Max = 10,
        [string]$ExpectedKind
    )
    if (-not $Ids) { return $null }
    $arr = @($Ids | Where-Object { $_ })
    if ($arr.Count -eq 0) { return $null }
    $resolved = foreach ($id in $arr) {
        $name = Get-DirObjectName -Id $id -ExpectedKind $ExpectedKind
        $kind = if ($script:DirObjectCache.ContainsKey($id)) { $script:DirObjectCache[$id].Kind } else { 'unknown' }
        if ($name -and $name -ne $id) { "$name [$kind]" } else { $id }
    }
    if ($Max -gt 0 -and $resolved.Count -gt $Max) {
        return ($resolved[0..($Max - 1)] -join '; ') + " ... (+$($resolved.Count - $Max) more)"
    }
    return ($resolved -join '; ')
}

# ------------------------------------------------------------- CA policy flatten
function ConvertTo-CAPolicyRow {
    # Flatten a CA policy graph object into a single-row, Excel-friendly
    # pscustomobject. Conditions that don't apply (e.g. user risk on a policy
    # that doesn't use it) come back as $null and render as empty cells.
    param([Parameter(Mandatory)]$Policy)

    $cond = $Policy.conditions
    $u    = $cond.users
    $apps = $cond.applications
    $grant = $Policy.grantControls
    $sess  = $Policy.sessionControls

    # Authentication strength resolution — grantControls.authenticationStrength is an inline object
    $authStrength = $null
    if ($grant -and $grant.authenticationStrength -and $grant.authenticationStrength.displayName) {
        $authStrength = "$($grant.authenticationStrength.displayName) ($($grant.authenticationStrength.policyType))"
    }

    # Session controls — present only when set
    $signInFreq = $null
    if ($sess -and $sess.signInFrequency -and $sess.signInFrequency.isEnabled) {
        $signInFreq = "$($sess.signInFrequency.value) $($sess.signInFrequency.type) (frequencyInterval=$($sess.signInFrequency.frequencyInterval))"
    }
    $persistentBrowser = $null
    if ($sess -and $sess.persistentBrowser -and $sess.persistentBrowser.isEnabled) {
        $persistentBrowser = $sess.persistentBrowser.mode
    }
    $appEnforced = $null
    if ($sess -and $sess.applicationEnforcedRestrictions -and $sess.applicationEnforcedRestrictions.isEnabled) {
        $appEnforced = $true
    }
    $cas = $null
    if ($sess -and $sess.cloudAppSecurity -and $sess.cloudAppSecurity.isEnabled) {
        $cas = $sess.cloudAppSecurity.cloudAppSecurityType
    }

    # Platforms / locations / client app types — empty arrays mean "all"
    $platformsInc = if ($cond.platforms) { Join-Strings $cond.platforms.includePlatforms } else { 'All' }
    $platformsExc = if ($cond.platforms) { Join-Strings $cond.platforms.excludePlatforms } else { $null }
    $locationsInc = if ($cond.locations) { Resolve-DirObjectIds $cond.locations.includeLocations -Max 20 } else { 'All' }
    $locationsExc = if ($cond.locations) { Resolve-DirObjectIds $cond.locations.excludeLocations -Max 20 } else { $null }
    $clientApps   = Join-Strings $cond.clientAppTypes

    [pscustomobject]@{
        DisplayName            = $Policy.displayName
        State                  = $Policy.state
        PolicyId               = $Policy.id
        CreatedDateTime        = $Policy.createdDateTime
        ModifiedDateTime       = $Policy.modifiedDateTime
        IncludeUsers           = Resolve-DirObjectIds $u.includeUsers -ExpectedKind 'user' -Max 20
        ExcludeUsers           = Resolve-DirObjectIds $u.excludeUsers -ExpectedKind 'user' -Max 20
        IncludeGroups          = Resolve-DirObjectIds $u.includeGroups -ExpectedKind 'group' -Max 20
        ExcludeGroups          = Resolve-DirObjectIds $u.excludeGroups -ExpectedKind 'group' -Max 20
        IncludeRoles           = Resolve-DirObjectIds $u.includeRoles -ExpectedKind 'role' -Max 20
        ExcludeRoles           = Resolve-DirObjectIds $u.excludeRoles -ExpectedKind 'role' -Max 20
        IncludeApplications    = Join-Strings $apps.includeApplications -Max 20
        ExcludeApplications    = Join-Strings $apps.excludeApplications -Max 20
        IncludeUserActions     = Join-Strings $apps.includeUserActions
        IncludeAuthContexts    = Join-Strings $apps.includeAuthenticationContextClassReferences
        ClientAppTypes         = $clientApps
        IncludePlatforms       = $platformsInc
        ExcludePlatforms       = $platformsExc
        IncludeLocations       = $locationsInc
        ExcludeLocations       = $locationsExc
        SignInRiskLevels       = Join-Strings $cond.signInRiskLevels
        UserRiskLevels         = Join-Strings $cond.userRiskLevels
        DeviceFilterMode       = if ($cond.devices) { $cond.devices.deviceFilter.mode } else { $null }
        DeviceFilterRule       = if ($cond.devices) { $cond.devices.deviceFilter.rule } else { $null }
        GrantOperator          = $grant.operator
        BuiltInControls        = Join-Strings $grant.builtInControls
        AuthenticationStrength = $authStrength
        TermsOfUse             = Join-Strings $grant.termsOfUse
        CustomAuthFactors      = Join-Strings $grant.customAuthenticationFactors
        SessionSignInFrequency = $signInFreq
        SessionPersistentBrowser = $persistentBrowser
        SessionAppEnforcedRestrictions = $appEnforced
        SessionCloudAppSecurity = $cas
    }
}

# ------------------------------------------------------------ SKU friendly names
# Static lookup keyed by skuPartNumber so workbook columns show readable plan
# names instead of GUIDs. Populated from Microsoft's published "Product names
# and service plan identifiers for licensing" reference. Anything not in this
# table falls back to the skuPartNumber from /v1.0/subscribedSkus, which is
# already much more readable than the raw GUID.
$script:SkuFriendlyNames = @{
    'AAD_BASIC'                                     = 'Microsoft Entra ID Basic'
    'AAD_PREMIUM'                                   = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'                                = 'Microsoft Entra ID P2'
    'ATP_ENTERPRISE'                                = 'Microsoft Defender for Office 365 (Plan 1)'
    'ATP_ENTERPRISE_FACULTY'                        = 'Microsoft Defender for Office 365 (Plan 1) — Faculty'
    'CRMSTANDARD'                                   = 'Dynamics 365 Customer Engagement Plan'
    'DEFENDER_ENDPOINT_P1'                          = 'Microsoft Defender for Endpoint Plan 1'
    'DESKLESSPACK'                                  = 'Office 365 F3'
    'DEVELOPERPACK'                                 = 'Office 365 E3 Developer'
    'DEVELOPERPACK_E5'                              = 'Microsoft 365 E5 Developer'
    'DYN365_ENTERPRISE_PLAN1'                       = 'Dynamics 365 Customer Engagement Plan'
    'DYN365_ENTERPRISE_SALES'                       = 'Dynamics 365 for Sales'
    'EMS'                                           = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                                    = 'Enterprise Mobility + Security E5'
    'ENTERPRISEPACK'                                = 'Office 365 E3'
    'ENTERPRISEPACKPLUS_FACULTY'                    = 'Office 365 A3 — Faculty'
    'ENTERPRISEPACK_FACULTY'                        = 'Office 365 A3 — Faculty'
    'ENTERPRISEPACK_STUDENT'                        = 'Office 365 A3 — Student'
    'ENTERPRISEPREMIUM'                             = 'Office 365 E5'
    'ENTERPRISEPREMIUM_FACULTY'                     = 'Office 365 A5 — Faculty'
    'ENTERPRISEPREMIUM_NOPSTNCONF'                  = 'Office 365 E5 (no Audio Conferencing)'
    'ENTERPRISEWITHSCAL'                            = 'Office 365 E4'
    'EOP_ENTERPRISE'                                = 'Exchange Online Protection'
    'EOP_ENTERPRISE_PREMIUM'                        = 'Exchange Online Protection Premium'
    'EQUIVIO_ANALYTICS'                             = 'Office 365 Advanced Compliance'
    'EXCHANGEARCHIVE'                               = 'Exchange Online Archiving for Exchange Server'
    'EXCHANGEARCHIVE_ADDON'                         = 'Exchange Online Archiving for Exchange Online'
    'EXCHANGEDESKLESS'                              = 'Exchange Online Kiosk'
    'EXCHANGEENTERPRISE'                            = 'Exchange Online (Plan 2)'
    'EXCHANGEENTERPRISE_FACULTY'                    = 'Exchange Online (Plan 2) — Faculty'
    'EXCHANGEONLINE_MULTIGEO'                       = 'Exchange Online Multi-Geo'
    'EXCHANGESTANDARD'                              = 'Exchange Online (Plan 1)'
    'EXCHANGESTANDARD_FACULTY'                      = 'Exchange Online (Plan 1) — Faculty'
    'EXCHANGE_S_DESKLESS'                           = 'Exchange Online Kiosk'
    'FLOW_FREE'                                     = 'Power Automate Free'
    'FLOW_PER_USER'                                 = 'Power Automate per User Plan'
    'IDENTITY_THREAT_PROTECTION'                    = 'Microsoft 365 E5 Security'
    'IDENTITY_THREAT_PROTECTION_FOR_EMS_E5'         = 'Microsoft 365 E5 Security for EMS E5'
    'INFORMATION_PROTECTION_COMPLIANCE'             = 'Microsoft 365 E5 Compliance'
    'INTUNE_A'                                      = 'Microsoft Intune Plan 1'
    'INTUNE_A_VL'                                   = 'Microsoft Intune Plan 1'
    'M365EDU_A1'                                    = 'Microsoft 365 A1'
    'M365EDU_A3_FACULTY'                            = 'Microsoft 365 A3 — Faculty'
    'M365EDU_A3_STUDENT'                            = 'Microsoft 365 A3 — Student'
    'M365EDU_A5_FACULTY'                            = 'Microsoft 365 A5 — Faculty'
    'M365EDU_A5_STUDENT'                            = 'Microsoft 365 A5 — Student'
    'M365_F1'                                       = 'Microsoft 365 F1'
    'M365_F1_COMM'                                  = 'Microsoft 365 F1 (Commercial)'
    'M365_SECURITY_COMPLIANCE_FOR_FLW'              = 'Microsoft 365 F5 Security + Compliance'
    'MCOEV'                                         = 'Microsoft 365 Phone System'
    'MCOMEETADV'                                    = 'Microsoft 365 Audio Conferencing'
    'MCOPSTN1'                                      = 'Microsoft 365 Domestic Calling Plan'
    'MCOPSTN2'                                      = 'Microsoft 365 Domestic and International Calling Plan'
    'MCOSTANDARD'                                   = 'Skype for Business Online (Plan 2)'
    'MEE_FACULTY'                                   = 'Minecraft: Education Edition — Faculty'
    'MEE_STUDENT'                                   = 'Minecraft: Education Edition — Student'
    'MIDSIZEPACK'                                   = 'Office 365 Midsize Business'
    'MS_TEAMS_IW'                                   = 'Microsoft Teams Trial'
    'OFFICESUBSCRIPTION'                            = 'Microsoft 365 Apps for Enterprise'
    'OFFICESUBSCRIPTION_FACULTY'                    = 'Microsoft 365 Apps for Enterprise — Faculty'
    'OFFICESUBSCRIPTION_STUDENT'                    = 'Microsoft 365 Apps for Enterprise — Student'
    'O365_BUSINESS'                                 = 'Microsoft 365 Apps for Business'
    'O365_BUSINESS_ESSENTIALS'                      = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM'                         = 'Microsoft 365 Business Standard'
    'POWERAPPS_PER_USER'                            = 'Power Apps per User Plan'
    'POWERFLOW_PER_USER'                            = 'Power Automate per User Plan'
    'POWER_BI_PREMIUM_PER_USER'                     = 'Power BI Premium Per User'
    'POWER_BI_PRO'                                  = 'Power BI Pro'
    'POWER_BI_STANDARD'                             = 'Power BI (free)'
    'PROJECTPREMIUM'                                = 'Project Plan 5'
    'PROJECTPROFESSIONAL'                           = 'Project Plan 3'
    'PROJECT_P1'                                    = 'Project Plan 1'
    'PROJECT_PLAN_3'                                = 'Project Plan 3'
    'RIGHTSMANAGEMENT'                              = 'Azure Information Protection Premium P1'
    'RIGHTSMANAGEMENT_ADHOC'                        = 'Rights Management Ad-Hoc'
    'SHAREPOINTENTERPRISE'                          = 'SharePoint Online (Plan 2)'
    'SHAREPOINTSTANDARD'                            = 'SharePoint Online (Plan 1)'
    'SMB_APPS'                                      = 'Microsoft 365 Apps for Business'
    'SMB_BUSINESS_ESSENTIALS'                       = 'Microsoft 365 Business Basic'
    'SMB_BUSINESS_PREMIUM'                          = 'Microsoft 365 Business Premium'
    'SPB'                                           = 'Microsoft 365 Business Premium'
    'SPE_E3'                                        = 'Microsoft 365 E3'
    'SPE_E3_USGOV_DOD'                              = 'Microsoft 365 E3 — US DoD'
    'SPE_E3_USGOV_GCCHIGH'                          = 'Microsoft 365 E3 — US GCC High'
    'SPE_E5'                                        = 'Microsoft 365 E5'
    'SPE_E5_NOPSTNCONF'                             = 'Microsoft 365 E5 (no Audio Conferencing)'
    'SPE_F1'                                        = 'Microsoft 365 F3'
    'STANDARDPACK'                                  = 'Office 365 E1'
    'STANDARDWOFFPACK'                              = 'Office 365 E2'
    'STANDARDWOFFPACK_FACULTY'                      = 'Office 365 A1 — Faculty'
    'STANDARDWOFFPACK_STUDENT'                      = 'Office 365 A1 — Student'
    'STREAM'                                        = 'Microsoft Stream'
    'TEAMS_EXPLORATORY'                             = 'Teams Exploratory'
    'TEAMS_FREE'                                    = 'Microsoft Teams (Free)'
    'THREAT_INTELLIGENCE'                           = 'Microsoft Defender for Office 365 (Plan 2)'
    'VISIOCLIENT'                                   = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'                             = 'Visio Plan 1'
    'VISIO_PLAN1_DEPT'                              = 'Visio Plan 1'
    'WACONEDRIVESTANDARD'                           = 'OneDrive for Business (Plan 1)'
    'WIN10_PRO_ENT_SUB'                             = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E5'                                  = 'Windows 10/11 Enterprise E5'
    'WIN_DEF_ATP'                                   = 'Microsoft Defender for Endpoint Plan 2'
}

# Caches populated once after Graph connects.
$script:SkuById      = @{}   # GUID skuId  -> friendly name (or skuPartNumber)
$script:LicenseByUpn = @{}   # lower(UPN)  -> array of assignedLicenses entries

function Initialize-SkuCatalog {
    # /subscribedSkus gives us every license SKU in this tenant. We map the
    # GUID to a friendly name (preferred) or to the skuPartNumber (fallback).
    $script:SkuById = @{}
    if (-not $graphConnected) { return }
    Write-Log "Fetching subscribed SKUs (license catalog)..."
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/subscribedSkus" -ErrorAction Stop
        $skus = @($resp.value)
        foreach ($s in $skus) {
            if (-not $s.skuId) { continue }
            $part = $s.skuPartNumber
            $name = if ($part -and $script:SkuFriendlyNames.ContainsKey($part)) {
                $script:SkuFriendlyNames[$part]
            } elseif ($part) {
                $part
            } else {
                $s.skuId
            }
            $script:SkuById[[string]$s.skuId] = $name
        }
        Write-Log "Subscribed SKUs indexed: $($script:SkuById.Count)"
    } catch {
        Write-Log "Subscribed SKU fetch failed — license columns will fall back to GUIDs: $($_.Exception.Message)" 'WARN'
    }
}

function Resolve-LicenseNames {
    # Takes a user's assignedLicenses collection (objects with .skuId) and
    # returns a "; "-joined string of friendly license names. Falls back to
    # skuPartNumber, then to the raw GUID, so the column is never blank if
    # licenses are assigned.
    param(
        [object]$AssignedLicenses,
        [int]$Max = 10
    )
    if (-not $AssignedLicenses) { return $null }
    $arr = @($AssignedLicenses | Where-Object { $_ -and $_.skuId })
    if ($arr.Count -eq 0) { return $null }
    $names = foreach ($l in $arr) {
        $sid = [string]$l.skuId
        if ($script:SkuById.ContainsKey($sid)) { $script:SkuById[$sid] } else { $sid }
    }
    return Join-Strings $names -Max $Max
}

function Initialize-UserLicenseCache {
    # One bulk pull of every user with assignedLicenses; lets the user-mailbox
    # pass attach license names without an extra Graph call per mailbox.
    $script:LicenseByUpn = @{}
    if (-not $graphConnected) { return }
    Write-Log "Fetching per-user license assignments (bulk)..."
    try {
        $uri  = "/v1.0/users?`$select=id,userPrincipalName,assignedLicenses&`$top=999"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $users = @($resp.value)
        while ($resp.'@odata.nextLink') {
            $resp  = Invoke-MgGraphRequest -Method GET -Uri $resp.'@odata.nextLink' -ErrorAction Stop
            $users += $resp.value
        }
        foreach ($u in $users) {
            if ($u -and $u.userPrincipalName) {
                $script:LicenseByUpn[$u.userPrincipalName.ToLower()] = @($u.assignedLicenses)
            }
        }
        Write-Log "User license cache: $($script:LicenseByUpn.Count) users"
    } catch {
        Write-Log "User license cache fetch failed: $($_.Exception.Message)" 'WARN'
    }
}

function Get-LicenseNamesForUpn {
    param([string]$Upn)
    if (-not $Upn) { return $null }
    $key = $Upn.ToLower()
    if (-not $script:LicenseByUpn.ContainsKey($key)) { return $null }
    return Resolve-LicenseNames -AssignedLicenses $script:LicenseByUpn[$key]
}

# -------------------------------------------------------------- module bootstrap
Write-Log "Starting O365 mail/group inventory."
Ensure-Module -Name 'ExchangeOnlineManagement'        -MinVersion '3.0.0'
Ensure-Module -Name 'Microsoft.Graph.Authentication'  -MinVersion '2.0.0'
Ensure-Module -Name 'Microsoft.Graph.Groups'          -MinVersion '2.0.0'
Ensure-Module -Name 'Microsoft.Graph.Users'           -MinVersion '2.0.0'
Ensure-Module -Name 'MicrosoftTeams'                  -MinVersion '5.0.0'
Ensure-Module -Name 'ImportExcel'                     -MinVersion '7.0.0'

# --------------------------------------------------------------------- connect
# Always delegated GDAP. The wrapper passes -TenantId, -DelegatedOrganization,
# and -ClientId per customer; the running user's GDAP-delegated rights in the
# customer tenant authorize every read.

try {
    Write-Log "Connecting to Exchange Online (delegated, $DelegatedOrganization)..."
    # The cmdlet set this script needs. Passed as -CommandName so EXO V3
    # explicitly imports them rather than relying on the autoload heuristic
    # (which sometimes fails under -DelegatedOrganization on PS7 — Get-
    # UnifiedGroup in particular has been observed missing without this).
    $exoCmds = @(
        'Get-Mailbox','Get-EXOMailbox','Get-MailboxStatistics','Get-EXOMailboxStatistics',
        'Get-MailboxFolderStatistics','Get-MailboxPermission','Get-EXOMailboxPermission',
        'Get-RecipientPermission','Get-EXORecipientPermission','Get-CalendarProcessing',
        'Get-Place','Get-DistributionGroup','Get-DistributionGroupMember',
        'Get-UnifiedGroup','Get-UnifiedGroupLinks','Get-Recipient','Get-EXORecipient'
    )
    Connect-ExchangeOnline -DelegatedOrganization $DelegatedOrganization -CommandName $exoCmds -ShowBanner:$false
} catch {
    $exoMsg = $_.Exception.Message
    Write-Log "Exchange Online connection failed: $exoMsg" 'ERROR'
    if ($exoMsg -match "role assigned to user .* isn't supported in this scenario|isn't supported in this scenario") {
        Write-Log "EXO rejected the partner-tenant operator's GDAP role assignment in this customer tenant. None of the Entra roles your GDAP relationship grants you here are recognized by Exchange Online. Fix path: extend the GDAP role template for this customer to include an EXO-recognized role — Global Reader covers the read-only inventory needs cleanly; Exchange Administrator works too if you also need write access. Done in Partner Center → Customers → this customer → Admin relationships → Edit roles, or by sending a fresh GDAP request with the wider template." 'WARN'
    }
    throw
}

$graphConnected = $true
try {
    Write-Log "Connecting to Microsoft Graph (tenant $TenantId)..."
    $graphScopes = [System.Collections.Generic.List[string]]::new()
    # Core inventory
    $graphScopes.AddRange([string[]]@(
        'Group.Read.All','GroupMember.Read.All','User.Read.All',
        'Team.ReadBasic.All','Channel.ReadBasic.All','Directory.Read.All',
        'Sites.Read.All','Reports.Read.All'
    ))
    # Conditional Access pass
    if (-not $SkipConditionalAccess) {
        $graphScopes.Add('Policy.Read.All')
    }
    # Admin posture pass
    if (-not $SkipAdminPosture) {
        $graphScopes.AddRange([string[]]@(
            'RoleManagement.Read.Directory',
            'AuditLog.Read.All',
            'UserAuthenticationMethod.Read.All',
            'Application.Read.All'
        ))
    }
    # Delegated multi-tenant: token issued for $TenantId. We deliberately do
    # NOT pass -ClientId — that would route through the caller's custom
    # partner-app clientId, which most customer tenants haven't authorized
    # (AADSTS90099: "The application has not been authorized in the tenant").
    # Standard GDAP role templates don't include Cloud Application
    # Administrator, so partner-side users can't admin-consent custom apps in
    # customer tenants on first use. Letting Connect-MgGraph fall back to the
    # default Microsoft Graph PowerShell client (pre-authorized in essentially
    # every tenant after the v2.0/adminconsent flow) sidesteps that. GDAP/
    # Lighthouse roles still propagate via the running user's identity, not
    # the app. $ClientId remains accepted for future cert-based app-only flows.
    Connect-MgGraph -TenantId $TenantId -Scopes $graphScopes.ToArray() -NoWelcome
} catch {
    $exMsg = $_.Exception.Message
    Write-Log "Microsoft Graph connection failed — security-groups, CA, and admin-posture sheets will be skipped: $exMsg" 'WARN'
    if ($exMsg -match 'AADSTS90099|AADSTS65001') {
        Write-Log "AADSTS90099/65001 = the Microsoft Graph PowerShell client doesn't have a service principal in this customer tenant, or the SP exists but the elevated scopes this script needs aren't admin-consented. There's no code-side fix." 'WARN'
        Write-Log "FIX: a Cloud Application Administrator in this customer tenant needs to grant admin consent for the full scope set this script requests. Microsoft Graph PowerShell is a dynamic-scope client, so the simple /adminconsent URL only grants its tiny static scope set — use the v2.0 /adminconsent URL with the explicit scope= parameter instead. Send the customer admin THIS URL:" 'WARN'
        $scopeList  = ($graphScopes.ToArray() | ForEach-Object { "https://graph.microsoft.com/$_" }) -join ' '
        $consentUrl = 'https://login.microsoftonline.com/' + $TenantId + '/v2.0/adminconsent?' + (
            'client_id=14d82eec-204b-4c2f-b7e8-296a70dab67e' +
            '&redirect_uri=' + [System.Net.WebUtility]::UrlEncode('https://login.microsoftonline.com/common/oauth2/nativeclient') +
            '&scope=' + [System.Net.WebUtility]::UrlEncode($scopeList) +
            '&state=' + [guid]::NewGuid().Guid
        )
        Write-Log "  $consentUrl" 'WARN'
        Write-Log "After they click Accept, Microsoft drops them on a 'This is not the right page' / phishing-warning page — that's the normal landing for first-party admin-consent flows; consent did land. They can close the tab. Future runs against this tenant will succeed." 'WARN'
    }
    $graphConnected = $false
}

$teamsConnected = $true
try {
    Write-Log "Connecting to Microsoft Teams (tenant $TenantId)..."
    Connect-MicrosoftTeams -TenantId $TenantId | Out-Null
} catch {
    $exMsg = $_.Exception.Message
    Write-Log "Microsoft Teams connection failed — Teams sheet will be skipped: $exMsg" 'WARN'
    if ($exMsg -match 'AADSTS90099|AADSTS65001') {
        Write-Log "Microsoft Teams PowerShell SP is missing or unconsented in this customer tenant. Often the Graph admin-consent URL above pre-creates everything Teams needs too — confirm by re-running. If Teams data is still empty, send a separate consent URL for the Teams client (12128f48-ec9e-42f0-b203-ea49fb6af367)." 'WARN'
    }
    $teamsConnected = $false
}

$ctx = if ($graphConnected) { Get-MgContext } else { $null }
if ($ctx) { Write-Log "Tenant: $($ctx.TenantId) | Signed in as: $($ctx.Account)" }

# ---------------------------------------------------- verified-domains cache
# Pull the customer tenant's verified domains once. Used to gate per-mailbox
# perms enumeration so we don't waste cycles (and generate 401 noise) calling
# Get-EXOMailboxPermission against UPNs that obviously can't have a mailbox
# in this tenant — most commonly the partner-tenant operator's own UPN, which
# EXO sometimes surfaces as a phantom recipient in delegated GDAP sessions.
$script:VerifiedDomains = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
if ($graphConnected) {
    try {
        $orgResp = Invoke-MgGraphRequest -Method GET -Uri '/v1.0/organization?$select=verifiedDomains' -ErrorAction Stop
        foreach ($org in @($orgResp.value)) {
            foreach ($d in @($org.verifiedDomains)) {
                if ($d -and $d.name) { [void]$script:VerifiedDomains.Add($d.name) }
            }
        }
        Write-Log "Verified domains for this tenant: $($script:VerifiedDomains -join ', ')"
    } catch {
        Write-Log "Could not fetch verified domains: $($_.Exception.Message). Perms enumeration will run unfiltered." 'WARN'
    }
}

function Test-IsCustomerTenantUpn {
    # Returns $true if the UPN's domain part is in the cached customer-tenant
    # verified-domains set. Returns $true (open) if the cache is empty (Graph
    # didn't connect or the org call failed) — better to attempt and fail
    # noisily than silently drop everything.
    param([string] $Upn)
    if ([string]::IsNullOrWhiteSpace($Upn)) { return $false }
    if ($script:VerifiedDomains.Count -eq 0) { return $true }
    $domain = ($Upn -split '@')[-1]
    return $script:VerifiedDomains.Contains($domain)
}

# ----------------------------------------------------------- license catalog
# Pull the SKU friendly-name map and the per-user license assignments once
# up front. After this, every mailbox / admin row resolves licenses out of
# in-memory hashtables without an extra Graph call.
Initialize-SkuCatalog
Initialize-UserLicenseCache

# --------------------------------------------------------------- shared mailboxes
Write-Log "Collecting shared mailboxes..."
$sharedMbxs = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -PropertySets All
$sharedPermsSkipped = 0

$sharedRows = foreach ($m in $sharedMbxs) {
    $stats = $null
    if (-not $SkipMailboxStats) {
        $stats = Try-Block { Get-EXOMailboxStatistics -Identity $m.UserPrincipalName -Properties LastLogonTime,LastUserActionTime } "stats for $($m.DisplayName)"
    }
    # Gate FullAccess + SendAs perms calls behind the verified-domains check.
    # SendOnBehalf comes from the mailbox object itself, no extra call — safe
    # to read regardless. See Test-IsCustomerTenantUpn for rationale.
    $fa = @(); $sa = @(); $sob = @()
    if (-not $SkipPermissions) {
        if (Test-IsCustomerTenantUpn $m.UserPrincipalName) {
            $fa = Try-Block {
                Get-EXOMailboxPermission -Identity $m.UserPrincipalName |
                    Where-Object { $_.User -notmatch 'NT AUTHORITY|S-1-5|SELF' -and
                                   $_.AccessRights -contains 'FullAccess' -and -not $_.Deny } |
                    Select-Object -ExpandProperty User
            } "FullAccess perms for $($m.DisplayName)"
            $sa = Try-Block {
                Get-EXORecipientPermission -Identity $m.UserPrincipalName |
                    Where-Object { $_.Trustee -notmatch 'NT AUTHORITY|S-1-5|SELF' -and
                                   $_.AccessRights -contains 'SendAs' } |
                    Select-Object -ExpandProperty Trustee
            } "SendAs perms for $($m.DisplayName)"
        } else {
            $sharedPermsSkipped++
        }
        $sob = @($m.GrantSendOnBehalfTo)
    }

    [pscustomobject]@{
        DisplayName            = $m.DisplayName
        PrimarySmtpAddress     = $m.PrimarySmtpAddress
        Alias                  = $m.Alias
        SmtpDomain             = Get-SmtpDomain $m.PrimarySmtpAddress
        WhenCreated            = $m.WhenCreated
        WhenChanged            = $m.WhenChanged
        HiddenFromGAL          = $m.HiddenFromAddressListsEnabled
        ForwardingAddress      = $m.ForwardingAddress
        ForwardingSmtpAddress  = $m.ForwardingSmtpAddress
        DeliverAndForward      = $m.DeliverToMailboxAndForward
        LitigationHold         = $m.LitigationHoldEnabled
        RetentionPolicy        = $m.RetentionPolicy
        ArchiveStatus          = $m.ArchiveStatus
        ItemCount              = if ($stats) { $stats.ItemCount } else { $null }
        TotalItemSize          = if ($stats) { Format-Size $stats.TotalItemSize } else { $null }
        LastLogonTime          = if ($stats) { $stats.LastLogonTime } else { $null }
        LastUserActionTime     = if ($stats) { $stats.LastUserActionTime } else { $null }
        FullAccessDelegates    = Join-Strings $fa
        SendAsDelegates        = Join-Strings $sa
        SendOnBehalfDelegates  = Join-Strings $sob
        AcceptMessagesOnlyFrom = Join-Strings $m.AcceptMessagesOnlyFromSendersOrMembers
        RequireSenderAuth      = $m.RequireSenderAuthenticationEnabled
        AllAliases             = Join-Strings ($m.EmailAddresses | Where-Object { $_ -match '^smtp:' })
        IsDirSynced            = $m.IsDirSynced
        ExchangeObjectId       = $m.ExchangeObjectId
    }
}
Write-Log "Shared mailboxes: $($sharedRows.Count) (perms skipped on $sharedPermsSkipped cross-tenant UPN(s))"

# ----------------------------------------------------------------- user mailboxes
$userRows = @()
$userPermsSkipped = 0
if (-not $SkipUserMailboxes) {
    Write-Log "Collecting user mailboxes (this may take a while)..."
    $userMbxs = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -PropertySets All
    $userRows = foreach ($m in $userMbxs) {
        $stats = $null
        if (-not $SkipMailboxStats) {
            $stats = Try-Block { Get-EXOMailboxStatistics -Identity $m.UserPrincipalName -Properties LastLogonTime,LastUserActionTime } "stats for $($m.UserPrincipalName)"
        }
        # Same delegate pattern as Shared Mailboxes — FullAccess, SendAs,
        # SendOnBehalf. Gated by Test-IsCustomerTenantUpn so we don't waste
        # cycles on cross-tenant phantom recipients (your partner-tenant
        # account, mostly). Skipped wholesale via -SkipPermissions on huge
        # tenants where you only want structural data.
        $fa = @(); $sa = @(); $sob = @()
        if (-not $SkipPermissions) {
            if (Test-IsCustomerTenantUpn $m.UserPrincipalName) {
                $fa = Try-Block {
                    Get-EXOMailboxPermission -Identity $m.UserPrincipalName |
                        Where-Object { $_.User -notmatch 'NT AUTHORITY|S-1-5|SELF' -and
                                       $_.AccessRights -contains 'FullAccess' -and -not $_.Deny } |
                        Select-Object -ExpandProperty User
                } "FullAccess perms for $($m.UserPrincipalName)"
                $sa = Try-Block {
                    Get-EXORecipientPermission -Identity $m.UserPrincipalName |
                        Where-Object { $_.Trustee -notmatch 'NT AUTHORITY|S-1-5|SELF' -and
                                       $_.AccessRights -contains 'SendAs' } |
                        Select-Object -ExpandProperty Trustee
                } "SendAs perms for $($m.UserPrincipalName)"
            } else {
                $userPermsSkipped++
            }
            $sob = @($m.GrantSendOnBehalfTo)
        }

        [pscustomobject]@{
            DisplayName           = $m.DisplayName
            UserPrincipalName     = $m.UserPrincipalName
            PrimarySmtpAddress    = $m.PrimarySmtpAddress
            SmtpDomain            = Get-SmtpDomain $m.PrimarySmtpAddress
            WhenCreated           = $m.WhenCreated
            HiddenFromGAL         = $m.HiddenFromAddressListsEnabled
            ForwardingAddress     = $m.ForwardingAddress
            ForwardingSmtpAddress = $m.ForwardingSmtpAddress
            DeliverAndForward     = $m.DeliverToMailboxAndForward
            LitigationHold        = $m.LitigationHoldEnabled
            RetentionPolicy       = $m.RetentionPolicy
            ArchiveStatus         = $m.ArchiveStatus
            ItemCount             = if ($stats) { $stats.ItemCount } else { $null }
            TotalItemSize         = if ($stats) { Format-Size $stats.TotalItemSize } else { $null }
            LastLogonTime         = if ($stats) { $stats.LastLogonTime } else { $null }
            LastUserActionTime    = if ($stats) { $stats.LastUserActionTime } else { $null }
            FullAccessDelegates   = Join-Strings $fa
            SendAsDelegates       = Join-Strings $sa
            SendOnBehalfDelegates = Join-Strings $sob
            LicenseNames          = Get-LicenseNamesForUpn -Upn $m.UserPrincipalName
            AllAliases            = Join-Strings ($m.EmailAddresses | Where-Object { $_ -match '^smtp:' })
            IsDirSynced           = $m.IsDirSynced
        }
    }
    Write-Log "User mailboxes: $($userRows.Count) (perms skipped on $userPermsSkipped cross-tenant UPN(s))"
}

# ----------------------------------------------------------- resource mailboxes
Write-Log "Collecting resource (room/equipment) mailboxes..."
$resMbxs = Get-EXOMailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited -PropertySets All
$resourceRows = foreach ($m in $resMbxs) {
    # Get-Place only supports RoomMailbox; calling it on EquipmentMailbox
    # surfaces a noisy host-stream error from EXO that leaks past Try-Block.
    # Equipment mailboxes have no Place metadata anyway, so gate the call.
    $place = if ($m.RecipientTypeDetails -eq 'RoomMailbox') {
        Try-Block { Get-Place -Identity $m.UserPrincipalName } "Get-Place for $($m.DisplayName)"
    }
    $cal   = Try-Block { Get-CalendarProcessing -Identity $m.UserPrincipalName } "Get-CalendarProcessing for $($m.DisplayName)"
    [pscustomobject]@{
        DisplayName        = $m.DisplayName
        PrimarySmtpAddress = $m.PrimarySmtpAddress
        SmtpDomain         = Get-SmtpDomain $m.PrimarySmtpAddress
        ResourceType       = $m.RecipientTypeDetails
        WhenCreated        = $m.WhenCreated
        HiddenFromGAL      = $m.HiddenFromAddressListsEnabled
        Capacity           = $place.Capacity
        Building           = $place.Building
        Floor              = $place.Floor
        AutomateProcessing = $cal.AutomateProcessing
        AllowConflicts     = $cal.AllowConflicts
        BookingWindowDays  = $cal.BookingWindowInDays
        DelegateForward    = $cal.ForwardRequestsToDelegates
        ResourceDelegates  = Join-Strings $cal.ResourceDelegates
        AllAliases         = Join-Strings ($m.EmailAddresses | Where-Object { $_ -match '^smtp:' })
    }
}
Write-Log "Resource mailboxes: $($resourceRows.Count)"

# ------------------------------------------------------------- distribution groups
Write-Log "Collecting distribution groups..."
$dgs = Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails 'MailUniversalDistributionGroup','RoomList'
$dgRows = foreach ($g in $dgs) {
    $gId     = Resolve-GroupIdentity $g
    $members = Try-Block { Get-DistributionGroupMember -Identity $gId -ResultSize Unlimited } "members for $($g.DisplayName) [$gId]"
    [pscustomobject]@{
        DisplayName               = $g.DisplayName
        PrimarySmtpAddress        = $g.PrimarySmtpAddress
        Alias                     = $g.Alias
        SmtpDomain                = Get-SmtpDomain $g.PrimarySmtpAddress
        RecipientTypeDetails      = $g.RecipientTypeDetails
        GroupType                 = $g.GroupType
        WhenCreated               = $g.WhenCreated
        WhenChanged               = $g.WhenChanged
        HiddenFromGAL             = $g.HiddenFromAddressListsEnabled
        ManagedBy                 = Join-Strings $g.ManagedBy
        MemberCount               = @($members).Count
        MembersPreview            = Join-Strings ($members | Select-Object -ExpandProperty DisplayName) -Max $GroupMemberPreviewCount
        RequireSenderAuth         = $g.RequireSenderAuthenticationEnabled
        AcceptMessagesOnlyFrom    = Join-Strings $g.AcceptMessagesOnlyFromSendersOrMembers
        ModerationEnabled         = $g.ModerationEnabled
        ModeratedBy               = Join-Strings $g.ModeratedBy
        MemberJoinRestriction     = $g.MemberJoinRestriction
        MemberDepartRestriction   = $g.MemberDepartRestriction
        IsDirSynced               = $g.IsDirSynced
        AllAliases                = Join-Strings ($g.EmailAddresses | Where-Object { $_ -match '^smtp:' })
        ExternalDirectoryObjectId = $g.ExternalDirectoryObjectId
    }
}
Write-Log "Distribution groups: $($dgRows.Count)"

# ----------------------------------------------------- mail-enabled security groups
Write-Log "Collecting mail-enabled security groups..."
$mesgs = Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails MailUniversalSecurityGroup
$mesgRows = foreach ($g in $mesgs) {
    $gId     = Resolve-GroupIdentity $g
    $members = Try-Block { Get-DistributionGroupMember -Identity $gId -ResultSize Unlimited } "members for $($g.DisplayName) [$gId]"
    [pscustomobject]@{
        DisplayName               = $g.DisplayName
        PrimarySmtpAddress        = $g.PrimarySmtpAddress
        Alias                     = $g.Alias
        SmtpDomain                = Get-SmtpDomain $g.PrimarySmtpAddress
        WhenCreated               = $g.WhenCreated
        WhenChanged               = $g.WhenChanged
        HiddenFromGAL             = $g.HiddenFromAddressListsEnabled
        ManagedBy                 = Join-Strings $g.ManagedBy
        MemberCount               = @($members).Count
        MembersPreview            = Join-Strings ($members | Select-Object -ExpandProperty DisplayName) -Max $GroupMemberPreviewCount
        RequireSenderAuth         = $g.RequireSenderAuthenticationEnabled
        AcceptMessagesOnlyFrom    = Join-Strings $g.AcceptMessagesOnlyFromSendersOrMembers
        IsDirSynced               = $g.IsDirSynced
        AllAliases                = Join-Strings ($g.EmailAddresses | Where-Object { $_ -match '^smtp:' })
        ExternalDirectoryObjectId = $g.ExternalDirectoryObjectId
    }
}
Write-Log "Mail-enabled security groups: $($mesgRows.Count)"

# ------------------------------------------------------------ SharePoint usage
# Pull the SP usage report once before walking M365 Groups / Teams. After
# this, every per-group SP lookup is an in-memory hashtable hit.
Initialize-SharePointUsageCache

# ------------------------------------------------------------------ M365 groups
Write-Log "Collecting Microsoft 365 (Unified) Groups..."

# Defensive: Get-UnifiedGroup is an EXO implicit-remoting cmdlet. On EXO V3
# with -DelegatedOrganization (GDAP delegated session) on PowerShell 7, this
# cmdlet sometimes fails to auto-import even though Connect-ExchangeOnline
# succeeded. Without this guard, the not-recognized error crashes the entire
# tenant run before the workbook export step. Skip the M365 Groups + Teams
# sections cleanly if Get-UnifiedGroup isn't loaded; downstream code that
# depends on $ugs already handles empty/null arrays.
$ugs = $null
if (Get-Command Get-UnifiedGroup -ErrorAction SilentlyContinue) {
    $ugs = Try-Block { Get-UnifiedGroup -ResultSize Unlimited } "Get-UnifiedGroup -ResultSize Unlimited"
} else {
    Write-Log "Get-UnifiedGroup not available in this EXO session — M365 Groups + Teams sheets will be empty. (EXO V3 + -DelegatedOrganization sometimes fails to import implicit-remoting cmdlets; this is a known EXO PowerShell limitation, not a permission issue.)" 'WARN'
}
if (-not $ugs) { $ugs = @() }

# Cache SP stats per group GUID so the Teams pass reuses the same lookup.
$spStatsByGroupId = @{}

$m365Rows = foreach ($g in $ugs) {
    $gId         = Resolve-GroupIdentity $g
    $owners      = Try-Block { Get-UnifiedGroupLinks -Identity $gId -LinkType Owners      -ResultSize Unlimited } "owners for $($g.DisplayName) [$gId]"
    $members     = Try-Block { Get-UnifiedGroupLinks -Identity $gId -LinkType Members     -ResultSize Unlimited } "members for $($g.DisplayName) [$gId]"
    $subscribers = Try-Block { Get-UnifiedGroupLinks -Identity $gId -LinkType Subscribers -ResultSize Unlimited } "subscribers for $($g.DisplayName) [$gId]"

    $isTeam = $false
    if ($g.ResourceProvisioningOptions -and ($g.ResourceProvisioningOptions -contains 'Team')) { $isTeam = $true }

    # SharePoint stats — site URL comes from the unified group itself
    # (no Graph call), file count + storage from the cached usage report.
    $sp = $null
    if (-not $SkipSharePointStats -and $g.SharePointSiteUrl) {
        $sp = Get-GroupSharePointStats `
            -SharePointSiteUrl      $g.SharePointSiteUrl `
            -SharePointDocumentsUrl $g.SharePointDocumentsUrl
        if ($sp -and $g.ExternalDirectoryObjectId) { $spStatsByGroupId[$g.ExternalDirectoryObjectId] = $sp }
    }

    # Group mailbox stats — Get-EXOMailboxStatistics works against the
    # group's primary SMTP. Skipped via -SkipMailboxStats just like the
    # user / shared mailbox passes.
    $mbxStats = $null
    if (-not $SkipMailboxStats -and $g.PrimarySmtpAddress) {
        $mbxStats = Try-Block {
            Get-EXOMailboxStatistics -Identity $g.PrimarySmtpAddress -Properties LastLogonTime,LastUserActionTime
        } "mailbox stats for $($g.DisplayName) [$($g.PrimarySmtpAddress)]"
    }

    [pscustomobject]@{
        DisplayName                  = $g.DisplayName
        PrimarySmtpAddress           = $g.PrimarySmtpAddress
        Alias                        = $g.Alias
        SmtpDomain                   = Get-SmtpDomain $g.PrimarySmtpAddress
        GroupId                      = $g.ExternalDirectoryObjectId
        WhenCreated                  = $g.WhenCreated
        WhenChanged                  = $g.WhenChanged
        AccessType                   = $g.AccessType
        IsTeamProvisioned            = $isTeam
        ResourceProvisioningOptions  = Join-Strings $g.ResourceProvisioningOptions
        HiddenFromGAL                = $g.HiddenFromAddressListsEnabled
        HiddenFromExchangeClients    = $g.HiddenFromExchangeClientsEnabled
        IsArchived                   = $g.IsArchived
        Classification               = $g.Classification
        SensitivityLabel             = $g.SensitivityLabel
        OwnerCount                   = @($owners).Count
        MemberCount                  = @($members).Count
        SubscriberCount              = @($subscribers).Count
        GuestCount                   = @($members | Where-Object { $_.RecipientTypeDetails -eq 'GuestMailUser' }).Count
        Owners                       = Join-Strings ($owners | Select-Object -ExpandProperty DisplayName)
        MembersPreview               = Join-Strings ($members | Select-Object -ExpandProperty DisplayName) -Max $GroupMemberPreviewCount
        AllowExternalSenders         = (-not $g.RequireSenderAuthenticationEnabled)
        AllAliases                   = Join-Strings ($g.EmailAddresses | Where-Object { $_ -match '^smtp:' })
        # Group mailbox stats
        MailboxItemCount             = if ($mbxStats) { $mbxStats.ItemCount }          else { $null }
        MailboxTotalItemSize         = if ($mbxStats) { Format-Size $mbxStats.TotalItemSize } else { $null }
        MailboxLastLogonTime         = if ($mbxStats) { $mbxStats.LastLogonTime }       else { $null }
        MailboxLastUserActionTime    = if ($mbxStats) { $mbxStats.LastUserActionTime }  else { $null }
        # SharePoint backing site
        SharePointSiteUrl            = if ($sp) { $sp.SharePointSiteUrl }      else { $g.SharePointSiteUrl }
        SharePointDocumentsUrl       = if ($sp) { $sp.SharePointDocumentsUrl } else { $g.SharePointDocumentsUrl }
        SharePointSiteId             = if ($sp) { $sp.SharePointSiteId }       else { $null }
        FileCount                    = if ($sp) { $sp.FileCount }              else { $null }
        ActiveFileCount              = if ($sp) { $sp.ActiveFileCount }        else { $null }
        StorageUsedMB                = if ($sp) { $sp.StorageUsedMB }          else { $null }
        StorageUsedBytes             = if ($sp) { $sp.StorageUsedBytes }       else { $null }
        StorageQuotaBytes            = if ($sp) { $sp.StorageQuotaBytes }      else { $null }
        LastSPActivityDate           = if ($sp) { $sp.LastSPActivityDate }     else { $null }
        SPStatsSource                = if ($sp) { $sp.SPStatsSource }          else { $null }
    }
}
Write-Log "M365 Groups: $($m365Rows.Count) (SharePoint enrichment: $(if ($SkipSharePointStats) {'SKIPPED'} else {"$($spStatsByGroupId.Count) sites enriched"}))"

# ----------------------------------------------------------------------- Teams
$teamRows = @()
if ($teamsConnected) {
    Write-Log "Collecting Microsoft Teams..."
    $teams = Try-Block { Get-Team } "Get-Team"
    $teamRows = foreach ($t in $teams) {
        # Defensive: Get-Team can return malformed records (null GroupId,
        # null DisplayName) for teams that are mid-deprovisioning or whose
        # underlying M365 Group was deleted out from under them. Indexing
        # $spStatsByGroupId with a null key below would crash the whole
        # script with "Index operation failed; the array index evaluated
        # to null" — losing the workbook export. Skip these.
        if (-not $t -or -not $t.GroupId) {
            Write-Log "Skipping malformed Team record (null GroupId): '$($t.DisplayName)'" 'WARN'
            continue
        }
        $channels = Try-Block { Get-TeamChannel -GroupId $t.GroupId } "channels for $($t.DisplayName)"
        $owners   = Try-Block { Get-TeamUser    -GroupId $t.GroupId -Role Owner  } "owners for $($t.DisplayName)"
        $members  = Try-Block { Get-TeamUser    -GroupId $t.GroupId -Role Member } "members for $($t.DisplayName)"
        $guests   = Try-Block { Get-TeamUser    -GroupId $t.GroupId -Role Guest  } "guests for $($t.DisplayName)"

        # Cross-reference to unified group for WhenCreated / SMTP
        $ug = $ugs | Where-Object { $_.ExternalDirectoryObjectId -eq $t.GroupId } | Select-Object -First 1

        # Reuse the SP stats collected during the M365 Groups pass. If
        # the cache missed (e.g. a Team whose underlying group wasn't
        # visible to Get-UnifiedGroup), look up by site URL on the fly
        # using whatever URL we can derive from $ug.
        $sp = $spStatsByGroupId[$t.GroupId]
        if (-not $sp -and -not $SkipSharePointStats -and $ug -and $ug.SharePointSiteUrl) {
            $sp = Get-GroupSharePointStats `
                -SharePointSiteUrl      $ug.SharePointSiteUrl `
                -SharePointDocumentsUrl $ug.SharePointDocumentsUrl
        }

        [pscustomobject]@{
            DisplayName            = $t.DisplayName
            GroupId                = $t.GroupId
            MailNickname           = $t.MailNickName
            PrimarySmtpAddress     = $ug.PrimarySmtpAddress
            SmtpDomain             = Get-SmtpDomain $ug.PrimarySmtpAddress
            WhenCreated            = $ug.WhenCreated
            Visibility             = $t.Visibility
            Description            = $t.Description
            IsArchived             = $t.Archived
            ChannelCount           = @($channels).Count
            StandardChannels       = @($channels | Where-Object { $_.MembershipType -eq 'Standard' }).Count
            PrivateChannels        = @($channels | Where-Object { $_.MembershipType -eq 'Private'  }).Count
            SharedChannels         = @($channels | Where-Object { $_.MembershipType -eq 'Shared'   }).Count
            OwnerCount             = @($owners).Count
            MemberCount            = @($members).Count
            GuestCount             = @($guests).Count
            Owners                 = Join-Strings ($owners | Select-Object -ExpandProperty Name)
            AllowGuestAccess       = $t.AllowGuestCreateUpdateChannels
            AllowGiphy             = $t.AllowGiphy
            SharePointSiteUrl      = if ($sp) { $sp.SharePointSiteUrl }      else { $ug.SharePointSiteUrl }
            SharePointDocumentsUrl = if ($sp) { $sp.SharePointDocumentsUrl } else { $ug.SharePointDocumentsUrl }
            SharePointSiteId       = if ($sp) { $sp.SharePointSiteId }       else { $null }
            FileCount              = if ($sp) { $sp.FileCount }              else { $null }
            ActiveFileCount        = if ($sp) { $sp.ActiveFileCount }        else { $null }
            StorageUsedMB          = if ($sp) { $sp.StorageUsedMB }          else { $null }
            StorageUsedBytes       = if ($sp) { $sp.StorageUsedBytes }       else { $null }
            StorageQuotaBytes      = if ($sp) { $sp.StorageQuotaBytes }      else { $null }
            LastSPActivityDate     = if ($sp) { $sp.LastSPActivityDate }     else { $null }
            SPStatsSource          = if ($sp) { $sp.SPStatsSource }          else { $null }
        }
    }
    Write-Log "Teams: $($teamRows.Count)"
}

# ------------------------------------------------------------ cloud security groups
$secRows = @()
if ($graphConnected) {
    Write-Log "Collecting cloud security groups (non-mail-enabled) via Graph..."
    $secGroups = Try-Block {
        Get-MgGroup -All -ConsistencyLevel eventual -CountVariable c `
            -Filter "securityEnabled eq true and mailEnabled eq false" `
            -Property "id,displayName,description,createdDateTime,onPremisesSyncEnabled,isAssignableToRole,visibility,membershipRule,groupTypes"
    } "Get-MgGroup security filter"

    $secRows = foreach ($g in $secGroups) {
        $owners = Try-Block {
            Get-MgGroupOwner -GroupId $g.Id -All |
                ForEach-Object { $_.AdditionalProperties['displayName'] ?? $_.AdditionalProperties['userPrincipalName'] ?? $_.Id }
        } "owners for $($g.DisplayName)"

        $memberCount = $null
        try {
            $allMembers  = Get-MgGroupMember -GroupId $g.Id -All
            $memberCount = @($allMembers).Count
        } catch {
            Write-Log "members for $($g.DisplayName) failed: $($_.Exception.Message)" 'WARN'
        }

        [pscustomobject]@{
            DisplayName        = $g.DisplayName
            GroupId            = $g.Id
            Description        = $g.Description
            CreatedDateTime    = $g.CreatedDateTime
            OnPremSynced       = [bool]$g.OnPremisesSyncEnabled
            IsAssignableToRole = [bool]$g.IsAssignableToRole
            Visibility         = $g.Visibility
            IsDynamic          = ($g.GroupTypes -contains 'DynamicMembership')
            MembershipRule     = $g.MembershipRule
            OwnerCount         = @($owners).Count
            Owners             = Join-Strings $owners
            MemberCount        = $memberCount
        }
    }
    Write-Log "Cloud security groups: $($secRows.Count)"
}

# ============================================================================
# CONDITIONAL ACCESS
# ============================================================================
# Pulls four artefacts: CA policies, named locations, authentication-strength
# policies, and the tenant Authentication Methods Policy. All via raw Graph
# requests against /v1.0 — no extra cmdlet modules required.
$caPolicyRows = @()
$caNamedLocationRows = @()
$caAuthStrengthRows = @()
$authMethodsPolicyRows = @()

if (-not $SkipConditionalAccess -and $graphConnected) {
    Write-Log "Collecting Conditional Access policies..."
    $caPolicies = Try-Block {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/identity/conditionalAccess/policies?`$top=999" -ErrorAction Stop
        $all = @($resp.value)
        while ($resp.'@odata.nextLink') {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $resp.'@odata.nextLink' -ErrorAction Stop
            $all += $resp.value
        }
        $all
    } "Get CA policies"
    if ($caPolicies) {
        $caPolicyRows = foreach ($p in $caPolicies) { ConvertTo-CAPolicyRow -Policy $p }
        Write-Log "CA policies: $($caPolicyRows.Count) (enabled: $(@($caPolicyRows | Where-Object {$_.State -eq 'enabled'}).Count); report-only: $(@($caPolicyRows | Where-Object {$_.State -eq 'enabledForReportingButNotEnforced'}).Count); disabled: $(@($caPolicyRows | Where-Object {$_.State -eq 'disabled'}).Count))"
    }

    Write-Log "Collecting CA named locations..."
    $caLocs = Try-Block {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/identity/conditionalAccess/namedLocations?`$top=999" -ErrorAction Stop
        @($resp.value)
    } "Get CA named locations"
    if ($caLocs) {
        $caNamedLocationRows = foreach ($l in $caLocs) {
            $type = ($l.'@odata.type' -replace '#microsoft\.graph\.','')
            $ipRanges = $null
            $countries = $null
            if ($type -eq 'ipNamedLocation') {
                $ipRanges = Join-Strings ($l.ipRanges | ForEach-Object {
                    if ($_.cidrAddress) { $_.cidrAddress } elseif ($_.lowerAddress) { "$($_.lowerAddress)-$($_.upperAddress)" } else { $null }
                })
            } elseif ($type -eq 'countryNamedLocation') {
                $countries = Join-Strings $l.countriesAndRegions
            }
            # Cache the named location ID so CA policies that reference it can resolve to a name.
            if ($l.id) { $script:DirObjectCache[$l.id] = [pscustomobject]@{ Id = $l.id; Kind = 'namedLocation'; Display = $l.displayName } }
            [pscustomobject]@{
                DisplayName       = $l.displayName
                LocationId        = $l.id
                Type              = $type
                IsTrusted         = $l.isTrusted
                IpRanges          = $ipRanges
                Countries         = $countries
                IncludeUnknownCountries = $l.includeUnknownCountriesAndRegions
                CountryLookupMethod = $l.countryLookupMethod
                CreatedDateTime   = $l.createdDateTime
                ModifiedDateTime  = $l.modifiedDateTime
            }
        }
        Write-Log "CA named locations: $($caNamedLocationRows.Count)"
    }

    Write-Log "Collecting CA authentication-strength policies..."
    # Two endpoints expose the same data — the older /policies/...
    # path BadRequests on some tenants, so we fall through to the newer
    # /identity/conditionalAccess/authenticationStrength/policies path.
    # Neither endpoint accepts $top, so we paginate via @odata.nextLink only.
    $authStrengths = Try-Block {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/policies/authenticationStrengthPolicies" -ErrorAction Stop
        $all = @($resp.value)
        while ($resp.'@odata.nextLink') {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $resp.'@odata.nextLink' -ErrorAction Stop
            $all += $resp.value
        }
        $all
    } "Get auth strength policies (legacy path)"
    if (-not $authStrengths) {
        $authStrengths = Try-Block {
            $resp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/identity/conditionalAccess/authenticationStrength/policies" -ErrorAction Stop
            $all = @($resp.value)
            while ($resp.'@odata.nextLink') {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $resp.'@odata.nextLink' -ErrorAction Stop
                $all += $resp.value
            }
            $all
        } "Get auth strength policies (CA path)"
    }
    if ($authStrengths) {
        $caAuthStrengthRows = foreach ($a in $authStrengths) {
            [pscustomobject]@{
                DisplayName               = $a.displayName
                PolicyType                = $a.policyType         # 'builtIn' or 'custom'
                RequirementsSatisfied     = $a.requirementsSatisfied
                AllowedCombinations       = Join-Strings $a.allowedCombinations
                Description               = $a.description
                CreatedDateTime           = $a.createdDateTime
                ModifiedDateTime          = $a.modifiedDateTime
            }
        }
        Write-Log "CA authentication-strength policies: $($caAuthStrengthRows.Count)"
    }

    Write-Log "Collecting tenant Authentication Methods Policy..."
    $authMethodsPolicy = Try-Block {
        Invoke-MgGraphRequest -Method GET -Uri "/v1.0/policies/authenticationMethodsPolicy" -ErrorAction Stop
    } "Get auth methods policy"
    if ($authMethodsPolicy -and $authMethodsPolicy.authenticationMethodConfigurations) {
        $authMethodsPolicyRows = foreach ($m in $authMethodsPolicy.authenticationMethodConfigurations) {
            $methodType = ($m.'@odata.type' -replace '#microsoft\.graph\.','' -replace 'AuthenticationMethodConfiguration','')
            $includeTargets = if ($m.includeTargets) {
                Join-Strings ($m.includeTargets | ForEach-Object { "$($_.targetType):$($_.id)$(if ($_.isRegistrationRequired) { ' (req)' } else { '' })" })
            } else { $null }
            $excludeTargets = if ($m.excludeTargets) {
                Join-Strings ($m.excludeTargets | ForEach-Object { "$($_.targetType):$($_.id)" })
            } else { $null }
            [pscustomobject]@{
                Method          = $methodType
                State           = $m.state
                IncludeTargets  = $includeTargets
                ExcludeTargets  = $excludeTargets
            }
        }
        Write-Log "Authentication-methods policy: $($authMethodsPolicyRows.Count) method configurations"
    }
} elseif ($SkipConditionalAccess) {
    Write-Log "Conditional Access pass skipped via -SkipConditionalAccess." 'WARN'
}

# ============================================================================
# ADMIN POSTURE
# ============================================================================
# Pulls active and (where licensed) PIM-eligible role assignments, then
# enriches each privileged user with MFA registration state, registered
# methods, sign-in activity, account state, and license summary. Service
# principals with role assignments are captured on a separate tab. Likely
# break-glass accounts (by name pattern OR by being a Global Admin with no
# recent sign-in) are surfaced on a dedicated review tab.
$dirRoleRows = @()
$adminUserRows = @()
$adminSpRows = @()
$breakGlassRows = @()

if (-not $SkipAdminPosture -and $graphConnected) {

    # ---- shared paginator for role-management endpoints. The
    # /roleManagement/directory/* endpoints reject $top=999 with a
    # BadRequest on at least some tenants; let Graph default the page
    # size and walk @odata.nextLink ourselves.
    function Get-GraphPagedValues {
        param([Parameter(Mandatory)][string]$Uri, [string]$Context = 'graph paged GET')
        Try-Block {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
            $all = @($resp.value)
            while ($resp.'@odata.nextLink') {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $resp.'@odata.nextLink' -ErrorAction Stop
                $all += $resp.value
            }
            ,$all   # comma operator forces single-array return — protects against
                    # PowerShell's pipeline-unrolling of length-1 arrays
        } $Context
    }

    # --- Role definitions (catalog of all ~100 directory roles) ---
    Write-Log "Collecting directory role definitions..."
    $roleDefs = @(Get-GraphPagedValues -Uri "/v1.0/roleManagement/directory/roleDefinitions" -Context "Get role definitions")
    Write-Log "Role definitions: $($roleDefs.Count)"
    $roleDefById = @{}
    foreach ($r in $roleDefs) {
        if (-not $r -or -not $r.id) { continue }
        $roleDefById[$r.id] = $r
        # Cache for Resolve-DirObjectIds
        $script:DirObjectCache[$r.id] = [pscustomobject]@{ Id = $r.id; Kind = 'roleDefinition'; Display = $r.displayName }
    }

    # --- Active role memberships (permanent + currently-activated PIM) ---
    #
    # /roleManagement/directory/roleAssignments?$expand=principal is the
    # unified Microsoft Graph endpoint and returns active role assignments
    # (direct + active-PIM-activated) in a single GDAP-safe call. Earlier
    # builds also walked /directoryRoles + /directoryRoles/{id}/members,
    # but that legacy pair returns BadRequest in delegated/GDAP contexts
    # for many built-in roles, producing one WARN per role. Removed.
    #
    # PIM-eligible roles (people who can activate but currently aren't) are
    # collected separately below from /roleEligibilityScheduleInstances.
    Write-Log "Collecting active directory-role assignments..."
    $activeAssignments = @(Get-GraphPagedValues -Uri "/v1.0/roleManagement/directory/roleAssignments?`$expand=principal" -Context "Get directory-role assignments")
    # Filter out malformed rows up front so downstream code doesn't have to.
    $activeAssignments = @($activeAssignments | Where-Object {
        $_ -and $_.principalId -and $_.roleDefinitionId
    })
    Write-Log "Active role assignments: $($activeAssignments.Count)"

    # --- PIM eligibility (current-state instances, not the recurring schedule).
    #     Returns empty on tenants without Entra ID P2 / PIM licensing. ---
    Write-Log "Collecting PIM-eligible role memberships (current instances)..."
    $eligibleAssignments = @(Get-GraphPagedValues -Uri "/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?`$expand=principal" -Context "Get PIM-eligible instances")
    if ($eligibleAssignments.Count -eq 0) {
        # Fall back to the schedules view in case instances aren't exposed
        # (some tenants surface only the schedule object).
        $eligibleAssignments = @(Get-GraphPagedValues -Uri "/v1.0/roleManagement/directory/roleEligibilitySchedules?`$expand=principal" -Context "Get PIM-eligible schedules (fallback)")
    }
    if ($eligibleAssignments.Count -eq 0) {
        Write-Log "PIM eligibility list empty — tenant may not have Entra ID P2 / PIM in use." 'INFO'
    } else {
        Write-Log "PIM-eligible memberships: $($eligibleAssignments.Count)"
    }

    # --- Per-user MFA registration details (one row per user, aggregate) ---
    Write-Log "Collecting MFA user-registration-details report..."
    $mfaReg = @(Get-GraphPagedValues -Uri "/v1.0/reports/authenticationMethods/userRegistrationDetails" -Context "Get MFA registration details")
    $mfaByUserId = @{}
    foreach ($r in $mfaReg) {
        if (-not $r -or -not $r.id) { continue }
        $mfaByUserId[$r.id] = $r
    }
    Write-Log "MFA registration entries: $($mfaReg.Count)"

    # --- Build a unified list of (user|sp, role, source) tuples ---
    function Get-RoleNameSafe {
        # Order of preference:
        #   1. expanded roleDefinition.displayName on the assignment
        #   2. the cached /roleDefinitions catalog by id
        #   3. the cross-skill DirObjectCache (other passes may have resolved
        #      the same role already)
        #   4. a one-shot Graph fetch against /roleDefinitions/{id}, which
        #      we then cache in both maps so subsequent rows are cheap
        # Falling back to the raw GUID is the absolute last resort and only
        # happens if Graph itself can't resolve the role.
        param($Assignment)
        if ($null -eq $Assignment) { return $null }
        if ($Assignment.roleDefinition -and $Assignment.roleDefinition.displayName) {
            return $Assignment.roleDefinition.displayName
        }
        $rid = $Assignment.roleDefinitionId
        if (-not $rid) { return $null }
        if ($roleDefById.ContainsKey($rid) -and $roleDefById[$rid].displayName) {
            return $roleDefById[$rid].displayName
        }
        if ($script:DirObjectCache.ContainsKey($rid) -and
            $script:DirObjectCache[$rid].Display -and
            $script:DirObjectCache[$rid].Display -ne $rid) {
            return $script:DirObjectCache[$rid].Display
        }
        # Final attempt: pull this single role definition from Graph.
        try {
            $r = Invoke-MgGraphRequest -Method GET `
                    -Uri "/v1.0/roleManagement/directory/roleDefinitions/$rid" `
                    -ErrorAction Stop
            if ($r -and $r.displayName) {
                $roleDefById[$rid] = [pscustomobject]@{
                    id          = $rid
                    displayName = $r.displayName
                    description = $r.description
                    isBuiltIn   = $r.isBuiltIn
                    isEnabled   = $r.isEnabled
                }
                $script:DirObjectCache[$rid] = [pscustomobject]@{
                    Id = $rid; Kind = 'roleDefinition'; Display = $r.displayName
                }
                return $r.displayName
            }
        } catch { }
        return $rid  # absolute last resort
    }

    $allAssignments = @()
    foreach ($a in $activeAssignments) {
        if (-not $a) { continue }
        $allAssignments += [pscustomobject]@{
            PrincipalId   = $a.principalId
            Principal     = $a.principal
            RoleId        = $a.roleDefinitionId
            RoleName      = Get-RoleNameSafe -Assignment $a
            Scope         = $a.directoryScopeId
            Source        = 'Active'
            AssignmentId  = $a.id
        }
    }
    foreach ($a in $eligibleAssignments) {
        if (-not $a) { continue }
        $allAssignments += [pscustomobject]@{
            PrincipalId   = $a.principalId
            Principal     = $a.principal
            RoleId        = $a.roleDefinitionId
            RoleName      = Get-RoleNameSafe -Assignment $a
            Scope         = $a.directoryScopeId
            Source        = 'Eligible (PIM)'
            AssignmentId  = $a.id
        }
    }

    # --- Roll up per-role member counts (active + eligible). Note: do NOT
    # use $matches as a variable name here — that's a PowerShell automatic
    # variable populated by every regex -match operation, and aliasing it
    # against your own array can produce surprises down the line. ---
    $roleAssignmentsByRoleId = @{}
    if ($allAssignments.Count -gt 0) {
        foreach ($g in ($allAssignments | Group-Object -Property RoleId)) {
            if ($g.Name) { $roleAssignmentsByRoleId[$g.Name] = @($g.Group) }
        }
    }

    $dirRoleRows = foreach ($rd in $roleDefs) {
        if (-not $rd -or -not $rd.id) { continue }
        $assignments = if ($roleAssignmentsByRoleId.ContainsKey($rd.id)) { $roleAssignmentsByRoleId[$rd.id] } else { @() }
        [pscustomobject]@{
            RoleName            = $rd.displayName
            RoleId              = $rd.id
            IsBuiltIn           = $rd.isBuiltIn
            IsEnabled           = $rd.isEnabled
            ActiveAssignments   = @($assignments | Where-Object { $_.Source -eq 'Active' }).Count
            EligibleAssignments = @($assignments | Where-Object { $_.Source -eq 'Eligible (PIM)' }).Count
            Description         = $rd.description
        }
    }
    # Filter to roles that actually have assignments to keep the tab readable;
    # full role catalog is also available via the Graph if Mike wants it.
    $dirRoleRows = @($dirRoleRows | Where-Object { $_.ActiveAssignments -gt 0 -or $_.EligibleAssignments -gt 0 } | Sort-Object @{Expression='ActiveAssignments';Descending=$true}, RoleName)
    Write-Log "Directory roles with at least one assignment: $($dirRoleRows.Count)"

    # --- Cache of per-user details so we don't refetch the same user when
    #     they hold multiple roles (very common) ---
    $userDetailCache = @{}
    function Get-AdminUserDetail {
        param([string]$UserId)
        if ($userDetailCache.ContainsKey($UserId)) { return $userDetailCache[$UserId] }

        $u = Try-Block {
            Invoke-MgGraphRequest -Method GET -Uri ("/v1.0/users/{0}?`$select=id,displayName,userPrincipalName,accountEnabled,onPremisesSyncEnabled,assignedLicenses,signInActivity,createdDateTime,userType" -f $UserId) -ErrorAction Stop
        } "Get user $UserId"

        $methods = Try-Block {
            $resp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/users/$UserId/authentication/methods" -ErrorAction Stop
            @($resp.value)
        } "Get auth methods for $UserId"

        $methodKinds = @($methods | ForEach-Object {
            ($_.'@odata.type' -replace '#microsoft\.graph\.','' -replace 'AuthenticationMethod','')
        } | Sort-Object -Unique)

        $detail = [pscustomobject]@{
            User    = $u
            Methods = $methodKinds
        }
        $userDetailCache[$UserId] = $detail
        return $detail
    }

    # --- Per-(user,role) admin rows ---
    # First, normalize the principal-type discriminator on every row.
    # When $expand=principal worked, $a.Principal.'@odata.type' is set
    # ('#microsoft.graph.user' / 'group' / 'servicePrincipal'). When the
    # expand failed (Principal is null), look up the principalId via
    # /directoryObjects/{id} so the row is still classified instead of
    # silently dropped. Get-DirObjectName already caches results.
    Write-Log "Classifying principals (resolving any unexpanded ones)..."
    foreach ($a in $allAssignments) {
        $kind = if ($a.Principal -and $a.Principal.'@odata.type') {
            $a.Principal.'@odata.type'
        } else {
            $null
        }
        if (-not $kind -and $a.PrincipalId) {
            # Trigger resolution + cache. Get-DirObjectName fills the cache.
            $null = Get-DirObjectName -Id $a.PrincipalId
            if ($script:DirObjectCache.ContainsKey($a.PrincipalId)) {
                $cached = $script:DirObjectCache[$a.PrincipalId]
                # Cache.Kind is e.g. 'user' / 'group' / 'servicePrincipal' /
                # 'unknown (likely deleted)'. Synthesize the same string the
                # filter below expects.
                if ($cached.Kind -and $cached.Kind -notmatch 'unknown') {
                    $kind = "#microsoft.graph.$($cached.Kind)"
                    # Synthesize a minimal Principal so downstream code can
                    # still read displayName / userPrincipalName.
                    $a.Principal = [pscustomobject]@{
                        '@odata.type'      = $kind
                        id                 = $a.PrincipalId
                        displayName        = $cached.Display
                        userPrincipalName  = $cached.Upn
                    }
                }
            }
        }
        # Stamp resolved kind for the partition below.
        $a | Add-Member -MemberType NoteProperty -Name PrincipalKind -Value $kind -Force
    }

    Write-Log "Enriching privileged users (sign-in activity + MFA methods)..."
    $userPrincipals  = @($allAssignments | Where-Object { $_.PrincipalKind -match 'user$' })
    $spPrincipals    = @($allAssignments | Where-Object { $_.PrincipalKind -match 'servicePrincipal$' })
    $groupPrincipals = @($allAssignments | Where-Object { $_.PrincipalKind -match 'group$' })
    $unresolved      = @($allAssignments | Where-Object { -not $_.PrincipalKind })
    Write-Log ("Privileged principals: users={0}, service principals={1}, groups={2}, unresolved={3}" -f $userPrincipals.Count, $spPrincipals.Count, $groupPrincipals.Count, $unresolved.Count)
    if ($unresolved.Count -gt 0) {
        Write-Log "Some assignments could not be classified (likely deleted principals still listed) — they will appear on the Admin Users tab as raw GUIDs." 'WARN'
        # Surface unresolved rows so they aren't silently dropped — they show
        # up as "[UNRESOLVED]" entries the operator can investigate.
        foreach ($a in $unresolved) {
            $adminUserRows += [pscustomobject]@{
                DisplayName            = "[UNRESOLVED] $($a.PrincipalId)"
                UserPrincipalName      = $null
                RoleName               = $a.RoleName
                AssignmentSource       = $a.Source
                DirectoryScope         = $a.Scope
                AccountEnabled         = $null
                UserType               = 'Unresolved'
                DirSynced              = $null
                CreatedDateTime        = $null
                LastInteractiveSignIn  = $null
                LastNonInteractiveSignIn = $null
                IsMfaCapable           = $null
                IsMfaRegistered        = $null
                IsPasswordlessCapable  = $null
                IsAdmin                = $null
                DefaultMfaMethod       = $null
                RegisteredMethods      = $null
                MethodsRegisteredCount = $null
                LicenseNames           = $null
                PrincipalId            = $a.PrincipalId
                AssignmentId           = $a.AssignmentId
            }
        }
    }

    foreach ($a in $userPrincipals) {
        $detail = Get-AdminUserDetail -UserId $a.PrincipalId
        $u = $detail.User
        $mfa = $mfaByUserId[$a.PrincipalId]
        # Resolve license SKU GUIDs to friendly product names (Office 365 E3,
        # Microsoft 365 E5, etc.) using the SubscribedSkus catalog. Falls back
        # to skuPartNumber, then to the raw GUID, so the column is never
        # silently empty when licenses are assigned.
        $licenses = if ($u -and $u.assignedLicenses) {
            Resolve-LicenseNames -AssignedLicenses $u.assignedLicenses -Max 10
        } else { $null }
        $lastInteractive = if ($u -and $u.signInActivity) { $u.signInActivity.lastSignInDateTime } else { $null }
        $lastNonInteractive = if ($u -and $u.signInActivity) { $u.signInActivity.lastNonInteractiveSignInDateTime } else { $null }

        $adminUserRows += [pscustomobject]@{
            DisplayName            = if ($u) { $u.displayName } else { $a.PrincipalId }
            UserPrincipalName      = if ($u) { $u.userPrincipalName } else { $null }
            RoleName               = $a.RoleName
            AssignmentSource       = $a.Source
            DirectoryScope         = $a.Scope
            AccountEnabled         = if ($u) { $u.accountEnabled } else { $null }
            UserType               = if ($u) { $u.userType } else { $null }
            DirSynced              = if ($u) { [bool]$u.onPremisesSyncEnabled } else { $null }
            CreatedDateTime        = if ($u) { $u.createdDateTime } else { $null }
            LastInteractiveSignIn  = $lastInteractive
            LastNonInteractiveSignIn = $lastNonInteractive
            IsMfaCapable           = if ($mfa) { $mfa.isMfaCapable } else { $null }
            IsMfaRegistered        = if ($mfa) { $mfa.isMfaRegistered } else { $null }
            IsPasswordlessCapable  = if ($mfa) { $mfa.isPasswordlessCapable } else { $null }
            IsAdmin                = if ($mfa) { $mfa.isAdmin } else { $null }
            DefaultMfaMethod       = if ($mfa) { $mfa.defaultMfaMethod } else { $null }
            RegisteredMethods      = Join-Strings $detail.Methods
            MethodsRegisteredCount = if ($mfa) { @($mfa.methodsRegistered).Count } else { @($detail.Methods).Count }
            LicenseNames           = $licenses
            PrincipalId            = $a.PrincipalId
            AssignmentId           = $a.AssignmentId
        }
    }

    # --- Service principals holding directory roles ---
    foreach ($a in $spPrincipals) {
        $sp = Try-Block { Invoke-MgGraphRequest -Method GET -Uri "/v1.0/servicePrincipals/$($a.PrincipalId)" -ErrorAction Stop } "Get SP $($a.PrincipalId)"
        $owners = Try-Block {
            $resp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/servicePrincipals/$($a.PrincipalId)/owners" -ErrorAction Stop
            @($resp.value | ForEach-Object { $_.displayName ?? $_.userPrincipalName ?? $_.id })
        } "Get SP owners $($a.PrincipalId)"
        $adminSpRows += [pscustomobject]@{
            DisplayName       = if ($sp) { $sp.displayName } else { $a.PrincipalId }
            AppId             = if ($sp) { $sp.appId } else { $null }
            ServicePrincipalId = $a.PrincipalId
            ServicePrincipalType = if ($sp) { $sp.servicePrincipalType } else { $null }
            AccountEnabled    = if ($sp) { $sp.accountEnabled } else { $null }
            RoleName          = $a.RoleName
            AssignmentSource  = $a.Source
            DirectoryScope    = $a.Scope
            HomepageUrl       = if ($sp) { $sp.homepage } else { $null }
            PublisherName     = if ($sp) { $sp.publisherName } else { $null }
            VerifiedPublisher = if ($sp -and $sp.verifiedPublisher) { $sp.verifiedPublisher.displayName } else { $null }
            Owners            = Join-Strings $owners -Max 10
            AssignmentId      = $a.AssignmentId
        }
    }

    # --- Groups holding directory roles (via role-assignable groups) ---
    foreach ($a in $groupPrincipals) {
        # Surface as a synthetic admin row so the Admin Users tab covers
        # group-based admin grants too. We don't expand into individual
        # member rows here; the cloud-security-groups tab already has the
        # member counts and can be cross-referenced by GroupId.
        $g = Try-Block { Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groups/$($a.PrincipalId)?`$select=id,displayName,isAssignableToRole,securityEnabled" -ErrorAction Stop } "Get group $($a.PrincipalId)"
        $adminUserRows += [pscustomobject]@{
            DisplayName            = if ($g) { "[GROUP] $($g.displayName)" } else { "[GROUP] $($a.PrincipalId)" }
            UserPrincipalName      = $null
            RoleName               = $a.RoleName
            AssignmentSource       = $a.Source
            DirectoryScope         = $a.Scope
            AccountEnabled         = $null
            UserType               = 'Group'
            DirSynced              = $null
            CreatedDateTime        = $null
            LastInteractiveSignIn  = $null
            LastNonInteractiveSignIn = $null
            IsMfaCapable           = $null
            IsMfaRegistered        = $null
            IsPasswordlessCapable  = $null
            IsAdmin                = $null
            DefaultMfaMethod       = $null
            RegisteredMethods      = $null
            MethodsRegisteredCount = $null
            LicenseNames           = $null
            PrincipalId            = $a.PrincipalId
            AssignmentId           = $a.AssignmentId
        }
    }

    Write-Log "Admin user rows (incl. groups): $($adminUserRows.Count); admin SP rows: $($adminSpRows.Count)"

    # --- Break-Glass Candidates ---
    # Two heuristics: name pattern match OR (Global Admin AND no sign-in within $InactiveAdminDays).
    Write-Log "Identifying break-glass candidates (pattern: $BreakGlassNamePattern; inactive threshold: ${InactiveAdminDays}d)..."
    $bgSeen = @{}
    foreach ($row in $adminUserRows) {
        if (-not $row.UserPrincipalName) { continue }  # skip groups
        $reasons = @()
        if ($row.DisplayName -match $BreakGlassNamePattern -or $row.UserPrincipalName -match $BreakGlassNamePattern) {
            $reasons += 'Name pattern match'
        }
        if ($row.RoleName -eq 'Global Administrator') {
            $cutoff = (Get-Date).AddDays(-$InactiveAdminDays)
            $li = if ($row.LastInteractiveSignIn) { try { [datetime]$row.LastInteractiveSignIn } catch { $null } } else { $null }
            $ni = if ($row.LastNonInteractiveSignIn) { try { [datetime]$row.LastNonInteractiveSignIn } catch { $null } } else { $null }
            $latest = @($li, $ni) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
            if (-not $latest -or $latest -lt $cutoff) {
                $reasons += "Global Admin with no sign-in in last ${InactiveAdminDays}d"
            }
        }
        if ($reasons.Count -gt 0) {
            $key = $row.UserPrincipalName
            if ($bgSeen.ContainsKey($key)) {
                # Merge reasons + roles
                $existing = $bgSeen[$key]
                $existing.Reasons = (@($existing.Reasons.Split('; ')) + $reasons | Sort-Object -Unique) -join '; '
                $existing.RolesHeld = (@($existing.RolesHeld.Split('; ')) + $row.RoleName | Sort-Object -Unique) -join '; '
            } else {
                $bgSeen[$key] = [pscustomobject]@{
                    DisplayName            = $row.DisplayName
                    UserPrincipalName      = $row.UserPrincipalName
                    AccountEnabled         = $row.AccountEnabled
                    RolesHeld              = $row.RoleName
                    Reasons                = ($reasons -join '; ')
                    LastInteractiveSignIn  = $row.LastInteractiveSignIn
                    LastNonInteractiveSignIn = $row.LastNonInteractiveSignIn
                    IsMfaRegistered        = $row.IsMfaRegistered
                    RegisteredMethods      = $row.RegisteredMethods
                    InactiveAdminDays      = $InactiveAdminDays
                    BreakGlassPattern      = $BreakGlassNamePattern
                }
            }
        }
    }
    $breakGlassRows = @($bgSeen.Values)
    Write-Log "Break-glass candidates: $($breakGlassRows.Count)"

} elseif ($SkipAdminPosture) {
    Write-Log "Admin posture pass skipped via -SkipAdminPosture." 'WARN'
}

# -------------------------------------------------------------------- export xlsx
Write-Log "Exporting workbook to: $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$summary = @(
    # Mail-flow inventory
    [pscustomobject]@{ Category = 'Shared Mailboxes';                   Count = @($sharedRows).Count }
    [pscustomobject]@{ Category = 'User Mailboxes';                     Count = @($userRows).Count }
    [pscustomobject]@{ Category = 'Resource Mailboxes (Room/Equipment)';Count = @($resourceRows).Count }
    [pscustomobject]@{ Category = 'Distribution Groups';                Count = @($dgRows).Count }
    [pscustomobject]@{ Category = 'Mail-Enabled Security Groups';       Count = @($mesgRows).Count }
    [pscustomobject]@{ Category = 'M365 Groups (Unified)';              Count = @($m365Rows).Count }
    [pscustomobject]@{ Category = '  -> Teams-provisioned';             Count = @($m365Rows | Where-Object { $_.IsTeamProvisioned }).Count }
    [pscustomobject]@{ Category = 'Microsoft Teams';                    Count = @($teamRows).Count }
    [pscustomobject]@{ Category = 'Cloud Security Groups (non-mail)';   Count = @($secRows).Count }
    # Conditional Access
    [pscustomobject]@{ Category = '— CA —';                              Count = $null }
    [pscustomobject]@{ Category = 'CA Policies (total)';                 Count = @($caPolicyRows).Count }
    [pscustomobject]@{ Category = '  -> enabled';                        Count = @($caPolicyRows | Where-Object { $_.State -eq 'enabled' }).Count }
    [pscustomobject]@{ Category = '  -> report-only';                    Count = @($caPolicyRows | Where-Object { $_.State -eq 'enabledForReportingButNotEnforced' }).Count }
    [pscustomobject]@{ Category = '  -> disabled';                       Count = @($caPolicyRows | Where-Object { $_.State -eq 'disabled' }).Count }
    [pscustomobject]@{ Category = 'CA Named Locations';                  Count = @($caNamedLocationRows).Count }
    [pscustomobject]@{ Category = '  -> trusted';                        Count = @($caNamedLocationRows | Where-Object { $_.IsTrusted }).Count }
    [pscustomobject]@{ Category = 'CA Authentication-Strength Policies'; Count = @($caAuthStrengthRows).Count }
    [pscustomobject]@{ Category = 'Authentication-Methods Policy entries';Count = @($authMethodsPolicyRows).Count }
    [pscustomobject]@{ Category = '  -> methods enabled';                Count = @($authMethodsPolicyRows | Where-Object { $_.State -eq 'enabled' }).Count }
    # Admin posture
    [pscustomobject]@{ Category = '— Admin posture —';                    Count = $null }
    [pscustomobject]@{ Category = 'Directory roles with assignments';     Count = @($dirRoleRows).Count }
    [pscustomobject]@{ Category = 'Admin assignments — active';           Count = @($adminUserRows | Where-Object { $_.AssignmentSource -eq 'Active' }).Count + @($adminSpRows | Where-Object { $_.AssignmentSource -eq 'Active' }).Count }
    [pscustomobject]@{ Category = 'Admin assignments — eligible (PIM)';   Count = @($adminUserRows | Where-Object { $_.AssignmentSource -eq 'Eligible (PIM)' }).Count + @($adminSpRows | Where-Object { $_.AssignmentSource -eq 'Eligible (PIM)' }).Count }
    [pscustomobject]@{ Category = 'Distinct human admins';                Count = @($adminUserRows | Where-Object { $_.UserPrincipalName } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count }
    [pscustomobject]@{ Category = '  -> Global Administrator role holders';Count = @($adminUserRows | Where-Object { $_.RoleName -eq 'Global Administrator' -and $_.UserPrincipalName } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count }
    [pscustomobject]@{ Category = '  -> admins NOT MFA-registered';       Count = @($adminUserRows | Where-Object { $_.UserPrincipalName -and $_.IsMfaRegistered -eq $false } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count }
    [pscustomobject]@{ Category = 'Admin role-holding service principals (distinct)'; Count = @($adminSpRows | Select-Object -ExpandProperty AppId -Unique).Count }
    [pscustomobject]@{ Category = 'Admin role-holding groups';            Count = @($adminUserRows | Where-Object { $_.UserType -eq 'Group' } | Select-Object -ExpandProperty PrincipalId -Unique).Count }
    [pscustomobject]@{ Category = 'Break-glass candidates';               Count = @($breakGlassRows).Count }
)

$meta = @(
    [pscustomobject]@{ Field = 'Run Timestamp';         Value = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
    [pscustomobject]@{ Field = 'Tenant ID';             Value = if ($ctx) { $ctx.TenantId } else { '(graph not connected)' } }
    [pscustomobject]@{ Field = 'Run As';                Value = if ($ctx) { $ctx.Account } else { (whoami) } }
    [pscustomobject]@{ Field = 'OutputPath';            Value = $OutputPath }
    [pscustomobject]@{ Field = 'SkipMailboxStats';      Value = [bool]$SkipMailboxStats }
    [pscustomobject]@{ Field = 'SkipPermissions';       Value = [bool]$SkipPermissions }
    [pscustomobject]@{ Field = 'SkipUserMailboxes';     Value = [bool]$SkipUserMailboxes }
    [pscustomobject]@{ Field = 'SkipSharePointStats';   Value = [bool]$SkipSharePointStats }
    [pscustomobject]@{ Field = 'SkipConditionalAccess'; Value = [bool]$SkipConditionalAccess }
    [pscustomobject]@{ Field = 'SkipAdminPosture';      Value = [bool]$SkipAdminPosture }
    [pscustomobject]@{ Field = 'BreakGlassNamePattern'; Value = $BreakGlassNamePattern }
    [pscustomobject]@{ Field = 'InactiveAdminDays';     Value = $InactiveAdminDays }
    [pscustomobject]@{ Field = 'GraphConnected';        Value = $graphConnected }
    [pscustomobject]@{ Field = 'TeamsConnected';        Value = $teamsConnected }
)

$xl = @{ AutoSize = $true; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true }

# Summary tab — metadata block on top, counts table below it.
$meta    | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -StartRow 1                                @xl
$summary | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -StartRow ($meta.Count + 3) -TableName 'Counts' -TableStyle Medium2

# Mail / group inventory tabs
if ($sharedRows)   { $sharedRows   | Export-Excel -Path $OutputPath -WorksheetName 'Shared Mailboxes'        @xl }
if ($userRows)     { $userRows     | Export-Excel -Path $OutputPath -WorksheetName 'User Mailboxes'          @xl }
if ($resourceRows) { $resourceRows | Export-Excel -Path $OutputPath -WorksheetName 'Resource Mailboxes'      @xl }
if ($dgRows)       { $dgRows       | Export-Excel -Path $OutputPath -WorksheetName 'Distribution Groups'     @xl }
if ($mesgRows)     { $mesgRows     | Export-Excel -Path $OutputPath -WorksheetName 'Mail-Enabled Sec Groups' @xl }
if ($m365Rows)     { $m365Rows     | Export-Excel -Path $OutputPath -WorksheetName 'M365 Groups'             @xl }
if ($teamRows)     { $teamRows     | Export-Excel -Path $OutputPath -WorksheetName 'Teams'                   @xl }
if ($secRows)      { $secRows      | Export-Excel -Path $OutputPath -WorksheetName 'Security Groups (cloud)' @xl }

# Conditional Access tabs
if ($caPolicyRows)         { $caPolicyRows         | Export-Excel -Path $OutputPath -WorksheetName 'CA Policies'          @xl }
if ($caNamedLocationRows)  { $caNamedLocationRows  | Export-Excel -Path $OutputPath -WorksheetName 'CA Named Locations'   @xl }
if ($caAuthStrengthRows)   { $caAuthStrengthRows   | Export-Excel -Path $OutputPath -WorksheetName 'CA Auth Strengths'    @xl }
if ($authMethodsPolicyRows){ $authMethodsPolicyRows| Export-Excel -Path $OutputPath -WorksheetName 'Auth Methods Policy'  @xl }

# Admin posture tabs
if ($dirRoleRows)    { $dirRoleRows    | Export-Excel -Path $OutputPath -WorksheetName 'Directory Roles'         @xl }
if ($adminUserRows)  { $adminUserRows  | Export-Excel -Path $OutputPath -WorksheetName 'Admin Users'             @xl }
if ($adminSpRows)    { $adminSpRows    | Export-Excel -Path $OutputPath -WorksheetName 'Admin Service Principals' @xl }
if ($breakGlassRows) { $breakGlassRows | Export-Excel -Path $OutputPath -WorksheetName 'Break-Glass Candidates'  @xl }

$script:RunLog     | Export-Excel -Path $OutputPath -WorksheetName 'Run Log'                                  @xl

Write-Log "Done. Workbook at: $OutputPath"

# --------------------------------------------------------------------- disconnect
Try-Block { Disconnect-ExchangeOnline -Confirm:$false } 'Disconnect-ExchangeOnline' | Out-Null
if ($graphConnected) { Try-Block { Disconnect-MgGraph } 'Disconnect-MgGraph' | Out-Null }
if ($teamsConnected) { Try-Block { Disconnect-MicrosoftTeams } 'Disconnect-MicrosoftTeams' | Out-Null }
