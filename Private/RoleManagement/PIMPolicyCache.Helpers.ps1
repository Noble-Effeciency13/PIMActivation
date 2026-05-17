<#
.SYNOPSIS
    Reads the first available property value from an object.

.DESCRIPTION
    Safely reads a list of candidate property names from either a hashtable or a PowerShell object.
    This keeps the policy-cache helpers tolerant of live objects, cloned role snapshots, and JSON
    objects that have been deserialized from disk. Null and blank string values are treated as missing.

.PARAMETER InputObject
    The object or hashtable to inspect.

.PARAMETER PropertyNames
    Candidate property names to check in order. The first non-empty value is returned.

.OUTPUTS
    System.Object. Returns the first matching value, or $null when no usable value exists.
#>
function Get-PIMPolicyObjectValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames
    )

    if ($null -eq $InputObject) { return $null }

    foreach ($propertyName in $PropertyNames) {
        if ($InputObject -is [hashtable] -and $InputObject.ContainsKey($propertyName)) {
            $value = $InputObject[$propertyName]
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { return $value }
        }

        if ($InputObject.PSObject.Properties[$propertyName]) {
            $value = $InputObject.$propertyName
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { return $value }
        }
    }

    return $null
}

<#
.SYNOPSIS
    Creates a stable SHA-256 hash for policy-cache values.

.DESCRIPTION
    Converts a string to a lowercase SHA-256 hash. The helper is used to create tenant-scoped cache
    folder names and build stable policy-content hashes without storing raw tenant identifiers in file
    paths.

.PARAMETER Value
    The string value to hash. Null values are normalized to an empty string.

.OUTPUTS
    System.String. A lowercase hexadecimal SHA-256 hash.
#>
function ConvertTo-PIMPolicyCacheHash {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Value
    )

    $normalizedValue = if ($null -eq $Value) { '' } else { $Value }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalizedValue)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha256.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

<#
.SYNOPSIS
    Converts common serialized boolean values to a Boolean.

.DESCRIPTION
    Normalizes Boolean-like values from live policy objects and JSON cache records. The function accepts
    native Boolean values and common string or numeric representations such as true, false, yes, no, 1,
    and 0. Unknown values resolve to $false.

.PARAMETER Value
    The value to normalize.

.OUTPUTS
    System.Boolean.
#>
function ConvertTo-PIMPolicyCacheBool {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }

    $valueText = [string]$Value
    if ([string]::IsNullOrWhiteSpace($valueText)) { return $false }
    if ($valueText -match '^(1|true|yes)$') { return $true }
    if ($valueText -match '^(0|false|no)$') { return $false }

    return $false
}

<#
.SYNOPSIS
    Converts a cache timestamp to UTC.

.DESCRIPTION
    Normalizes DateTime and string timestamp values to UTC. PowerShell can deserialize ISO JSON timestamps
    back as DateTime objects, and string conversion can become culture-sensitive, so this helper keeps cache
    freshness checks consistent across locales.

.PARAMETER Value
    The timestamp value to convert.

.PARAMETER DefaultValue
    The UTC fallback value to use when the input cannot be parsed. Defaults to the current UTC time.

.OUTPUTS
    System.DateTime. A UTC DateTime value.
#>
function ConvertTo-PIMPolicyCacheUtcDateTime {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value,

        [DateTime]$DefaultValue = ([DateTime]::UtcNow)
    )

    if ($null -eq $Value) { return $DefaultValue.ToUniversalTime() }
    if ($Value -is [DateTime]) { return $Value.ToUniversalTime() }

    $valueText = [string]$Value
    if ([string]::IsNullOrWhiteSpace($valueText)) { return $DefaultValue.ToUniversalTime() }

    try {
        $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
        return ([DateTime]::Parse($valueText, [System.Globalization.CultureInfo]::InvariantCulture, $styles)).ToUniversalTime()
    }
    catch {
        try { return ([DateTime]::Parse($valueText)).ToUniversalTime() } catch { return $DefaultValue.ToUniversalTime() }
    }
}

<#
.SYNOPSIS
    Resolves the tenant scope for persistent policy cache files.

.DESCRIPTION
    Builds a cache scope from the current Graph tenant. Policy metadata is tenant-wide, so no account
    information is used for the persistent cache path or metadata.

.OUTPUTS
    PSCustomObject. Contains tenant values and hash-derived cache scope identifiers.
#>
function Get-PIMPolicyCacheScope {
    [CmdletBinding()]
    param()

    $tenantId = $null
    if ((Get-Variable -Name 'CurrentTenantId' -Scope Script -ErrorAction SilentlyContinue) -and -not [string]::IsNullOrWhiteSpace([string]$script:CurrentTenantId)) {
        $tenantId = [string]$script:CurrentTenantId
    }
    elseif ((Get-Variable -Name 'GraphContext' -Scope Script -ErrorAction SilentlyContinue) -and $script:GraphContext -and $script:GraphContext.PSObject.Properties['TenantId'] -and $script:GraphContext.TenantId) {
        $tenantId = [string]$script:GraphContext.TenantId
    }
    else {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($graphContext -and $graphContext.PSObject.Properties['TenantId'] -and $graphContext.TenantId) {
                $tenantId = [string]$graphContext.TenantId
            }
        }
        catch { }
    }

    if ([string]::IsNullOrWhiteSpace($tenantId)) { $tenantId = 'unknown-tenant' }

    $tenantHash = (ConvertTo-PIMPolicyCacheHash -Value $tenantId).Substring(0, 32)

    return [PSCustomObject]@{
        TenantId   = $tenantId
        TenantHash = $tenantHash
        ScopeHash  = $tenantHash
        ScopeType  = 'Tenant'
    }
}

<#
.SYNOPSIS
    Gets the folder used for the current policy-cache scope.

.DESCRIPTION
    Resolves the per-tenant policy-cache directory under %LOCALAPPDATA%\PIMActivation\PolicyCache.
    The final folder name is a hash of the tenant scope, avoiding raw tenant identifiers in the local
    file-system path.

.PARAMETER Create
    Creates the directory when it does not already exist.

.OUTPUTS
    System.String. The full cache directory path.
#>
function Get-PIMPolicyCacheStorePath {
    [CmdletBinding()]
    param(
        [switch]$Create
    )

    $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    if ([string]::IsNullOrWhiteSpace($localAppData)) { $localAppData = $env:LOCALAPPDATA }
    if ([string]::IsNullOrWhiteSpace($localAppData)) { throw 'LOCALAPPDATA could not be resolved.' }

    $scope = Get-PIMPolicyCacheScope
    $cacheRoot = Join-Path (Join-Path $localAppData 'PIMActivation') 'PolicyCache'
    $cachePath = Join-Path $cacheRoot $scope.ScopeHash

    if ($Create -and -not (Test-Path -Path $cachePath)) {
        New-Item -Path $cachePath -ItemType Directory -Force | Out-Null
    }

    return $cachePath
}

<#
.SYNOPSIS
    Gets the persistent policy-cache JSON file path.

.DESCRIPTION
    Returns the policy-cache file path for the current tenant scope. The file stores only sanitized policy
    requirement metadata, not tokens, authentication-context tokens, display-name lookups, or activation
    request data.

.PARAMETER Create
    Creates the scoped cache directory before returning the file path.

.OUTPUTS
    System.String. The full path to policy-cache.json.
#>
function Get-PIMPolicyCacheFilePath {
    [CmdletBinding()]
    param(
        [switch]$Create
    )

    return (Join-Path (Get-PIMPolicyCacheStorePath -Create:$Create) 'policy-cache.json')
}

<#
.SYNOPSIS
    Builds the in-memory policy-cache key for a role.

.DESCRIPTION
    Produces the same cache-key shape used by the existing policy code for Entra roles, PIM groups, and
    Azure Resource roles. This centralizes key construction so disk cache, UI refresh, and policy fetch
    paths all address the same policy entries.

.PARAMETER Role
    A role object, cloned role snapshot, or hashtable containing role identity properties.

.OUTPUTS
    System.String. The cache key, or $null when the role identity cannot be determined.
#>
function Get-PIMPolicyCacheKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Role
    )

    $roleType = [string](Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('Type'))
    if ($roleType -eq 'Azure') { $roleType = 'AzureResource' }

    switch ($roleType) {
        'Group' {
            $groupId = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('GroupId', 'ResourceId', 'Id')
            if ($groupId) { return "Group_$groupId" }
        }
        'Entra' {
            $roleId = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('RoleDefinitionId', 'Id')
            if ($roleId) { return "Entra_$roleId" }
        }
        'AzureResource' {
            $roleDefinitionId = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('RoleDefinitionId', 'Id')
            if ($roleDefinitionId) { return "AzureResource_$roleDefinitionId" }
        }
        default {
            $roleId = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('RoleDefinitionId', 'GroupId', 'ResourceId', 'Id')
            if ($roleType -and $roleId) { return "${roleType}_$roleId" }
        }
    }

    return $null
}

<#
.SYNOPSIS
    Converts policy information into a sanitized persistent cache record.

.DESCRIPTION
    Copies only non-secret policy metadata into a JSON-friendly ordered dictionary. Tokens, activation
    payloads, justifications, ticket values, and runtime authentication details are deliberately excluded.
    A content hash is included so refreshed policy metadata can be compared without storing raw source objects.

.PARAMETER CacheKey
    The policy-cache key associated with the policy information.

.PARAMETER PolicyInfo
    The live policy information object to persist.

.OUTPUTS
    System.Collections.Specialized.OrderedDictionary. A JSON-safe policy-cache record.
#>
function ConvertTo-PIMPolicyCacheRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CacheKey,

        [Parameter(Mandatory)]
        [object]$PolicyInfo
    )

    $fetchedAtValue = Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('PIMCacheFetchedAt', 'FetchedAt', 'CachedAt')
    $fetchedAt = ConvertTo-PIMPolicyCacheUtcDateTime -Value $fetchedAtValue

    $maxDurationValue = Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('MaxDuration')
    $maxDuration = 8
    if ($maxDurationValue) {
        try { $maxDuration = [int]$maxDurationValue } catch { $maxDuration = 8 }
    }

    $hashPayload = [ordered]@{
        MaxDuration                      = $maxDuration
        RequiresMfa                      = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('RequiresMfa', 'RequiresMFA'))
        RequiresJustification            = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('RequiresJustification'))
        RequiresTicket                   = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('RequiresTicket'))
        RequiresApproval                 = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('RequiresApproval'))
        RequiresAuthenticationContext    = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('RequiresAuthenticationContext'))
        AuthenticationContextId          = Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('AuthenticationContextId')
        AuthenticationContextDisplayName = Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('AuthenticationContextDisplayName')
        AuthenticationContextDescription = Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('AuthenticationContextDescription')
        PolicyUnavailable                = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('PolicyUnavailable'))
    }

    $hash = ConvertTo-PIMPolicyCacheHash -Value ($hashPayload | ConvertTo-Json -Depth 8 -Compress)

    return [ordered]@{
        CacheKey                         = $CacheKey
        MaxDuration                      = $hashPayload.MaxDuration
        RequiresMfa                      = $hashPayload.RequiresMfa
        RequiresJustification            = $hashPayload.RequiresJustification
        RequiresTicket                   = $hashPayload.RequiresTicket
        RequiresApproval                 = $hashPayload.RequiresApproval
        RequiresAuthenticationContext    = $hashPayload.RequiresAuthenticationContext
        AuthenticationContextId          = $hashPayload.AuthenticationContextId
        AuthenticationContextDisplayName = $hashPayload.AuthenticationContextDisplayName
        AuthenticationContextDescription = $hashPayload.AuthenticationContextDescription
        AuthenticationContextDetails     = $null
        PolicyUnavailable                = $hashPayload.PolicyUnavailable
        FetchedAt                        = $fetchedAt.ToString('o')
        Hash                             = $hash
    }
}

<#
.SYNOPSIS
    Converts a persisted policy record back to runtime policy information.

.DESCRIPTION
    Rehydrates sanitized policy metadata from disk into the same shape expected by the UI and activation
    dialog. The returned object is marked as disk-sourced and keeps the cache timestamp/hash for freshness
    checks and background revalidation.

.PARAMETER Record
    The deserialized JSON policy record.

.OUTPUTS
    PSCustomObject. A policy information object suitable for $script:PolicyCache.
#>
function ConvertFrom-PIMPolicyCacheRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Record
    )

    $maxDurationValue = Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('MaxDuration')
    $maxDuration = 8
    if ($maxDurationValue) {
        try { $maxDuration = [int]$maxDurationValue } catch { $maxDuration = 8 }
    }

    return [PSCustomObject]@{
        MaxDuration                      = $maxDuration
        RequiresMfa                      = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('RequiresMfa', 'RequiresMFA'))
        RequiresJustification            = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('RequiresJustification'))
        RequiresTicket                   = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('RequiresTicket'))
        RequiresApproval                 = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('RequiresApproval'))
        RequiresAuthenticationContext    = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('RequiresAuthenticationContext'))
        AuthenticationContextId          = Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('AuthenticationContextId')
        AuthenticationContextDisplayName = Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('AuthenticationContextDisplayName')
        AuthenticationContextDescription = Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('AuthenticationContextDescription')
        AuthenticationContextDetails     = $null
        PolicyUnavailable                = ConvertTo-PIMPolicyCacheBool -Value (Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('PolicyUnavailable'))
        PIMCacheFetchedAt                = Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('FetchedAt', 'PIMCacheFetchedAt')
        PIMCacheHash                     = Get-PIMPolicyObjectValue -InputObject $Record -PropertyNames @('Hash', 'PIMCacheHash')
        PIMCacheSource                   = 'Disk'
    }
}

<#
.SYNOPSIS
    Tests whether a cached policy entry is still fresh.

.DESCRIPTION
    Compares the policy entry timestamp against the configured stale threshold. This is used by the
    background refresh path to decide which disk-loaded policies should be revalidated after the UI has
    rendered.

.PARAMETER PolicyInfo
    The policy information object to check.

.PARAMETER MaxAgeHours
    Optional freshness threshold. When omitted or zero, the module-level PIMPolicyCacheStaleAfterHours
    value is used, with a default of 24 hours.

.OUTPUTS
    System.Boolean. True when the policy timestamp is inside the freshness window.
#>
function Test-PIMPolicyCacheEntryFresh {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$PolicyInfo,

        [int]$MaxAgeHours = 0
    )

    if ($MaxAgeHours -le 0) {
        $MaxAgeHours = if ((Get-Variable -Name 'PIMPolicyCacheStaleAfterHours' -Scope Script -ErrorAction SilentlyContinue) -and $script:PIMPolicyCacheStaleAfterHours) {
            [int]$script:PIMPolicyCacheStaleAfterHours
        }
        else { 24 }
    }

    $fetchedAtValue = Get-PIMPolicyObjectValue -InputObject $PolicyInfo -PropertyNames @('PIMCacheFetchedAt', 'FetchedAt', 'CachedAt')
    if (-not $fetchedAtValue) { return $false }

    try {
        $fetchedAt = ConvertTo-PIMPolicyCacheUtcDateTime -Value $fetchedAtValue
        return (([DateTime]::UtcNow - $fetchedAt).TotalHours -lt $MaxAgeHours)
    }
    catch {
        return $false
    }
}

<#
.SYNOPSIS
    Loads the persistent PIM policy cache for the current signed-in scope.

.DESCRIPTION
    Reads sanitized policy metadata from disk and merges it into the module's in-memory cache. The cache
    file is accepted when its hashed tenant scope matches the current session. Old records beyond the
    maximum age window are ignored.

.PARAMETER Force
    Reloads the cache even if it was already loaded for the current scope.

.PARAMETER MaxAgeDays
    Maximum persisted record age in days. When omitted or zero, PIMPolicyCacheMaxAgeDays is used, with a
    default of 30 days.

.OUTPUTS
    PSCustomObject. Includes whether a cache was loaded and how many policies and authentication contexts
    were imported.
#>
function Import-PIMPolicyCache {
    [CmdletBinding()]
    param(
        [switch]$Force,

        [int]$MaxAgeDays = 0
    )

    if ($MaxAgeDays -le 0) {
        $MaxAgeDays = if ((Get-Variable -Name 'PIMPolicyCacheMaxAgeDays' -Scope Script -ErrorAction SilentlyContinue) -and $script:PIMPolicyCacheMaxAgeDays) {
            [int]$script:PIMPolicyCacheMaxAgeDays
        }
        else { 30 }
    }

    $scope = Get-PIMPolicyCacheScope
    if (-not $Force -and (Get-Variable -Name 'PIMPolicyCacheLoadedForScope' -Scope Script -ErrorAction SilentlyContinue) -and $script:PIMPolicyCacheLoadedForScope -eq $scope.ScopeHash) {
        return [PSCustomObject]@{ Loaded = $false; PolicyCount = 0; AuthenticationContextCount = 0; Reason = 'AlreadyLoaded' }
    }

    $cacheFile = Get-PIMPolicyCacheFilePath
    if (-not (Test-Path -Path $cacheFile)) {
        $script:PIMPolicyCacheLoadedForScope = $scope.ScopeHash
        return [PSCustomObject]@{ Loaded = $false; PolicyCount = 0; AuthenticationContextCount = 0; Reason = 'Missing' }
    }

    try {
        $cacheData = Get-Content -Path $cacheFile -Raw -ErrorAction Stop | ConvertFrom-Json -Depth 20 -ErrorAction Stop
        $scopeMatches = $cacheData -and $cacheData.PSObject.Properties['ScopeHash'] -and $cacheData.ScopeHash -eq $scope.ScopeHash
        $tenantMatches = $cacheData -and $cacheData.PSObject.Properties['TenantHash'] -and $cacheData.TenantHash -eq $scope.TenantHash
        if (-not $scopeMatches -and -not $tenantMatches) {
            return [PSCustomObject]@{ Loaded = $false; PolicyCount = 0; AuthenticationContextCount = 0; Reason = 'ScopeMismatch' }
        }

        if (-not (Get-Variable -Name 'PolicyCache' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:PolicyCache) { $script:PolicyCache = @{} }
        if (-not (Get-Variable -Name 'AuthenticationContextCache' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:AuthenticationContextCache) { $script:AuthenticationContextCache = @{} }

        $cutoff = [DateTime]::UtcNow.AddDays(-1 * $MaxAgeDays)
        $policyCount = 0
        $contextCount = 0

        if ($cacheData.PSObject.Properties['Policies'] -and $cacheData.Policies) {
            foreach ($policyProperty in $cacheData.Policies.PSObject.Properties) {
                $record = $policyProperty.Value
                $fetchedAtValue = Get-PIMPolicyObjectValue -InputObject $record -PropertyNames @('FetchedAt')
                if ($fetchedAtValue) {
                    try {
                        if ((ConvertTo-PIMPolicyCacheUtcDateTime -Value $fetchedAtValue) -lt $cutoff) { continue }
                    }
                    catch { continue }
                }

                $script:PolicyCache[$policyProperty.Name] = ConvertFrom-PIMPolicyCacheRecord -Record $record
                $policyCount++
            }
        }

        $script:PIMPolicyCacheLoadedForScope = $scope.ScopeHash
        Write-Verbose "Loaded persistent PIM policy cache: $policyCount policies"
        return [PSCustomObject]@{ Loaded = $true; PolicyCount = $policyCount; AuthenticationContextCount = $contextCount; Reason = 'Loaded' }
    }
    catch {
        Write-Verbose "Failed to load persistent PIM policy cache: $($_.Exception.Message)"
        return [PSCustomObject]@{ Loaded = $false; PolicyCount = 0; AuthenticationContextCount = 0; Reason = 'Error'; Error = $_.Exception.Message }
    }
}

<#
.SYNOPSIS
    Saves sanitized PIM policy metadata for the current tenant scope.

.DESCRIPTION
    Persists the in-memory policy cache to the scoped local policy-cache file. Only policy requirement
    metadata is written. Access tokens, refresh tokens, auth-context tokens, activation request bodies,
    justifications, and ticket values are never included.

.OUTPUTS
    PSCustomObject. Includes save status, counts, and the cache path when successful.
#>
function Save-PIMPolicyCache {
    [CmdletBinding()]
    param()

    try {
        if (-not (Get-Variable -Name 'PolicyCache' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:PolicyCache) { return $null }
        if (-not (Get-Variable -Name 'AuthenticationContextCache' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:AuthenticationContextCache) { $script:AuthenticationContextCache = @{} }

        $scope = Get-PIMPolicyCacheScope
        $cacheFile = Get-PIMPolicyCacheFilePath -Create
        $policies = [ordered]@{}
        $contexts = [ordered]@{}

        foreach ($cacheKey in @($script:PolicyCache.Keys | Sort-Object)) {
            $policyInfo = $script:PolicyCache[$cacheKey]
            if ($null -eq $policyInfo) { continue }

            $record = ConvertTo-PIMPolicyCacheRecord -CacheKey $cacheKey -PolicyInfo $policyInfo
            $policies[$cacheKey] = $record

            try {
                $policyInfo | Add-Member -NotePropertyName PIMCacheFetchedAt -NotePropertyValue $record.FetchedAt -Force
                $policyInfo | Add-Member -NotePropertyName PIMCacheHash -NotePropertyValue $record.Hash -Force
            }
            catch { }
        }

        $cacheData = [ordered]@{
            SchemaVersion          = 1
            ScopeType              = $scope.ScopeType
            ScopeHash              = $scope.ScopeHash
            TenantHash             = $scope.TenantHash
            SavedAt                = [DateTime]::UtcNow.ToString('o')
            StaleAfterHours        = if ((Get-Variable -Name 'PIMPolicyCacheStaleAfterHours' -Scope Script -ErrorAction SilentlyContinue) -and $script:PIMPolicyCacheStaleAfterHours) { [int]$script:PIMPolicyCacheStaleAfterHours } else { 24 }
            Policies               = $policies
            AuthenticationContexts = $contexts
        }

        $temporaryFile = "$cacheFile.tmp"
        $cacheData | ConvertTo-Json -Depth 20 | Set-Content -Path $temporaryFile -Encoding UTF8 -Force
        Move-Item -Path $temporaryFile -Destination $cacheFile -Force

        Write-Verbose "Saved persistent PIM policy cache: $($policies.Count) policies"
        return [PSCustomObject]@{ Saved = $true; PolicyCount = $policies.Count; AuthenticationContextCount = $contexts.Count; Path = $cacheFile }
    }
    catch {
        Write-Verbose "Failed to save persistent PIM policy cache: $($_.Exception.Message)"
        return [PSCustomObject]@{ Saved = $false; Error = $_.Exception.Message }
    }
}

<#
.SYNOPSIS
    Creates a lightweight role snapshot for background policy refresh.

.DESCRIPTION
    Copies only the role identity fields needed to refresh policy metadata. The snapshot avoids passing
    UI controls, Graph model objects, or activation-time inputs into the background job.

.PARAMETER Role
    The role object to snapshot.

.OUTPUTS
    PSCustomObject. A minimal role identity object.
#>
function New-PIMPolicyRoleSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Role
    )

    return [PSCustomObject]@{
        Type             = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('Type')
        DisplayName      = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('DisplayName', 'Name')
        Id               = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('Id')
        RoleDefinitionId = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('RoleDefinitionId')
        GroupId          = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('GroupId')
        ResourceId       = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('ResourceId')
        FullScope        = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('FullScope')
        Scope            = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('Scope')
        SubscriptionId   = Get-PIMPolicyObjectValue -InputObject $Role -PropertyNames @('SubscriptionId')
    }
}

<#
.SYNOPSIS
    Refreshes policy-cache entries for a set of roles.

.DESCRIPTION
    Fetches fresh policies for Entra, PIM group, and Azure Resource roles, updates the in-memory policy
    cache, and persists the sanitized cache back to disk. The helper can force refresh by removing existing
    in-memory entries before fetching.

.PARAMETER Roles
    The role objects or role snapshots whose policy metadata should be refreshed.

.PARAMETER ForceRefresh
    Removes matching in-memory entries before fetching, ensuring the source services are queried again.

.PARAMETER DisableParallelProcessing
    Disables parallel policy fetching for paths that support it.

.PARAMETER ThrottleLimit
    Maximum concurrency used by parallel policy fetch paths.

.OUTPUTS
    PSCustomObject. Counts refreshed policies and reports whether the cache was saved.
#>
function Update-PIMPolicyCacheForRoles {
    [CmdletBinding()]
    param(
        [object[]]$Roles = @(),

        [switch]$ForceRefresh,

        [switch]$DisableParallelProcessing,

        [int]$ThrottleLimit = 10
    )

    if (-not $Roles) {
        return [PSCustomObject]@{ PoliciesUpdated = 0; AuthenticationContextsUpdated = 0 }
    }

    if (-not (Get-Variable -Name 'PolicyCache' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:PolicyCache) { $script:PolicyCache = @{} }
    if (-not (Get-Variable -Name 'AuthenticationContextCache' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:AuthenticationContextCache) { $script:AuthenticationContextCache = @{} }

    $policyResult = @{}
    $entraRoleIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $groupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $azureRoles = [System.Collections.ArrayList]::new()
    $seenAzureKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($role in @($Roles)) {
        if ($null -eq $role) { continue }

        $cacheKey = Get-PIMPolicyCacheKey -Role $role
        if ($ForceRefresh -and $cacheKey -and $script:PolicyCache.ContainsKey($cacheKey)) {
            $script:PolicyCache.Remove($cacheKey)
        }

        $roleType = [string](Get-PIMPolicyObjectValue -InputObject $role -PropertyNames @('Type'))
        if ($roleType -eq 'Azure') { $roleType = 'AzureResource' }

        switch ($roleType) {
            'Entra' {
                $roleId = Get-PIMPolicyObjectValue -InputObject $role -PropertyNames @('RoleDefinitionId', 'Id')
                if ($roleId) { [void]$entraRoleIds.Add([string]$roleId) }
            }
            'Group' {
                $groupId = Get-PIMPolicyObjectValue -InputObject $role -PropertyNames @('GroupId', 'ResourceId', 'Id')
                if ($groupId) { [void]$groupIds.Add([string]$groupId) }
            }
            'AzureResource' {
                if ($cacheKey -and $seenAzureKeys.Add($cacheKey)) {
                    [void]$azureRoles.Add((New-PIMPolicyRoleSnapshot -Role $role))
                }
            }
        }
    }

    if ($entraRoleIds.Count -gt 0) {
        Get-PIMPoliciesBatch -RoleIds @($entraRoleIds) -Type 'Entra' -PolicyCache $policyResult -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
    }

    if ($groupIds.Count -gt 0) {
        Get-PIMPoliciesBatch -GroupIds @($groupIds) -Type 'Group' -PolicyCache $policyResult -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
    }

    foreach ($azureRole in @($azureRoles)) {
        $cacheKey = Get-PIMPolicyCacheKey -Role $azureRole
        if (-not $cacheKey) { continue }

        try {
            $roleDefinitionId = Get-PIMPolicyObjectValue -InputObject $azureRole -PropertyNames @('RoleDefinitionId', 'Id')
            $subscriptionId = Get-PIMPolicyObjectValue -InputObject $azureRole -PropertyNames @('SubscriptionId')
            $scope = Get-PIMPolicyObjectValue -InputObject $azureRole -PropertyNames @('FullScope', 'Scope')
            if (-not $roleDefinitionId) { continue }

            $azurePolicy = Get-AzureResourcePIMPolicy -RoleDefinitionId $roleDefinitionId -SubscriptionId $subscriptionId -Scope $scope
            if ($azurePolicy) {
                $policyResult[$cacheKey] = $azurePolicy
                $script:PolicyCache[$cacheKey] = $azurePolicy
            }
        }
        catch {
            Write-Verbose "Failed to refresh Azure policy cache for ${cacheKey}: $($_.Exception.Message)"
        }
    }

    foreach ($cacheKey in @($policyResult.Keys)) {
        $script:PolicyCache[$cacheKey] = $policyResult[$cacheKey]
    }

    $saveResult = Save-PIMPolicyCache
    return [PSCustomObject]@{
        PoliciesUpdated               = $policyResult.Count
        AuthenticationContextsUpdated = 0
        Saved                         = if ($saveResult) { $saveResult.Saved } else { $false }
    }
}

<#
.SYNOPSIS
    Updates eligible-role UI columns from the policy cache.

.DESCRIPTION
    Re-applies policy requirements from $script:PolicyCache to the eligible roles ListView. This lets the
    UI update after a background refresh completes without rebuilding the entire form or role list.

.PARAMETER Form
    The main PIM Activation Windows Forms form containing the lstEligible control.

.OUTPUTS
    None.
#>
function Update-PIMPolicyColumnsFromCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form
    )

    if (-not $Form -or $Form.IsDisposed) { return }
    if (-not (Get-Variable -Name 'PolicyCache' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:PolicyCache -or $script:PolicyCache.Count -eq 0) { return }

    $eligibleMatches = $Form.Controls.Find('lstEligible', $true)
    if (-not $eligibleMatches -or $eligibleMatches.Count -eq 0) { return }

    $eligibleListView = $eligibleMatches[0]
    if (-not $eligibleListView -or $eligibleListView.IsDisposed) { return }

    $eligibleListView.BeginUpdate()
    try {
        foreach ($item in $eligibleListView.Items) {
            if (-not $item -or -not $item.Tag) { continue }

            $role = $item.Tag
            $cacheKey = Get-PIMPolicyCacheKey -Role $role
            if (-not $cacheKey -or -not $script:PolicyCache.ContainsKey($cacheKey)) { continue }

            $policyInfo = $script:PolicyCache[$cacheKey]
            try { $role | Add-Member -NotePropertyName PolicyInfo -NotePropertyValue $policyInfo -Force } catch { }
            $item.Tag = $role

            $policyUnavailable = $policyInfo -and $policyInfo.PSObject.Properties['PolicyUnavailable'] -and $policyInfo.PolicyUnavailable
            if ($item.SubItems.Count -gt 4) { $item.SubItems[4].Text = if ($policyUnavailable) { 'N/A' } elseif ($policyInfo -and $policyInfo.MaxDuration) { "$($policyInfo.MaxDuration)h" } else { '8h' } }
            if ($item.SubItems.Count -gt 5) { $item.SubItems[5].Text = if ($policyUnavailable) { 'N/A' } elseif ($policyInfo -and $policyInfo.RequiresMfa) { 'Yes' } else { 'No' } }
            if ($item.SubItems.Count -gt 6) {
                $item.SubItems[6].Text = if ($policyUnavailable) { 'N/A' }
                    elseif ($policyInfo -and $policyInfo.RequiresAuthenticationContext) {
                        $contextId = if ($policyInfo.PSObject.Properties['AuthenticationContextId']) { $policyInfo.AuthenticationContextId } else { $null }
                        if ($contextId) { "Required ($contextId)" } else { 'Required' }
                    }
                    else { 'No' }
            }
            if ($item.SubItems.Count -gt 7) { $item.SubItems[7].Text = if ($policyUnavailable) { 'N/A' } elseif ($policyInfo -and $policyInfo.RequiresJustification) { 'Required' } else { 'No' } }
            if ($item.SubItems.Count -gt 8) { $item.SubItems[8].Text = if ($policyUnavailable) { 'N/A' } elseif ($policyInfo -and $policyInfo.RequiresTicket) { 'Yes' } else { 'No' } }
            if ($item.SubItems.Count -gt 9) { $item.SubItems[9].Text = if ($policyUnavailable) { 'N/A' } elseif ($policyInfo -and $policyInfo.RequiresApproval) { 'Required' } else { 'No' } }

            if ($policyInfo) {
                if ($policyInfo.RequiresApproval) {
                    $item.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                }
                elseif ($policyInfo.RequiresAuthenticationContext) {
                    $item.ForeColor = [System.Drawing.Color]::FromArgb(0, 78, 146)
                }
                else {
                    $item.ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
                }
            }
        }
    }
    finally {
        $eligibleListView.EndUpdate()
    }
}

<#
.SYNOPSIS
    Starts a background refresh for stale disk-loaded policy-cache entries.

.DESCRIPTION
    Identifies cached role policies that are present but stale, starts a thread job to refresh those policy
    entries, saves the updated sanitized cache, and uses a Windows Forms timer to merge refreshed metadata
    back into the current UI when the job finishes. Fresh entries are skipped so startup remains fast.

.PARAMETER Form
    The main PIM Activation Windows Forms form to update after the background refresh completes.

.PARAMETER DisableParallelProcessing
    Disables parallel policy fetching in the background refresh job.

.PARAMETER ThrottleLimit
    Maximum concurrency used by supported background policy fetch paths.

.OUTPUTS
    None.

.NOTES
    The background job imports the module in a separate thread and relies on the current Graph/Azure session
    being available to the process. If that context is unavailable, the refresh fails quietly and the already
    loaded cache remains a startup performance hint only.
#>
function Start-PIMPolicyCacheBackgroundRefresh {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form,

        [switch]$DisableParallelProcessing,

        [int]$ThrottleLimit = 10
    )

    try {
        if (-not $Form -or $Form.IsDisposed) { return }
        if (-not (Get-Command Start-ThreadJob -ErrorAction SilentlyContinue)) {
            Write-Verbose 'Start-ThreadJob is unavailable; background policy cache refresh skipped.'
            return
        }

        if ((Get-Variable -Name 'PIMPolicyBackgroundRefreshJob' -Scope Script -ErrorAction SilentlyContinue) -and $script:PIMPolicyBackgroundRefreshJob) {
            if ($script:PIMPolicyBackgroundRefreshJob.State -in @('NotStarted', 'Running')) { return }
        }

        $roleCandidates = @()
        if ((Get-Variable -Name 'CachedEligibleRoles' -Scope Script -ErrorAction SilentlyContinue) -and $script:CachedEligibleRoles) { $roleCandidates += @($script:CachedEligibleRoles) }
        if ((Get-Variable -Name 'CachedActiveRoles' -Scope Script -ErrorAction SilentlyContinue) -and $script:CachedActiveRoles) { $roleCandidates += @($script:CachedActiveRoles) }
        if ($roleCandidates.Count -eq 0) { return }

        $refreshSnapshots = [System.Collections.ArrayList]::new()
        $seenKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($role in $roleCandidates) {
            if ($null -eq $role) { continue }

            $cacheKey = Get-PIMPolicyCacheKey -Role $role
            if (-not $cacheKey -or -not $seenKeys.Add($cacheKey)) { continue }
            if (-not $script:PolicyCache.ContainsKey($cacheKey)) { continue }
            if (Test-PIMPolicyCacheEntryFresh -PolicyInfo $script:PolicyCache[$cacheKey]) { continue }

            [void]$refreshSnapshots.Add((New-PIMPolicyRoleSnapshot -Role $role))
        }

        if ($refreshSnapshots.Count -eq 0) { return }

        $modulePath = Join-Path $script:ModuleRoot 'PIMActivation.psd1'
        $scope = Get-PIMPolicyCacheScope
        $currentUserId = if ((Get-Variable -Name 'CurrentUser' -Scope Script -ErrorAction SilentlyContinue) -and $script:CurrentUser -and $script:CurrentUser.PSObject.Properties['Id']) { [string]$script:CurrentUser.Id } else { $null }
        $currentUserPrincipalName = if ((Get-Variable -Name 'CurrentUser' -Scope Script -ErrorAction SilentlyContinue) -and $script:CurrentUser -and $script:CurrentUser.PSObject.Properties['UserPrincipalName']) { [string]$script:CurrentUser.UserPrincipalName }
            elseif ((Get-Variable -Name 'CurrentGraphUser' -Scope Script -ErrorAction SilentlyContinue) -and -not [string]::IsNullOrWhiteSpace([string]$script:CurrentGraphUser)) { [string]$script:CurrentGraphUser }
            elseif ((Get-Variable -Name 'GraphContext' -Scope Script -ErrorAction SilentlyContinue) -and $script:GraphContext -and $script:GraphContext.PSObject.Properties['Account']) { [string]$script:GraphContext.Account }
            else { $null }
        $disableParallel = [bool]$DisableParallelProcessing

        Write-Verbose "Starting background PIM policy cache refresh for $($refreshSnapshots.Count) stale policies"
        $job = Start-ThreadJob -Name 'PIMPolicyCacheRefresh' -ArgumentList @($modulePath, @($refreshSnapshots), $scope.TenantId, $currentUserId, $currentUserPrincipalName, $disableParallel, $ThrottleLimit) -ScriptBlock {
            param(
                [string]$ModulePath,
                [object[]]$RoleSnapshots,
                [string]$TenantId,
                [string]$CurrentUserId,
                [string]$CurrentUserPrincipalName,
                [bool]$DisableParallel,
                [int]$Throttle
            )

            $VerbosePreference = 'SilentlyContinue'
            try {
                Import-Module $ModulePath -Force -ErrorAction Stop
                & (Get-Module PIMActivation) {
                    param(
                        [object[]]$InnerRoleSnapshots,
                        [string]$InnerTenantId,
                        [string]$InnerCurrentUserId,
                        [string]$InnerCurrentUserPrincipalName,
                        [bool]$InnerDisableParallel,
                        [int]$InnerThrottle
                    )

                    $script:CurrentTenantId = $InnerTenantId
                    $script:CurrentUser = [PSCustomObject]@{
                        Id                = $InnerCurrentUserId
                        UserPrincipalName = $InnerCurrentUserPrincipalName
                    }
                    $script:GraphContext = [PSCustomObject]@{
                        TenantId = $InnerTenantId
                        Account  = $InnerCurrentUserPrincipalName
                    }

                    Import-PIMPolicyCache -Force | Out-Null
                    $refreshResult = Update-PIMPolicyCacheForRoles -Roles $InnerRoleSnapshots -ForceRefresh -DisableParallelProcessing:$InnerDisableParallel -ThrottleLimit $InnerThrottle
                    return [PSCustomObject]@{
                        Success                       = $true
                        PoliciesUpdated               = $refreshResult.PoliciesUpdated
                        AuthenticationContextsUpdated = $refreshResult.AuthenticationContextsUpdated
                    }
                } @($RoleSnapshots) $TenantId $CurrentUserId $CurrentUserPrincipalName $DisableParallel $Throttle
            }
            catch {
                return [PSCustomObject]@{
                    Success = $false
                    Error   = $_.Exception.Message
                }
            }
        }

        $script:PIMPolicyBackgroundRefreshJob = $job
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 1500
        $timer.Add_Tick({
            try {
                if (-not (Get-Variable -Name 'PIMPolicyBackgroundRefreshJob' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:PIMPolicyBackgroundRefreshJob) {
                    $this.Stop()
                    $this.Dispose()
                    return
                }

                $activeJob = $script:PIMPolicyBackgroundRefreshJob
                if ($activeJob.State -notin @('Completed', 'Failed', 'Stopped')) { return }

                $this.Stop()
                $jobOutput = Receive-Job -Job $activeJob -ErrorAction SilentlyContinue
                Remove-Job -Job $activeJob -Force -ErrorAction SilentlyContinue
                $script:PIMPolicyBackgroundRefreshJob = $null

                $successfulOutput = @($jobOutput | Where-Object { $_ -and $_.PSObject.Properties['Success'] -and $_.Success } | Select-Object -First 1)
                if ($successfulOutput.Count -gt 0) {
                    Import-PIMPolicyCache -Force | Out-Null
                    Update-PIMPolicyColumnsFromCache -Form $Form
                    Write-Verbose "Background PIM policy cache refresh completed. Policies updated: $($successfulOutput[0].PoliciesUpdated)"
                }
                else {
                    $errorOutput = @($jobOutput | Where-Object { $_ -and $_.PSObject.Properties['Error'] } | Select-Object -First 1)
                    if ($errorOutput.Count -gt 0) {
                        Write-Verbose "Background PIM policy cache refresh skipped or failed: $($errorOutput[0].Error)"
                    }
                }

                $this.Dispose()
            }
            catch {
                Write-Verbose "Failed while completing background policy cache refresh: $($_.Exception.Message)"
                try { $this.Stop(); $this.Dispose() } catch { }
            }
        })

        $script:PIMPolicyBackgroundRefreshTimer = $timer
        $Form.Add_FormClosed({
            try {
                if ((Get-Variable -Name 'PIMPolicyBackgroundRefreshTimer' -Scope Script -ErrorAction SilentlyContinue) -and $script:PIMPolicyBackgroundRefreshTimer) {
                    $script:PIMPolicyBackgroundRefreshTimer.Stop()
                    $script:PIMPolicyBackgroundRefreshTimer.Dispose()
                    $script:PIMPolicyBackgroundRefreshTimer = $null
                }
                if ((Get-Variable -Name 'PIMPolicyBackgroundRefreshJob' -Scope Script -ErrorAction SilentlyContinue) -and $script:PIMPolicyBackgroundRefreshJob -and $script:PIMPolicyBackgroundRefreshJob.State -in @('NotStarted', 'Running')) {
                    Stop-Job -Job $script:PIMPolicyBackgroundRefreshJob -ErrorAction SilentlyContinue
                    Remove-Job -Job $script:PIMPolicyBackgroundRefreshJob -Force -ErrorAction SilentlyContinue
                    $script:PIMPolicyBackgroundRefreshJob = $null
                }
            }
            catch { }
        })
        $timer.Start()
    }
    catch {
        Write-Verbose "Unable to start background PIM policy cache refresh: $($_.Exception.Message)"
    }
}