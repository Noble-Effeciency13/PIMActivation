function Get-AzureResourceRoles {
    <#
    .SYNOPSIS
        Retrieves Azure Resource PIM eligible and active roles for the calling user
        using the ARM root-scope endpoints with the asTarget() filter.

    .DESCRIPTION
        Calls the ARM PIM schedule-instance endpoints at the root tenant scope
        ("/providers/Microsoft.Authorization/...") with the $filter=asTarget()
        clause. asTarget() returns every PIM eligibility (or activation) the
        calling user holds tenant-wide - across management groups, subscriptions,
        resource groups, and individual resources - without requiring any
        pre-existing Azure role at the queried scope (avoids the historic
        Azure Resource PIM catch-22 where principalId filtering required
        Microsoft.Authorization/roleAssignments/read at the scope).

        Two endpoints are used:
            * roleEligibilityScheduleInstances - eligible (PIM) assignments
            * roleAssignmentScheduleInstances  - currently active assignments
              (both PIM-activated and permanent role assignments that ARM has
              materialized as schedule instances)

        Display names for role definitions, subscriptions, and management
        groups are resolved on demand through ARM REST so that Az.Resources
        is not required.

    .PARAMETER UserId
        UPN of the user. Used to (re)establish the Az context if necessary.

    .PARAMETER UserObjectId
        Azure AD object ID of the user. Used to classify Direct vs Inherited
        assignments returned by asTarget().

    .PARAMETER IncludeActive
        Include active (currently-assigned) Azure Resource roles.

    .PARAMETER IncludeEligible
        Include eligible (PIM) Azure Resource roles.

    .PARAMETER SubscriptionIds
        Optional client-side filter. When provided, results are limited to
        scopes under any of the supplied subscription IDs (or matching
        management-group / tenant scopes are kept as-is).

    .PARAMETER OnlyDirtyManagementGroups
        Reserved for delta-refresh callers. Currently a no-op for the
        root-scope endpoint because asTarget() already returns everything
        the user can see in one call.

    .PARAMETER DisableParallelProcessing
        Reserved for backward compatibility with prior implementations.
        Has no effect on the root-scope endpoint (single ARM call per
        endpoint, paginated).

    .PARAMETER ThrottleLimit
        Reserved for backward compatibility. Unused.

    .EXAMPLE
        Get-AzureResourceRoles -UserId 'user@contoso.com' -UserObjectId '...' -IncludeEligible -IncludeActive

    .NOTES
        Requires Az.Accounts (for Get-AzContext / Connect-AzAccount / token
        acquisition). Az.Resources is not required.
        ARM API versions:
            * Schedule instances : 2020-10-01
            * Role definitions   : 2022-04-01
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$UserObjectId,

        [switch]$IncludeActive,
        [switch]$IncludeEligible,

        [string[]]$SubscriptionIds,
        [switch]$OnlyDirtyManagementGroups,
        [switch]$DisableParallelProcessing,
        [int]$ThrottleLimit = 10
    )

    Write-Verbose "Get-AzureResourceRoles: using ARM root-scope query with asTarget() for principal $UserObjectId"

    if (-not $IncludeActive -and -not $IncludeEligible) {
        Write-Verbose "Neither IncludeActive nor IncludeEligible specified - nothing to fetch"
        return @()
    }

    # OnlyDirtyManagementGroups is a no-op for the root-scope endpoint - asTarget()
    # already returns everything the principal can see across MGs in one call.
    if ($OnlyDirtyManagementGroups) {
        Write-Verbose "OnlyDirtyManagementGroups is a no-op for the ARM root-scope endpoint; returning empty"
        return @()
    }

    # ----- 1. Ensure Azure context --------------------------------------------------
    try {
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $azContext) {
            Write-Verbose "No Az context - connecting as $UserId"
            Connect-AzAccount -AccountId $UserId -ErrorAction Stop | Out-Null
            $azContext = Get-AzContext -ErrorAction Stop
        }
        elseif ($azContext.Account.Id -ne $UserId -and $azContext.Account.Id -ne $UserObjectId) {
            Write-Verbose "Az context mismatch - reconnecting as $UserId"
            Connect-AzAccount -AccountId $UserId -ErrorAction Stop | Out-Null
            $azContext = Get-AzContext -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Failed to establish Azure context: $($_.Exception.Message)"
        return @()
    }

    # ----- 2. Acquire ARM bearer token ---------------------------------------------
    $armToken = $null
    $headers  = $null
    try {
        try {
            $tokenObj = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop
            if ($tokenObj.Token -is [System.Security.SecureString]) {
                $armToken = [System.Net.NetworkCredential]::new('', $tokenObj.Token).Password
            }
            elseif ($tokenObj.Token -is [string] -and $tokenObj.Token.Length -gt 0) {
                $armToken = $tokenObj.Token
            }
            else {
                $secureToken = (Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -AsSecureString -ErrorAction Stop).Token
                $armToken    = [System.Net.NetworkCredential]::new('', $secureToken).Password
            }
            if ([string]::IsNullOrEmpty($armToken)) { throw 'Token was null or empty after extraction.' }
        }
        catch {
            Write-Warning "Failed to acquire ARM access token: $($_.Exception.Message)"
            return @()
        }

        $headers = @{
            'Authorization' = "Bearer $armToken"
            'Content-Type'  = 'application/json'
        }

        # ----- 3. Display-name caches ----------------------------------------------
        $roleDefCache = @{}
        $subNameCache = @{}
        $mgNameCache  = @{}

        # ----- 4. Helpers ----------------------------------------------------------

        # Paginate a single ARM list endpoint and return all items.
        $invokeArmList = {
            param([string]$Uri)
            $items   = [System.Collections.ArrayList]::new()
            $nextUri = $Uri
            while ($nextUri) {
                try {
                    $resp = Invoke-RestMethod -Uri $nextUri -Headers $headers -Method Get -ErrorAction Stop
                    if ($resp.value) {
                        foreach ($item in $resp.value) { [void]$items.Add($item) }
                    }
                    $nextUri = if ($resp.PSObject.Properties.Name -contains 'nextLink') { $resp.nextLink } else { $null }
                }
                catch {
                    Write-Verbose "ARM list failed ($nextUri): $($_.Exception.Message)"
                    break
                }
            }
            return ,$items
        }

        # Resolve a role definition (full ARM path or bare GUID) to its friendly name.
        $resolveRoleDefName = {
            param([string]$RoleDefId)
            if ([string]::IsNullOrEmpty($RoleDefId)) { return 'Unknown' }
            $guid = $RoleDefId
            if ($guid -match '/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})') { $guid = $matches[1] }
            if ($roleDefCache.ContainsKey($guid)) { return $roleDefCache[$guid] }
            try {
                $rdPath = if ($RoleDefId.StartsWith('/')) { $RoleDefId } else { "/providers/Microsoft.Authorization/roleDefinitions/$guid" }
                $rdUri  = "https://management.azure.com$rdPath`?api-version=2022-04-01"
                $rdResp = Invoke-RestMethod -Uri $rdUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                $name   = if ($rdResp -and $rdResp.properties -and $rdResp.properties.roleName) { $rdResp.properties.roleName } else { "Unknown ($guid)" }
            }
            catch {
                $name = "Unknown ($guid)"
            }
            $roleDefCache[$guid] = $name
            return $name
        }

        # Resolve a subscription display name via ARM REST (does not require Az.Resources).
        $resolveSubName = {
            param([string]$SubId)
            if ([string]::IsNullOrEmpty($SubId)) { return $null }
            if ($subNameCache.ContainsKey($SubId)) { return $subNameCache[$SubId] }
            $name = $SubId
            try {
                $sUri  = "https://management.azure.com/subscriptions/$SubId`?api-version=2022-12-01"
                $sResp = Invoke-RestMethod -Uri $sUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                if ($sResp -and $sResp.displayName) { $name = $sResp.displayName }
            }
            catch { }
            $subNameCache[$SubId] = $name
            return $name
        }

        # Resolve a management-group display name via ARM REST.
        $resolveMgName = {
            param([string]$MgId)
            if ([string]::IsNullOrEmpty($MgId)) { return $null }
            if ($mgNameCache.ContainsKey($MgId)) { return $mgNameCache[$MgId] }
            $name = $MgId
            try {
                $mUri  = "https://management.azure.com/providers/Microsoft.Management/managementGroups/$MgId`?api-version=2020-05-01"
                $mResp = Invoke-RestMethod -Uri $mUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                if ($mResp -and $mResp.properties -and $mResp.properties.displayName) { $name = $mResp.properties.displayName }
            }
            catch { }
            $mgNameCache[$MgId] = $name
            return $name
        }

        # Resolve an ARM scope string to a structured info hashtable.
        $getScopeInfo = {
            param([string]$Scope)

            if ([string]::IsNullOrEmpty($Scope) -or $Scope -eq '/') {
                return @{
                    ScopeType        = 'Tenant'
                    ResourceDisplay  = '/'
                    SubscriptionId   = $null
                    SubscriptionName = $null
                    ScopeDisplayName = 'Tenant Root'
                }
            }
            if ($Scope -match '^/providers/Microsoft\.Management/managementGroups/([^/]+)$') {
                $mgId = $matches[1]
                $dn   = & $resolveMgName $mgId
                return @{
                    ScopeType        = 'Management Group'
                    ResourceDisplay  = $dn
                    SubscriptionId   = $null
                    SubscriptionName = $null
                    ScopeDisplayName = "MG: $dn"
                }
            }
            if ($Scope -match '^/subscriptions/([a-fA-F0-9\-]{36})/resourceGroups/([^/]+)/providers/.+$') {
                $subId   = $matches[1]
                $resName = ($Scope -split '/')[-1]
                $subName = & $resolveSubName $subId
                return @{
                    ScopeType        = 'Resource'
                    ResourceDisplay  = $resName
                    SubscriptionId   = $subId
                    SubscriptionName = $subName
                    ScopeDisplayName = "Resource: $resName"
                }
            }
            if ($Scope -match '^/subscriptions/([a-fA-F0-9\-]{36})/resourceGroups/([^/]+)$') {
                $subId   = $matches[1]
                $rgName  = $matches[2]
                $subName = & $resolveSubName $subId
                return @{
                    ScopeType        = 'Resource Group'
                    ResourceDisplay  = $rgName
                    SubscriptionId   = $subId
                    SubscriptionName = $subName
                    ScopeDisplayName = "RG: $rgName"
                }
            }
            if ($Scope -match '^/subscriptions/([a-fA-F0-9\-]{36})$') {
                $subId   = $matches[1]
                $subName = & $resolveSubName $subId
                return @{
                    ScopeType        = 'Subscription'
                    ResourceDisplay  = $subName
                    SubscriptionId   = $subId
                    SubscriptionName = $subName
                    ScopeDisplayName = "Sub: $subName"
                }
            }
            return @{
                ScopeType        = 'Unknown'
                ResourceDisplay  = $Scope
                SubscriptionId   = $null
                SubscriptionName = $null
                ScopeDisplayName = $Scope
            }
        }

        # ----- 5. Query root-scope endpoints with asTarget() -----------------------
        $apiVer  = '2020-10-01'
        $baseUri = 'https://management.azure.com/providers/Microsoft.Authorization'

        $eligibleRaw = @()
        $activeRaw   = @()

        if ($IncludeEligible) {
            $uri = "$baseUri/roleEligibilityScheduleInstances?api-version=$apiVer&`$filter=asTarget()"
            Write-Verbose "Querying eligible role schedule instances: $uri"
            $eligibleRaw = & $invokeArmList $uri
            Write-Verbose "asTarget() returned $($eligibleRaw.Count) eligible Azure Resource role instance(s)"
        }

        if ($IncludeActive) {
            $uri = "$baseUri/roleAssignmentScheduleInstances?api-version=$apiVer&`$filter=asTarget()"
            Write-Verbose "Querying active role schedule instances: $uri"
            $activeRaw = & $invokeArmList $uri
            Write-Verbose "asTarget() returned $($activeRaw.Count) active Azure Resource role instance(s)"
        }

        # ----- 6. Project raw items into role objects ------------------------------
        $allRoles = [System.Collections.ArrayList]::new()

        # Safe property accessor - StrictMode-friendly. Returns $null if the
        # property is missing instead of throwing.
        $getProp = {
            param($Object, [string]$Name)
            if ($null -eq $Object) { return $null }
            $p = $Object.PSObject.Properties[$Name]
            if ($p) { return $p.Value }
            return $null
        }

        $projectInstance = {
            param($Item, [string]$Status)

            $props = & $getProp $Item 'properties'
            if (-not $props) { return $null }

            $scope        = & $getProp $props 'scope'
            $roleDefId    = & $getProp $props 'roleDefinitionId'
            $principalId  = & $getProp $props 'principalId'
            $principalTyp = & $getProp $props 'principalType'
            $startDt      = & $getProp $props 'startDateTime'
            $endDt        = & $getProp $props 'endDateTime'
            $assignType   = & $getProp $props 'assignmentType'   # 'Activated' or 'Assigned' (active endpoint)
            $memberTypeAp = & $getProp $props 'memberType'       # 'Direct' / 'Inherited' / 'Group'

            $scopeInfo = & $getScopeInfo $scope
            $roleName  = & $resolveRoleDefName $roleDefId

            # Direct vs Inherited - prefer ARM-supplied memberType when present,
            # otherwise derive from principal identity vs caller.
            $memberType = if ($memberTypeAp) {
                $memberTypeAp
            }
            elseif ($principalTyp -eq 'Group') {
                'Group'
            }
            elseif ($principalId -and $principalId -ne $UserObjectId) {
                'Inherited'
            }
            else {
                'Direct'
            }

            $formatted = if ($scopeInfo.ScopeDisplayName) { $scopeInfo.ScopeDisplayName } else { $scope }

            [PSCustomObject]@{
                RoleId               = $Item.name
                RoleDefinitionId     = $roleDefId
                DisplayName          = $roleName
                ResourceName         = $scopeInfo.ResourceDisplay
                ResourceDisplayName  = $scopeInfo.ResourceDisplay
                ScopeDisplayName     = $scopeInfo.ScopeDisplayName
                Type                 = 'AzureResource'
                Status               = $Status
                MemberType           = $memberType
                SubscriptionId       = $scopeInfo.SubscriptionId
                SubscriptionName     = $scopeInfo.SubscriptionName
                FullScope            = $scope
                ObjectId             = $principalId
                ObjectType           = $principalTyp
                StartDateTime        = $startDt
                EndDateTime          = $endDt
                Scope                = $scopeInfo.ScopeType
                FormattedScope       = $formatted
                AssignmentType       = $assignType
                # Used by reduced-scope Azure activation. For eligible instances ARM
                # returns the parent eligibility schedule ID in roleEligibilityScheduleId.
                LinkedRoleEligibilityScheduleId = (& $getProp $props 'roleEligibilityScheduleId')
            }
        }

        foreach ($item in $eligibleRaw) {
            $obj = & $projectInstance $item 'Eligible'
            if ($obj) { [void]$allRoles.Add($obj) }
        }
        foreach ($item in $activeRaw) {
            $obj = & $projectInstance $item 'Active'
            if ($obj) { [void]$allRoles.Add($obj) }
        }

        Write-Verbose "Projected $($allRoles.Count) Azure Resource role object(s) from ARM root-scope query"

        # ----- 7. Optional client-side subscription filter -------------------------
        if ($SubscriptionIds -and $SubscriptionIds.Count -gt 0) {
            $subFilter = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($s in $SubscriptionIds) { if ($s) { [void]$subFilter.Add($s) } }
            $filtered = [System.Collections.ArrayList]::new()
            foreach ($r in $allRoles) {
                # Keep roles with no subscription (tenant / MG scopes) plus those in the filter set.
                if (-not $r.SubscriptionId -or $subFilter.Contains($r.SubscriptionId)) {
                    [void]$filtered.Add($r)
                }
            }
            Write-Verbose "Filtered by SubscriptionIds: kept $($filtered.Count) of $($allRoles.Count) role(s)"
            $allRoles = $filtered
        }

        # ----- 8. Deduplicate ------------------------------------------------------
        $uniqueRoles = [System.Collections.ArrayList]::new()
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($r in $allRoles) {
            $defGuid = $r.RoleDefinitionId
            if ($defGuid -and $defGuid -match '/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})') { $defGuid = $matches[1] }
            $key = "{0}|{1}|{2}|{3}" -f $defGuid, $r.FullScope, $r.Status, $r.ObjectId
            if ($seen.Add($key)) { [void]$uniqueRoles.Add($r) }
        }

        $activeCount   = @($uniqueRoles | Where-Object { $_.Status -eq 'Active' }).Count
        $eligibleCount = @($uniqueRoles | Where-Object { $_.Status -eq 'Eligible' }).Count
        Write-Verbose "Azure Resource roles after dedupe: $($uniqueRoles.Count) total ($eligibleCount eligible, $activeCount active)"

        # Return as concrete array
        return ,@($uniqueRoles)
    }
    catch {
        Write-Warning "Failed to retrieve Azure Resource roles: $($_.Exception.Message)"
        return @()
    }
    finally {
        # Best-effort scrub of plaintext bearer token from memory.
        if (Get-Variable -Name 'armToken' -Scope 0 -ErrorAction SilentlyContinue) { $armToken = $null }
        if (Get-Variable -Name 'headers'  -Scope 0 -ErrorAction SilentlyContinue) { $headers  = $null }
    }
}
