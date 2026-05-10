function Get-AzureResourceRoles {
    <#
    .SYNOPSIS
        Retrieves Azure resource PIM-eligible and active roles for a user using the ARM root-scope endpoint.

    .DESCRIPTION
        Uses the ARM Management API at root scope ("/") with a principalId filter to enumerate all
        role eligibility schedules and role assignment schedules for the signed-in user.
        This approach does NOT require any standing Azure RBAC permissions – it relies only on the
        user being able to read their own assignments, which the ARM API permits unconditionally.

        Role definition display names and scope display names (subscription/MG names) are resolved
        lazily and cached for the duration of the call.

    .PARAMETER UserId
        The user principal name (UPN) used to establish the Az context if one is not already present.

    .PARAMETER UserObjectId
        The Azure AD object ID of the user.  Used as the principalId filter in ARM REST calls.

    .PARAMETER IncludeActive
        When specified, queries roleAssignmentSchedules at root scope.

    .PARAMETER IncludeEligible
        When specified, queries roleEligibilitySchedules at root scope.

    .PARAMETER SubscriptionIds
        Optional.  When provided, only roles whose scope matches one of these subscription IDs are
        returned.  Used by the delta-refresh path to limit results to recently activated scopes.

    .PARAMETER OnlyDirtyManagementGroups
        When set, only returns roles scoped to management groups listed in
        $script:DirtyManagementGroups.  Used by the delta-refresh path for MG activation events.

    .PARAMETER DisableParallelProcessing
        Accepted for backward-compatibility; not used in this implementation (no subscription loop).

    .PARAMETER ThrottleLimit
        Accepted for backward-compatibility; not used in this implementation.

    .EXAMPLE
        Get-AzureResourceRoles -UserId "user@domain.com" -UserObjectId "12345678-..." -IncludeEligible -IncludeActive

    .NOTES
        Requires Az.Accounts (for Get-AzContext / Connect-AzAccount / token acquisition).
        Does NOT require Az.Resources.
        ARM API version: 2020-10-01 for PIM schedule endpoints.
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

    Write-Verbose "Get-AzureResourceRoles: using ARM root-scope query for principal $UserObjectId"

    # ── 1. Ensure Azure context ──────────────────────────────────────────────
    try {
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $azContext) {
            Write-Verbose "No Az context – connecting as $UserId"
            Connect-AzAccount -AccountId $UserId -ErrorAction Stop | Out-Null
            $azContext = Get-AzContext -ErrorAction Stop
        }
        elseif ($azContext.Account.Id -ne $UserId -and $azContext.Account.Id -ne $UserObjectId) {
            Write-Verbose "Az context mismatch – reconnecting as $UserId"
            Connect-AzAccount -AccountId $UserId -ErrorAction Stop | Out-Null
            $azContext = Get-AzContext -ErrorAction Stop
        }
    }
    catch {
        throw "Failed to establish Azure context: $($_.Exception.Message)"
    }

    # ── 2. Acquire ARM bearer token ──────────────────────────────────────────
    try {
        $tokenObj = Get-AzAccessToken -ResourceUrl "https://management.azure.com/" -ErrorAction Stop

        # Az.Accounts 2.13+ returns a SecureString; older versions return plain string
        if ($tokenObj.Token -is [System.Security.SecureString]) {
            $armToken = [System.Net.NetworkCredential]::new('', $tokenObj.Token).Password
        }
        elseif ($tokenObj.Token -is [string] -and $tokenObj.Token.Length -gt 0) {
            $armToken = $tokenObj.Token
        }
        else {
            # Fallback: Az.Accounts 2.17+ supports -AsSecureString explicitly
            $secureToken = (Get-AzAccessToken -ResourceUrl "https://management.azure.com/" -AsSecureString -ErrorAction Stop).Token
            $armToken = [System.Net.NetworkCredential]::new('', $secureToken).Password
        }

        if ([string]::IsNullOrEmpty($armToken)) {
            throw "Token was null or empty after extraction."
        }
    }
    catch {
        throw "Failed to acquire ARM access token: $($_.Exception.Message)"
    }

    $headers = @{
        'Authorization' = "Bearer $armToken"
        'Content-Type'  = 'application/json'
    }

    # ── 3. Local caches for display-name resolution ──────────────────────────
    $roleDefCache = @{}   # guid -> friendly role name
    $subNameCache = @{}   # subId -> subscription display name
    $mgNameCache  = @{}   # mgId  -> management-group display name

    $allRoles = [System.Collections.ArrayList]::new()

    # ── 4. Helpers ────────────────────────────────────────────────────────────

    # Paginate through a single ARM list endpoint and return all items.
    $InvokeArmList = {
        param([string]$Uri)
        $items   = [System.Collections.ArrayList]::new()
        $nextUri = $Uri
        while ($nextUri) {
            try {
                $resp    = Invoke-RestMethod -Uri $nextUri -Headers $headers -Method Get -ErrorAction Stop
                if ($resp.value) {
                    foreach ($item in $resp.value) { $items.Add($item) | Out-Null }
                }
                $nextUri = if ($resp.PSObject.Properties.Name -contains 'nextLink') {
                    $resp.nextLink
                } else { $null }
            }
            catch {
                Write-Verbose "ARM list failed ($nextUri): $($_.Exception.Message)"
                break
            }
        }
        return ,$items
    }

    # Resolve a roleDefinitionId (full ARM path or bare GUID) to its display name.
    # Uses the ARM REST API directly (via the bearer token already acquired) so that
    # Az.Resources is NOT required and no standing subscription permissions are needed.
    $ResolveRoleDefName = {
        param([string]$RoleDefId)
        $guid = $RoleDefId
        if ($guid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
            $guid = $matches[1]
        }
        if ($roleDefCache.ContainsKey($guid)) { return $roleDefCache[$guid] }
        try {
            # Prefer the original full path; fall back to the global built-in role path
            $rdPath = if ($RoleDefId -match "^/") { $RoleDefId } else { "/providers/Microsoft.Authorization/roleDefinitions/$guid" }
            $rdUri  = "https://management.azure.com$rdPath`?api-version=2022-04-01"
            $rdResp = Invoke-RestMethod -Uri $rdUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
            $name   = if ($rdResp -and $rdResp.properties -and $rdResp.properties.roleName) {
                          $rdResp.properties.roleName
                      } else { "Unknown ($guid)" }
        }
        catch { $name = "Unknown ($guid)" }
        $roleDefCache[$guid] = $name
        return $name
    }

    # Resolve an ARM scope string to a structured info hashtable.
    $GetScopeInfo = {
        param([string]$Scope)
        if ([string]::IsNullOrEmpty($Scope) -or $Scope -eq '/') {
            return @{
                ScopeType       = 'Tenant'
                ResourceDisplay = '/'
                SubscriptionId  = $null
                ScopeDisplayName = 'Tenant Root'
            }
        }
        if ($Scope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
            $mgId = $matches[1]
            if (-not $mgNameCache.ContainsKey($mgId)) {
                try {
                    $mg = Get-AzManagementGroup -GroupId $mgId -ErrorAction SilentlyContinue
                    $mgNameCache[$mgId] = if ($mg -and $mg.DisplayName) { $mg.DisplayName } else { $mgId }
                }
                catch { $mgNameCache[$mgId] = $mgId }
            }
            $dn = $mgNameCache[$mgId]
            return @{
                ScopeType        = 'Management Group'
                ResourceDisplay  = $dn
                SubscriptionId   = $null
                ScopeDisplayName = "MG: $dn"
            }
        }
        if ($Scope -match "^/subscriptions/([a-fA-F0-9\-]{36})/resourceGroups/([^/]+)/(.+)$") {
            $subId = $matches[1]; $rgName = $matches[2]; $resName = ($Scope -split '/')[-1]
            if (-not $subNameCache.ContainsKey($subId)) {
                try {
                    $sub = Get-AzSubscription -SubscriptionId $subId -ErrorAction SilentlyContinue
                    $subNameCache[$subId] = if ($sub -and $sub.Name) { $sub.Name } else { $subId }
                }
                catch { $subNameCache[$subId] = $subId }
            }
            return @{
                ScopeType        = 'Resource'
                ResourceDisplay  = $resName
                SubscriptionId   = $subId
                ScopeDisplayName = "Resource: $resName"
            }
        }
        if ($Scope -match "^/subscriptions/([a-fA-F0-9\-]{36})/resourceGroups/([^/]+)$") {
            $subId = $matches[1]; $rgName = $matches[2]
            if (-not $subNameCache.ContainsKey($subId)) {
                try {
                    $sub = Get-AzSubscription -SubscriptionId $subId -ErrorAction SilentlyContinue
                    $subNameCache[$subId] = if ($sub -and $sub.Name) { $sub.Name } else { $subId }
                }
                catch { $subNameCache[$subId] = $subId }
            }
            return @{
                ScopeType        = 'Resource Group'
                ResourceDisplay  = $rgName
                SubscriptionId   = $subId
                ScopeDisplayName = "RG: $rgName"
            }
        }
        if ($Scope -match "^/subscriptions/([a-fA-F0-9\-]{36})$") {
            $subId = $matches[1]
            if (-not $subNameCache.ContainsKey($subId)) {
                try {
                    $sub = Get-AzSubscription -SubscriptionId $subId -ErrorAction SilentlyContinue
                    $subNameCache[$subId] = if ($sub -and $sub.Name) { $sub.Name } else { $subId }
                }
                catch { $subNameCache[$subId] = $subId }
            }
            return @{
                ScopeType        = 'Subscription'
                ResourceDisplay  = $subNameCache[$subId]
                SubscriptionId   = $subId
                ScopeDisplayName = 'Subscription'
            }
        }
        # Fallback for any deeper resource scope that starts with /subscriptions/
        if ($Scope -match "^/subscriptions/([a-fA-F0-9\-]{36})") {
            $subId = $matches[1]
            if (-not $subNameCache.ContainsKey($subId)) {
                try {
                    $sub = Get-AzSubscription -SubscriptionId $subId -ErrorAction SilentlyContinue
                    $subNameCache[$subId] = if ($sub -and $sub.Name) { $sub.Name } else { $subId }
                }
                catch { $subNameCache[$subId] = $subId }
            }
            $resName = ($Scope -split '/')[-1]
            return @{
                ScopeType        = 'Resource'
                ResourceDisplay  = $resName
                SubscriptionId   = $subId
                ScopeDisplayName = "Resource: $resName"
            }
        }
        return @{
            ScopeType        = 'Unknown'
            ResourceDisplay  = $Scope
            SubscriptionId   = $null
            ScopeDisplayName = $Scope
        }
    }

    # Decide whether a given ARM scope passes the caller's SubscriptionIds / OnlyDirtyManagementGroups filter.
    $TestScopeFilter = {
        param([string]$Scope)
        if ($SubscriptionIds -and $SubscriptionIds.Count -gt 0) {
            foreach ($sid in $SubscriptionIds) {
                if (-not [string]::IsNullOrEmpty($sid) -and $Scope -match [regex]::Escape($sid)) {
                    return $true
                }
            }
            return $false
        }
        if ($OnlyDirtyManagementGroups) {
            $dirtyMgs = @()
            if (Get-Variable -Name 'DirtyManagementGroups' -Scope Script -ErrorAction SilentlyContinue) {
                $dirtyMgs = @($script:DirtyManagementGroups | Where-Object { $_ } | Select-Object -Unique)
            }
            if ($Scope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                return ($dirtyMgs -contains $matches[1])
            }
            return $false
        }
        return $true
    }

    # ── 5. Eligible roles via roleEligibilitySchedules ───────────────────────
    if ($IncludeEligible) {
        Write-Verbose "Querying ARM root: roleEligibilitySchedules for principal $UserObjectId"
        $uri = "https://management.azure.com/providers/Microsoft.Authorization/roleEligibilitySchedules" +
               "?api-version=2020-10-01&`$filter=principalId eq '$UserObjectId'"

        $schedules = & $InvokeArmList -Uri $uri
        Write-Verbose "  Received $($schedules.Count) eligible schedule(s)"

        foreach ($schedule in $schedules) {
            try {
                $scope = $schedule.properties.scope
                if (-not (& $TestScopeFilter -Scope $scope)) { continue }

                $roleDefId  = $schedule.properties.roleDefinitionId
                $roleName   = & $ResolveRoleDefName -RoleDefId $roleDefId
                $scopeInfo  = & $GetScopeInfo -Scope $scope
                $eligibilityScheduleId = if ($schedule.PSObject.Properties['id']) { $schedule.id } else { $null }
                $eligibilityScheduleName = if ($schedule.PSObject.Properties['name']) { $schedule.name } else { $null }

                $roleObj = [PSCustomObject]@{
                    Type                = 'AzureResource'
                    DisplayName         = $roleName
                    Status              = 'Eligible'
                    Assignment          = $schedule
                    RoleDefinitionId    = $roleDefId
                    SubscriptionId      = $scopeInfo.SubscriptionId
                    SubscriptionName    = if ($scopeInfo.SubscriptionId -and $subNameCache.ContainsKey($scopeInfo.SubscriptionId)) {
                                             $subNameCache[$scopeInfo.SubscriptionId]
                                         } else { $null }
                    ResourceName        = $scopeInfo.ResourceDisplay
                    ResourceDisplayName = $scopeInfo.ResourceDisplay
                    Scope               = $scopeInfo.ScopeType
                    FullScope           = $scope
                    MemberType          = if ($schedule.properties.memberType) { $schedule.properties.memberType } else { 'Direct' }
                    EndDateTime         = $schedule.properties.endDateTime
                    ScopeDisplayName    = $scopeInfo.ScopeDisplayName
                    Id                  = $roleDefId
                    ObjectId            = $UserObjectId
                    PrincipalId         = $UserObjectId
                    EligibilityScheduleId   = $eligibilityScheduleId
                    EligibilityScheduleName = $eligibilityScheduleName
                    OriginalEligibilityScope = $scope
                }
                $allRoles.Add($roleObj) | Out-Null
                Write-Verbose "  Eligible: $roleName @ $($scopeInfo.ScopeDisplayName)"
            }
            catch {
                Write-Verbose "  Failed to process eligible schedule: $($_.Exception.Message)"
            }
        }
    }

    # ── 6. Active roles ───────────────────────────────────────────────────────
    # Two sources are required for a complete picture:
    #   A) roleAssignmentScheduleInstances – PIM-managed assignments (both
    #      time-limited activations and permanent assignments configured via PIM).
    #      status is 'Provisioned' (not 'Active') for active instances.
    #   B) roleAssignments – ALL direct RBAC assignments including non-PIM ones
    #      (e.g. Owner assigned directly in the portal).
    # We merge both sources and deduplicate on (scope + roleDefinitionId GUID).
    if ($IncludeActive) {

        # Track (scope|rdGuid) pairs already added to avoid duplicates
        $seenActiveKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        # ── A) PIM schedule instances ─────────────────────────────────────────
        Write-Verbose "Querying ARM root: roleAssignmentScheduleInstances for principal $UserObjectId"
        $uri = "https://management.azure.com/providers/Microsoft.Authorization/roleAssignmentScheduleInstances" +
               "?api-version=2020-10-01&`$filter=principalId eq '$UserObjectId'"

        $activeSchedules = & $InvokeArmList -Uri $uri
        Write-Verbose "  Received $($activeSchedules.Count) schedule instance(s)"

        # Statuses that indicate the assignment is NOT currently in effect
        $terminatedStatuses = @('Revoked','Canceled','TimedOut','Denied','AdminDenied','PendingRevocation','Expired')

        foreach ($schedule in $activeSchedules) {
            try {
                $scope = $schedule.properties.scope
                if (-not (& $TestScopeFilter -Scope $scope)) { continue }

                $status = $schedule.properties.status
                if ($status -and $status -in $terminatedStatuses) {
                    Write-Verbose "  Skipping instance with terminal status '$status' (scope: $scope)"
                    continue
                }

                $roleDefId = $schedule.properties.roleDefinitionId
                $rdGuid    = if ($roleDefId -match '/roleDefinitions/([a-fA-F0-9\-]{36})') { $matches[1] } else { $roleDefId }
                $dedupeKey = "$scope|$rdGuid"
                if (-not $seenActiveKeys.Add($dedupeKey)) { continue }

                $roleName  = & $ResolveRoleDefName -RoleDefId $roleDefId
                $scopeInfo = & $GetScopeInfo -Scope $scope

                $roleObj = [PSCustomObject]@{
                    Type                = 'AzureResource'
                    DisplayName         = $roleName
                    Status              = 'Active'
                    Assignment          = $schedule
                    RoleDefinitionId    = $roleDefId
                    SubscriptionId      = $scopeInfo.SubscriptionId
                    SubscriptionName    = if ($scopeInfo.SubscriptionId -and $subNameCache.ContainsKey($scopeInfo.SubscriptionId)) { $subNameCache[$scopeInfo.SubscriptionId] } else { $null }
                    ResourceName        = $scopeInfo.ResourceDisplay
                    ResourceDisplayName = $scopeInfo.ResourceDisplay
                    Scope               = $scopeInfo.ScopeType
                    FullScope           = $scope
                    MemberType          = if ($schedule.properties.memberType) { $schedule.properties.memberType } else { 'Direct' }
                    StartDateTime       = $schedule.properties.startDateTime
                    EndDateTime         = $schedule.properties.endDateTime
                    ScopeDisplayName    = $scopeInfo.ScopeDisplayName
                    ScheduleId          = $schedule.name
                    Id                  = $roleDefId
                    ObjectId            = $UserObjectId
                    PrincipalId         = $UserObjectId
                    AssignmentType      = if ($schedule.properties.assignmentType) { $schedule.properties.assignmentType } else { 'Assigned' }
                }
                $allRoles.Add($roleObj) | Out-Null
                Write-Verbose "  Active (PIM): $roleName @ $($scopeInfo.ScopeDisplayName) [assignmentType=$($schedule.properties.assignmentType)]"
            }
            catch {
                Write-Verbose "  Failed to process schedule instance: $($_.Exception.Message)"
            }
        }

        # ── B) Direct roleAssignments (non-PIM and PIM-activated) ────────────
        # The assignedTo() filter returns all assignments for the principal,
        # including inherited group-based ones via transitivity.
        Write-Verbose "Querying ARM root: roleAssignments for principal $UserObjectId"
        $uri2 = "https://management.azure.com/providers/Microsoft.Authorization/roleAssignments" +
                "?api-version=2022-04-01&`$filter=assignedTo('$UserObjectId')"

        $directAssignments = & $InvokeArmList -Uri $uri2
        Write-Verbose "  Received $($directAssignments.Count) role assignment(s)"

        foreach ($ra in $directAssignments) {
            try {
                $scope = $ra.properties.scope
                if (-not (& $TestScopeFilter -Scope $scope)) { continue }

                $roleDefId = $ra.properties.roleDefinitionId
                $rdGuid    = if ($roleDefId -match '/roleDefinitions/([a-fA-F0-9\-]{36})') { $matches[1] } else { $roleDefId }
                $dedupeKey = "$scope|$rdGuid"
                # Skip if already captured via schedule instances
                if (-not $seenActiveKeys.Add($dedupeKey)) { continue }

                $roleName  = & $ResolveRoleDefName -RoleDefId $roleDefId
                $scopeInfo = & $GetScopeInfo -Scope $scope

                # assignedTo() includes group-inherited roles; detect them by principalId
                $isGroupInherited = $ra.properties.principalId -and
                                    $ra.properties.principalId -ne $UserObjectId

                $roleObj = [PSCustomObject]@{
                    Type                = 'AzureResource'
                    DisplayName         = $roleName
                    Status              = 'Active'
                    Assignment          = $ra
                    RoleDefinitionId    = $roleDefId
                    SubscriptionId      = $scopeInfo.SubscriptionId
                    SubscriptionName    = if ($scopeInfo.SubscriptionId -and $subNameCache.ContainsKey($scopeInfo.SubscriptionId)) { $subNameCache[$scopeInfo.SubscriptionId] } else { $null }
                    ResourceName        = $scopeInfo.ResourceDisplay
                    ResourceDisplayName = $scopeInfo.ResourceDisplay
                    Scope               = $scopeInfo.ScopeType
                    FullScope           = $scope
                    MemberType          = if ($isGroupInherited) { 'Group' } else { 'Direct' }
                    StartDateTime       = $ra.properties.createdOn
                    EndDateTime         = $null   # permanent assignment
                    ScopeDisplayName    = $scopeInfo.ScopeDisplayName
                    ScheduleId          = $ra.name
                    Id                  = $roleDefId
                    ObjectId            = $UserObjectId
                    PrincipalId         = $UserObjectId
                    AssignmentType      = 'Assigned'
                }
                $allRoles.Add($roleObj) | Out-Null
                Write-Verbose "  Active (direct): $roleName @ $($scopeInfo.ScopeDisplayName)"
            }
            catch {
                Write-Verbose "  Failed to process role assignment: $($_.Exception.Message)"
            }
        }
    }

    Write-Verbose "Get-AzureResourceRoles: returning $($allRoles.Count) role(s) total"
    return $allRoles.ToArray()
}
