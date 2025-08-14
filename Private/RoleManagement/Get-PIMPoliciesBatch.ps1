function Get-PIMPoliciesBatch {
    <#
    .SYNOPSIS
        Retrieves policies for multiple roles in batch operations.
    
    .DESCRIPTION
        Fetches role management policies for multiple roles at once to improve performance.
        Uses batch API operations and intelligent filtering to minimize Graph API calls.
    
    .PARAMETER RoleIds
        Array of Entra ID role definition IDs to fetch policies for.
    
    .PARAMETER GroupIds
        Array of group IDs to fetch policies for.
    
    .PARAMETER Type
        The type of roles being processed (Entra or Group).
    
    .PARAMETER PolicyCache
        Hashtable to store the fetched policies in.
    
    .EXAMPLE
        Get-PIMPoliciesBatch -RoleIds $roleIds -Type 'Entra' -PolicyCache $cache
        Fetches policies for the specified Entra roles and stores them in the cache.
    
    .NOTES
        This function uses batch operations to significantly reduce the number of API calls
        required to fetch role management policies.
    #>
    [CmdletBinding()]
    param(
        [string[]]$RoleIds = @(),
        [string[]]$GroupIds = @(),
        [ValidateSet('Entra', 'Group')]
        [string]$Type,
        [Parameter(Mandatory)]
        [hashtable]$PolicyCache
    )
    
    Write-Verbose "Starting batch policy fetch for $Type roles"
    
    # Ensure we have arrays to work with for input parameters
    if (-not $RoleIds) {
        $RoleIds = @()
    } elseif ($RoleIds -isnot [array]) {
        $RoleIds = @($RoleIds)
    }
    
    if (-not $GroupIds) {
        $GroupIds = @()
    } elseif ($GroupIds -isnot [array]) {
        $GroupIds = @($GroupIds)
    }
    
    try {
        if ($Type -eq 'Entra' -and $RoleIds.Count -gt 0) {
            Write-Verbose "Fetching policies for $($RoleIds.Count) Entra roles"
            
            # Filter out roles that already have cached policies
            $uncachedRoleIds = [System.Collections.ArrayList]::new()
            foreach ($roleId in $RoleIds) {
                $cacheKey = "Entra_$roleId"
                if (-not $script:PolicyCache.ContainsKey($cacheKey)) {
                    [void]$uncachedRoleIds.Add($roleId)
                } else {
                    Write-Verbose "Using cached policy for Entra role: $roleId"
                    # Copy from script cache to local cache for this batch operation
                    $PolicyCache[$cacheKey] = $script:PolicyCache[$cacheKey]
                }
            }
            
            # Only fetch policies for roles not in cache
            if ($uncachedRoleIds.Count -gt 0) {
                Write-Verbose "Fetching $($uncachedRoleIds.Count) uncached Entra role policies from Graph API"

                # Prepare base filter and chunking to avoid Graph InvalidFilter when too many OR predicates
                $filterBase = "scopeId eq '/' and scopeType eq 'DirectoryRole'"
                $roleIdsToQuery = $uncachedRoleIds | Sort-Object -Unique
                $chunkSize = 15
                $policyAssignments = [System.Collections.ArrayList]::new()
                $skipBatchMapping = $false

                try {
                    if ($roleIdsToQuery.Count -le $chunkSize) {
                        $orFilter = ($roleIdsToQuery | ForEach-Object { "roleDefinitionId eq '$_'" }) -join ' or '
                        $filter = "$filterBase and ($orFilter)"
                        $assignmentParams = @{ Filter = $filter; All = $true }
                        $result = Get-MgPolicyRoleManagementPolicyAssignment @assignmentParams -ErrorAction Stop
                        if ($result) { $result | ForEach-Object { [void]$policyAssignments.Add($_) } }
                    }
                    else {
                        Write-Verbose "Role count exceeds $chunkSize. Querying in chunks..."
                        for ($i = 0; $i -lt $roleIdsToQuery.Count; $i += $chunkSize) {
                            $chunk = $roleIdsToQuery[$i..([Math]::Min($i + $chunkSize - 1, $roleIdsToQuery.Count - 1))]
                            $chunkFilter = "$filterBase and (" + (($chunk | ForEach-Object { "roleDefinitionId eq '$_'" }) -join ' or ') + ")"
                            $chunkParams = @{ Filter = $chunkFilter; All = $true }
                            $chunkResult = Get-MgPolicyRoleManagementPolicyAssignment @chunkParams -ErrorAction Stop
                            if ($chunkResult) { $chunkResult | ForEach-Object { [void]$policyAssignments.Add($_) } }
                        }
                    }
                }
                catch {
                    if ($_.Exception.Message -match 'Invalid(Filter|Resource)') {
                        Write-Verbose "Filter rejected by Graph (InvalidFilter). Falling back to broad fetch + local filter."
                        $broadParams = @{ Filter = $filterBase; All = $true }
                        $allAssignments = Get-MgPolicyRoleManagementPolicyAssignment @broadParams -ErrorAction Stop
                        if (-not $allAssignments) { $allAssignments = @() } elseif ($allAssignments -isnot [array]) { $allAssignments = @($allAssignments) }
                        $policyAssignments = $allAssignments | Where-Object { $roleIdsToQuery -contains $_.RoleDefinitionId }
                    } else {
                        Write-Warning "Failed to fetch Entra policy assignments: $_"
                        # Fall back to individual fetches if batch fails with other errors
                        foreach ($roleId in $roleIdsToQuery) {
                            try {
                                $singleParams = @{ Filter = "$filterBase and roleDefinitionId eq '$roleId'"; All = $true }
                                $assignment = Get-MgPolicyRoleManagementPolicyAssignment @singleParams -ErrorAction Stop | Select-Object -First 1
                                if ($assignment) {
                                    $policyParams = @{ UnifiedRoleManagementPolicyId = $assignment.PolicyId; ExpandProperty = 'rules' }
                                    $policy = Get-MgPolicyRoleManagementPolicy @policyParams
                                    $policyInfo = ConvertTo-PolicyInfo -Policy $policy
                                    $cacheKey = "Entra_$roleId"
                                    $PolicyCache[$cacheKey] = $policyInfo
                                    # Also cache in script-level cache for future use
                                    $script:PolicyCache[$cacheKey] = $policyInfo
                                }
                            }
                            catch {
                                Write-Verbose "Failed to fetch policy for role $roleId : $_"
                                continue
                            }
                        }
                        $skipBatchMapping = $true
                    }
                }

                if (-not $skipBatchMapping) {
                    # Ensure we have arrays to work with; flatten ArrayList if used
                    if (-not $policyAssignments -or $policyAssignments.Count -eq 0) {
                        $policyAssignments = @()
                    } else {
                        $policyAssignments = @($policyAssignments | ForEach-Object { $_ })
                    }
                    Write-Verbose "Found $($policyAssignments.Count) policy assignments"

                    # Get unique policy IDs
                    $uniquePolicyIds = $policyAssignments | Select-Object -ExpandProperty PolicyId -Unique
                    if (-not $uniquePolicyIds) {
                        $uniquePolicyIds = @()
                    } elseif ($uniquePolicyIds -isnot [array]) {
                        $uniquePolicyIds = @($uniquePolicyIds)
                    }
                    Write-Verbose "Processing $($uniquePolicyIds.Count) unique policies"

                    # Batch fetch all policies with expanded rules
                    foreach ($policyId in $uniquePolicyIds) {
                        try {
                            $policy = Get-MgPolicyRoleManagementPolicy -UnifiedRoleManagementPolicyId $policyId -ExpandProperty "rules"

                            # Process policy rules
                            $policyInfo = ConvertTo-PolicyInfo -Policy $policy

                            # Map policy to all roles that use it
                            $applicableRoles = $policyAssignments | Where-Object { $_.PolicyId -eq $policyId }
                            foreach ($assignment in $applicableRoles) {
                                $cacheKey = "Entra_$($assignment.RoleDefinitionId)"
                                $PolicyCache[$cacheKey] = $policyInfo
                                # Also cache in script-level cache for future use
                                $script:PolicyCache[$cacheKey] = $policyInfo
                                Write-Verbose "Cached policy for Entra role: $($assignment.RoleDefinitionId)"
                            }
                        }
                        catch {
                            Write-Warning "Failed to fetch policy $policyId : $_"
                            continue
                        }
                    }
                }
            } else {
                Write-Verbose "All Entra role policies found in cache"
            }
        }
        
        if ($Type -eq 'Group' -and $GroupIds.Count -gt 0) {
            Write-Verbose "Fetching policies for $($GroupIds.Count) groups"
            
            # Filter out groups that already have cached policies
            $uncachedGroupIds = [System.Collections.ArrayList]::new()
            foreach ($groupId in $GroupIds) {
                $cacheKey = "Group_$groupId"
                if (-not $script:PolicyCache.ContainsKey($cacheKey)) {
                    [void]$uncachedGroupIds.Add($groupId)
                } else {
                    Write-Verbose "Using cached policy for Group: $groupId"
                    # Copy from script cache to local cache for this batch operation
                    $PolicyCache[$cacheKey] = $script:PolicyCache[$cacheKey]
                }
            }
            
            # Only fetch policies for groups not in cache
            if ($uncachedGroupIds.Count -gt 0) {
                Write-Verbose "Fetching $($uncachedGroupIds.Count) uncached group policies from Graph API"

                $filterBaseGroup = "scopeType eq 'Group'"
                $groupIdsToQuery = $uncachedGroupIds | Sort-Object -Unique
                $chunkSize = 15
                $allGroupAssignments = [System.Collections.ArrayList]::new()
                $fallbackToPerGroup = $false

                try {
                    if ($groupIdsToQuery.Count -le $chunkSize) {
                        $orFilter = ($groupIdsToQuery | ForEach-Object { "scopeId eq '$_'" }) -join ' or '
                        $filter = "$filterBaseGroup and ($orFilter)"
                        $uri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=$filter"
                        $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
                        $items = @($response.value)
                        while ($response.'@odata.nextLink') {
                            $response = Invoke-MgGraphRequest -Uri $response.'@odata.nextLink' -Method GET -ErrorAction Stop
                            if ($response.value) { $items += @($response.value) }
                        }
                        foreach ($it in $items) {
                            $norm = [PSCustomObject]@{
                                Id = $it.id
                                PolicyId = $it.policyId
                                RoleDefinitionId = $it.roleDefinitionId
                                ScopeId = $it.scopeId
                                ScopeType = $it.scopeType
                            }
                            [void]$allGroupAssignments.Add($norm)
                        }
                    } else {
                        Write-Verbose "Group count exceeds $chunkSize. Querying in chunks..."
                        for ($i = 0; $i -lt $groupIdsToQuery.Count; $i += $chunkSize) {
                            $chunk = $groupIdsToQuery[$i..([Math]::Min($i + $chunkSize - 1, $groupIdsToQuery.Count - 1))]
                            $chunkFilter = "$filterBaseGroup and (" + (($chunk | ForEach-Object { "scopeId eq '$_'" }) -join ' or ') + ")"
                            $uri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=$chunkFilter"
                            $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
                            $items = @($response.value)
                            while ($response.'@odata.nextLink') {
                                $response = Invoke-MgGraphRequest -Uri $response.'@odata.nextLink' -Method GET -ErrorAction Stop
                                if ($response.value) { $items += @($response.value) }
                            }
                            foreach ($it in $items) {
                                $norm = [PSCustomObject]@{
                                    Id = $it.id
                                    PolicyId = $it.policyId
                                    RoleDefinitionId = $it.roleDefinitionId
                                    ScopeId = $it.scopeId
                                    ScopeType = $it.scopeType
                                }
                                [void]$allGroupAssignments.Add($norm)
                            }
                        }
                    }
                }
                catch {
                    if ($_.Exception.Message -match 'Invalid(Filter|Resource)') {
                        Write-Verbose "Group filter rejected (InvalidFilter). Falling back to per-group fetch."
                        $fallbackToPerGroup = $true
                    } else {
                        Write-Warning "Failed to batch fetch group policy assignments: $_"
                        $fallbackToPerGroup = $true
                    }
                }
                if ($fallbackToPerGroup) {
                    foreach ($groupId in $groupIdsToQuery) {
                        try {
                            $singleParams = @{ Filter = "scopeId eq '$groupId' and scopeType eq 'Group'"; All = $true }
                            $assignments = Get-MgPolicyRoleManagementPolicyAssignment @singleParams -ErrorAction Stop
                            if (-not $assignments) { $assignments = @() } elseif ($assignments -isnot [array]) { $assignments = @($assignments) }

                            $assignment = $assignments | Where-Object { $_.RoleDefinitionId -eq 'member' } | Select-Object -First 1
                # If batch returned no assignments and no exception, fall back to per-group queries
                if (-not $fallbackToPerGroup -and ($allGroupAssignments.Count -eq 0) -and ($groupIdsToQuery.Count -gt 0)) {
                    Write-Verbose "No group policy assignments returned by batch query; falling back to per-group fetch."
                    $fallbackToPerGroup = $true
                }
                            if (-not $assignment) { $assignment = $assignments | Where-Object { $_.RoleDefinitionId -eq 'owner' } | Select-Object -First 1 }

                            if ($assignment) {
                                $policyParams = @{ UnifiedRoleManagementPolicyId = $assignment.PolicyId; ExpandProperty = 'rules' }
                                $policy = Get-MgPolicyRoleManagementPolicy @policyParams
                                $policyInfo = ConvertTo-PolicyInfo -Policy $policy

                                $cacheKey = "Group_$groupId"
                                $PolicyCache[$cacheKey] = $policyInfo
                                $script:PolicyCache[$cacheKey] = $policyInfo
                                Write-Verbose "Cached policy for group: $groupId"
                            } else {
                                Write-Verbose "No policy assignment found for group: $groupId"
                            }
                        }
                        catch {
                            Write-Warning "Failed to fetch policy for group $groupId : $_"
                            continue
                        }
                    }
                } else {
                    # Process batched group assignments
                    if (-not $allGroupAssignments) { $allGroupAssignments = @() }
                    Write-Verbose "Found $($allGroupAssignments.Count) group policy assignments"

                    $uniquePolicyIds = $allGroupAssignments | Select-Object -ExpandProperty PolicyId -Unique
                    if (-not $uniquePolicyIds) { $uniquePolicyIds = @() } elseif ($uniquePolicyIds -isnot [array]) { $uniquePolicyIds = @($uniquePolicyIds) }
                    Write-Verbose "Processing $($uniquePolicyIds.Count) unique group policies"

                    # Fetch all referenced policies once
                    $policyMap = @{}
                    foreach ($policyId in $uniquePolicyIds) {
                        try {
                            $policyParams = @{ UnifiedRoleManagementPolicyId = $policyId; ExpandProperty = 'rules' }
                            $policy = Get-MgPolicyRoleManagementPolicy @policyParams
                            $policyMap[$policyId] = ConvertTo-PolicyInfo -Policy $policy
                        }
                        catch {
                            Write-Warning "Failed to fetch group policy $policyId : $_"
                            continue
                        }
                    }

                    # Map to each group (prefer member over owner)
                    foreach ($groupId in $groupIdsToQuery) {
                        $gAssign = $allGroupAssignments | Where-Object { $_.ScopeId -eq $groupId }
                        if ($gAssign) {
                            $assignment = $gAssign | Where-Object { $_.RoleDefinitionId -eq 'member' } | Select-Object -First 1
                            if (-not $assignment) { $assignment = $gAssign | Where-Object { $_.RoleDefinitionId -eq 'owner' } | Select-Object -First 1 }
                            if ($assignment -and $policyMap.ContainsKey($assignment.PolicyId)) {
                                $cacheKey = "Group_$groupId"
                                $PolicyCache[$cacheKey] = $policyMap[$assignment.PolicyId]
                                $script:PolicyCache[$cacheKey] = $policyMap[$assignment.PolicyId]
                                Write-Verbose "Cached policy for group: $groupId"
                            } else {
                                Write-Verbose "No member/owner policy assignment found for group: $groupId"
                            }
                        } else {
                            Write-Verbose "No policy assignments returned for group: $groupId"
                        }
                    }
                }
            } else {
                Write-Verbose "All group policies found in cache"
            }
        }
        
        Write-Verbose "Completed batch policy fetch for $Type"
    }
    catch {
        Write-Warning "Failed to batch fetch policies: $_"
        throw
    }
}

function ConvertTo-PolicyInfo {
    <#
    .SYNOPSIS
        Converts a Graph API policy object to a standardized policy info object.
    
    .PARAMETER Policy
        The policy object returned from the Graph API.
    
    .OUTPUTS
        PSCustomObject with standardized policy information.
    #>
    param(
        [Parameter(Mandatory)]
        $Policy
    )
    
    $policyInfo = [PSCustomObject]@{
        MaxDuration = 8
        RequiresMfa = $false
        RequiresJustification = $false
        RequiresTicket = $false
        RequiresApproval = $false
        RequiresAuthenticationContext = $false
        AuthenticationContextId = $null
        AuthenticationContextDisplayName = $null
        AuthenticationContextDescription = $null
        AuthenticationContextDetails = $null
    }
    
    if (-not $Policy.Rules) {
        Write-Verbose "Policy has no rules, returning defaults"
        return $policyInfo
    }
    
    foreach ($rule in $Policy.Rules) {
        $ruleType = $rule.AdditionalProperties['@odata.type'] ?? $rule.'@odata.type'
        
        switch ($ruleType) {
            '#microsoft.graph.unifiedRoleManagementPolicyExpirationRule' {
                if ($rule.AdditionalProperties.maximumDuration -or $rule.maximumDuration) {
                    $duration = $rule.AdditionalProperties.maximumDuration ?? $rule.maximumDuration
                    try {
                        $timespan = [System.Xml.XmlConvert]::ToTimeSpan($duration)
                        $policyInfo.MaxDuration = [int]$timespan.TotalHours
                        Write-Verbose "Set max duration to $($policyInfo.MaxDuration) hours"
                    }
                    catch {
                        Write-Verbose "Could not parse duration: $duration"
                    }
                }
            }
            '#microsoft.graph.unifiedRoleManagementPolicyEnablementRule' {
                $enabledRules = @($rule.AdditionalProperties.enabledRules ?? $rule.enabledRules ?? @())
                $policyInfo.RequiresJustification = 'Justification' -in $enabledRules
                $policyInfo.RequiresTicket = 'Ticketing' -in $enabledRules
                $policyInfo.RequiresMfa = 'MultiFactorAuthentication' -in $enabledRules
                $policyInfo.RequiresAuthenticationContext = 'AuthenticationContext' -in $enabledRules
                Write-Verbose "Enablement rules: MFA=$($policyInfo.RequiresMfa), Justification=$($policyInfo.RequiresJustification), Ticket=$($policyInfo.RequiresTicket), AuthContext=$($policyInfo.RequiresAuthenticationContext)"
            }
            '#microsoft.graph.unifiedRoleManagementPolicyApprovalRule' {
                $setting = $rule.AdditionalProperties.setting ?? $rule.setting
                if ($setting -and $setting.isApprovalRequired) {
                    $policyInfo.RequiresApproval = $true
                    Write-Verbose "Approval required: true"
                }
            }
            '#microsoft.graph.unifiedRoleManagementPolicyAuthenticationContextRule' {
                if (($rule.AdditionalProperties.isEnabled ?? $rule.isEnabled) -and 
                    ($rule.AdditionalProperties.claimValue ?? $rule.claimValue)) {
                    $policyInfo.RequiresAuthenticationContext = $true
                    $policyInfo.AuthenticationContextId = $rule.AdditionalProperties.claimValue ?? $rule.claimValue
                    Write-Verbose "Authentication context required: $($policyInfo.AuthenticationContextId)"
                }
            }
        }
    }
    
    return $policyInfo
}

function Get-AuthenticationContextsBatch {
    <#
    .SYNOPSIS
        Retrieves authentication contexts in batch for better performance.
    
    .PARAMETER ContextIds
        Array of authentication context IDs to fetch.
    
    .PARAMETER ContextCache
        Hashtable to store the fetched contexts in.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ContextIds,
        
        [Parameter(Mandatory)]
        [hashtable]$ContextCache
    )
    
    # Ensure we have arrays to work with for input parameters
    if (-not $ContextIds) {
        $ContextIds = @()
    } elseif ($ContextIds -isnot [array]) {
        $ContextIds = @($ContextIds)
    }
    
    Write-Verbose "Batch fetching $($ContextIds.Count) authentication contexts"
    
    foreach ($contextId in $ContextIds) {
        try {
            # Skip if already cached locally
            if ($ContextCache.ContainsKey($contextId)) {
                continue
            }
            
            # Check script-level cache first
            if ($script:AuthenticationContextCache.ContainsKey($contextId)) {
                Write-Verbose "Using cached authentication context: $contextId"
                $ContextCache[$contextId] = $script:AuthenticationContextCache[$contextId]
                continue
            }
            
            # Fetch from Graph API if not in any cache
            $context = Get-MgIdentityConditionalAccessAuthenticationContextClassReference -AuthenticationContextClassReferenceId $contextId
            if ($context) {
                $ContextCache[$contextId] = $context
                # Also cache in script-level cache for future use
                $script:AuthenticationContextCache[$contextId] = $context
                Write-Verbose "Cached authentication context: $contextId - $($context.DisplayName)"
            }
        }
        catch {
            Write-Warning "Failed to fetch authentication context $contextId : $_"
            continue
        }
    }
    
    Write-Verbose "Completed batch authentication context fetch"
}
