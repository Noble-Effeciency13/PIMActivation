function Get-AzureResourceRoles {
    <#
    .SYNOPSIS
        Retrieves Azure resource roles for a user from all accessible subscriptions.

    .DESCRIPTION
        Gets both active and eligible Azure resource roles using Azure PowerShell modules.
        Utilizes silent SSO through Connect-AzAccount with the Graph context user principal.

        The function iterates through subscriptions in the current tenant only and retrieves 
        PIM-eligible and PIM-activated role assignments for the specified user. Results are 
        formatted to align with Entra portal display conventions.

    .PARAMETER UserId
        The user ID (UPN) to retrieve roles for.

    .PARAMETER UserObjectId
        The Azure AD object ID of the user.

    .EXAMPLE
        Get-AzureResourceRoles -UserId "user@domain.com" -UserObjectId "12345678-1234-1234-1234-123456789abc"

        Retrieves all Azure resource roles for the specified user in the current tenant.

    .NOTES
        Requires Az.Accounts (5.1.0+) and Az.Resources (6.0.0+).
        Returns role objects with [Azure] prefix and Entra portal-aligned column mappings.
        Only processes subscriptions within the current authenticated tenant.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$UserObjectId,
        
        # Control which schedules to enumerate to avoid unnecessary work during active-only refreshes
        [switch]$IncludeActive,
        [switch]$IncludeEligible,
        
        # Optional: limit processing to specific subscription Ids for delta refresh scenarios
        [string[]]$SubscriptionIds,
        
        # When set, perform ONLY explicit active queries for script:DirtyManagementGroups
        # and skip tenant root discovery, general MG enumeration, and subscription loops.
        [switch]$OnlyDirtyManagementGroups,
        
        # Enable parallel processing of subscriptions (requires PowerShell 7+)
        [switch]$DisableParallelProcessing,
        
        # Maximum concurrent subscriptions to process in parallel
        [int]$ThrottleLimit = 10
    )
    
    Write-Verbose "Getting Azure resource roles for user: $UserId (ObjectId: $UserObjectId)"
    
    try {
        # Ensure Azure connection with silent SSO
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $azContext) {
            Write-Verbose "No Azure context found, establishing connection using SSO for: $UserId"
            Connect-AzAccount -AccountId $UserId -ErrorAction Stop | Out-Null
            $azContext = Get-AzContext
        }
        else {
            Write-Verbose "Existing Azure context found for account: $($azContext.Account.Id)"
            # Check if the account matches (could be UPN or object ID)
            if ($azContext.Account.Id -ne $UserId -and $azContext.Account.Id -ne $UserObjectId) {
                Write-Verbose "Azure context account mismatch, re-establishing connection for: $UserId"
                Connect-AzAccount -AccountId $UserId -ErrorAction Stop | Out-Null
                $azContext = Get-AzContext -ErrorAction Stop
            }
        }
        
        Write-Verbose "Azure context established successfully"
        
        $allRoles = [System.Collections.ArrayList]::new()
        # Initialize to avoid undefined variable usage in explicit MG query paths
        $tenantRootActive = @()
        
        # Get current tenant ID for filtering
        $currentTenantId = $azContext.Tenant.Id
        Write-Verbose "Current Azure tenant: $currentTenantId"
        
        # Get all accessible subscriptions and filter strictly to the HOME/current tenant
        Write-Verbose "Retrieving subscriptions for home tenant $currentTenantId..."
        $allSubscriptions = Get-AzSubscription -ErrorAction Stop

        # Filter to current tenant only (prefer HomeTenantId when available) and enabled subscriptions
        $subscriptions = $allSubscriptions | Where-Object { 
            $_.State -eq 'Enabled' -and $_.HomeTenantId -eq $currentTenantId
        }

        # If specific SubscriptionIds were provided, further filter to only those
        if ($SubscriptionIds -and $SubscriptionIds.Count -gt 0) {
            $subIdSet = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($sid in $SubscriptionIds) { if ($sid) { $null = $subIdSet.Add($sid) } }
            $subscriptions = @($subscriptions | Where-Object { $subIdSet.Contains($_.Id) })
            Write-Verbose "Delta mode: restricting processing to subscriptions: $(@($subscriptions | ForEach-Object { $_.Id }) -join ', ')"
        }
        
        # Fix Count property issue - handle single object vs array
        $subscriptionCount = if ($subscriptions) {
            if ($subscriptions -is [array]) { $subscriptions.Count } else { 1 }
        } else { 0 }
        
        $allSubscriptionCount = if ($allSubscriptions -is [array]) { $allSubscriptions.Count } else { 1 }
        
        Write-Verbose "Found $subscriptionCount accessible subscriptions in home tenant (filtered from $allSubscriptionCount total)"
        
        if ($subscriptionCount -eq 0) {
            Write-Verbose "No subscriptions found in home tenant"
            return $allRoles.ToArray()
        }
        
        # Convert single subscription to array for consistent processing
        if ($subscriptions -isnot [array]) {
            $subscriptions = @($subscriptions)
        }

        # Initialize management group discovery collection
        $script:DiscoveredManagementGroups = [System.Collections.Generic.HashSet[string]]::new()
        
        # If OnlyDirtyManagementGroups is requested, skip tenant root/MG discovery and subscription processing
        if ($OnlyDirtyManagementGroups) {
            Write-Verbose "OnlyDirtyManagementGroups mode enabled - skipping tenant root and subscription processing; querying explicit dirty MGs only"
        }

        # Process tenant root and management group assignments first (if not in delta mode and not MG-only)
        if (-not $SubscriptionIds -and -not $OnlyDirtyManagementGroups) {
            Write-Verbose "Checking tenant root and management group assignments..."
            
            # Check tenant root assignments (scope "/")
            try {
                Write-Verbose "Fetching tenant root assignments"
                if ($IncludeActive) {
                    $tenantRootActive = @()
                    $tra = Get-AzRoleAssignment -Scope "/" -ObjectId $UserObjectId -ErrorAction SilentlyContinue
                    if ($tra) { $tenantRootActive = ($tra -is [array]) ? $tra : @($tra) }
                    if ($tenantRootActive -and $tenantRootActive.Count -gt 0) {
                        foreach ($assignment in @($tenantRootActive)) {
                            # Create properly structured Azure role object
                            $tenantRole = [PSCustomObject]@{
                                Type = 'AzureResource'
                                DisplayName = $assignment.RoleDefinitionName
                                Status = 'Active'
                                Assignment = $assignment
                                RoleDefinitionId = $assignment.RoleDefinitionId
                                SubscriptionId = $null
                                SubscriptionName = $null
                                ResourceName = '/'
                                ResourceDisplayName = '/'
                                Scope = 'Tenant'
                                FullScope = $assignment.Scope
                                MemberType = 'Inherited'
                                EndDateTime = $null
                                ScopeDisplayName = 'Tenant Root'
                                Id = $assignment.RoleDefinitionId
                                ObjectId = $assignment.ObjectId
                                PrincipalId = $assignment.ObjectId
                            }
                            $allRoles.Add($tenantRole) | Out-Null
                        }
                        Write-Verbose "Found $(@($tenantRootActive).Count) tenant root active assignments"
                    }
                }
                
                if ($IncludeEligible -and (Get-Command Get-AzRoleEligibilitySchedule -ErrorAction SilentlyContinue)) {
                    Write-Verbose "Querying tenant root eligible assignments for user: $UserObjectId"
                    $tenantRootEligible = Get-AzRoleEligibilitySchedule -Scope "/" -Filter "principalId eq '$UserObjectId'" -ErrorAction SilentlyContinue
                    Write-Verbose "Raw tenant root eligible query returned: $(@($tenantRootEligible).Count) items"
                    if ($tenantRootEligible) {
                        foreach ($schedule in @($tenantRootEligible)) {
                            Write-Verbose "Processing tenant root eligible schedule: RoleDefId=$($schedule.RoleDefinitionId), Scope=$($schedule.Scope), Status=$($schedule.Status)"
                            
                            # Get role definition name with better error handling
                            $roleName = "Unknown Role"
                            $roleDefinition = $null
                            try {
                                # Extract GUID if full path
                                $roleDefId = $schedule.RoleDefinitionId
                                if ($roleDefId -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
                                    $roleDefId = $matches[1]
                                }
                                $roleDefinition = Get-AzRoleDefinition -Id $roleDefId -ErrorAction Stop
                                if ($roleDefinition -and $roleDefinition.Name) {
                                    $roleName = $roleDefinition.Name
                                }
                            } catch {
                                Write-Verbose "Could not resolve role definition for ID: $($schedule.RoleDefinitionId), error: $($_.Exception.Message)"
                                # Try alternative approach using the subscription context
                                try {
                                    if ($subscriptions -and $subscriptions.Count -gt 0) {
                                        Set-AzContext -Subscription $subscriptions[0].Id -ErrorAction Stop
                                        $roleDefinition = Get-AzRoleDefinition -Id $roleDefId -ErrorAction Stop
                                        if ($roleDefinition -and $roleDefinition.Name) {
                                            $roleName = $roleDefinition.Name
                                        }
                                    }
                                } catch {
                                    Write-Verbose "Alternative role definition lookup also failed for: $($schedule.RoleDefinitionId)"
                                }
                            }
                            
                        # Process management group roles found in tenant root query and collect MG info
                        $isDuplicate = $false
                        $isMGRole = $false
                        if ($roleName -ne "Unknown Role") {
                            # Check if this role is from a management group - process it here instead of skipping
                            if ($schedule.Scope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                                $mgName = $matches[1]
                                Write-Verbose "Processing management group role '$roleName' found in tenant root query - MG: $mgName"
                                
                                # Create management group role object
                                $mgRole = [PSCustomObject]@{
                                    Type = 'AzureResource'
                                    DisplayName = $roleName
                                    Status = 'Eligible'
                                    Assignment = $schedule
                                    RoleDefinitionId = $schedule.RoleDefinitionId
                                    SubscriptionId = $null
                                    SubscriptionName = $null
                                    ResourceName = $mgName
                                    ResourceDisplayName = $mgName
                                    Scope = 'Management Group'
                                    FullScope = $schedule.Scope
                                    MemberType = 'Direct'
                                    EndDateTime = $schedule.EndDateTime
                                    ScopeDisplayName = "MG: $mgName"
                                    Id = $schedule.RoleDefinitionId
                                    ObjectId = $schedule.PrincipalId
                                    PrincipalId = $schedule.PrincipalId
                                }
                                $allRoles.Add($mgRole) | Out-Null
                                Write-Verbose "Added management group role: $roleName on MG $mgName"
                                
                                # Add this MG to our discovery list for active role checking
                                if (-not $script:DiscoveredManagementGroups) {
                                    $script:DiscoveredManagementGroups = [System.Collections.Generic.HashSet[string]]::new()
                                }
                                $script:DiscoveredManagementGroups.Add($mgName) | Out-Null
                                
                                $isMGRole = $true
                                $isDuplicate = $true  # Skip normal tenant root processing
                            }
                            # Check if this is really at tenant root
                            elseif ($schedule.Scope -eq "/" -or $schedule.Scope -eq "") {
                                # Verify the role actually exists at tenant root by re-checking with more specific filter
                                try {
                                    $verifyRole = Get-AzRoleEligibilitySchedule -Scope "/" -Filter "principalId eq '$UserObjectId' and roleDefinitionId eq '$($schedule.RoleDefinitionId)'" -ErrorAction SilentlyContinue
                                    $verifyCount = @($verifyRole).Count
                                    Write-Verbose "Verification check for '$roleName' at tenant root returned $verifyCount results"
                                    
                                    if (-not $verifyRole -or $verifyCount -eq 0) {
                                        $isDuplicate = $true
                                        Write-Verbose "Skipping phantom tenant root role '$roleName' - not found in verification check"
                                    }
                                    else {
                                        # Check if we'll find this same role at subscription level to avoid true duplicates
                                        foreach ($sub in $subscriptions) {
                                            try {
                                                $subRoles = Get-AzRoleAssignment -Scope "/subscriptions/$($sub.Id)" -ObjectId $UserObjectId -ErrorAction SilentlyContinue
                                                if ($subRoles | Where-Object { $_.RoleDefinitionId -eq $schedule.RoleDefinitionId }) {
                                                    $isDuplicate = $true
                                                    Write-Verbose "Skipping duplicate tenant root role '$roleName' - found at subscription level"
                                                    break
                                                }
                                            } catch { }
                                        }
                                    }
                                } catch {
                                    Write-Verbose "Could not verify tenant root role '$roleName': $($_.Exception.Message)"
                                    # If verification fails, be conservative and exclude it
                                    $isDuplicate = $true
                                }
                            }
                            else {
                                # For subscription-scoped roles that appear in tenant root query, skip them
                                if ($schedule.Scope -match "^/subscriptions/") {
                                    $isDuplicate = $true
                                    Write-Verbose "Skipping subscription-scoped role '$roleName' found in tenant root query - will be processed in subscription loop. Scope: $($schedule.Scope)"
                                } else {
                                    Write-Verbose "Tenant root query returned role with unexpected scope '$($schedule.Scope)' - processing anyway"
                                }
                            }
                        }
                        
                        if (-not $isDuplicate) {
                                # Create properly structured Azure role object
                                $tenantRole = [PSCustomObject]@{
                                    Type = 'AzureResource'
                                    DisplayName = $roleName
                                    Status = 'Eligible'
                                    Assignment = $schedule
                                    RoleDefinitionId = $schedule.RoleDefinitionId
                                    SubscriptionId = $null
                                    SubscriptionName = $null
                                    ResourceName = '/'
                                    ResourceDisplayName = '/'
                                    Scope = 'Tenant'
                                    FullScope = $schedule.Scope
                                    MemberType = 'Inherited'
                                    EndDateTime = $schedule.EndDateTime
                                    ScopeDisplayName = 'Tenant Root'
                                    Id = $schedule.RoleDefinitionId
                                    ObjectId = $schedule.PrincipalId
                                    PrincipalId = $schedule.PrincipalId
                                }
                                $allRoles.Add($tenantRole) | Out-Null
                            }
                        }
                        Write-Verbose "Found $(@($tenantRootEligible).Count) tenant root eligible assignments"
                    }
                }
            }
            catch {
                Write-Verbose "Failed to fetch tenant root assignments: $($_.Exception.Message)"
            }

            # Check management group assignments
            try {
                Write-Verbose "Fetching management group assignments"
                
                # Get all management groups the user has access to
                $managementGroups = @()
                $mgIds = @()
                
                # First, collect MG IDs from roles found in tenant root query
                if ($script:DiscoveredManagementGroups -and $script:DiscoveredManagementGroups.Count -gt 0) {
                    $mgIds += @($script:DiscoveredManagementGroups)
                    Write-Verbose "Found $($script:DiscoveredManagementGroups.Count) management group IDs from tenant root roles: $(@($script:DiscoveredManagementGroups) -join ', ')"
                }
                
                try {
                    # Try to get management groups the user has roles in by checking role assignments
                    Write-Verbose "Attempting to discover management groups via direct enumeration..."
                    
                    # First try direct enumeration
                    $allMGs = @(Get-AzManagementGroup -ErrorAction SilentlyContinue | Where-Object { $_.Name })
                    Write-Verbose "Found $($allMGs.Count) management groups via direct enumeration"
                    
                    # Add MG IDs from direct enumeration
                    foreach ($mg in $allMGs) {
                        if ($mg.Name -notin $mgIds) {
                            $mgIds += $mg.Name
                        }
                    }
                    
                    # Create management group objects for all discovered IDs
                    foreach ($mgId in $mgIds) {
                        try {
                            $mgDetails = $allMGs | Where-Object { $_.Name -eq $mgId } | Select-Object -First 1
                            if (-not $mgDetails) {
                                # Try to get details for MG ID found in roles
                                try {
                                    $mgDetails = Get-AzManagementGroup -GroupId $mgId -ErrorAction SilentlyContinue
                                    if ($mgDetails) {
                                        Write-Verbose "Retrieved management group details for $mgId`: $($mgDetails.DisplayName)"
                                    }
                                } catch {
                                    Write-Verbose "Could not get details for management group $mgId`: $($_.Exception.Message)"
                                    # Create a minimal object
                                    $mgDetails = [PSCustomObject]@{
                                        Name = $mgId
                                        DisplayName = $mgId
                                    }
                                }
                            }
                            
                            if ($mgDetails) {
                                $managementGroups += $mgDetails
                                Write-Verbose "Added management group to processing list: $($mgDetails.DisplayName) ($($mgDetails.Name))"
                            }
                        } catch {
                            Write-Verbose "Could not process management group '$mgId': $($_.Exception.Message)"
                        }
                    }
                    
                    Write-Verbose "Found $($managementGroups.Count) management groups with role assignments to check"
                } catch {
                    Write-Verbose "Could not enumerate management groups: $($_.Exception.Message)"
                }
                
                foreach ($mg in $managementGroups) {
                    try {
                        $mgScope = "/providers/Microsoft.Management/managementGroups/$($mg.Name)"
                        Write-Verbose "Checking management group: $($mg.DisplayName) ($($mg.Name))"
                        
                        # Check active assignments in this management group
                        if ($IncludeActive) {
                            $mgActive = Get-AzRoleAssignment -Scope $mgScope -ObjectId $UserObjectId -ErrorAction SilentlyContinue
                            if ($mgActive) {
                                foreach ($assignment in @($mgActive)) {
                                    $mgRole = [PSCustomObject]@{
                                        Type = 'AzureResource'
                                        DisplayName = $assignment.RoleDefinitionName
                                        Status = 'Active'
                                        Assignment = $assignment
                                        RoleDefinitionId = $assignment.RoleDefinitionId
                                        SubscriptionId = $null
                                        SubscriptionName = $null
                                        ResourceName = $mg.DisplayName
                                        ResourceDisplayName = $mg.DisplayName
                                        Scope = 'Management Group'
                                        FullScope = $assignment.Scope
                                        MemberType = 'Direct'
                                        EndDateTime = $null
                                        ScopeDisplayName = "MG: $($mg.DisplayName)"
                                        Id = $assignment.RoleDefinitionId
                                        ObjectId = $assignment.ObjectId
                                        PrincipalId = $assignment.ObjectId
                                    }
                                    $allRoles.Add($mgRole) | Out-Null
                                    Write-Verbose "Found active management group role: $($assignment.RoleDefinitionName) on $($mg.DisplayName)"
                                }
                            }
                        }
                        
                        # Check eligible assignments in this management group
                        if ($IncludeEligible -and (Get-Command Get-AzRoleEligibilitySchedule -ErrorAction SilentlyContinue)) {
                            $mgEligible = Get-AzRoleEligibilitySchedule -Scope $mgScope -Filter "principalId eq '$UserObjectId'" -ErrorAction SilentlyContinue
                            if ($mgEligible) {
                                foreach ($schedule in @($mgEligible)) {
                                    # Verify this schedule actually belongs to this management group
                                    if ($schedule.Scope -ne $mgScope) {
                                        Write-Verbose "Skipping schedule with scope '$($schedule.Scope)' - doesn't match current MG scope '$mgScope'"
                                        continue
                                    }
                                    
                                    # Get role definition name
                                    $roleName = "Unknown Role"
                                    try {
                                        # Extract GUID if full path
                                        $roleDefId = $schedule.RoleDefinitionId
                                        if ($roleDefId -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
                                            $roleDefId = $matches[1]
                                        }
                                        $roleDef = Get-AzRoleDefinition -Id $roleDefId -ErrorAction SilentlyContinue
                                        if ($roleDef) { $roleName = $roleDef.Name }
                                    } catch {
                                        Write-Verbose "Could not resolve role definition for '$($schedule.RoleDefinitionId)': $($_.Exception.Message)"
                                    }
                                    
                                    $mgRole = [PSCustomObject]@{
                                        Type = 'AzureResource'
                                        DisplayName = $roleName
                                        Status = 'Eligible'
                                        Assignment = $schedule
                                        RoleDefinitionId = $schedule.RoleDefinitionId
                                        SubscriptionId = $null
                                        SubscriptionName = $null
                                        ResourceName = $mg.DisplayName
                                        ResourceDisplayName = $mg.DisplayName
                                        Scope = 'Management Group'
                                        FullScope = $schedule.Scope
                                        MemberType = 'Direct'
                                        EndDateTime = $schedule.EndDateTime
                                        ScopeDisplayName = "MG: $($mg.DisplayName)"
                                        Id = $schedule.RoleDefinitionId
                                        ObjectId = $schedule.PrincipalId
                                        PrincipalId = $schedule.PrincipalId
                                    }
                                    $allRoles.Add($mgRole) | Out-Null
                                    Write-Verbose "Found eligible management group role: $roleName on $($mg.DisplayName) with scope $($schedule.Scope)"
                                }
                            }
                        }
                    } catch {
                        Write-Verbose "Failed to process management group '$($mg.DisplayName)': $($_.Exception.Message)"
                    }
                }
            }
            catch {
                Write-Verbose "Management group processing failed: $($_.Exception.Message)"
            }
        }

        # Even in delta mode (SubscriptionIds provided), include explicit MG active refresh for any dirty MGs
        if ($IncludeActive -and (Get-Variable -Name 'DirtyManagementGroups' -Scope Script -ErrorAction SilentlyContinue) -and $script:DirtyManagementGroups.Count -gt 0) {
            $dirtyMgs = @($script:DirtyManagementGroups | Select-Object -Unique)
            Write-Verbose "Performing explicit active query for discovered management groups: $($dirtyMgs -join ', ')"
            foreach ($mgName in $dirtyMgs) {
                try {
                    $mgScope = "/providers/Microsoft.Management/managementGroups/$mgName"
                    $mgActive = Get-AzRoleAssignment -Scope $mgScope -ObjectId $UserObjectId -ErrorAction SilentlyContinue
                    if ($mgActive) {
                        foreach ($assignment in @($mgActive)) {
                            # Skip tenant-root inherited duplicates
                            $isTenantDup = $false
                            foreach ($t in @($tenantRootActive)) {
                                if ($t.RoleDefinitionId -eq $assignment.RoleDefinitionId -and $t.ObjectId -eq $assignment.ObjectId) { $isTenantDup = $true; break }
                            }
                            if ($isTenantDup) { Write-Verbose "Skipping inherited management group role '$($assignment.RoleDefinitionName)' (provided by tenant root) on ${mgName}"; continue }

                            $mgRole = [PSCustomObject]@{
                                Type = 'AzureResource'
                                DisplayName = $assignment.RoleDefinitionName
                                Status = 'Active'
                                Assignment = $assignment
                                RoleDefinitionId = $assignment.RoleDefinitionId
                                SubscriptionId = $null
                                SubscriptionName = $null
                                ResourceName = $mgName
                                ResourceDisplayName = $mgName
                                Scope = 'Management Group'
                                FullScope = $assignment.Scope
                                MemberType = 'Direct'
                                EndDateTime = $null
                                ScopeDisplayName = 'Management Group'
                                Id = $assignment.RoleDefinitionId
                                ObjectId = $assignment.ObjectId
                                PrincipalId = $assignment.ObjectId
                            }
                            $allRoles.Add($mgRole) | Out-Null
                            Write-Verbose "Found active management group role via explicit query: $($assignment.RoleDefinitionName) on ${mgName}"
                        }
                    }
                } catch { Write-Verbose "Explicit MG active query failed for ${mgName}: $($_.Exception.Message)" }
            }
        }

        # Fallback: Explicitly query active assignments for discovered management groups when direct enumeration yielded none
        try {
            if (-not $OnlyDirtyManagementGroups -and $IncludeActive -and ($script:DiscoveredManagementGroups -and $script:DiscoveredManagementGroups.Count -gt 0)) {
                Write-Verbose "Performing explicit active query for discovered management groups: $(@($script:DiscoveredManagementGroups) -join ', ')"
                foreach ($mgName in $script:DiscoveredManagementGroups) {
                    $mgScope = "/providers/Microsoft.Management/managementGroups/$mgName"
                    try {
                        $mgActive = Get-AzRoleAssignment -Scope $mgScope -ObjectId $UserObjectId -ErrorAction SilentlyContinue
                        if ($mgActive) {
                            foreach ($assignment in @($mgActive)) {
                                # If the same role is active at tenant root for this principal, treat MG visibility as inherited and skip
                                $inheritedFromTenant = $false
                                foreach ($t in @($tenantRootActive)) {
                                    if ($t.RoleDefinitionId -eq $assignment.RoleDefinitionId -and $t.ObjectId -eq $assignment.ObjectId) {
                                        $inheritedFromTenant = $true
                                        break
                                    }
                                }
                                if ($inheritedFromTenant) {
                                    Write-Verbose "Skipping inherited management group role '$($assignment.RoleDefinitionName)' (provided by tenant root) on ${mgName}"
                                    continue
                                }
                                # Attempt to enrich activation window (Start/End) for MG scope using assignment schedules
                                $startTime = $null; $endTime = $null
                                try {
                                    if (Get-Command Get-AzRoleAssignmentSchedule -ErrorAction SilentlyContinue) {
                                        $roleDefGuid = $assignment.RoleDefinitionId
                                        if ($roleDefGuid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefGuid = $matches[1] }
                                        $roleDefPath = "$mgScope/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"
                                        $filter = "principalId eq '$UserObjectId' and roleDefinitionId eq '$roleDefPath' and status eq 'Active'"
                                        $sched = Get-AzRoleAssignmentSchedule -Scope $mgScope -Filter $filter -ErrorAction SilentlyContinue
                                        $schedArr = @(); if ($sched) { $schedArr = ($sched -is [array]) ? $sched : @($sched) }
                                        if ($schedArr.Count -gt 0) {
                                            $selected = $schedArr | Sort-Object {
                                                $end = $null
                                                if ($_.PSObject.Properties["EndDateTime"]) { $end = $_.EndDateTime }
                                                elseif ($_.PSObject.Properties["ScheduleInfo"]) { $end = $_.ScheduleInfo.Expiration.EndDateTime }
                                                if (-not $end) { [datetime]::MinValue } else { [datetime]$end }
                                            } -Descending | Select-Object -First 1

                                            if ($selected.PSObject.Properties["StartDateTime"]) { $startTime = $selected.StartDateTime }
                                            if ($selected.PSObject.Properties["EndDateTime"])   { $endTime   = $selected.EndDateTime }
                                            if (-not $startTime -and $selected.PSObject.Properties["ScheduleInfo"]) { $startTime = $selected.ScheduleInfo.StartDateTime }
                                            if (-not $endTime   -and $selected.PSObject.Properties["ScheduleInfo"]) { $endTime   = $selected.ScheduleInfo.Expiration.EndDateTime }
                                        }
                                    }
                                } catch { Write-Verbose "Failed to enrich MG active role time window: $($_.Exception.Message)" }

                                $mgRole = [PSCustomObject]@{
                                    Type = 'AzureResource'
                                    DisplayName = $assignment.RoleDefinitionName
                                    Status = 'Active'
                                    Assignment = $assignment
                                    RoleDefinitionId = $assignment.RoleDefinitionId
                                    SubscriptionId = $null
                                    SubscriptionName = $null
                                    ResourceName = $mgName
                                    ResourceDisplayName = $mgName
                                    Scope = 'Management Group'
                                    FullScope = $assignment.Scope
                                    MemberType = 'Direct'
                                    StartDateTime = $startTime
                                    EndDateTime = $endTime
                                    ScopeDisplayName = 'Management Group'
                                    Id = $assignment.RoleDefinitionId
                                    ObjectId = $assignment.ObjectId
                                    PrincipalId = $assignment.ObjectId
                                }
                                $allRoles.Add($mgRole) | Out-Null
                                Write-Verbose "Found active management group role via explicit query: $($assignment.RoleDefinitionName) on ${mgName}"
                            }
                        }
                        else {
                            Write-Verbose "No active assignments found for management group $mgName via explicit query"
                        }
                    }
                    catch {
                        Write-Verbose "Explicit MG query failed for ${mgName}: $($_.Exception.Message)"
                    }
                }
            }
        }
        catch {
            Write-Verbose "Explicit MG active query block failed: $($_.Exception.Message)"
        }

        # Optional quick scan scoped properly (never returns early)
        if (-not $OnlyDirtyManagementGroups) {
            $subsToCheck = $subscriptions | Select-Object -First ([Math]::Min(3, $subscriptions.Count))
            foreach ($checkSub in $subsToCheck) {
                try {
                    Write-Verbose "Quick role check on subscription: $($checkSub.Name)"
                    Select-AzSubscription -SubscriptionId $checkSub.Id -Tenant $currentTenantId -ErrorAction SilentlyContinue | Out-Null

                    # Scoped active + eligible checks using ObjectId
                    $quickActive   = Get-AzRoleAssignment -Scope "/subscriptions/$($checkSub.Id)" -ObjectId $UserObjectId -ErrorAction SilentlyContinue
                    $quickEligible = $null
                    if (Get-Command Get-AzRoleEligibilitySchedule -ErrorAction SilentlyContinue) {
                        $quickEligible = Get-AzRoleEligibilitySchedule -Scope "/subscriptions/$($checkSub.Id)" -Filter "principalId eq '$UserObjectId'" -ErrorAction SilentlyContinue
                    }
                    $ca = @($quickActive).Count
                    $ce = @($quickEligible).Count
                    Write-Verbose "Quick scan counts for '$($checkSub.Name)': Active=$ca, Eligible=$ce"
                }
                catch {
                    Write-Verbose "Quick role check failed for subscription $($checkSub.Name): $($_.Exception.Message)"
                }
            }
        }
        
        # Process subscriptions with progress feedback
        if (-not $OnlyDirtyManagementGroups) {
        
        # Check if parallel processing is requested and supported
        $useParallel = -not $DisableParallelProcessing -and $PSVersionTable.PSVersion.Major -ge 7 -and $subscriptions.Count -gt 1
        
        if ($useParallel) {
            Write-Verbose "Using parallel processing for $($subscriptions.Count) subscriptions (ThrottleLimit: $ThrottleLimit)"
            
            # Create thread-safe collection for results and progress tracking
            $allSubscriptionRoles = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
            $processedCount = [System.Collections.Concurrent.ConcurrentDictionary[string,bool]]::new()
            $startTime = Get-Date
            
            # Process subscriptions in parallel
            $subscriptions | ForEach-Object -Parallel {
                $subscription = $_
                $localRoles = [System.Collections.ArrayList]::new()
                
                # Import required variables into parallel scope
                $currentTenantIdLocal = $using:currentTenantId
                $UserObjectIdLocal = $using:UserObjectId
                $IncludeActiveLocal = $using:IncludeActive
                $IncludeEligibleLocal = $using:IncludeEligible
                $allSubscriptionRolesLocal = $using:allSubscriptionRoles
                $processedCountLocal = $using:processedCount
                
                try {
                    Write-Verbose "[Parallel] Starting subscription: $($subscription.Name) ($($subscription.Id))"
                    
                    # Select subscription context with tenant scoping BEFORE any queries
                    Select-AzSubscription -SubscriptionId $subscription.Id -Tenant $currentTenantIdLocal -ErrorAction SilentlyContinue | Out-Null
                
                    # Get active role assignments at subscription scope (optional)
                    $activeAssignments = @()
                    if ($IncludeActiveLocal) {
                        Write-Verbose "[Parallel] Fetching active role assignments for subscription: $($subscription.Name)"
                        try {
                            $activeAssignments = Get-AzRoleAssignment -Scope "/subscriptions/$($subscription.Id)" -ObjectId $UserObjectIdLocal -ErrorAction SilentlyContinue
                            if ($activeAssignments -and ($activeAssignments -isnot [array])) { $activeAssignments = @($activeAssignments) }
                        }
                        catch {
                            Write-Verbose "[Parallel] Active assignment query failed on $($subscription.Id): $($_.Exception.Message)"
                            $activeAssignments = @()
                        }

                        # REST fallback when Az cmdlet returns no results
                        if (-not $activeAssignments -or @($activeAssignments).Count -eq 0) {
                            Write-Verbose "[Parallel] No active assignments via Az cmdlet for $($subscription.Name); attempting REST fallback"
                            try {
                                $context = Get-AzContext -ErrorAction SilentlyContinue
                                if ($context -and $context.Account.Id) {
                                    $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
                                        $context.Account, $context.Environment, $context.Tenant.Id, $null, "Never", $null, "https://management.azure.com/"
                                    ).AccessToken
                                    $headers = @{
                                        'Authorization' = "Bearer $token"
                                        'Content-Type'  = 'application/json'
                                    }
                                    $uri = "https://management.azure.com/subscriptions/$($subscription.Id)/providers/Microsoft.Authorization/roleAssignments?api-version=2022-04-01&`$filter=principalId eq '$UserObjectIdLocal'"
                                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                                if ($response -and $response.value) {
                                    $restAssignments = @()
                                    foreach ($item in $response.value) {
                                        $roleDefPath = $item.properties.roleDefinitionId
                                        $roleDefGuid = $roleDefPath
                                        if ($roleDefGuid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefGuid = $matches[1] }
                                        $roleDef = $null; $roleName = "Unknown Role"
                                        try {
                                            $roleDef = Get-AzRoleDefinition -Id $roleDefGuid -ErrorAction SilentlyContinue
                                            if ($roleDef) { $roleName = $roleDef.Name }
                                        } catch { }

                                        $restAssignments += [PSCustomObject]@{
                                            RoleAssignmentId   = $item.name
                                            RoleDefinitionId   = $roleDefPath
                                            RoleDefinitionName = $roleName
                                            Scope              = $item.properties.scope
                                            ObjectId           = $item.properties.principalId
                                            ObjectType         = 'User'
                                        }
                                    }
                                    $activeAssignments = $restAssignments
                                    Write-Verbose "REST fallback returned $(@($activeAssignments).Count) active assignments for $($subscription.Name)"
                                }
                            }
                            }
                            catch {
                                Write-Verbose "[Parallel] REST fallback for active assignments failed on $($subscription.Id): $($_.Exception.Message)"
                            }
                        }
                    }

                    # Do NOT use -IncludeInherited (parameter not available across Az versions). Skip inherited enrichment to prevent errors.

                    if ($IncludeActiveLocal) { Write-Verbose "[Parallel] Found $(@($activeAssignments).Count) active role assignments" }

                    # Get eligible role assignments (PIM) at subscription scope using principalId filter (optional)
                    $eligibleAssignments = @()
                    if ($IncludeEligibleLocal) {
                        Write-Verbose "[Parallel] Fetching eligible role assignments for subscription: $($subscription.Name)"
                        try {
                            if (Get-Command Get-AzRoleEligibilitySchedule -ErrorAction SilentlyContinue) {
                                $eligibleAssignments = Get-AzRoleEligibilitySchedule -Scope "/subscriptions/$($subscription.Id)" -Filter "principalId eq '$UserObjectIdLocal'" -ErrorAction SilentlyContinue
                                if ($eligibleAssignments -and ($eligibleAssignments -isnot [array])) { $eligibleAssignments = @($eligibleAssignments) }
                                Write-Verbose "[Parallel] Found $(@($eligibleAssignments).Count) eligible assignments via Az.Resources"
                            }
                            else {
                                Write-Verbose "[Parallel] Az.Resources PIM cmdlets not available, using REST API fallback"
                                $context = Get-AzContext -ErrorAction SilentlyContinue
                                if ($context -and $context.Account.Id) {
                                    $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
                                        $context.Account, $context.Environment, $context.Tenant.Id, $null, "Never", $null, "https://management.azure.com/"
                                    ).AccessToken
                                    $headers = @{
                                        'Authorization' = "Bearer $token"
                                        'Content-Type'  = 'application/json'
                                    }
                                    $uri = "https://management.azure.com/subscriptions/$($subscription.Id)/providers/Microsoft.Authorization/roleEligibilitySchedules?api-version=2020-10-01&`$filter=principalId eq '$UserObjectIdLocal'"
                                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                                if ($response -and $response.value) {
                                    $eligibleAssignments = @($response.value)
                                    Write-Verbose "Found $(@($eligibleAssignments).Count) eligible assignments via REST API"
                                }
                            }
                        }
                        }
                        catch {
                            Write-Verbose "[Parallel] Could not retrieve eligible assignments: $($_.Exception.Message)"
                            $eligibleAssignments = @()
                        }
                    }
                
                    # Process active assignments
                    foreach ($assignment in @($activeAssignments)) {
                        # Skip tenant-root scoped entries ("/") that are duplicated on every subscription
                        if ($assignment.Scope -eq "/") {
                            Write-Verbose "[Parallel] Skipping duplicate tenant root active assignment '$($assignment.RoleDefinitionName)' surfaced in subscription '$($subscription.Name)'"
                            continue
                        }
                        Write-Verbose "[Parallel] Processing active assignment: $($assignment.RoleDefinitionName) at scope $($assignment.Scope)"
                    
                    # Simplified scope info calculation for parallel processing
                    $scopeInfo = @{
                        ResourceDisplay = if ($assignment.Scope -match "^/subscriptions/([^/]+)$") { $subscription.Name }
                                         elseif ($assignment.Scope -match "^/subscriptions/([^/]+)/resourceGroups/([^/]+)$") { $matches[2] }
                                         elseif ($assignment.Scope -eq "/" -or $assignment.Scope -eq "") { "/" }
                                         else { $assignment.Scope }
                        ScopeType = if ($assignment.Scope -match "^/subscriptions/([^/]+)$") { "Subscription" }
                                   elseif ($assignment.Scope -match "^/subscriptions/([^/]+)/resourceGroups/([^/]+)$") { "Resource Group" }
                                   elseif ($assignment.Scope -eq "/" -or $assignment.Scope -eq "") { "Tenant" }
                                   else { "Resource" }
                    }
                    
                    # Simplified member type detection for parallel processing
                    $memberType = if ($assignment.ObjectType -eq 'Group') { 
                        "Group" 
                    } elseif ($assignment.Scope -eq "/subscriptions/$($subscription.Id)" -or $assignment.Scope -match "^/subscriptions/$($subscription.Id)/resourceGroups/") { 
                        "Direct" 
                    } else { 
                        "Inherited" 
                    }
                    
                    # If assignment scope doesn't match current subscription scope, mark as inherited
                    if ($assignment.Scope -ne "/subscriptions/$($subscription.Id)" -and 
                        $assignment.Scope -notmatch "^/subscriptions/$($subscription.Id)/resourceGroups/" -and
                        $assignment.Scope -notmatch "^/subscriptions/$($subscription.Id)/resourceGroups/.+") {
                        Write-Verbose "Assignment scope $($assignment.Scope) indicates inheritance -> Inherited"
                        $memberType = "Inherited"
                    }
                    
                    $roleObject = [PSCustomObject]@{
                        RoleId             = "$($assignment.RoleAssignmentId)-$($subscription.Id)"
                        RoleDefinitionId   = $assignment.RoleDefinitionId
                        DisplayName        = $assignment.RoleDefinitionName
                        ResourceName       = $scopeInfo.ResourceDisplay
                        ResourceDisplayName= $scopeInfo.ResourceDisplay
                        ScopeDisplayName   = $scopeInfo.ScopeType
                        Type               = "AzureResource"
                        Status             = "Active"
                        MemberType         = $memberType
                        SubscriptionId     = $subscription.Id
                        SubscriptionName   = $subscription.Name
                        FullScope          = $assignment.Scope
                        ObjectId           = $assignment.ObjectId
                        ObjectType         = $assignment.ObjectType
                        StartDateTime      = $null
                        EndDateTime        = $null
                        Scope              = $scopeInfo.ScopeType
                        FormattedScope     = "$($scopeInfo.ResourceDisplay) ($($scopeInfo.ScopeType))"
                    }
                    
                    # Try enrich time window via schedules (best-effort)
                    try {
                        if (-not $roleObject.EndDateTime -and (Get-Command Get-AzRoleAssignmentSchedule -ErrorAction SilentlyContinue)) {
                            $roleDefGuid = $assignment.RoleDefinitionId
                            if ($roleDefGuid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefGuid = $matches[1] }
                            $scopePathPrefix = if ($assignment.Scope -match "^/providers/Microsoft\.Management/managementGroups/[^/]+$") { $assignment.Scope } else { "/subscriptions/$($subscription.Id)" }
                            $roleDefPath = "$scopePathPrefix/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"

                            $filter = "principalId eq '$UserObjectIdLocal' and roleDefinitionId eq '$roleDefPath' and status eq 'Active'"
                            $sched = Get-AzRoleAssignmentSchedule -Scope $assignment.Scope -Filter $filter -ErrorAction SilentlyContinue
                            $schedArr = @(); if ($sched) { $schedArr = ($sched -is [array]) ? $sched : @($sched) }
                            if ($schedArr.Count -gt 0) {
                                $selected = $schedArr | Sort-Object {
                                    $end = $null
                                    if ($_.PSObject.Properties["EndDateTime"]) { $end = $_.EndDateTime }
                                    elseif ($_.PSObject.Properties["ScheduleInfo"]) { $end = $_.ScheduleInfo.Expiration.EndDateTime }
                                    if (-not $end) { [datetime]::MinValue } else { [datetime]$end }
                                } -Descending | Select-Object -First 1

                                $startTime = $null; $endTime = $null
                                if ($selected.PSObject.Properties["StartDateTime"]) { $startTime = $selected.StartDateTime }
                                if ($selected.PSObject.Properties["EndDateTime"])   { $endTime   = $selected.EndDateTime }
                                if (-not $startTime -and $selected.PSObject.Properties["ScheduleInfo"]) { $startTime = $selected.ScheduleInfo.StartDateTime }
                                if (-not $endTime   -and $selected.PSObject.Properties["ScheduleInfo"]) { $endTime   = $selected.ScheduleInfo.Expiration.EndDateTime }

                                if ($startTime) { $roleObject.StartDateTime = $startTime }
                                if ($endTime)   { $roleObject.EndDateTime   = $endTime }
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Failed to enrich Azure active role time window: $($_.Exception.Message)"
                    }

                        Write-Verbose "[Parallel] Created active role object: $($roleObject.DisplayName) | Resource: $($roleObject.ResourceName) | Scope: $($roleObject.Scope) | MemberType: $($roleObject.MemberType) | FullScope: $($roleObject.FullScope)"
                        $localRoles.Add($roleObject) | Out-Null
                    }
                
                    # Process eligible assignments
                    if ($IncludeEligibleLocal) {
                        foreach ($eligibleAssignment in @($eligibleAssignments)) {
                    # Get role definition details
                    $roleDefinition = $null
                    $roleDefinitionName = "Unknown Role"
                    
                    # Handle different property names from different sources
                    $roleDefId = $eligibleAssignment.RoleDefinitionId ?? $eligibleAssignment.roleDefinitionId ?? $eligibleAssignment.properties.roleDefinitionId
                    
                    try {
                        if ($roleDefId) {
                            # Extract GUID if full path
                            if ($roleDefId -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
                                $roleDefId = $matches[1]
                            }
                            
                            $roleDefinition = Get-AzRoleDefinition -Id $roleDefId -ErrorAction SilentlyContinue
                            if ($roleDefinition) {
                                $roleDefinitionName = $roleDefinition.Name
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Could not retrieve role definition for ID: $roleDefId - $($_.Exception.Message)"
                    }
                    
                    # Determine scope for eligible assignment
                    $assignmentScope = $eligibleAssignment.Scope ?? 
                                       $eligibleAssignment.scope ?? 
                                       $eligibleAssignment.properties.scope ?? 
                                       $eligibleAssignment.DirectoryScopeId ?? 
                                       "/subscriptions/$($subscription.Id)"
                    
                    # Simplified scope info calculation for parallel processing
                    $scopeInfo = @{
                        ResourceDisplay = if ($assignmentScope -match "^/subscriptions/([^/]+)$") { $subscription.Name }
                                         elseif ($assignmentScope -match "^/subscriptions/([^/]+)/resourceGroups/([^/]+)$") { $matches[2] }
                                         elseif ($assignmentScope -eq "/" -or $assignmentScope -eq "") { "/" }
                                         elseif ($assignmentScope -match "^/providers/Microsoft\.Management/managementGroups/(.+)$") { $matches[1] }
                                         else { $assignmentScope }
                        ScopeType = if ($assignmentScope -match "^/subscriptions/([^/]+)$") { "Subscription" }
                                   elseif ($assignmentScope -match "^/subscriptions/([^/]+)/resourceGroups/([^/]+)$") { "Resource Group" }
                                   elseif ($assignmentScope -eq "/" -or $assignmentScope -eq "") { "Tenant" }
                                   elseif ($assignmentScope -match "^/providers/Microsoft\.Management/managementGroups/") { "Management group" }
                                   else { "Resource" }
                    }
                    
                    # Enhanced scope display logic
                    $scopeDisplayName = $scopeInfo.ScopeType
                    $resourceDisplayName = $scopeInfo.ResourceDisplay
                    
                    if ($assignmentScope -match "^/subscriptions/([^/]+)$") {
                        $scopeDisplayName = $subscription.Name
                        $resourceDisplayName = $subscription.Name
                    }
                    elseif ($assignmentScope -match "^/providers/Microsoft\.Management/managementGroups/(.+)$") {
                        $mgName = $matches[1]
                        try {
                            $mgInfo = Get-AzManagementGroup -GroupId $mgName -ErrorAction SilentlyContinue
                            if ($mgInfo -and $mgInfo.DisplayName) {
                                $scopeDisplayName = $mgInfo.DisplayName
                                $resourceDisplayName = $mgInfo.DisplayName
                            } else {
                                $scopeDisplayName = $mgName
                                $resourceDisplayName = $mgName
                            }
                        }
                        catch {
                            $scopeDisplayName = $mgName
                            $resourceDisplayName = $mgName
                        }
                    }
                    elseif ($assignmentScope -match "^/subscriptions/([^/]+)/resourceGroups/(.+)$") {
                        $scopeDisplayName = $subscription.Name
                        $resourceDisplayName = $subscription.Name
                    }
                    elseif ($assignmentScope -match "^/subscriptions/([^/]+)/") {
                        $scopeDisplayName = $subscription.Name  
                        $resourceDisplayName = $subscription.Name
                    }
                    
                        # Simplified member type detection for eligible assignments in parallel processing
                        $memberType = if ($assignmentScope -eq "/subscriptions/$($subscription.Id)" -or $assignmentScope -match "^/subscriptions/$($subscription.Id)/resourceGroups/") { 
                            "Direct" 
                        } else { 
                            "Inherited" 
                        }
                    
                    Write-Verbose "Final member type for eligible ${roleDefinitionName}: $memberType (Scope: $assignmentScope)"
                    Write-Verbose "Scope display will show: $scopeDisplayName"
                    
                    $roleObject = [PSCustomObject]@{
                        RoleId              = "$($eligibleAssignment.Id ?? $eligibleAssignment.id ?? $eligibleAssignment.name)-$($subscription.Id)"
                        RoleDefinitionId    = $roleDefId
                        DisplayName         = $roleDefinitionName
                        ResourceName        = $resourceDisplayName
                        ResourceDisplayName = $resourceDisplayName
                        ScopeDisplayName    = $scopeDisplayName
                        Type                = "AzureResource"
                        Status              = "Eligible"
                        MemberType          = $memberType
                        SubscriptionId      = $subscription.Id
                        SubscriptionName    = $subscription.Name
                        FullScope           = $assignmentScope
                        ObjectId            = $UserObjectId
                        StartDateTime       = $eligibleAssignment.StartDateTime ?? 
                                              $eligibleAssignment.properties.startDateTime ??
                                              $eligibleAssignment.ScheduleInfo.StartDateTime
                        EndDateTime         = $eligibleAssignment.EndDateTime ?? 
                                              $eligibleAssignment.properties.endDateTime ??
                                              $eligibleAssignment.ScheduleInfo.Expiration.EndDateTime
                        Scope               = $scopeDisplayName
                        FormattedScope      = $scopeDisplayName
                    }
                    
                            $localRoles.Add($roleObject) | Out-Null
                        }
                    }
                
                    
                    # Add all local roles to thread-safe collection
                    foreach ($role in $localRoles) {
                        $allSubscriptionRolesLocal.Add($role)
                    }
                    
                    # Track completion for progress reporting
                    $processedCountLocal.TryAdd($subscription.Id, $true) | Out-Null
                    $currentProgress = $processedCountLocal.Count
                    $totalSubs = $using:subscriptions.Count
                    
                    Write-Verbose "[Parallel] ✅ Completed subscription: $($subscription.Name) with $($localRoles.Count) roles ($currentProgress/$totalSubs completed)"
                }
                catch {
                    Write-Verbose "[Parallel] Failed to process subscription '$($subscription.Name)': $($_.Exception.Message)"
                }
            } -ThrottleLimit $ThrottleLimit
            
            # Convert thread-safe collection to regular array and add to main collection
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalSeconds
            Write-Verbose "Parallel processing completed in $([Math]::Round($duration, 1))s. Collected $($allSubscriptionRoles.Count) roles from $($subscriptions.Count) subscriptions"
            foreach ($role in $allSubscriptionRoles) {
                $allRoles.Add($role) | Out-Null
            }
        }
        else {
            # Use sequential processing (original logic)
            if ($DisableParallelProcessing) {
                Write-Verbose "Parallel processing disabled by user. Using sequential processing for $($subscriptions.Count) subscriptions"
            } else {
                Write-Verbose "Parallel processing requested but not supported (PowerShell $($PSVersionTable.PSVersion.Major) or single subscription). Using sequential processing."
            }
            
            $processed = 0
            foreach ($subscription in $subscriptions) {
                $processed++
                Write-Verbose "Processing subscription $processed/$($subscriptions.Count): $($subscription.Name) ($($subscription.Id))"
                
                try {
                    # Select subscription context with tenant scoping BEFORE any queries
                    Select-AzSubscription -SubscriptionId $subscription.Id -Tenant $currentTenantId -ErrorAction SilentlyContinue | Out-Null
                    
                    # Get active role assignments at subscription scope (optional)
                    $activeAssignments = @()
                    if ($IncludeActive) {
                        Write-Verbose "Fetching active role assignments for subscription: $($subscription.Name)"
                        try {
                            $activeAssignments = Get-AzRoleAssignment -Scope "/subscriptions/$($subscription.Id)" -ObjectId $UserObjectId -ErrorAction SilentlyContinue
                            if ($activeAssignments -and ($activeAssignments -isnot [array])) { $activeAssignments = @($activeAssignments) }
                        }
                        catch {
                            Write-Verbose "Active assignment query failed on $($subscription.Id): $($_.Exception.Message)"
                            $activeAssignments = @()
                        }

                        # REST fallback when Az cmdlet returns no results
                        if (-not $activeAssignments -or @($activeAssignments).Count -eq 0) {
                            Write-Verbose "No active assignments via Az cmdlet for $($subscription.Name); attempting REST fallback"
                            try {
                                $context = Get-AzContext -ErrorAction SilentlyContinue
                                if ($context -and $context.Account.Id) {
                                    $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
                                        $context.Account, $context.Environment, $context.Tenant.Id, $null, "Never", $null, "https://management.azure.com/"
                                    ).AccessToken
                                    $headers = @{
                                        'Authorization' = "Bearer $token"
                                        'Content-Type'  = 'application/json'
                                    }
                                    $uri = "https://management.azure.com/subscriptions/$($subscription.Id)/providers/Microsoft.Authorization/roleAssignments?api-version=2022-04-01&`$filter=principalId eq '$UserObjectId'"
                                    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                                    if ($response -and $response.value) {
                                        $restAssignments = @()
                                        foreach ($item in $response.value) {
                                            $roleDefPath = $item.properties.roleDefinitionId
                                            $roleDefGuid = $roleDefPath
                                            if ($roleDefGuid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefGuid = $matches[1] }
                                            $roleDef = $null; $roleName = "Unknown Role"
                                            try {
                                                $roleDef = Get-AzRoleDefinition -Id $roleDefGuid -ErrorAction SilentlyContinue
                                                if ($roleDef) { $roleName = $roleDef.Name }
                                            } catch { }

                                            $restAssignments += [PSCustomObject]@{
                                                RoleAssignmentId   = $item.name
                                                RoleDefinitionId   = $roleDefPath
                                                RoleDefinitionName = $roleName
                                                Scope              = $item.properties.scope
                                                ObjectId           = $item.properties.principalId
                                                ObjectType         = 'User'
                                            }
                                        }
                                        $activeAssignments = $restAssignments
                                        Write-Verbose "REST fallback returned $(@($activeAssignments).Count) active assignments for $($subscription.Name)"
                                    }
                                }
                            }
                            catch {
                                Write-Verbose "REST fallback for active assignments failed on $($subscription.Id): $($_.Exception.Message)"
                            }
                        }
                    }

                    # Do NOT use -IncludeInherited (parameter not available across Az versions). Skip inherited enrichment to prevent errors.

                    if ($IncludeActive) { Write-Verbose "Found $(@($activeAssignments).Count) active role assignments" }

                    # Get eligible role assignments (PIM) at subscription scope using principalId filter (optional)
                    $eligibleAssignments = @()
                    if ($IncludeEligible) {
                        Write-Verbose "Fetching eligible role assignments for subscription: $($subscription.Name)"
                        try {
                            if (Get-Command Get-AzRoleEligibilitySchedule -ErrorAction SilentlyContinue) {
                                $eligibleAssignments = Get-AzRoleEligibilitySchedule -Scope "/subscriptions/$($subscription.Id)" -Filter "principalId eq '$UserObjectId'" -ErrorAction SilentlyContinue
                                if ($eligibleAssignments -and ($eligibleAssignments -isnot [array])) { $eligibleAssignments = @($eligibleAssignments) }
                                Write-Verbose "Found $(@($eligibleAssignments).Count) eligible assignments via Az.Resources"
                            }
                            else {
                                Write-Verbose "Az.Resources PIM cmdlets not available, using REST API fallback"
                                $context = Get-AzContext -ErrorAction SilentlyContinue
                                if ($context -and $context.Account.Id) {
                                    $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
                                        $context.Account, $context.Environment, $context.Tenant.Id, $null, "Never", $null, "https://management.azure.com/"
                                    ).AccessToken
                                    $headers = @{
                                        'Authorization' = "Bearer $token"
                                        'Content-Type'  = 'application/json'
                                    }
                                    $uri = "https://management.azure.com/subscriptions/$($subscription.Id)/providers/Microsoft.Authorization/roleEligibilitySchedules?api-version=2020-10-01&`$filter=principalId eq '$UserObjectId'"
                                    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                                    if ($response -and $response.value) {
                                        $eligibleAssignments = @($response.value)
                                        Write-Verbose "Found $(@($eligibleAssignments).Count) eligible assignments via REST API"
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Could not retrieve eligible assignments: $($_.Exception.Message)"
                            $eligibleAssignments = @()
                        }
                    }
                    
                    # Process active assignments
                    foreach ($assignment in @($activeAssignments)) {
                        # Skip tenant-root scoped entries ("/") that are duplicated on every subscription
                        if ($assignment.Scope -eq "/") {
                            Write-Verbose "Skipping duplicate tenant root active assignment '$($assignment.RoleDefinitionName)' surfaced in subscription '$($subscription.Name)'"
                            continue
                        }
                        Write-Verbose "Processing active assignment: $($assignment.RoleDefinitionName) at scope $($assignment.Scope)"
                        
                        # Determine scope type and display names
                        $scopeInfo = Get-AzureScopeInfo -Scope $assignment.Scope
                        
                        # Member type detection
                        $memberType = Get-AzureMemberType -AssignmentScope $assignment.Scope -CurrentSubscriptionId $subscription.Id -PrincipalType $assignment.ObjectType -IsEligible $false -ObjectId $assignment.ObjectId
                        
                        # If assignment scope doesn't match current subscription scope, mark as inherited
                        if ($assignment.Scope -ne "/subscriptions/$($subscription.Id)" -and 
                            $assignment.Scope -notmatch "^/subscriptions/$($subscription.Id)/resourceGroups/" -and
                            $assignment.Scope -notmatch "^/subscriptions/$($subscription.Id)/resourceGroups/.+") {
                            Write-Verbose "Assignment scope $($assignment.Scope) indicates inheritance -> Inherited"
                            $memberType = "Inherited"
                        }
                        
                        $roleObject = [PSCustomObject]@{
                            RoleId             = "$($assignment.RoleAssignmentId)-$($subscription.Id)"
                            RoleDefinitionId   = $assignment.RoleDefinitionId
                            DisplayName        = $assignment.RoleDefinitionName
                            ResourceName       = $scopeInfo.ResourceDisplay
                            ResourceDisplayName= $scopeInfo.ResourceDisplay
                            ScopeDisplayName   = $scopeInfo.ScopeType
                            Type               = "AzureResource"
                            Status             = "Active"
                            MemberType         = $memberType
                            SubscriptionId     = $subscription.Id
                            SubscriptionName   = $subscription.Name
                            FullScope          = $assignment.Scope
                            ObjectId           = $assignment.ObjectId
                            ObjectType         = $assignment.ObjectType
                            StartDateTime      = $null
                            EndDateTime        = $null
                            Scope              = $scopeInfo.ScopeType
                            FormattedScope     = "$($scopeInfo.ResourceDisplay) ($($scopeInfo.ScopeType))"
                        }
                        
                        # Try enrich time window via schedules (best-effort)
                        try {
                            if (-not $roleObject.EndDateTime -and (Get-Command Get-AzRoleAssignmentSchedule -ErrorAction SilentlyContinue)) {
                                $roleDefGuid = $assignment.RoleDefinitionId
                                if ($roleDefGuid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefGuid = $matches[1] }
                                $scopePathPrefix = if ($assignment.Scope -match "^/providers/Microsoft\.Management/managementGroups/[^/]+$") { $assignment.Scope } else { "/subscriptions/$($subscription.Id)" }
                                $roleDefPath = "$scopePathPrefix/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"

                                $filter = "principalId eq '$UserObjectId' and roleDefinitionId eq '$roleDefPath' and status eq 'Active'"
                                $sched = Get-AzRoleAssignmentSchedule -Scope $assignment.Scope -Filter $filter -ErrorAction SilentlyContinue
                                $schedArr = @(); if ($sched) { $schedArr = ($sched -is [array]) ? $sched : @($sched) }
                                if ($schedArr.Count -gt 0) {
                                    $selected = $schedArr | Sort-Object {
                                        $end = $null
                                        if ($_.PSObject.Properties["EndDateTime"]) { $end = $_.EndDateTime }
                                        elseif ($_.PSObject.Properties["ScheduleInfo"]) { $end = $_.ScheduleInfo.Expiration.EndDateTime }
                                        if (-not $end) { [datetime]::MinValue } else { [datetime]$end }
                                    } -Descending | Select-Object -First 1

                                    $startTime = $null; $endTime = $null
                                    if ($selected.PSObject.Properties["StartDateTime"]) { $startTime = $selected.StartDateTime }
                                    if ($selected.PSObject.Properties["EndDateTime"])   { $endTime   = $selected.EndDateTime }
                                    if (-not $startTime -and $selected.PSObject.Properties["ScheduleInfo"]) { $startTime = $selected.ScheduleInfo.StartDateTime }
                                    if (-not $endTime   -and $selected.PSObject.Properties["ScheduleInfo"]) { $endTime   = $selected.ScheduleInfo.Expiration.EndDateTime }

                                    if ($startTime) { $roleObject.StartDateTime = $startTime }
                                    if ($endTime)   { $roleObject.EndDateTime   = $endTime }
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Failed to enrich Azure active role time window: $($_.Exception.Message)"
                        }

                        Write-Verbose "Created active role object: $($roleObject.DisplayName) | Resource: $($roleObject.ResourceName) | Scope: $($roleObject.Scope) | MemberType: $($roleObject.MemberType) | FullScope: $($roleObject.FullScope)"
                        $allRoles.Add($roleObject) | Out-Null
                    }
                    
                    # Process eligible assignments
                    if ($IncludeEligible) {
                        foreach ($eligibleAssignment in @($eligibleAssignments)) {
                        # Get role definition details
                        $roleDefinition = $null
                        $roleDefinitionName = "Unknown Role"
                        
                        # Handle different property names from different sources
                        $roleDefId = $eligibleAssignment.RoleDefinitionId ?? $eligibleAssignment.roleDefinitionId ?? $eligibleAssignment.properties.roleDefinitionId
                        
                        try {
                            if ($roleDefId) {
                                # Extract GUID if full path
                                if ($roleDefId -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
                                    $roleDefId = $matches[1]
                                }
                                
                                $roleDefinition = Get-AzRoleDefinition -Id $roleDefId -ErrorAction SilentlyContinue
                                if ($roleDefinition) {
                                    $roleDefinitionName = $roleDefinition.Name
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Could not retrieve role definition for ID: $roleDefId - $($_.Exception.Message)"
                        }
                        
                        # Determine scope for eligible assignment
                        $assignmentScope = $eligibleAssignment.Scope ?? 
                                           $eligibleAssignment.scope ?? 
                                           $eligibleAssignment.properties.scope ?? 
                                           $eligibleAssignment.DirectoryScopeId ?? 
                                           "/subscriptions/$($subscription.Id)"
                        
                        $scopeInfo = Get-AzureScopeInfo -Scope $assignmentScope
                        
                        # Enhanced scope display logic
                        $scopeDisplayName = $scopeInfo.ScopeType
                        $resourceDisplayName = $scopeInfo.ResourceDisplay
                        
                        if ($assignmentScope -match "^/subscriptions/([^/]+)$") {
                            $scopeDisplayName = $subscription.Name
                            $resourceDisplayName = $subscription.Name
                        }
                        elseif ($assignmentScope -match "^/providers/Microsoft\.Management/managementGroups/(.+)$") {
                            $mgName = $matches[1]
                            try {
                                $mgInfo = Get-AzManagementGroup -GroupId $mgName -ErrorAction SilentlyContinue
                                if ($mgInfo -and $mgInfo.DisplayName) {
                                    $scopeDisplayName = $mgInfo.DisplayName
                                    $resourceDisplayName = $mgInfo.DisplayName
                                } else {
                                    $scopeDisplayName = $mgName
                                    $resourceDisplayName = $mgName
                                }
                            }
                            catch {
                                $scopeDisplayName = $mgName
                                $resourceDisplayName = $mgName
                            }
                        }
                        elseif ($assignmentScope -match "^/subscriptions/([^/]+)/resourceGroups/(.+)$") {
                            $scopeDisplayName = $subscription.Name
                            $resourceDisplayName = $subscription.Name
                        }
                        elseif ($assignmentScope -match "^/subscriptions/([^/]+)/") {
                            $scopeDisplayName = $subscription.Name  
                            $resourceDisplayName = $subscription.Name
                        }
                        
                        # Member type detection for eligible assignments
                        $memberType = Get-AzureMemberType -AssignmentScope $assignmentScope -CurrentSubscriptionId $subscription.Id -PrincipalType "User" -IsEligible $true -ObjectId $UserObjectId
                        
                        # Inheritance check
                        if ($assignmentScope -ne "/subscriptions/$($subscription.Id)" -and 
                            $assignmentScope -notmatch "^/subscriptions/$($subscription.Id)/resourceGroups/" -and
                            $assignmentScope -notmatch "^/subscriptions/$($subscription.Id)/resourceGroups/.+") {
                            Write-Verbose "Eligible assignment scope $assignmentScope indicates inheritance -> Inherited"
                            $memberType = "Inherited"
                        }
                        
                        Write-Verbose "Final member type for eligible ${roleDefinitionName}: $memberType (Scope: $assignmentScope)"
                        Write-Verbose "Scope display will show: $scopeDisplayName"
                        
                        $roleObject = [PSCustomObject]@{
                            RoleId              = "$($eligibleAssignment.Id ?? $eligibleAssignment.id ?? $eligibleAssignment.name)-$($subscription.Id)"
                            RoleDefinitionId    = $roleDefId
                            DisplayName         = $roleDefinitionName
                            ResourceName        = $resourceDisplayName
                            ResourceDisplayName = $resourceDisplayName
                            ScopeDisplayName    = $scopeDisplayName
                            Type                = "AzureResource"
                            Status              = "Eligible"
                            MemberType          = $memberType
                            SubscriptionId      = $subscription.Id
                            SubscriptionName    = $subscription.Name
                            FullScope           = $assignmentScope
                            ObjectId            = $UserObjectId
                            StartDateTime       = $eligibleAssignment.StartDateTime ?? 
                                                  $eligibleAssignment.properties.startDateTime ??
                                                  $eligibleAssignment.ScheduleInfo.StartDateTime
                            EndDateTime         = $eligibleAssignment.EndDateTime ?? 
                                                  $eligibleAssignment.properties.endDateTime ??
                                                  $eligibleAssignment.ScheduleInfo.Expiration.EndDateTime
                            Scope               = $scopeDisplayName
                            FormattedScope      = $scopeDisplayName
                        }
                        
                            $allRoles.Add($roleObject) | Out-Null
                        }
                    }
                    
                }
                catch {
                    Write-Warning "Failed to process subscription '$($subscription.Name)': $($_.Exception.Message)"
                    continue
                }
            }
        }
        }
        
        Write-Verbose "Found $($allRoles.Count) Azure resource role assignments across $($subscriptions.Count) subscriptions"
        
        # Debug: Show details of all roles before deduplication
        if ($allRoles.Count -gt 0) {
            Write-Verbose "Roles before deduplication:"
            for ($i = 0; $i -lt $allRoles.Count; $i++) {
                $r = $allRoles[$i]
                Write-Verbose "  [$i] $($r.DisplayName) | Status: $($r.Status) | Scope: $($r.FullScope) | Type: $($r.Type)"
            }
        }
        
        # Deduplicate roles based on RoleDefinitionId + Scope + Status
        $uniqueRoles = [System.Collections.ArrayList]::new()
        $seenRoles = [System.Collections.Generic.HashSet[string]]::new()

        try {
            foreach ($role in $allRoles) {
                # Create a more specific unique key that includes subscription ID for subscription-scoped roles
                $scopeKey = $role.FullScope
                if ($role.SubscriptionId -and $role.FullScope -match "^/subscriptions/") {
                    $scopeKey = "$(($role.FullScope))_sub:$(($role.SubscriptionId))"
                }

                $uniqueKey = "$(($role.RoleDefinitionId))_$scopeKey_$(($role.Status))"
                Write-Verbose "Processing role: $($role.DisplayName) with key: $uniqueKey"

                if (-not $seenRoles.Contains($uniqueKey)) {
                    $seenRoles.Add($uniqueKey) | Out-Null
                    $uniqueRoles.Add($role) | Out-Null
                    Write-Verbose "  Added unique role: $($role.DisplayName)"
                }
                else {
                    Write-Verbose "  Skipping duplicate role: $($role.DisplayName) at scope $($role.FullScope) with status $($role.Status) (key: $uniqueKey)"
                }
            }

            Write-Verbose "After deduplication: $($uniqueRoles.Count) unique roles (removed $($allRoles.Count - $uniqueRoles.Count) duplicates)"

            $activeCount = ($uniqueRoles | Where-Object { $_.Status -eq 'Active' }).Count
            $eligibleCount = ($uniqueRoles | Where-Object { $_.Status -eq 'Eligible' }).Count
            Write-Verbose "Breakdown: $activeCount active roles, $eligibleCount eligible roles"

            if ($uniqueRoles.Count -eq 0) {
                Write-Warning "No Azure Resource role assignments found after deduplication. Returning pre-deduped results to avoid upstream null/empty handling issues."
                $uniqueRoles = $allRoles
            }
        }
        catch {
            Write-Warning "Deduplication failed: $($_.Exception.Message). Returning pre-deduped role list."
            $uniqueRoles = $allRoles
        }

        # Always return a concrete array (never null) to upstream callers
        $resultRoles = @()
        try {
            if ($uniqueRoles -is [System.Collections.IEnumerable]) {
                $resultRoles = $uniqueRoles.ToArray()
            } else {
                # Fallback: wrap any single role object
                if ($uniqueRoles) { $resultRoles = @($uniqueRoles) } else { $resultRoles = @() }
            }
        } catch {
            Write-Verbose "Failed to materialize Azure roles array: $($_.Exception.Message)"
            $resultRoles = @()
        }

        Write-Verbose "Returning $($resultRoles.Count) Azure resource roles to caller"
        return $resultRoles
    }
    catch {
        # Return empty array (not null) on failure to keep upstream logic stable
        $msg = $_.Exception.Message
        Write-Warning "Failed to retrieve Azure resource roles: $msg"
        return @()
    }
}