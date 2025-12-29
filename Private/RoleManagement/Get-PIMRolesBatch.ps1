function Get-PIMRolesBatch {
    <#
    .SYNOPSIS
        Retrieves all PIM roles in batch operations for enhanced performance.
    
    .DESCRIPTION
        Fetches eligible and active roles with their associated data in minimal API calls.
        Includes role definitions, policies, and authentication contexts. This function
        significantly improves performance by batching API requests instead of individual calls.
    
    .PARAMETER UserId
        The ID of the user to fetch roles for.
    
    .PARAMETER IncludeEntraRoles
        Switch to include Entra ID directory roles in the batch fetch.
    
    .PARAMETER IncludeGroups
        Switch to include PIM-enabled security groups in the batch fetch.
    
    .PARAMETER IncludeAzureResources
        Switch to include Azure resource roles in the batch fetch.

    .PARAMETER SplashForm
        Optional splash screen form to update during batch operations.
        Used to display granular progress information during role data retrieval.

    .PARAMETER SkipAzureRefresh
        When set, reuses cached Azure resource roles instead of re-fetching them.
        Useful for post-activation refreshes where Azure RBAC doesn’t change.

    .EXAMPLE
        $result = Get-PIMRolesBatch -UserId $user.Id -IncludeEntraRoles -IncludeGroups
        Fetches all Entra roles and groups for the user in batch operations.

    .EXAMPLE
        $result = Get-PIMRolesBatch -UserId $user.Id -IncludeEntraRoles -IncludeGroups -SplashForm $splash
        Fetches roles with progress updates to the splash screen.

    .OUTPUTS
        PSCustomObject
        Returns an object with the following properties:
        - EligibleRoles: Array of eligible role objects
        - ActiveRoles: Array of active role objects
        - RoleDefinitions: Hashtable of cached role definitions
        - Policies: Hashtable of cached policy information
        - AuthenticationContexts: Hashtable of cached authentication contexts
    
    .NOTES
        This function uses batch API operations to minimize the number of Graph API calls,
        significantly improving performance for users with many PIM-eligible roles.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,
        [switch]$IncludeEntraRoles,
        [switch]$IncludeGroups,
        [switch]$IncludeAzureResources,
        [PSCustomObject]$SplashForm,
        [switch]$SkipAzureRefresh,
        [switch]$ActiveOnly,
        
        # Enable parallel processing for Azure subscriptions (requires PowerShell 7+)
        [switch]$DisableParallelProcessing,
        
        # Maximum concurrent operations for parallel processing
        [int]$ThrottleLimit = 10
    )
    
    Write-Verbose "Starting batch PIM role fetch for user: $UserId (ActiveOnly: $ActiveOnly)"
    
    # Ensure Azure cache exists
    if (-not (Get-Variable -Name 'AzureRolesCache' -Scope Script -ErrorAction SilentlyContinue)) { $script:AzureRolesCache = @() }
    if (-not (Get-Variable -Name 'AzureRolesCacheTime' -Scope Script -ErrorAction SilentlyContinue)) { $script:AzureRolesCacheTime = $null }

    $result = [PSCustomObject]@{
        EligibleRoles          = [System.Collections.ArrayList]@()
        ActiveRoles            = [System.Collections.ArrayList]@()
        RoleDefinitions        = @{}
        Policies               = @{}
        AuthenticationContexts = @{}
    }
    
    # Batch fetch Entra ID roles if requested
    if ($IncludeEntraRoles) {
        Write-Verbose "Batch fetching Entra ID roles..."
            
        # Update progress: Starting Entra roles fetch
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            $SplashForm.UpdateStatus("Fetching Entra ID active roles...", 86)
        }
            
        try {
            # Get eligible assignments only when not ActiveOnly
            $eligibleAssignments = @()
            if (-not $ActiveOnly) {
                $eligibleParams = @{
                    Filter         = "principalId eq '$UserId'"
                    ExpandProperty = 'roleDefinition'
                    Select         = 'id,principalId,roleDefinitionId,directoryScopeId,startDateTime,endDateTime,memberType,roleDefinition'
                    All            = $true
                }
                $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance @eligibleParams
            }
                
            # Update progress: Active roles fetched, starting eligible
            if (-not $ActiveOnly) {
                if ($SplashForm -and -not $SplashForm.IsDisposed) {
                    $SplashForm.UpdateStatus("Fetching Entra ID eligible roles...", 88)
                }
            }
                
            # Ensure we have arrays to work with
            if (-not $eligibleAssignments) {
                $eligibleAssignments = @()
            }
            elseif ($eligibleAssignments -isnot [array]) {
                $eligibleAssignments = @($eligibleAssignments)
            }
            if (-not $ActiveOnly) { Write-Verbose "Found $($eligibleAssignments.Count) eligible Entra assignments" }
                
            # Get all active assignments with expanded properties in one call
            $activeParams = @{
                Filter         = "principalId eq '$UserId'"
                ExpandProperty = 'roleDefinition'
                Select         = 'id,principalId,roleDefinitionId,directoryScopeId,startDateTime,endDateTime,memberType,roleDefinition'
                All            = $true
            }
            $activeAssignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance @activeParams
                
            # Ensure we have arrays to work with
            if (-not $activeAssignments) {
                $activeAssignments = @()
            }
            elseif ($activeAssignments -isnot [array]) {
                $activeAssignments = @($activeAssignments)
            }
            Write-Verbose "Found $($activeAssignments.Count) active Entra assignments"
                
            # Update progress: Processing Entra roles
            if ($SplashForm -and -not $SplashForm.IsDisposed) {
                $SplashForm.UpdateStatus("Processing Entra ID role assignments...", 89)
            }
                
            # Process eligible roles
            if (-not $ActiveOnly) {
                Write-Verbose "Processing $($eligibleAssignments.Count) eligible Entra assignments..."
                foreach ($assignment in $eligibleAssignments) {
                    try {
                        Write-Verbose "Processing eligible assignment: $($assignment.RoleDefinition.DisplayName)"
                        
                        $scopeDisplayName = $null
                        try { 
                            $scopeDisplayName = Get-ScopeDisplayName -Scope $assignment.DirectoryScopeId 
                        } 
                        catch { 
                            Write-Warning "Failed to get scope display name for $($assignment.RoleDefinition.DisplayName): $_"
                            $scopeDisplayName = "Directory" 
                        }
                        
                        # Check if this is a group-derived role by examining MemberType
                        $isGroupDerived = $assignment.MemberType -eq 'Inherited'
                        
                        $roleObj = [PSCustomObject]@{
                            Type             = 'Entra'
                            DisplayName      = $assignment.RoleDefinition.DisplayName
                            Status           = 'Eligible'
                            Assignment       = $assignment
                            DirectoryScopeId = $assignment.DirectoryScopeId
                            MemberType       = $assignment.MemberType
                            RoleDefinitionId = $assignment.RoleDefinitionId
                            ResourceName     = "Entra ID Directory"  # Default resource name for Entra roles
                            Scope            = $scopeDisplayName
                            Id               = $assignment.RoleDefinitionId
                            IsGroupDerived   = $isGroupDerived
                        }
                        
                        Write-Verbose "Created role object for: $($roleObj.DisplayName)"
                        
                        try {
                            $result.EligibleRoles.Add($roleObj) | Out-Null
                            Write-Verbose "Added eligible role: $($roleObj.DisplayName)"
                        }
                        catch {
                            Write-Warning "Failed to add eligible role to result: $_"
                        }
                        
                        # Cache role definition
                        if (-not $result.RoleDefinitions.ContainsKey($assignment.RoleDefinitionId)) {
                            $result.RoleDefinitions[$assignment.RoleDefinitionId] = $assignment.RoleDefinition
                        }
                    }
                    catch {
                        Write-Warning "Failed to process eligible assignment '$($assignment.RoleDefinition.DisplayName)': $_"
                    }
                }
            }
                
            # Process active roles (ALWAYS - moved outside the ActiveOnly check)
            Write-Verbose "Processing $($activeAssignments.Count) active Entra assignments..."
            foreach ($assignment in $activeAssignments) {
                try {
                    Write-Verbose "Processing active assignment: $($assignment.RoleDefinition.DisplayName)"
                        
                    $scopeDisplayName = $null
                    try { 
                        $scopeDisplayName = Get-ScopeDisplayName -Scope $assignment.DirectoryScopeId 
                    } 
                    catch { 
                        Write-Warning "Failed to get scope display name for $($assignment.RoleDefinition.DisplayName): $_"
                        $scopeDisplayName = "Directory" 
                    }
                        
                    # Check if this is a group-derived role by examining MemberType
                    Write-Verbose "Assignment MemberType for $($assignment.RoleDefinition.DisplayName): '$($assignment.MemberType)'"
                    $isGroupDerived = $assignment.MemberType -eq 'Group'
                        
                    # Safely get EndDateTime from assignment
                    $effectiveEndDateTime = $null
                    try {
                        if ($assignment.PSObject.Properties['EndDateTime']) {
                            $effectiveEndDateTime = $assignment.EndDateTime
                        }
                    }
                    catch {
                        Write-Verbose "EndDateTime property not available for Entra role $($assignment.RoleDefinition.DisplayName)"
                    }
                        
                    # If this appears to be a group-derived role and has no specific end time, 
                    # we'll try to match it with active group memberships later
                    $roleObj = [PSCustomObject]@{
                        Type             = 'Entra'
                        DisplayName      = $assignment.RoleDefinition.DisplayName
                        Status           = 'Active'
                        Assignment       = $assignment
                        DirectoryScopeId = $assignment.DirectoryScopeId
                        EndDateTime      = $effectiveEndDateTime
                        MemberType       = $assignment.MemberType
                        RoleDefinitionId = $assignment.RoleDefinitionId
                        ResourceName     = "Entra ID Directory"  # Default resource name for Entra roles
                        Scope            = $scopeDisplayName
                        Id               = $assignment.RoleDefinitionId
                        ScheduleId       = $assignment.Id
                        IsGroupDerived   = $isGroupDerived
                    }
                        
                    Write-Verbose "Created active role object for: $($roleObj.DisplayName) (Group-derived: $isGroupDerived)"
                        
                    try {
                        $result.ActiveRoles.Add($roleObj) | Out-Null
                        Write-Verbose "Added active role: $($roleObj.DisplayName)"
                    }
                    catch {
                        Write-Warning "Failed to add active role to result: $_"
                    }
                }
                catch {
                    Write-Warning "Failed to process active assignment '$($assignment.RoleDefinition.DisplayName)': $_"
                }
            }
        }
        catch {
            Write-Warning "Failed to batch fetch Entra roles: $_"
            # Continue with other role types
        }
    }
        
    # Batch fetch PIM Groups if requested
    if ($IncludeGroups) {
        Write-Verbose "Batch fetching PIM group memberships..."
        
        # Update progress: Starting group fetch
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            $SplashForm.UpdateStatus("Fetching PIM group memberships...", 90)
        }
        
        try {
            # Get eligible group memberships only when not ActiveOnly
            $eligibleGroups = @()
            if (-not $ActiveOnly) {
                $eligibleGroupParams = @{
                    Filter         = "principalId eq '$UserId'"
                    ExpandProperty = 'group'
                    Select         = 'id,principalId,groupId,accessId,startDateTime,endDateTime,group'
                    All            = $true
                }
                $eligibleGroups = Get-MgIdentityGovernancePrivilegedAccessGroupEligibilityScheduleInstance @eligibleGroupParams
            }
            
            # Ensure we have arrays to work with
            if (-not $eligibleGroups) {
                $eligibleGroups = @()
            }
            elseif ($eligibleGroups -isnot [array]) {
                $eligibleGroups = @($eligibleGroups)
            }
            if (-not $ActiveOnly) { Write-Verbose "Found $($eligibleGroups.Count) eligible group memberships" }
            
            $activeGroupParams = @{
                Filter         = "principalId eq '$UserId'"
                ExpandProperty = 'group'
                Select         = 'id,principalId,groupId,accessId,startDateTime,endDateTime,group'
                All            = $true
            }
            $activeGroups = Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleInstance @activeGroupParams
            
            # Ensure we have arrays to work with
            if (-not $activeGroups) {
                $activeGroups = @()
            }
            elseif ($activeGroups -isnot [array]) {
                $activeGroups = @($activeGroups)
            }
            Write-Verbose "Found $($activeGroups.Count) active group memberships"
            
            # Update progress: Processing group memberships
            if ($SplashForm -and -not $SplashForm.IsDisposed) {
                $SplashForm.UpdateStatus("Processing PIM group assignments...", 91)
            }
            
            # Process eligible group memberships
            if (-not $ActiveOnly) {
                Write-Verbose "Processing $($eligibleGroups.Count) eligible group memberships..."
                foreach ($membership in $eligibleGroups) {
                    try {
                        Write-Verbose "Processing eligible group: $($membership.Group.DisplayName)"
                    
                        try {
                            $memberType = Get-MembershipType -Assignment $membership -RoleType 'Group'
                        }
                        catch {
                            Write-Warning "Failed to get membership type for $($membership.Group.DisplayName): $_"
                            $memberType = "Member"
                        }
                    
                        # Determine group scope based on role-assignable status and AU membership
                        $groupScope = "Directory"  # Default scope for groups
                        try {
                            # Check if group is role-assignable (indicates directory-level scope)
                            if ($membership.Group.PSObject.Properties['IsAssignableToRole'] -and $membership.Group.IsAssignableToRole) {
                                $groupScope = "Directory"
                            }
                            # Check if group is member of any administrative units
                            elseif ($membership.GroupId) {
                                try {
                                    $groupDetails = Get-MgGroup -GroupId $membership.GroupId -Select "displayName,memberOf" -ErrorAction SilentlyContinue
                                    if ($groupDetails -and $groupDetails.MemberOf) {
                                        # Check if any memberOf entries are administrative units
                                        foreach ($memberOfId in $groupDetails.MemberOf) {
                                            try {
                                                $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $memberOfId -ErrorAction SilentlyContinue
                                                if ($au) {
                                                    $groupScope = "AU: $($au.DisplayName)"
                                                    break
                                                }
                                            }
                                            catch {
                                                # Not an AU, continue checking
                                                continue
                                            }
                                        }
                                    }
                                }
                                catch {
                                    Write-Verbose "Failed to check group administrative unit membership for $($membership.Group.DisplayName): $_"
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Failed to determine group scope for $($membership.Group.DisplayName): $_"
                        }
                    
                        # Fetch roles provided by this group membership
                        $providedRoles = [System.Collections.ArrayList]::new()
                        try {
                            Write-Verbose "Fetching roles provided by group: $($membership.Group.DisplayName)"
                            $groupRoleAssignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($membership.GroupId)'" -ExpandProperty "roleDefinition" -Select "id,principalId,roleDefinitionId,directoryScopeId,roleDefinition" -ErrorAction SilentlyContinue
                        
                            if ($groupRoleAssignments) {
                                # Ensure we have an array
                                if ($groupRoleAssignments -isnot [array]) {
                                    $groupRoleAssignments = @($groupRoleAssignments)
                                }
                            
                                foreach ($roleAssignment in $groupRoleAssignments) {
                                    try {
                                        $scopeDisplayName = $null
                                        try { $scopeDisplayName = Get-ScopeDisplayName -Scope $roleAssignment.DirectoryScopeId } catch { $scopeDisplayName = "Directory" }
                                        $providedRole = [PSCustomObject]@{
                                            RoleDefinitionId = $roleAssignment.RoleDefinitionId
                                            DisplayName      = $roleAssignment.RoleDefinition.DisplayName
                                            Description      = $roleAssignment.RoleDefinition.Description
                                            DirectoryScopeId = $roleAssignment.DirectoryScopeId
                                            ScopeDisplayName = $scopeDisplayName
                                        }
                                        $null = $providedRoles.Add($providedRole)
                                        Write-Verbose "Group $($membership.Group.DisplayName) provides role: $($providedRole.DisplayName)"
                                    }
                                    catch {
                                        Write-Warning "Failed to process role assignment for group $($membership.Group.DisplayName): $_"
                                    }
                                }
                                Write-Verbose "Group $($membership.Group.DisplayName) provides $($providedRoles.Count) roles"
                            }
                        }
                        catch {
                            Write-Warning "Failed to fetch roles provided by group $($membership.Group.DisplayName): $_"
                        }
                    
                        $roleObj = [PSCustomObject]@{
                            Type          = 'Group'
                            DisplayName   = $membership.Group.DisplayName
                            Status        = 'Eligible'
                            Assignment    = $membership
                            GroupId       = $membership.GroupId
                            MemberType    = $memberType
                            ResourceName  = $membership.Group.DisplayName
                            Scope         = $groupScope
                            Id            = $membership.GroupId
                            AccessId      = $membership.AccessId
                            ProvidedRoles = $providedRoles
                        }
                    
                        Write-Verbose "Created eligible group object for: $($roleObj.DisplayName)"
                    
                        try {
                            $result.EligibleRoles.Add($roleObj) | Out-Null
                            Write-Verbose "Added eligible group: $($roleObj.DisplayName)"
                        }
                        catch {
                            Write-Warning "Failed to add eligible group to result: $_"
                        }
                    }
                    catch {
                        Write-Warning "Failed to process eligible group '$($membership.Group.DisplayName)': $_"
                    }
                }
            }
        
            # Process active group memberships (ALWAYS - moved outside ActiveOnly check)
            Write-Verbose "Processing $($activeGroups.Count) active group memberships..."
            foreach ($membership in $activeGroups) {
                try {
                    Write-Verbose "Processing active group: $($membership.Group.DisplayName)"
                
                    try {
                        $memberType = Get-MembershipType -Assignment $membership -RoleType 'Group'
                    }
                    catch {
                        Write-Warning "Failed to get membership type for $($membership.Group.DisplayName): $_"
                        $memberType = "Member"
                    }
                
                    # Determine group scope based on role-assignable status and AU membership
                    $groupScope = "Directory"  # Default scope for groups
                    try {
                        # Check if group is role-assignable (indicates directory-level scope)
                        if ($membership.Group.PSObject.Properties['IsAssignableToRole'] -and $membership.Group.IsAssignableToRole) {
                            $groupScope = "Directory"
                        }
                        # Check if group is member of any administrative units
                        elseif ($membership.GroupId) {
                            try {
                                $groupDetails = Get-MgGroup -GroupId $membership.GroupId -Select "displayName,memberOf" -ErrorAction SilentlyContinue
                                if ($groupDetails -and $groupDetails.MemberOf) {
                                    # Check if any memberOf entries are administrative units
                                    foreach ($memberOfId in $groupDetails.MemberOf) {
                                        try {
                                            $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $memberOfId -ErrorAction SilentlyContinue
                                            if ($au) {
                                                $groupScope = "AU: $($au.DisplayName)"
                                                break
                                            }
                                        }
                                        catch {
                                            # Not an AU, continue checking
                                            continue
                                        }
                                    }
                                }
                            }
                            catch {
                                Write-Verbose "Failed to check group administrative unit membership for $($membership.Group.DisplayName): $_"
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Failed to determine group scope for $($membership.Group.DisplayName): $_"
                    }
                
                    # Fetch roles provided by this group membership
                    $providedRoles = [System.Collections.ArrayList]::new()
                    try {
                        Write-Verbose "Fetching roles provided by active group: $($membership.Group.DisplayName)"
                        $groupRoleAssignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($membership.GroupId)'" -ExpandProperty "roleDefinition" -Select "id,principalId,roleDefinitionId,directoryScopeId,roleDefinition" -ErrorAction SilentlyContinue
                    
                        if ($groupRoleAssignments) {
                            # Ensure we have an array
                            if ($groupRoleAssignments -isnot [array]) {
                                $groupRoleAssignments = @($groupRoleAssignments)
                            }
                        
                            foreach ($roleAssignment in $groupRoleAssignments) {
                                try {
                                    $scopeDisplayName = $null
                                    try { $scopeDisplayName = Get-ScopeDisplayName -Scope $roleAssignment.DirectoryScopeId } catch { $scopeDisplayName = "Directory" }
                                    $providedRole = [PSCustomObject]@{
                                        RoleDefinitionId = $roleAssignment.RoleDefinitionId
                                        DisplayName      = $roleAssignment.RoleDefinition.DisplayName
                                        Description      = $roleAssignment.RoleDefinition.Description
                                        DirectoryScopeId = $roleAssignment.DirectoryScopeId
                                        ScopeDisplayName = $scopeDisplayName
                                    }
                                    $null = $providedRoles.Add($providedRole)
                                    Write-Verbose "Active group $($membership.Group.DisplayName) provides role: $($providedRole.DisplayName)"
                                }
                                catch {
                                    Write-Warning "Failed to process role assignment for active group $($membership.Group.DisplayName): $_"
                                }
                            }
                            Write-Verbose "Active group $($membership.Group.DisplayName) provides $($providedRoles.Count) roles"
                        }
                    }
                    catch {
                        Write-Warning "Failed to fetch roles provided by active group $($membership.Group.DisplayName): $_"
                    }
                
                    # Safely get EndDateTime from membership
                    $membershipEndDateTime = $null
                    try {
                        if ($membership.PSObject.Properties['EndDateTime']) {
                            $membershipEndDateTime = $membership.EndDateTime
                        }
                    }
                    catch {
                        Write-Verbose "EndDateTime property not available for group $($membership.Group.DisplayName)"
                    }
                
                    $roleObj = [PSCustomObject]@{
                        Type          = 'Group'
                        DisplayName   = $membership.Group.DisplayName
                        Status        = 'Active'
                        Assignment    = $membership
                        GroupId       = $membership.GroupId
                        EndDateTime   = $membershipEndDateTime
                        MemberType    = $memberType
                        ResourceName  = $membership.Group.DisplayName
                        Scope         = $groupScope
                        Id            = $membership.GroupId
                        ScheduleId    = $membership.Id
                        AccessId      = $membership.AccessId
                        ProvidedRoles = $providedRoles
                    }
                
                    Write-Verbose "Created active group object for: $($roleObj.DisplayName)"
                
                    try {
                        $result.ActiveRoles.Add($roleObj) | Out-Null
                        Write-Verbose "Added active group: $($roleObj.DisplayName)"
                    }
                    catch {
                        Write-Warning "Failed to add active group to result: $_"
                    }
                }
                catch {
                    Write-Warning "Failed to process active group '$($membership.Group.DisplayName)': $_"
                }
            }
        }
        catch {
            Write-Warning "Failed to batch fetch group roles: $_"
            # Continue with other role types
        }
    }

        
    # Batch fetch Azure Resource roles if requested
    if ($IncludeAzureResources) {
        Write-Verbose "Batch fetching Azure Resource roles..."

        # Safely test cache and SkipAzureRefresh
        $hasAzureCache = ($script:AzureRolesCache -is [System.Collections.IEnumerable]) -and ($script:AzureRolesCache.Count -gt 0)
        # Maintain a per-subscription dirty set to support delta refresh after activation
        if (-not (Get-Variable -Name 'DirtyAzureSubscriptions' -Scope Script -ErrorAction SilentlyContinue)) {
            $script:DirtyAzureSubscriptions = @()
        }

        if ($SkipAzureRefresh -and $hasAzureCache) {
            Write-Verbose "Using cached Azure resource roles (SkipAzureRefresh requested)"
            
            # Update progress: Processing Azure resources from cache
            if ($SplashForm -and -not $SplashForm.IsDisposed) {
                $SplashForm.UpdateStatus("Processing Azure Resource roles...", 92)
            }
            
            $dirtySubs = @(); if ($script:DirtyAzureSubscriptions) { $dirtySubs = @($script:DirtyAzureSubscriptions | Select-Object -Unique) }
            $dirtyMgs  = @(); if ((Get-Variable -Name 'DirtyManagementGroups' -Scope Script -ErrorAction SilentlyContinue)) { $dirtyMgs = @($script:DirtyManagementGroups | Select-Object -Unique) }
            if ($dirtySubs.Count -gt 0) {
                Write-Verbose "Delta refresh: updating Azure cache for subscriptions: $(@($dirtySubs) -join ', ')"
                try {
                    # Determine user principal for Azure call
                    $currentUser = $script:CurrentUser
                    $userPrincipalName = if ($currentUser -and $currentUser.UserPrincipalName) { $currentUser.UserPrincipalName } else { (Get-MgUser -UserId $UserId -Select "userPrincipalName" -ErrorAction SilentlyContinue).UserPrincipalName }
                    $deltaRoles = Get-AzureResourceRoles -UserId $userPrincipalName -UserObjectId $UserId -IncludeActive:$true -IncludeEligible:$false -SubscriptionIds $dirtySubs -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
                    if (-not $deltaRoles) { $deltaRoles = @() } elseif ($deltaRoles -isnot [array]) { $deltaRoles = @($deltaRoles) }

                    # Rebuild cache by replacing roles for dirty subscriptions
                    $cleanCache = @()
                    foreach ($r in $script:AzureRolesCache) {
                        $sid = if ($r.PSObject.Properties['SubscriptionId']) { $r.SubscriptionId } else { $null }
                        if ($sid -and ($dirtySubs -contains $sid)) {
                            # skip old entries for dirty sub
                            continue
                        }
                        $cleanCache += $r
                    }
                    $script:AzureRolesCache = @($cleanCache + $deltaRoles)
                    $script:AzureRolesCacheTime = Get-Date
                    # Clear dirty set once applied
                    $script:DirtyAzureSubscriptions = @()
                }
                catch { Write-Verbose "Delta refresh failed, will proceed with existing cache: $($_.Exception.Message)" }
            }

            # Targeted MG delta refresh when management groups are marked dirty
            if ($dirtyMgs.Count -gt 0) {
                Write-Verbose "Delta refresh: updating Azure cache for management groups: $(@($dirtyMgs) -join ', ')"
                try {
                    # Determine user principal for Azure call
                    $currentUser = $script:CurrentUser
                    $userPrincipalName = if ($currentUser -and $currentUser.UserPrincipalName) { $currentUser.UserPrincipalName } else { (Get-MgUser -UserId $UserId -Select "userPrincipalName" -ErrorAction SilentlyContinue).UserPrincipalName }

                    # Invoke Azure role fetch limited to explicit DirtyManagementGroups only (no full discovery)
                    $deltaMgRoles = Get-AzureResourceRoles -UserId $userPrincipalName -UserObjectId $UserId -IncludeActive:$true -IncludeEligible:$false -SubscriptionIds $null -OnlyDirtyManagementGroups:$true -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
                    if (-not $deltaMgRoles) { $deltaMgRoles = @() } elseif ($deltaMgRoles -isnot [array]) { $deltaMgRoles = @($deltaMgRoles) }

                    # Remove cache entries for dirty MG scopes
                    $cleanCache = @()
                    foreach ($r in $script:AzureRolesCache) {
                        $fs = if ($r.PSObject.Properties['FullScope']) { $r.FullScope } elseif ($r.PSObject.Properties['Scope']) { $r.Scope } else { $null }
                        $isMgScope = $false
                        if ($fs -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                            $mgName = $matches[1]
                            if ($dirtyMgs -contains $mgName) { $isMgScope = $true }
                        }
                        if ($isMgScope) { continue }
                        $cleanCache += $r
                    }

                    # Keep only refreshed roles for dirty MGs
                    $deltaFiltered = @()
                    foreach ($dr in $deltaMgRoles) {
                        $dfs = if ($dr.PSObject.Properties['FullScope']) { $dr.FullScope } elseif ($dr.PSObject.Properties['Scope']) { $dr.Scope } else { $null }
                        if ($dfs -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                            $dmg = $matches[1]
                            if ($dirtyMgs -contains $dmg) { $deltaFiltered += $dr }
                        }
                    }

                    $script:AzureRolesCache = @($cleanCache + $deltaFiltered)
                    $script:AzureRolesCacheTime = Get-Date
                    # Clear dirty MGs once applied
                    $script:DirtyManagementGroups = @()
                }
                catch { Write-Verbose "MG delta refresh failed, will proceed with existing cache: $($_.Exception.Message)" }
            }

            # Apply any pending Azure Active overrides (immediate post-activation visibility)
                try {
                    if (-not (Get-Variable -Name 'AzureActiveOverrides' -Scope Script -ErrorAction SilentlyContinue)) { $script:AzureActiveOverrides = @{} }
                    if ($script:AzureActiveOverrides -and $script:AzureActiveOverrides.Count -gt 0) {
                    Write-Verbose "Applying Azure active overrides to cache for immediate visibility"
                    foreach ($role in $script:AzureRolesCache) {
                        try {
                            $fs = if ($role.PSObject.Properties['FullScope']) { $role.FullScope } elseif ($role.PSObject.Properties['Scope']) { $role.Scope } else { $null }
                            $rd = if ($role.PSObject.Properties['RoleDefinitionId']) { $role.RoleDefinitionId } else { $null }
                            if ($fs -and $rd) {
                                $key = "$fs|$rd"
                                if ($script:AzureActiveOverrides.ContainsKey($key)) {
                                    $ov = $script:AzureActiveOverrides[$key]
                                    # Mark as Active and set EndDateTime from override
                                    $role | Add-Member -NotePropertyName Status -NotePropertyValue 'Active' -Force
                                    if ($ov -and $ov.PSObject.Properties['EndDateTime']) {
                                        $role | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $ov.EndDateTime -Force
                                    }
                                }
                            }
                        } catch {}
                    }
                }
            } catch { Write-Verbose "Failed to apply Azure active overrides: $($_.Exception.Message)" }

            # Process cached Azure roles and add them to result arrays
            $addedEligibleCount = 0
            $addedActiveCount = 0
            $cleanedCache = [System.Collections.ArrayList]::new()
            
            foreach ($role in $script:AzureRolesCache) {
                # Normalize type and ensure required properties for UI consistency
                try {
                    if ($role.PSObject.Properties['Type'] -and $role.Type -eq 'Azure') { $role.Type = 'AzureResource' }
                    # Ensure FullScope exists for policy/UI keys
                    if (-not $role.PSObject.Properties['FullScope']) {
                        $fs = $null
                        if ($role.PSObject.Properties['FullScope']) { $fs = $role.FullScope }
                        elseif ($role.PSObject.Properties['Scope']) { $fs = $role.Scope }
                        elseif ($role.PSObject.Properties['DirectoryScopeId']) { $fs = $role.DirectoryScopeId }
                        if ($fs) { $role | Add-Member -NotePropertyName FullScope -NotePropertyValue $fs -Force }
                    }
                    # Ensure EndDateTime exists for active roles (many UI paths rely on it)
                    if (-not $role.PSObject.Properties['EndDateTime']) {
                        $end = $null
                        if ($role.PSObject.Properties['EndDateTime']) { $end = $role.EndDateTime }
                        elseif ($role.PSObject.Properties['Expiry']) { $end = $role.Expiry }
                        elseif ($role.PSObject.Properties['Expiration']) { $end = $role.Expiration }
                        elseif ($role.PSObject.Properties['ScheduleEnd']) { $end = $role.ScheduleEnd }
                        if ($end) { $role | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $end -Force }
                    }
                    # Ensure a displayable resource name exists
                    if (-not $role.PSObject.Properties['ResourceDisplayName'] -or -not $role.ResourceDisplayName) {
                        $rd = $null
                        if ($role.PSObject.Properties['ResourceName'] -and $role.ResourceName) { $rd = $role.ResourceName }
                        elseif ($role.PSObject.Properties['SubscriptionName'] -and $role.SubscriptionName) { $rd = $role.SubscriptionName }
                        elseif ($role.PSObject.Properties['ScopeDisplayName'] -and $role.ScopeDisplayName) { $rd = $role.ScopeDisplayName }
                        if ($rd) { $role | Add-Member -NotePropertyName ResourceDisplayName -NotePropertyValue $rd -Force }
                    }
                } catch {}
                
                # Check if active role has expired before processing (using same logic as UI)
                $isExpired = $false
                if ($role.Status -eq 'Active' -and $role.PSObject.Properties['EndDateTime'] -and $role.EndDateTime) {
                    try {
                        # Use same parsing logic as UI code
                        $endTime = if ($role.EndDateTime -is [DateTime]) {
                            $role.EndDateTime
                        }
                        else {
                            [DateTime]::Parse($role.EndDateTime)
                        }
                        
                        # Ensure UTC comparison (same as UI code)
                        if ($endTime.Kind -ne [DateTimeKind]::Utc) {
                            $endTime = $endTime.ToUniversalTime()
                        }
                        
                        # Calculate time remaining (same as UI code)
                        $now = [DateTime]::UtcNow
                        $remaining = $endTime - $now
                        
                        if ($remaining.TotalMinutes -le 0) {
                            $isExpired = $true
                            Write-Verbose "Detected expired Azure role '$($role.DisplayName)' on scope '$($role.FullScope)' (expired: $($role.EndDateTime))"
                        }
                    }
                    catch {
                        Write-Verbose "Could not parse EndDateTime for role '$($role.DisplayName)': $($role.EndDateTime) - treating as non-expired"
                    }
                }
                
                # Add non-expired roles to cleaned cache and results
                if (-not $isExpired) {
                    $cleanedCache.Add($role) | Out-Null
                    
                    # Honor ActiveOnly mode when adding cached Azure roles to result
                    if ($role.Status -eq 'Eligible' -and -not $ActiveOnly) { 
                        $result.EligibleRoles.Add($role) | Out-Null
                        $addedEligibleCount++
                    }
                    elseif ($role.Status -eq 'Active') {
                        $result.ActiveRoles.Add($role) | Out-Null
                        $addedActiveCount++
                    }
                }
                # Note: Expired roles are not added to cleanedCache or results, effectively removing them
            }
            
            # Update cache with cleaned entries (expired active roles removed)
            $expiredCount = $script:AzureRolesCache.Count - $cleanedCache.Count
            if ($expiredCount -gt 0) {
                Write-Verbose "Removed $expiredCount expired role(s) from Azure cache"
                $script:AzureRolesCache = $cleanedCache.ToArray()
                $script:AzureRolesCacheTime = Get-Date
            }
            
            Write-Verbose "Debug: Total Azure cache contains $($script:AzureRolesCache.Count) roles"
            Write-Verbose "Debug: ActiveOnly mode = $ActiveOnly"
            Write-Verbose "Debug: Eligible roles in cache: $((@($script:AzureRolesCache | Where-Object { $_.Status -eq 'Eligible' })).Count)"
            Write-Verbose "Debug: Active roles in cache: $((@($script:AzureRolesCache | Where-Object { $_.Status -eq 'Active' })).Count)"
            Write-Verbose "Added $addedEligibleCount eligible and $addedActiveCount active Azure resource roles from cache to result arrays"
            Write-Verbose "Current result counts after Azure cache processing: $($result.EligibleRoles.Count) eligible, $($result.ActiveRoles.Count) active"
        }
        else {
            # Update progress: Azure resources fetch
            if ($SplashForm -and -not $SplashForm.IsDisposed) {
                $SplashForm.UpdateStatus("Fetching Azure Resource roles...", 92)
            }
        
            try {
                # Determine user UPN for Azure authentication
                $currentUser = $script:CurrentUser
                $userPrincipalName = if ($currentUser -and $currentUser.UserPrincipalName) {
                    $currentUser.UserPrincipalName
                }
                else {
                    # Fallback: try to get user info from Graph
                    try {
                        $userInfo = Get-MgUser -UserId $UserId -Select "userPrincipalName" -ErrorAction SilentlyContinue
                        $userInfo.UserPrincipalName
                    }
                    catch {
                        Write-Warning "Could not determine user principal name for Azure authentication"
                        $UserId  # Use UserId as fallback
                    }
                }
            
                # Get Azure resource roles; skip eligible enumeration when ActiveOnly is true
                # If we have a dirty subscription set, restrict enumeration to those for delta
                $subIds = $null
                if ($script:DirtyAzureSubscriptions -and $script:DirtyAzureSubscriptions.Count -gt 0) { $subIds = @($script:DirtyAzureSubscriptions | Select-Object -Unique) }
                $azureRoles = Get-AzureResourceRoles -UserId $userPrincipalName -UserObjectId $UserId -IncludeActive:$true -IncludeEligible:$(-not $ActiveOnly) -SubscriptionIds $subIds -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
                
                # Ensure we have an array even if function returns null
                if (-not $azureRoles) { 
                    Write-Verbose "Get-AzureResourceRoles returned null - setting to empty array"
                    $azureRoles = @() 
                }
                elseif ($azureRoles -isnot [array]) { 
                    Write-Verbose "Get-AzureResourceRoles returned single object - converting to array"
                    $azureRoles = @($azureRoles) 
                }
                
                Write-Verbose "Azure roles returned from Get-AzureResourceRoles: $($azureRoles.Count) total"

                # Normalize type naming to match batch/UI expectations
                # Some implementations return Type='Azure'; the batch logic historically used 'AzureResource'.
                # We'll harmonize by converting any 'Azure' to 'AzureResource' here.
                foreach ($role in $azureRoles) {
                    try {
                        if ($role.PSObject.Properties['Type'] -and $role.Type -eq 'Azure') {
                            $role.Type = 'AzureResource'
                        }
                        # Ensure FullScope is available for downstream policy keying/UI
                        if (-not $role.PSObject.Properties['FullScope']) {
                            $fullScopeValue = $null
                            if ($role.PSObject.Properties['FullScope']) { $fullScopeValue = $role.FullScope }
                            elseif ($role.PSObject.Properties['Scope']) { $fullScopeValue = $role.Scope }
                            elseif ($role.PSObject.Properties['DirectoryScopeId']) { $fullScopeValue = $role.DirectoryScopeId }
                            if ($fullScopeValue) { $role | Add-Member -NotePropertyName FullScope -NotePropertyValue $fullScopeValue -Force }
                        }
                        # Ensure EndDateTime is present for active roles for immediate UI visibility
                        if (-not $role.PSObject.Properties['EndDateTime']) {
                            $end2 = $null
                            if ($role.PSObject.Properties['EndDateTime']) { $end2 = $role.EndDateTime }
                            elseif ($role.PSObject.Properties['Expiry']) { $end2 = $role.Expiry }
                            elseif ($role.PSObject.Properties['Expiration']) { $end2 = $role.Expiration }
                            elseif ($role.PSObject.Properties['ScheduleEnd']) { $end2 = $role.ScheduleEnd }
                            if ($end2) { $role | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $end2 -Force }
                        }
                    }
                    catch {
                        Write-Verbose "Could not normalize Azure role type: $($_.Exception.Message)"
                    }
                }
            
                if ($azureRoles.Count -gt 0) {
                    # Cache the full set to reuse on post-activation refresh
                    # Merge into cache: replace entries for any processed subscriptions to keep cache coherent
                    if ($subIds -and $subIds.Count -gt 0 -and ($script:AzureRolesCache -is [System.Collections.IEnumerable])) {
                        $existing = @()
                        foreach ($r in $script:AzureRolesCache) {
                            $sid = if ($r.PSObject.Properties['SubscriptionId']) { $r.SubscriptionId } else { $null }
                            if ($sid -and ($subIds -contains $sid)) { continue }
                            $existing += $r
                        }
                        $script:AzureRolesCache = @($existing + $azureRoles)
                    }
                    else {
                        $script:AzureRolesCache = @($azureRoles)
                    }
                    $script:AzureRolesCacheTime = Get-Date

                    # Separate eligible and active Azure roles - ensure arrays
                    $eligibleAzureRoles = @($azureRoles | Where-Object { $_.Status -eq 'Eligible' })
                    $activeAzureRoles = @($azureRoles | Where-Object { $_.Status -eq 'Active' })
                    
                    # Apply any Azure Active overrides before separation to ensure immediate display
                    try {
                        if (-not (Get-Variable -Name 'AzureActiveOverrides' -Scope Script -ErrorAction SilentlyContinue)) { $script:AzureActiveOverrides = @{} }
                        if ($script:AzureActiveOverrides -and $script:AzureActiveOverrides.Count -gt 0) {
                            Write-Verbose "Applying Azure active overrides before separation"
                            foreach ($role in $azureRoles) {
                                try {
                                    $fs = if ($role.PSObject.Properties['FullScope']) { $role.FullScope } elseif ($role.PSObject.Properties['Scope']) { $role.Scope } else { $null }
                                    $rd = if ($role.PSObject.Properties['RoleDefinitionId']) { $role.RoleDefinitionId } else { $null }
                                    if ($fs -and $rd) {
                                        $key = "$fs|$rd"
                                        if ($script:AzureActiveOverrides.ContainsKey($key)) {
                                            $ov = $script:AzureActiveOverrides[$key]
                                            $role | Add-Member -NotePropertyName Status -NotePropertyValue 'Active' -Force
                                            if ($ov -and $ov.PSObject.Properties['EndDateTime']) {
                                                $role | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $ov.EndDateTime -Force
                                            }
                                        }
                                    }
                                } catch {}
                            }
                        }
                    } catch { Write-Verbose "Failed to apply Azure active overrides pre-separation: $($_.Exception.Message)" }

                    Write-Verbose "Azure roles separation: $($eligibleAzureRoles.Count) eligible, $($activeAzureRoles.Count) active"
                    if ($azureRoles.Count -gt 0) {
                        $statusTypes = $azureRoles | Group-Object -Property Status | ForEach-Object { "$($_.Name): $($_.Count)" }
                        Write-Verbose "Azure roles by status: $($statusTypes -join ', ')"
                    }
                
                    # Add to result arrays, honoring ActiveOnly refresh
                    if (-not $ActiveOnly) {
                        foreach ($role in $eligibleAzureRoles) { $result.EligibleRoles.Add($role) | Out-Null }
                    }
                    foreach ($role in $activeAzureRoles) { $result.ActiveRoles.Add($role) | Out-Null }

                    if (-not $ActiveOnly) {
                        Write-Verbose "Added $($eligibleAzureRoles.Count) eligible and $($activeAzureRoles.Count) active Azure resource roles"
                    } else {
                        Write-Verbose "Active-only refresh: added $($activeAzureRoles.Count) active Azure resource roles (skipped eligible)"
                    }
                }
            }
            catch {
                Write-Warning "Failed to batch fetch Azure resource roles: $_"
                # Continue processing
            }
        }
    }
        
    if (-not $ActiveOnly) {
        # Batch fetch all unique policies (Entra + Groups + Azure)
        $uniqueRoleIds = [System.Collections.ArrayList]::new()
        $null = $uniqueRoleIds.AddRange(@($result.EligibleRoles | Where-Object { $_.Type -eq 'Entra' } | Select-Object -ExpandProperty RoleDefinitionId -Unique))
        $uniqueGroupIds = [System.Collections.ArrayList]::new()
        $null = $uniqueGroupIds.AddRange(@($result.EligibleRoles | Where-Object { $_.Type -eq 'Group' } | Select-Object -ExpandProperty GroupId -Unique))
        # Add Azure resource roles to policy collection
        $uniqueAzureRoles = [System.Collections.ArrayList]::new()
        # Support both historical 'AzureResource' and normalized 'Azure' types
        $azureResourceRoles = @($result.EligibleRoles | Where-Object { $_.Type -in @('AzureResource','Azure') })
        foreach ($azureRole in $azureResourceRoles) {
            if ($azureRole.PSObject.Properties['RoleDefinitionId'] -and $azureRole.PSObject.Properties['FullScope']) {
                $null = $uniqueAzureRoles.Add([PSCustomObject]@{
                    RoleDefinitionId = $azureRole.RoleDefinitionId
                    FullScope        = $azureRole.FullScope 
                    SubscriptionId   = $azureRole.SubscriptionId
                })
            }
        }

        # Update progress: Starting policy fetch
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            $SplashForm.UpdateStatus("Fetching role policies...", 92)
        }

        # Ensure we have arrays and handle null/empty cases
        if (-not $uniqueRoleIds) {
            $uniqueRoleIds = @()
        }
        elseif ($uniqueRoleIds -isnot [array]) {
            $uniqueRoleIds = @($uniqueRoleIds)
        }

        if (-not $uniqueGroupIds) {
            $uniqueGroupIds = @()
        }
        elseif ($uniqueGroupIds -isnot [array]) {
            $uniqueGroupIds = @($uniqueGroupIds)
        }

        if (-not $uniqueAzureRoles) {
            $uniqueAzureRoles = @()
        }
        elseif ($uniqueAzureRoles -isnot [array]) {
            $uniqueAzureRoles = @($uniqueAzureRoles)
        }

        if ($uniqueRoleIds.Count -gt 0) {
            Write-Verbose "Batch fetching $($uniqueRoleIds.Count) Entra role policies"
            try {
                Get-PIMPoliciesBatch -RoleIds $uniqueRoleIds -Type 'Entra' -PolicyCache $result.Policies -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
            }
            catch {
                Write-Warning "Failed to batch fetch Entra role policies: $_"
            }
        }

        if ($uniqueGroupIds.Count -gt 0) {
            Write-Verbose "Batch fetching $($uniqueGroupIds.Count) group policies"
            try {
                Get-PIMPoliciesBatch -GroupIds $uniqueGroupIds -Type 'Group' -PolicyCache $result.Policies -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
            }
            catch {
                Write-Warning "Failed to batch fetch group policies: $_"
            }
        }

        # Fetch Azure resource policies
        if ($uniqueAzureRoles.Count -gt 0) {
            Write-Verbose "Fetching policies for $($uniqueAzureRoles.Count) Azure resource roles..."
            try {
                foreach ($azureRole in $uniqueAzureRoles) {
                    $cacheKey = "AzureResource_$($azureRole.FullScope)_$($azureRole.RoleDefinitionId)"
                    
                    if (-not $result.Policies.ContainsKey($cacheKey)) {
                        # Get Azure resource policy using the helper function
                        $azurePolicy = Get-AzureResourcePIMPolicy -RoleDefinitionId $azureRole.RoleDefinitionId -SubscriptionId $azureRole.SubscriptionId -Scope $azureRole.FullScope
                        
                        if ($azurePolicy) {
                            $result.Policies[$cacheKey] = $azurePolicy
                            
                            # Also cache in script-level cache
                            if (Get-Variable -Name 'PolicyCache' -Scope Script -ErrorAction SilentlyContinue) {
                                $script:PolicyCache[$cacheKey] = $azurePolicy
                            }
                            
                            Write-Verbose "Cached Azure resource policy for role: $($azureRole.RoleDefinitionId) in scope: $($azureRole.FullScope)"
                        }
                    }
                }
            }
            catch {
                Write-Warning "Failed to fetch Azure resource policies: $_"
            }
        }

        # Batch fetch authentication contexts if any policies require them
        $contextIds = @($result.Policies.Values | 
            Where-Object { $_.RequiresAuthenticationContext -and $_.AuthenticationContextId } | 
            Select-Object -ExpandProperty AuthenticationContextId -Unique)

        # Ensure we have arrays and handle null/empty cases
        if (-not $contextIds) {
            $contextIds = @()
        }
        elseif ($contextIds -isnot [array]) {
            $contextIds = @($contextIds)
        }

        if ($contextIds.Count -gt 0) {
            Write-Verbose "Batch fetching $($contextIds.Count) authentication contexts"
            try {
                Get-AuthenticationContextsBatch -ContextIds $contextIds -ContextCache $result.AuthenticationContexts
            }
            catch {
                Write-Warning "Failed to batch fetch authentication contexts: $_"
            }
        }

        # Update progress: Attaching policies
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            $SplashForm.UpdateStatus("Attaching policies to roles...", 96)
        }

        # Attach PolicyInfo to role objects for direct UI access
        Write-Verbose "Attaching policy information to role objects..."
    }
        
    # Attach policies to eligible roles
    foreach ($role in $result.EligibleRoles) {
        $cacheKey = if ($role.Type -eq 'Group') { 
            "Group_$($role.GroupId)" 
        }
        elseif ($role.Type -in @('AzureResource','Azure')) {
            "AzureResource_$($role.FullScope)_$($role.RoleDefinitionId)"  # Use FullScope
        }
        else { 
            "Entra_$($role.RoleDefinitionId)" 
        }

        if ($result.Policies.ContainsKey($cacheKey)) {
            $policyInfo = $result.Policies[$cacheKey]

            # Add PolicyInfo as a property to the role object
            $role | Add-Member -NotePropertyName PolicyInfo -NotePropertyValue $policyInfo -Force
            Write-Verbose "Attached policy info to $($role.DisplayName)"
        }
        else {
            # Create default policy for roles without cached policies
            $defaultPolicy = [PSCustomObject]@{
                MaxDuration                      = 8
                RequiresMfa                      = $false
                RequiresJustification            = if ($role.Type -eq 'AzureResource') { $true } else { $false }
                RequiresTicket                   = $false
                RequiresApproval                 = $false
                RequiresAuthenticationContext    = $false
                AuthenticationContextId          = $null
                AuthenticationContextDisplayName = $null
                AuthenticationContextDescription = $null
                AuthenticationContextDetails     = $null
            }
            $role | Add-Member -NotePropertyName PolicyInfo -NotePropertyValue $defaultPolicy -Force
            Write-Verbose "Attached default policy to role: $($role.DisplayName)"
        }
    }

    # Attach policies to active roles (they may also need policy info for display)
    foreach ($role in $result.ActiveRoles) {
        $cacheKey = if ($role.Type -eq 'Group') { 
            "Group_$($role.GroupId)" 
        }
        elseif ($role.Type -in @('AzureResource','Azure')) {
            "AzureResource_$($role.FullScope)_$($role.RoleDefinitionId)"  # Use FullScope
        }
        else { 
            "Entra_$($role.RoleDefinitionId)" 
        }
    
        if ($result.Policies.ContainsKey($cacheKey)) {
            $policyInfo = $result.Policies[$cacheKey]
        
            # Enhance policy with authentication context details if needed
            if ($policyInfo.RequiresAuthenticationContext -and $policyInfo.AuthenticationContextId -and 
                $result.AuthenticationContexts.ContainsKey($policyInfo.AuthenticationContextId)) {
                $authContext = $result.AuthenticationContexts[$policyInfo.AuthenticationContextId]
                $policyInfo.AuthenticationContextDisplayName = $authContext.DisplayName
                $policyInfo.AuthenticationContextDescription = $authContext.Description
                $policyInfo.AuthenticationContextDetails = $authContext
            }
        
            # Add PolicyInfo as a property to the role object
            $role | Add-Member -NotePropertyName PolicyInfo -NotePropertyValue $policyInfo -Force
            Write-Verbose "Attached policy info to active role $($role.DisplayName)"
        }
    }
        
    # Cross-reference group-derived Entra roles with their providing groups for proper attribution
    Write-Verbose "Cross-referencing group-derived Entra roles with providing groups..."

    # Ensure we get arrays from the filters
    $activeGroupRoles = @($result.ActiveRoles   | Where-Object { $_.Type -eq 'Group' })
    $eligibleGroupRoles = @($result.EligibleRoles | Where-Object { $_.Type -eq 'Group' })
    $allActiveEntraRoles = @($result.ActiveRoles  | Where-Object { $_.Type -eq 'Entra' })

    Write-Verbose "Found $($activeGroupRoles.Count) active groups, $($eligibleGroupRoles.Count) eligible groups, and $($allActiveEntraRoles.Count) active Entra roles"

    # Combine all groups (active and eligible) for role mapping
    $allGroups = [System.Collections.ArrayList]::new()
    $null = $allGroups.AddRange($activeGroupRoles)
    $null = $allGroups.AddRange($eligibleGroupRoles)

    if ($allGroups.Count -gt 0) {
        # Debug: Show what groups we have and their provided roles
        foreach ($group in $allGroups) {
            $groupStatus = if ($activeGroupRoles -contains $group) { "Active" } else { "Eligible" }
        
            # Safely get EndDateTime for verbose logging
            $groupExpiration = "N/A"
            try {
                if ($group.PSObject.Properties['EndDateTime'] -and $group.EndDateTime) {
                    $groupExpiration = $group.EndDateTime
                }
                else {
                    $groupExpiration = "Permanent"
                }
            }
            catch {
                $groupExpiration = "Unknown"
            }
        
            Write-Verbose "$groupStatus group '$($group.DisplayName)' expires at: $groupExpiration"
            if ($group.ProvidedRoles -and $group.ProvidedRoles.Count -gt 0) {
                Write-Verbose "  Provides $($group.ProvidedRoles.Count) roles:"
                foreach ($pr in $group.ProvidedRoles) {
                    Write-Verbose "    - $($pr.DisplayName) (ID: $($pr.RoleDefinitionId))"
                }
            }
            else {
                Write-Verbose "  No provided roles found for this group"
            }
        }
    
        # Build a map of roles provided by groups, prioritizing active groups
        $roleProvidedByGroups = @{}
    
        foreach ($group in $allGroups) {
            try {
                if ($group.ProvidedRoles -and $group.ProvidedRoles.Count -gt 0) {
                    foreach ($providedRole in $group.ProvidedRoles) {
                        try {
                            $roleId = $providedRole.RoleDefinitionId
                        
                            if (-not $roleProvidedByGroups.ContainsKey($roleId)) {
                                $roleProvidedByGroups[$roleId] = [System.Collections.ArrayList]::new()
                            }
                        
                            # Store group info with this role, marking if it's active
                            $isActiveGroup = $activeGroupRoles -contains $group
                        
                            # Safely get group EndDateTime
                            $groupEndDateTime = $null
                            if ($group.PSObject.Properties['EndDateTime']) {
                                $groupEndDateTime = $group.EndDateTime
                            }
                        
                            $null = $roleProvidedByGroups[$roleId].Add([PSCustomObject]@{
                                    Group              = $group
                                    ProvidedRole       = $providedRole
                                    IsActiveGroup      = $isActiveGroup
                                    HasExpiration      = [bool]$groupEndDateTime
                                    ExpirationDateTime = $groupEndDateTime
                                    Priority           = if ($isActiveGroup) { 1 } else { 2 }  # Active groups get priority
                                })
                        }
                        catch {
                            Write-Warning "Failed to process provided role '$($providedRole.DisplayName)' for group '$($group.DisplayName)': $_"
                        }
                    }
                }
            }
            catch {
                Write-Warning "Failed to process roles for group '$($group.DisplayName)': $_"
            }
        }
    
        Write-Verbose "Built role-to-group mapping for $($roleProvidedByGroups.Keys.Count) unique roles"
    
        # Process each Entra role to attribute it to the correct group
        # Group same roles together to handle multiple instances intelligently
        $roleGroups = $allActiveEntraRoles | Group-Object -Property RoleDefinitionId
    
        foreach ($roleGroup in $roleGroups) {
            $roleId = $roleGroup.Name
            $roleInstances = $roleGroup.Group
        
            Write-Verbose "Processing $($roleInstances.Count) instance(s) of role ID: $roleId"
        
            # Check if any active group provides this role
            if ($roleProvidedByGroups.ContainsKey($roleId)) {
                $allProvidingGroups = $roleProvidedByGroups[$roleId]
                # ONLY consider ACTIVE groups for attribution
                $providingGroups = $allProvidingGroups | Where-Object { $_.IsActiveGroup -eq $true }
            
                if ($providingGroups -and $providingGroups.Count -gt 0) {
                    Write-Verbose "  Role '$($roleInstances[0].DisplayName)' is provided by $($providingGroups.Count) ACTIVE group(s)"
                }
                else {
                    Write-Verbose "  Role '$($roleInstances[0].DisplayName)' has potential providing groups but none are currently active - treating as direct assignment"
                    $providingGroups = @()  # Clear the array to indicate no active groups
                }
            
                # PRIORITY: First, handle any inherited roles (these definitely come from groups)
                # Since MemberType might not reliably indicate 'Inherited', we'll use a smarter approach:
                # If a role is provided by an active group and the expiration time matches, treat it as inherited
                $inheritedRoles = [System.Collections.ArrayList]::new()
                $directRoles = [System.Collections.ArrayList]::new()
            
                foreach ($role in $roleInstances) {
                    $isLikelyInherited = $false
                
                    # Primary check: MemberType = 'Group' indicates a group-derived role
                    if ($role.MemberType -eq 'Group') {
                        $isLikelyInherited = $true
                        Write-Verbose "    Role '$($role.DisplayName)' marked as inherited by MemberType 'Group'"
                    }
                    # Secondary check: If MemberType is not 'Group' but expiration exactly matches an active group
                    elseif ($providingGroups -and $providingGroups.Count -gt 0) {
                        foreach ($groupInfo in $providingGroups) {
                            # Only consider expiration matching for non-'Group' MemberTypes if the match is very specific
                            if ($role.EndDateTime -and $groupInfo.ExpirationDateTime -and 
                                $role.EndDateTime -eq $groupInfo.ExpirationDateTime) {
                                $isLikelyInherited = $true
                                Write-Verbose "    Role '$($role.DisplayName)' likely inherited from group '$($groupInfo.Group.DisplayName)' (exact expiration match: $($role.EndDateTime))"
                                break
                            }
                            # Special case: both role and group have no expiration (permanent) AND role has no other assignment path
                            elseif (-not $role.EndDateTime -and -not $groupInfo.ExpirationDateTime -and $role.MemberType -ne 'Direct') {
                                $isLikelyInherited = $true
                                Write-Verbose "    Role '$($role.DisplayName)' likely inherited from group '$($groupInfo.Group.DisplayName)' (both permanent and no direct assignment)"
                                break
                            }
                        }
                    }
                
                    if ($isLikelyInherited) {
                        $null = $inheritedRoles.Add($role)
                    }
                    else {
                        $null = $directRoles.Add($role)
                    }
                
                    Write-Verbose "    Role '$($role.DisplayName)' classified as: $(if ($isLikelyInherited) { 'INHERITED' } else { 'DIRECT' }) (MemberType: '$($role.MemberType)', EndDateTime: $($role.EndDateTime))"
                }
            
                Write-Verbose "  Found $($inheritedRoles.Count) inherited role(s) and $($directRoles.Count) direct role(s)"
            
                # Process inherited roles first - these MUST be attributed to groups
                $usedGroups = [System.Collections.ArrayList]::new()  # Track which groups have been used for attribution
                foreach ($inheritedRole in $inheritedRoles) {
                    if ($providingGroups -and $providingGroups.Count -gt 0) {
                        # Try to find the best matching group for this inherited role
                        $bestGroup = $null
                    
                        # First priority: Find group with exact expiration match
                        if ($inheritedRole.EndDateTime) {
                            $bestGroup = $providingGroups | Where-Object { 
                                $_.ExpirationDateTime -eq $inheritedRole.EndDateTime 
                            } | Select-Object -First 1
                            if ($bestGroup) {
                                Write-Verbose "    Matched inherited role '$($inheritedRole.DisplayName)' to group '$($bestGroup.Group.DisplayName)' by expiration: $($inheritedRole.EndDateTime)"
                            }
                        }
                    
                        # Second priority: If no expiration match, prefer unused groups
                        if (-not $bestGroup) {
                            $availableGroups = $providingGroups | Where-Object { $usedGroups -notcontains $_.Group.DisplayName }
                            if ($availableGroups) {
                                $bestGroup = $availableGroups | Sort-Object Priority | Select-Object -First 1
                                Write-Verbose "    Matched inherited role '$($inheritedRole.DisplayName)' to unused group '$($bestGroup.Group.DisplayName)'"
                            }
                            else {
                                # All groups used, fall back to priority sorting
                                $bestGroup = $providingGroups | Sort-Object Priority | Select-Object -First 1
                                Write-Verbose "    All groups used, matched inherited role '$($inheritedRole.DisplayName)' to highest priority group '$($bestGroup.Group.DisplayName)'"
                            }
                        }
                    
                        # Track this group as used
                        $null = $usedGroups.Add($bestGroup.Group.DisplayName)
                    
                        Write-Verbose "  Attributing inherited role '$($inheritedRole.DisplayName)' to group '$($bestGroup.Group.DisplayName)'"
                    
                        try {
                            # Update the inherited role with group information
                            if ($bestGroup.ExpirationDateTime) {
                                if ($inheritedRole.PSObject.Properties['EndDateTime']) {
                                    $inheritedRole.EndDateTime = $bestGroup.ExpirationDateTime
                                }
                                else {
                                    $inheritedRole | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $bestGroup.ExpirationDateTime -Force
                                }
                                Write-Verbose "    Updated inherited role expiration to match group: $($bestGroup.ExpirationDateTime)"
                            }
                            else {
                                # Group has no expiration (permanent)
                                if ($inheritedRole.PSObject.Properties['EndDateTime']) {
                                    $inheritedRole.EndDateTime = $null
                                }
                                else {
                                    $inheritedRole | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $null -Force
                                }
                                Write-Verbose "    Updated inherited role to permanent (no expiration)"
                            }
                        
                            # Update resource attribution for inherited role
                            $inheritedRole.ResourceName = "Entra ID (via Group: $($bestGroup.Group.DisplayName))"
                            $inheritedRole | Add-Member -NotePropertyName SourceGroup -NotePropertyValue $bestGroup.Group.DisplayName -Force
                            $inheritedRole | Add-Member -NotePropertyName SourceGroupId -NotePropertyValue $bestGroup.Group.GroupId -Force
                            $inheritedRole | Add-Member -NotePropertyName IsGroupAttributed -NotePropertyValue $true -Force
                        
                            Write-Verbose "  ✓ Successfully attributed inherited role '$($inheritedRole.DisplayName)' to group '$($bestGroup.Group.DisplayName)'"
                        }
                        catch {
                            Write-Warning "Failed to update inherited role '$($inheritedRole.DisplayName)' with group attribution: $_"
                        }
                    }
                    else {
                        Write-Warning "Inherited role '$($inheritedRole.DisplayName)' found but no active providing groups available for attribution"
                    }
                }
            
                # Now handle direct roles if there are any and remaining groups
                # Only attribute direct roles to groups if there are NO inherited roles of the same type
                if ($directRoles.Count -eq 1 -and $inheritedRoles.Count -eq 0 -and $providingGroups -and $providingGroups.Count -gt 0) {
                    $entraRole = $directRoles[0]  # Use directRoles instead of roleInstances
                
                    # Safely get the current EndDateTime from the Entra role
                    $currentEndDateTime = $null
                    if ($entraRole.PSObject.Properties['EndDateTime']) {
                        $currentEndDateTime = $entraRole.EndDateTime
                    }
                    Write-Verbose "  Single direct role - Current Entra role expiration: $currentEndDateTime"
                
                    # Find the best group for attribution
                    try {
                        $sortProperties = @(
                            @{Expression      = {
                                    try { $_.Priority } catch { 2 }
                                }; Descending = $false
                            }
                            @{Expression      = {
                                    try {
                                        $groupExp = if ($_.PSObject.Properties['ExpirationDateTime']) { $_.ExpirationDateTime } else { $null }
                                        if ($groupExp -eq $currentEndDateTime) { 0 } else { 1 }
                                    }
                                    catch { 1 }
                                }; Descending = $false
                            }
                            @{Expression      = {
                                    try { if ($_.PSObject.Properties['HasExpiration']) { $_.HasExpiration } else { $false } } catch { $false }
                                }; Descending = $false
                            }
                        )
                        $bestGroup = $providingGroups | Sort-Object $sortProperties | 
                        Select-Object -First 1
                    }
                    catch {
                        Write-Warning "Failed to sort providing groups for role '$($entraRole.DisplayName)': $_"
                        $bestGroup = $providingGroups | Select-Object -First 1
                    }
                
                    if ($bestGroup) {
                        Write-Verbose "  Best providing group for single direct role '$($entraRole.DisplayName)': '$($bestGroup.Group.DisplayName)' (Active: $($bestGroup.IsActiveGroup), expires: $($bestGroup.ExpirationDateTime))"
                    
                        # Update the Entra role with group information
                        try {
                            if ($bestGroup.ExpirationDateTime) {
                                # Add or update the EndDateTime property
                                if ($entraRole.PSObject.Properties['EndDateTime']) {
                                    $entraRole.EndDateTime = $bestGroup.ExpirationDateTime
                                }
                                else {
                                    $entraRole | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $bestGroup.ExpirationDateTime -Force
                                }
                                Write-Verbose "    Updated single direct role expiration to match group: $($bestGroup.ExpirationDateTime)"
                            }
                            else {
                                # Group has no expiration (permanent), ensure role shows as permanent
                                if ($entraRole.PSObject.Properties['EndDateTime']) {
                                    $entraRole.EndDateTime = $null
                                }
                                else {
                                    $entraRole | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $null -Force
                                }
                                Write-Verbose "    Updated single direct role to permanent (no expiration)"
                            }
                        
                            # Update resource attribution
                            $entraRole.ResourceName = "Entra ID (via Group: $($bestGroup.Group.DisplayName))"
                            $entraRole | Add-Member -NotePropertyName SourceGroup -NotePropertyValue $bestGroup.Group.DisplayName -Force
                            $entraRole | Add-Member -NotePropertyName SourceGroupId -NotePropertyValue $bestGroup.Group.GroupId -Force
                            $entraRole | Add-Member -NotePropertyName IsGroupAttributed -NotePropertyValue $true -Force
                        
                            Write-Verbose "  ✓ Updated single direct Entra role '$($entraRole.DisplayName)' attribution to group '$($bestGroup.Group.DisplayName)'"
                        }
                        catch {
                            Write-Warning "Failed to update single direct Entra role '$($entraRole.DisplayName)' with group attribution: $_"
                        }
                    }
                }
                # If we have multiple direct roles, try to match each to the most appropriate group
                elseif ($directRoles.Count -gt 1 -and $providingGroups -and $providingGroups.Count -gt 0) {
                    Write-Verbose "  Multiple direct role instances ($($directRoles.Count)) - attempting intelligent matching"
                
                    # Create a list of unassigned groups for this role
                    $availableGroups = $providingGroups | ForEach-Object { 
                        [PSCustomObject]@{
                            GroupInfo = $_
                            Assigned  = $false
                        }
                    }
                
                    # Sort direct role instances by their current expiration (nulls last)
                    $sortedDirectRoles = $directRoles | Sort-Object @(
                        @{ Expression     = {
                                try {
                                    if ($_.PSObject.Properties['EndDateTime'] -and $_.EndDateTime) {
                                        $_.EndDateTime
                                    }
                                    else {
                                        [DateTime]::MaxValue  # Put permanent/null expirations at the end
                                    }
                                }
                                catch {
                                    [DateTime]::MaxValue
                                }
                            }; Descending = $false
                        }
                    )
                
                    # Try to match each direct role instance to the best available group
                    foreach ($entraRole in $sortedDirectRoles) {
                        # Safely get the current EndDateTime from the Entra role
                        $currentEndDateTime = $null
                        if ($entraRole.PSObject.Properties['EndDateTime']) {
                            $currentEndDateTime = $entraRole.EndDateTime
                        }
                    
                        Write-Verbose "    Matching role instance with expiration: $currentEndDateTime"
                    
                        # Find the best available group for this specific instance
                        $bestAvailableGroup = $null
                    
                        # If we have unassigned groups, try to find the best match
                        $unassignedGroups = $availableGroups | Where-Object { -not $_.Assigned }
                    
                        if ($unassignedGroups) {
                            # Prefer active groups and then expiration characteristics
                            $sortProps = @(
                                @{Expression      = {
                                        # Priority 1: Active groups first
                                        try { $_.GroupInfo.Priority } catch { 2 }
                                    }; Descending = $false
                                }
                                @{Expression      = {
                                        # Priority 2: exact expiration match or temporal preference
                                        if (-not $currentEndDateTime) {
                                            if ($_.GroupInfo.PSObject.Properties['HasExpiration'] -and $_.GroupInfo.HasExpiration) { 0 } else { 1 }
                                        }
                                        else {
                                            if ($_.GroupInfo.ExpirationDateTime -eq $currentEndDateTime) { 0 } else { 1 }
                                        }
                                    }; Descending = $false
                                }
                                @{Expression      = {
                                        # Priority 3: prefer shorter expirations first
                                        try {
                                            if ($_.GroupInfo.ExpirationDateTime) {
                                                [datetime]$_.GroupInfo.ExpirationDateTime
                                            }
                                            else {
                                                [datetime]::MaxValue  # Permanent comes last
                                            }
                                        }
                                        catch {
                                            [datetime]::MaxValue
                                        }
                                    }; Descending = $false
                                }
                            )
                            $bestAvailableGroup = $unassignedGroups | Sort-Object $sortProps | Select-Object -First 1
                        
                            if ($bestAvailableGroup) {
                                $matchReason = "best priority-based match"
                            
                                # Check if it's an exact expiration match
                                if ($bestAvailableGroup.GroupInfo.ExpirationDateTime -eq $currentEndDateTime) {
                                    $matchReason = "exact expiration match"
                                }
                                # Check if it's a temporal vs permanent preference match
                                elseif (-not $currentEndDateTime -and $bestAvailableGroup.GroupInfo.HasExpiration) {
                                    $matchReason = "temporal group preference (role has no expiration, preferring group with expiration)"
                                }
                            
                                Write-Verbose "      Using $matchReason`: $($bestAvailableGroup.GroupInfo.Group.DisplayName)"
                            }
                        }
                    
                        if ($bestAvailableGroup) {
                            # Mark this group as assigned
                            $bestAvailableGroup.Assigned = $true
                            $groupInfo = $bestAvailableGroup.GroupInfo
                        
                            Write-Verbose "      Assigning group '$($groupInfo.Group.DisplayName)' (expires: $($groupInfo.ExpirationDateTime)) to role instance"
                        
                            # Update the Entra role with group information
                            try {
                                # Store the original EndDateTime before any modifications
                                $originalEndDateTime = $entraRole.EndDateTime
                            
                                # Only update EndDateTime if this role is ONLY available through groups
                                $shouldUpdateEndDateTime = $false
                            
                                if ($entraRole.MemberType -eq 'Group') {
                                    # This role is only available through group membership
                                    $shouldUpdateEndDateTime = $true
                                    Write-Verbose "        Role is only available through groups - will update expiration"
                                }
                                elseif (-not $originalEndDateTime -and $groupInfo.ExpirationDateTime) {
                                    # Role has no expiration but group does - preserve permanent direct assignment
                                    Write-Verbose "        Role has no expiration but group does - preserving permanent direct assignment"
                                }
                                elseif ($originalEndDateTime -and $groupInfo.ExpirationDateTime) {
                                    # Both role and group have expiration times - preserve direct assignment time
                                    Write-Verbose "        Role has both direct ($originalEndDateTime) and group ($($groupInfo.ExpirationDateTime)) assignments - preserving direct assignment time"
                                }
                            
                                if ($shouldUpdateEndDateTime) {
                                    if ($groupInfo.ExpirationDateTime) {
                                        # Add or update the EndDateTime property
                                        if ($entraRole.PSObject.Properties['EndDateTime']) {
                                            $entraRole.EndDateTime = $groupInfo.ExpirationDateTime
                                        }
                                        else {
                                            $entraRole | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $groupInfo.ExpirationDateTime -Force
                                        }
                                        Write-Verbose "        Updated instance expiration to match group: $($groupInfo.ExpirationDateTime)"
                                    }
                                    else {
                                        # Group has no expiration (permanent), ensure role shows as permanent
                                        if ($entraRole.PSObject.Properties['EndDateTime']) {
                                            $entraRole.EndDateTime = $null
                                        }
                                        else {
                                            $entraRole | Add-Member -NotePropertyName EndDateTime -NotePropertyValue $null -Force
                                        }
                                        Write-Verbose "        Updated instance to permanent (no expiration)"
                                    }
                                }
                                else {
                                    Write-Verbose "        Preserving original EndDateTime: $originalEndDateTime (direct assignment takes precedence)"
                                }
                            
                                # Update resource attribution - but only mark as group-attributed if it's ONLY from groups
                                if ($shouldUpdateEndDateTime) {
                                    $entraRole.ResourceName = "Entra ID (via Group: $($groupInfo.Group.DisplayName))"
                                    $entraRole | Add-Member -NotePropertyName IsGroupAttributed -NotePropertyValue $true -Force
                                }
                                else {
                                    # Role has multiple sources - keep as directory role but track group info
                                    $entraRole.ResourceName = "Entra ID Directory"
                                    $entraRole | Add-Member -NotePropertyName IsGroupAttributed -NotePropertyValue $false -Force
                                }
                            
                                # Always add source group info for reference
                                $entraRole | Add-Member -NotePropertyName SourceGroup -NotePropertyValue $groupInfo.Group.DisplayName -Force
                                $entraRole | Add-Member -NotePropertyName SourceGroupId -NotePropertyValue $groupInfo.Group.GroupId -Force
                            
                                Write-Verbose "      ✓ Updated Entra role instance '$($entraRole.DisplayName)' attribution (Group-only: $shouldUpdateEndDateTime)"
                            }
                            catch {
                                Write-Warning "Failed to update Entra role instance '$($entraRole.DisplayName)' with group attribution: $_"
                            }
                        }
                        else {
                            Write-Verbose "      No available groups left for this role instance - leaving unattributed"
                        }
                    }
                }
                # If no active providing groups, treat ONLY DIRECT instances as direct assignments
                # (inherited roles should keep their group attribution regardless)
                if (-not $providingGroups -or $providingGroups.Count -eq 0) {
                    Write-Verbose "  No active providing groups - treating $($directRoles.Count) direct instance(s) as direct assignments"
                    foreach ($entraRole in $directRoles) {
                        # Ensure direct assignment properties only for non-inherited roles
                        if ($entraRole.MemberType -ne 'Group') {
                            $entraRole.ResourceName = "Entra ID Directory"
                            $entraRole | Add-Member -NotePropertyName IsGroupAttributed -NotePropertyValue $false -Force
                            Write-Verbose "  ✓ Marked direct role '$($entraRole.DisplayName)' as direct assignment (no active providing groups)"
                        }
                        else {
                            Write-Verbose "  ✓ Skipped inherited role '$($entraRole.DisplayName)' - keeping existing group attribution"
                        }
                    }
                }
            }
            else {
                # This role type is not provided by any group - likely directly assigned
                foreach ($entraRole in $roleInstances) {
                    Write-Verbose "  Entra role '$($entraRole.DisplayName)' is not provided by any group (direct assignment)"
                }
            }
        }
    }

    # Final progress update
    if ($SplashForm -and -not $SplashForm.IsDisposed) {
        $SplashForm.UpdateStatus("Batch processing complete!", 98)
    }

    Write-Verbose "Batch fetch completed: $($result.EligibleRoles.Count) eligible, $($result.ActiveRoles.Count) active roles"
    
    # Safety net: ensure Azure active roles from cache are present in ActiveOnly refresh results
    try {
        if ($ActiveOnly -and $IncludeAzureResources) {
            $hasAzureCache = ($script:AzureRolesCache -is [System.Collections.IEnumerable]) -and ($script:AzureRolesCache.Count -gt 0)
            if ($hasAzureCache) {
                # Normalize types and build keys for comparison
                $existingKeys = @()
                foreach ($ar in $result.ActiveRoles) {
                    try {
                        $t = if ($ar.PSObject.Properties['Type']) { $ar.Type } else { $null }
                        $rd = if ($ar.PSObject.Properties['RoleDefinitionId']) { $ar.RoleDefinitionId } else { $null }
                        $fs = if ($ar.PSObject.Properties['FullScope']) { $ar.FullScope } elseif ($ar.PSObject.Properties['DirectoryScopeId']) { $ar.DirectoryScopeId } else { $null }
                        if ($t -and $rd -and $fs) { $existingKeys += "$t|$rd|$fs" }
                    } catch {}
                }

                $addedCount = 0
                foreach ($role in $script:AzureRolesCache) {
                    try {
                        # Normalize type for comparison
                        if ($role.PSObject.Properties['Type'] -and $role.Type -eq 'Azure') { $role.Type = 'AzureResource' }
                        if ($role.Status -ne 'Active') { continue }
                        $rd = if ($role.PSObject.Properties['RoleDefinitionId']) { $role.RoleDefinitionId } else { $null }
                        $fs = if ($role.PSObject.Properties['FullScope']) { $role.FullScope } elseif ($role.PSObject.Properties['Scope']) { $role.Scope } else { $null }
                        if (-not $role.PSObject.Properties['FullScope'] -and $fs) { $role | Add-Member -NotePropertyName FullScope -NotePropertyValue $fs -Force }
                        $key = "AzureResource|$rd|$fs"
                        if ($key -and ($existingKeys -notcontains $key)) {
                            $null = $result.ActiveRoles.Add($role)
                            $existingKeys += $key
                            $addedCount++
                        }
                    } catch {}
                }
                if ($addedCount -gt 0) { Write-Verbose "Safety net appended $addedCount Azure active role(s) from cache into ActiveOnly result" }
            }
        }
    } catch { Write-Verbose "Safety net merge failed: $($_.Exception.Message)" }

    Write-Verbose "Final counts: $($result.EligibleRoles.Count) eligible, $($result.ActiveRoles.Count) active roles"
    return $result
}