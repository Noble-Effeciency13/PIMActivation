function Update-PIMRolesList {
    <#
    .SYNOPSIS
        Updates the PIM roles lists in the Windows Forms UI.
    
    .DESCRIPTION
        Refreshes both active and eligible role lists in the PIM activation form.
        This function handles the UI updates for displaying PIM roles, including:
        - Fetching role data from Azure
        - Populating ListView controls with role information
        - Applying visual styling based on role status
        - Updating role count labels
        - Managing loading status during refresh operations
    
    .PARAMETER Form
        The Windows Forms form object containing the role list views.
        Must contain ListView controls named 'lstActive' and 'lstEligible'.
    
    .PARAMETER RefreshActive
        Switch to refresh the active roles list.
        When specified, updates the list of currently active PIM role assignments.
    
    .PARAMETER RefreshEligible
        Switch to refresh the eligible roles list.
        When specified, updates the list of available PIM roles that can be activated.
    
    .PARAMETER SplashForm
        Optional splash screen form to update during loading operations.
        Used to display progress information during role data retrieval.
    
    .EXAMPLE
        Update-PIMRolesList -Form $mainForm -RefreshActive -RefreshEligible
        Refreshes both active and eligible role lists in the specified form.
    
    .EXAMPLE
        Update-PIMRolesList -Form $mainForm -RefreshEligible -SplashForm $splash
        Refreshes only the eligible roles list while updating the splash screen progress.
    
    .NOTES
        This function is part of the PIM Activation module's UI layer.
        It relies on Get-PIMActiveRoles and Get-PIMEligibleRoles for data retrieval.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form,
        
        [switch]$RefreshActive,
        
        [switch]$RefreshEligible,
        
        [PSCustomObject]$SplashForm
    )
    
    Write-Verbose "Starting Update-PIMRolesList - Active: $RefreshActive, Eligible: $RefreshEligible"
    
    # Show operation splash if not provided (for manual refresh)
    $ownSplash = $false
    if (-not $PSBoundParameters.ContainsKey('SplashForm') -or -not $SplashForm) {
        $SplashForm = Show-OperationSplash -Title "Refreshing Roles" -InitialMessage "Fetching role data..." -ShowProgressBar $true
        $ownSplash = $true
    }

    try {
        # Check if we can use cached role data
        $useCachedData = $false
        $currentTime = Get-Date
        
        if ($script:LastRoleFetchTime -and $script:CachedEligibleRoles -and $script:CachedActiveRoles) {
            $cacheAge = ($currentTime - $script:LastRoleFetchTime).TotalMinutes
            if ($cacheAge -lt $script:RoleCacheValidityMinutes) {
                $useCachedData = $true
                Write-Verbose "Using cached role data (age: $([Math]::Round($cacheAge, 1)) minutes)"
            } else {
                Write-Verbose "Cache expired (age: $([Math]::Round($cacheAge, 1)) minutes), fetching fresh data"
            }
        } else {
            Write-Verbose "No valid cached data available, performing fresh fetch"
        }
        
        # Use batch fetching if both active and eligible roles are requested AND cache is invalid
        $batchResult = $null
        if (($RefreshActive -and $RefreshEligible) -and -not $useCachedData) {
            Write-Verbose "Performing batch role and policy fetch for both active and eligible roles..."
            
            # Update splash screen
            if ($SplashForm -and -not $SplashForm.IsDisposed) {
                if ($ownSplash) {
                    $SplashForm.UpdateStatus("Starting batch role and policy fetch...", 5)
                } else {
                    Update-LoadingStatus -SplashForm $SplashForm -Status "Starting batch role and policy fetch..." -Progress 5
                }
            }
            
            # Batch fetch everything
            $batchParams = @{
                UserId = $script:CurrentUser.Id
                IncludeEntraRoles = $script:IncludeEntraRoles
                IncludeGroups = $script:IncludeGroups
                IncludeAzureResources = $script:IncludeAzureResources
                SplashForm = $SplashForm
            }
            
            $batchResult = Get-PIMRolesBatch @batchParams
            
            # Cache the fetched role data
            $script:CachedEligibleRoles = $batchResult.EligibleRoles
            $script:CachedActiveRoles = $batchResult.ActiveRoles
            $script:LastRoleFetchTime = $currentTime
            Write-Verbose "Cached $($batchResult.EligibleRoles.Count) eligible and $($batchResult.ActiveRoles.Count) active roles"
            
            # Update the script-level policy cache (merge instead of replace to maintain cache persistence)
            if ($batchResult.Policies) {
                $newPolicyCount = 0
                $updatedPolicyCount = 0
                foreach ($key in $batchResult.Policies.Keys) {
                    if ($script:PolicyCache.ContainsKey($key)) {
                        $script:PolicyCache[$key] = $batchResult.Policies[$key]
                        $updatedPolicyCount++
                    } else {
                        $script:PolicyCache[$key] = $batchResult.Policies[$key]
                        $newPolicyCount++
                    }
                }
                Write-Verbose "Policy cache updated: $newPolicyCount new, $updatedPolicyCount updated (total: $($script:PolicyCache.Keys.Count))"
            }
            
            # Update authentication context cache (merge instead of replace)
            if ($batchResult.AuthenticationContexts) {
                $newAuthCount = 0
                $updatedAuthCount = 0
                foreach ($key in $batchResult.AuthenticationContexts.Keys) {
                    if ($script:AuthenticationContextCache.ContainsKey($key)) {
                        $script:AuthenticationContextCache[$key] = $batchResult.AuthenticationContexts[$key]
                        $updatedAuthCount++
                    } else {
                        $script:AuthenticationContextCache[$key] = $batchResult.AuthenticationContexts[$key]
                        $newAuthCount++
                    }
                }
                Write-Verbose "Auth context cache updated: $newAuthCount new, $updatedAuthCount updated (total: $($script:AuthenticationContextCache.Keys.Count))"
            }
        }
        
        # Process active roles if requested
        if ($RefreshActive) {
            Write-Verbose "Refreshing active roles list"
            
            try {
                # Update splash screen progress
                if ($SplashForm -and -not $SplashForm.IsDisposed) {
                    if ($ownSplash) {
                        $progressValue = if ($batchResult) { 96 } else { 25 }
                        $statusMessage = if ($batchResult) { "Processing active roles..." } else { "Fetching active roles..." }
                        $SplashForm.UpdateStatus($statusMessage, $progressValue)
                    } else {
                        # Continuing from Initialize-PIMForm (85%) - progress to higher values
                        $progressValue = if ($batchResult) { 96 } else { 87 }
                        $statusMessage = if ($batchResult) { "Processing active roles..." } else { "Fetching active roles..." }
                        Update-LoadingStatus -SplashForm $SplashForm -Status $statusMessage -Progress $progressValue
                    }
                }
                
                # Locate the active roles ListView control
                $activeListView = $Form.Controls.Find("lstActive", $true)[0]
                
                if ($activeListView) {
                    # Suspend UI updates for better performance
                    $activeListView.BeginUpdate()
                    
                    try {
                        # Clear existing items
                        $activeListView.Items.Clear()
                        
                        # Get active roles from batch result, cached data, or individual fetch
                        $activeRoles = if ($batchResult) {
                            Write-Verbose "Using active roles from batch result"
                            $batchResult.ActiveRoles
                        } elseif ($useCachedData) {
                            Write-Verbose "Using cached active roles data"
                            $script:CachedActiveRoles
                        } else {
                            # Retrieve active role assignments using traditional method
                            Write-Verbose "Fetching active roles from Azure (individual fetch)"
                            Get-PIMActiveRoles
                        }
                        
                        # Ensure we have an array to work with
                        if ($null -eq $activeRoles) {
                            $activeRoles = @()
                        }
                        elseif ($activeRoles -isnot [array]) {
                            $activeRoles = @($activeRoles)
                        }
                        
                        Write-Verbose "Processing $($activeRoles.Count) active roles"
                        
                        # Update splash screen with role count
                        if ($PSBoundParameters.ContainsKey('SplashForm') -and $SplashForm -and -not $SplashForm.IsDisposed) {
                            Update-LoadingStatus -SplashForm $SplashForm -Status "Processing $($activeRoles.Count) active roles..." -Progress 75
                        }
                        
                        $itemIndex = 0
                        foreach ($role in $activeRoles) {
                            try {
                                # Create new ListView item
                                $item = New-Object System.Windows.Forms.ListViewItem
                                
                                # Column 0: Role Type
                                $typePrefix = switch ($role.Type) {
                                    'Entra' { '[Entra]' }
                                    'Group' { '[Group]' }
                                    'AzureResource' { '[Azure]' }
                                    default { "[$($role.Type)]" }
                                }
                                $item.Text = $typePrefix
                                
                                # Column 1: Role Name
                                $item.SubItems.Add($role.DisplayName) | Out-Null
                                
                                # Column 2: Resource - show assignment source
                                $resourceName = "Entra ID Directory"  # Default for Entra roles
                                if ($role.Type -eq 'Entra') {
                                    # Check if this is scoped to an Administrative Unit
                                    if ($role.PSObject.Properties['Scope'] -and $role.Scope -and 
                                        $role.Scope -ne "Directory" -and $role.Scope -ne "Unknown Scope" -and
                                        $role.Scope.StartsWith("AU: ")) {
                                        # Administrative Unit scope - show AU name as resource
                                        $resourceName = $role.Scope  # e.g., "AU: Finance Department"
                                    } 
                                    # Check if ResourceName was explicitly set by batch processing (group attribution)
                                    elseif ($role.PSObject.Properties['ResourceName'] -and $role.ResourceName -and 
                                            $role.ResourceName -ne "Entra ID Directory") {
                                        $resourceName = $role.ResourceName  # e.g., "Entra ID (via Group: GroupName)"
                                    } 
                                    else {
                                        # Default case - direct assignment to directory
                                        $resourceName = "Entra ID Directory"
                                    }
                                } elseif ($role.Type -eq 'Group') {
                                    # For groups, show the group name
                                    if ($role.PSObject.Properties['ResourceName'] -and $role.ResourceName) {
                                        $resourceName = $role.ResourceName
                                    } else {
                                        $resourceName = "PIM Group"
                                    }
                                } elseif ($role.Type -eq 'AzureResource') {
                                    $resourceName = if ($role.PSObject.Properties['ResourceDisplayName'] -and $role.ResourceDisplayName) { 
                                        $role.ResourceDisplayName 
                                    } else { 
                                        "Azure Resource" 
                                    }
                                }
                                $item.SubItems.Add($resourceName) | Out-Null
                                
                                # Column 3: Scope - improve scope detection
                                $scopeDisplay = "Directory"  # Default for Entra roles
                                if ($role.Type -eq 'Entra') {
                                    # For Entra roles, always show "Directory" 
                                    # (AU information is now shown in Resource column)
                                    $scopeDisplay = "Directory"
                                } elseif ($role.Type -eq 'Group') {
                                    # For groups, show the determined scope (Directory or AU)
                                    if ($role.PSObject.Properties['Scope'] -and $role.Scope -and $role.Scope -ne "Group") {
                                        $scopeDisplay = $role.Scope
                                    } else {
                                        $scopeDisplay = "Directory"  # Default for groups
                                    }
                                } elseif ($role.Type -eq 'AzureResource') {
                                    $scopeDisplay = if ($role.PSObject.Properties['ScopeDisplayName'] -and $role.ScopeDisplayName -and $role.ScopeDisplayName -ne "Unknown Scope") {
                                        $role.ScopeDisplayName
                                    } else {
                                        "Subscription/Resource"
                                    }
                                } else {
                                    # Fallback: use the scope if available and not "Unknown Scope"
                                    if ($role.PSObject.Properties['Scope'] -and $role.Scope -and $role.Scope -ne "Unknown Scope") {
                                        $scopeDisplay = $role.Scope
                                    }
                                }
                                $item.SubItems.Add($scopeDisplay) | Out-Null
                                
                                # Column 4: Member Type - show correct role type for each context
                                $memberType = "Direct"  # Default for Entra roles
                                if ($role.Type -eq 'Group') {
                                    # For groups, show Member or Owner based on AccessId
                                    if ($role.PSObject.Properties['AccessId'] -and $role.AccessId) {
                                        $memberType = switch ($role.AccessId.ToLower()) {
                                            'member' { 'Member' }
                                            'owner' { 'Owner' }
                                            default { $role.AccessId }
                                        }
                                    } elseif ($role.PSObject.Properties['MemberType'] -and $role.MemberType) {
                                        $memberType = $role.MemberType
                                    }
                                } elseif ($role.Type -eq 'Entra') {
                                    # For Entra roles, show assignment method (Direct or Group)
                                    if ($role.PSObject.Properties['MemberType'] -and $role.MemberType -and $role.MemberType -ne "Unknown") {
                                        $memberType = $role.MemberType
                                    }
                                } else {
                                    # For other types, use existing logic
                                    if ($role.PSObject.Properties['MemberType'] -and $role.MemberType -and $role.MemberType -ne "Unknown") {
                                        $memberType = $role.MemberType
                                    } elseif ($role.PSObject.Properties['AssignmentType'] -and $role.AssignmentType) {
                                        $memberType = $role.AssignmentType
                                    }
                                }
                                $item.SubItems.Add($memberType) | Out-Null
                                
                                # Column 5: Expiration Time
                                $expiresText = "Permanent"  # Default for roles without expiration
                                if ($role.EndDateTime) {
                                    try {
                                        # Parse the expiration time
                                        $endTime = if ($role.EndDateTime -is [DateTime]) {
                                            $role.EndDateTime
                                        } else {
                                            [DateTime]::Parse($role.EndDateTime)
                                        }
                                        
                                        # Ensure UTC comparison
                                        if ($endTime.Kind -ne [DateTimeKind]::Utc) {
                                            $endTime = $endTime.ToUniversalTime()
                                        }
                                        
                                        # Display in local time
                                        $expiresText = $endTime.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
                                        
                                        # Calculate time remaining
                                        $now = [DateTime]::UtcNow
                                        $remaining = $endTime - $now
                                        
                                        Write-Verbose "Role '$($role.DisplayName)' expires in $([Math]::Round($remaining.TotalMinutes, 0)) minutes"
                                        
                                        # Apply color coding based on time remaining
                                        if ($remaining.TotalMinutes -le 30) {
                                            $item.ForeColor = [System.Drawing.Color]::FromArgb(242, 80, 34)  # Red - expiring soon
                                        }
                                        elseif ($remaining.TotalHours -le 2) {
                                            $item.ForeColor = [System.Drawing.Color]::FromArgb(255, 140, 0)  # Orange - warning
                                        }
                                        else {
                                            $item.ForeColor = [System.Drawing.Color]::FromArgb(0, 78, 146)  # Dark blue - normal
                                        }
                                    }
                                    catch {
                                        Write-Verbose "Failed to parse expiration time for role '$($role.DisplayName)': $_"
                                        $expiresText = "Parse Error"
                                    }
                                }
                                $item.SubItems.Add($expiresText) | Out-Null
                                
                                # Apply alternating row colors for better readability
                                if ($itemIndex % 2 -eq 1) {
                                    $item.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)  # Light gray
                                }
                                else {
                                    $item.BackColor = [System.Drawing.Color]::White
                                }
                                
                                # Store the role object for later retrieval
                                $item.Tag = $role
                                
                                # Add item to ListView
                                $activeListView.Items.Add($item) | Out-Null
                                $itemIndex++
                            }
                            catch {
                                Write-Warning "Failed to add active role '$($role.DisplayName)' to list: $_"
                            }
                        }
                    }
                    finally {
                        # Resume UI updates
                        $activeListView.EndUpdate()
                    }
                    
                    # Auto-size columns to fit content
                    foreach ($column in $activeListView.Columns) {
                        $column.Width = -2  # Auto-size to content
                    }
                    
                    # Ensure column headers are fully visible
                    $graphics = $activeListView.CreateGraphics()
                    $font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

                    for ($i = 0; $i -lt $activeListView.Columns.Count; $i++) {
                        $column = $activeListView.Columns[$i]
                        $headerText = $column.Text
                        
                        # Calculate minimum width based on header text
                        $textSize = $graphics.MeasureString($headerText, $font)
                        $minWidth = [int]$textSize.Width + 20  # Add padding
                        
                        # Ensure specific columns have adequate width
                        if ($headerText -eq "Max Duration" -or $headerText -eq "Justification") {
                            $minWidth = [Math]::Max($minWidth, 110)
                        }
                        
                        if ($column.Width -lt $minWidth) {
                            $column.Width = $minWidth
                        }
                    }
                    
                    # Update the active role count label
                    $activePanel = $Form.Controls.Find('pnlActive', $true)[0]
                    if ($activePanel) {
                        # Find the header panel containing the count label
                        $headerPanel = $activePanel.Controls | Where-Object { 
                            $_ -is [System.Windows.Forms.Panel] -and 
                            $_.BackColor.ToArgb() -eq [System.Drawing.Color]::FromArgb(0, 120, 212).ToArgb() 
                        }
                        if ($headerPanel) {
                            $lblCount = $headerPanel.Controls['lblActiveCount']
                            if ($lblCount) {
                                $lblCount.Text = "$($activeRoles.Count) roles active"
                                Write-Verbose "Updated active role count: $($activeRoles.Count)"
                            }
                        }
                    }
                }
            }
            catch {
                Write-Error "Failed to update active roles list: $_"
                
                # Display error in the ListView
                $activeListView = $Form.Controls.Find("lstActive", $true)[0]
                if ($activeListView) {
                    $activeListView.Items.Clear()
                    $errorItem = New-Object System.Windows.Forms.ListViewItem
                    $errorItem.Text = "Error"
                    $errorItem.SubItems.Add("Failed to load active roles") | Out-Null
                    $errorItem.SubItems.Add($_.ToString()) | Out-Null
                    $errorItem.SubItems.Add("") | Out-Null  # Empty scope
                    $errorItem.SubItems.Add("") | Out-Null  # Empty member type
                    $errorItem.SubItems.Add("") | Out-Null  # Empty expires
                    $errorItem.ForeColor = [System.Drawing.Color]::Red
                    $activeListView.Items.Add($errorItem) | Out-Null
                }
            }
        }
        
        # Process eligible roles if requested
        if ($RefreshEligible) {
            Write-Verbose "Refreshing eligible roles list"
            
            try {
                # Update splash screen progress
                if ($SplashForm -and -not $SplashForm.IsDisposed) {
                    if ($ownSplash) {
                        $statusMessage = if ($batchResult) { "Processing eligible roles..." } else { "Fetching eligible roles and policies..." }
                        $progressValue = if ($batchResult) { 97 } else { 75 }
                        $SplashForm.UpdateStatus($statusMessage, $progressValue)
                    } else {
                        # Continuing from Initialize-PIMForm - progress to higher values
                        $progressValue = if ($batchResult) { 97 } else { 90 }
                        $statusMessage = if ($batchResult) { "Processing eligible roles..." } else { "Fetching eligible roles and policies..." }
                        Update-LoadingStatus -SplashForm $SplashForm -Status $statusMessage -Progress $progressValue
                    }
                }
                
                # Locate the eligible roles ListView control
                $eligibleListView = $Form.Controls.Find("lstEligible", $true)[0]
                
                if ($eligibleListView) {
                    # Suspend UI updates for better performance
                    $eligibleListView.BeginUpdate()
                    
                    try {
                        # Clear existing items
                        $eligibleListView.Items.Clear()
                        
                        # Get eligible roles from batch result, cached data, or individual fetch
                        $eligibleRoles = if ($batchResult) {
                            Write-Verbose "Using eligible roles from batch result"
                            Write-Verbose "Batch result contains $($batchResult.EligibleRoles.Count) eligible roles"
                            $batchResult.EligibleRoles
                        } elseif ($useCachedData) {
                            Write-Verbose "Using cached eligible roles data"
                            Write-Verbose "Cache contains $($script:CachedEligibleRoles.Count) eligible roles"
                            $script:CachedEligibleRoles
                        } else {
                            # Retrieve eligible role assignments using traditional method
                            Write-Verbose "Fetching eligible roles and policies from Azure (individual fetch)"
                            Get-PIMEligibleRoles
                        }
                        
                        # Get pending requests to show which roles have pending activations
                        try {
                            # First, ensure we have a Graph connection
                            $graphContext = Get-MgContext
                            if (-not $graphContext) {
                                # Try to reconnect using PIM services
                                $connectionResult = Connect-PIMServices -IncludeEntraRoles -IncludeGroups
                            }
                            
                            $pendingRequests = if ($script:CurrentUser -and $script:CurrentUser.Id) {
                                Get-PIMPendingRequests -UserId $script:CurrentUser.Id
                            } else {
                                Get-PIMPendingRequests
                            }
                        } catch {
                            Write-Warning "Failed to retrieve pending requests: $($_.Exception.Message)"
                            $pendingRequests = @()
                        }
                        if (-not $pendingRequests) {
                            $pendingRequests = @()
                        } elseif ($pendingRequests -isnot [array]) {
                            $pendingRequests = @($pendingRequests)
                        }
                        Write-Verbose "Found $($pendingRequests.Count) pending role requests"
                        
                        # Debug: List pending request details
                        foreach ($pr in $pendingRequests) {
                            if ($pr.Type -eq 'Group') {
                                Write-Verbose "Pending request: Type=$($pr.Type), GroupId=$($pr.GroupId), RoleName=$($pr.RoleName)"
                            } else {
                                Write-Verbose "Pending request: Type=$($pr.Type), RoleDefinitionId=$($pr.RoleDefinitionId), RoleName=$($pr.RoleName)"
                            }
                        }
                        
                        # Ensure we have an array to work with
                        if ($null -eq $eligibleRoles) {
                            $eligibleRoles = @()
                        }
                        elseif ($eligibleRoles -isnot [array]) {
                            $eligibleRoles = @($eligibleRoles)
                        }
                        
                        Write-Verbose "Processing $($eligibleRoles.Count) eligible roles"
                        
                        # Debug: Show details of first few roles
                        for ($i = 0; $i -lt [Math]::Min(3, $eligibleRoles.Count); $i++) {
                            $debugRole = $eligibleRoles[$i]
                            Write-Verbose "Role $($i + 1): DisplayName='$($debugRole.DisplayName)', Type='$($debugRole.Type)', Scope='$($debugRole.Scope)'"
                        }
                        
                        # Update splash screen with role count
                        if ($PSBoundParameters.ContainsKey('SplashForm') -and $SplashForm -and -not $SplashForm.IsDisposed) {
                            Update-LoadingStatus -SplashForm $SplashForm -Status "Processing $($eligibleRoles.Count) eligible roles..." -Progress 85
                        }
                        
                        $itemIndex = 0
                        $totalRoles = $eligibleRoles.Count
                        $progressBase = 85
                        $progressRange = 10  # Progress range: 85-95%
                        
                        foreach ($role in $eligibleRoles) {
                            try {
                                # Update progress for each role (simplified progress for batch results)
                                if ($PSBoundParameters.ContainsKey('SplashForm') -and $SplashForm -and -not $SplashForm.IsDisposed -and $totalRoles -gt 0) {
                                    $currentProgress = $progressBase + [int](($itemIndex / $totalRoles) * $progressRange)
                                    $statusMessage = if ($batchResult) { "Processing role $($role.DisplayName)..." } else { "Fetching policy for $($role.DisplayName)..." }
                                    Update-LoadingStatus -SplashForm $SplashForm -Status $statusMessage -Progress $currentProgress
                                }
                                
                                # Retrieve or use existing policy information
                                $policyInfo = $null
                                if ($role.PSObject.Properties['PolicyInfo'] -and $role.PolicyInfo) {
                                    # Role already has policy info attached (from batch fetch)
                                    $policyInfo = $role.PolicyInfo
                                    Write-Verbose "Using attached policy info for $($role.DisplayName)"
                                }
                                else {
                                    # Fallback: Try cache lookup only if we don't have attached policy info
                                    Write-Verbose "No attached policy info for $($role.DisplayName), checking cache..."
                                    
                                    if ($script:PolicyCache -and $script:PolicyCache.Count -gt 0) {
                                        # Get policy from script-level cache
                                        $roleId = if ($role.Type -eq 'Group') { 
                                            if ($role.PSObject.Properties['GroupId']) { $role.GroupId } else { $null }
                                        } elseif ($role.PSObject.Properties['RoleDefinitionId']) {
                                            $role.RoleDefinitionId
                                        } elseif ($role.PSObject.Properties['Id']) {
                                            $role.Id
                                        } else {
                                            $null
                                        }
                                        
                                        if ($roleId) {
                                            $cacheKey = "Entra_$roleId"
                                            if ($role.Type -eq 'Group') {
                                                $cacheKey = "Group_$roleId"
                                            }
                                            
                                            if ($script:PolicyCache.ContainsKey($cacheKey)) {
                                                $policyInfo = $script:PolicyCache[$cacheKey]
                                                Write-Verbose "Using cached policy for $($role.DisplayName) (key: $cacheKey)"
                                            }
                                            else {
                                                Write-Verbose "Policy not found in cache for key: $cacheKey"
                                                # Only fetch individually if not in a batch operation
                                                if (-not $batchResult) {
                                                    Write-Verbose "Fetching policy individually for $($role.DisplayName)"
                                                    $policyInfo = Get-PIMRolePolicy -Role $role
                                                } else {
                                                    Write-Verbose "Skipping individual policy fetch for $($role.DisplayName) - batch operation should have provided policy"
                                                }
                                            }
                                        } else {
                                            Write-Verbose "Could not determine role ID for cache lookup: $($role.DisplayName)"
                                            # Only fetch individually if not in a batch operation
                                            if (-not $batchResult) {
                                                Write-Verbose "Fetching policy individually for $($role.DisplayName)"
                                                $policyInfo = Get-PIMRolePolicy -Role $role
                                            } else {
                                                Write-Verbose "Skipping individual policy fetch for $($role.DisplayName) - batch operation should have provided policy"
                                            }
                                        }
                                    }
                                    else {
                                        # Individual fetch (fallback for partial refreshes when no cache available)
                                        if (-not $batchResult) {
                                            Write-Verbose "No policy cache available, fetching individually for $($role.DisplayName)"
                                            $policyInfo = Get-PIMRolePolicy -Role $role
                                        } else {
                                            Write-Verbose "No policy cache and batch operation - this shouldn't happen"
                                        }
                                    }
                                }
                                
                                # If we have cached authentication context info, enhance the policy
                                if ($policyInfo -and $policyInfo.RequiresAuthenticationContext -and $policyInfo.AuthenticationContextId) {
                                    if ($script:AuthenticationContextCache -and $script:AuthenticationContextCache.ContainsKey($policyInfo.AuthenticationContextId)) {
                                        $authContext = $script:AuthenticationContextCache[$policyInfo.AuthenticationContextId]
                                        $policyInfo.AuthenticationContextDisplayName = $authContext.DisplayName
                                        $policyInfo.AuthenticationContextDescription = $authContext.Description
                                        $policyInfo.AuthenticationContextDetails = $authContext
                                    }
                                }
                                
                                Write-Verbose "Creating ListView item for role: $($role.DisplayName) (Type: $($role.Type))"
                                
                                # Create new ListView item
                                $item = New-Object System.Windows.Forms.ListViewItem
                                
                                # Column 0: Role Name with type prefix and pending status
                                $typePrefix = switch ($role.Type) {
                                    'Entra' { '[Entra]' }
                                    'Group' { '[Group]' }
                                    'AzureResource' { '[Azure]' }
                                    default { "[$($role.Type)]" }
                                }
                                
                                # Check if this role has a pending activation request
                                $hasPendingRequest = $false
                                if ($role.Type -eq 'Entra') {
                                    # Check if the role has the required property
                                    if ($role.PSObject.Properties['RoleDefinitionId']) {
                                        $pendingMatch = $pendingRequests | Where-Object { 
                                            $_.Type -eq 'Entra' -and 
                                            $_.PSObject.Properties['RoleDefinitionId'] -and 
                                            $_.RoleDefinitionId -eq $role.RoleDefinitionId 
                                        } | Select-Object -First 1
                                        $hasPendingRequest = [bool]$pendingMatch
                                        if ($pendingMatch) {
                                            Write-Verbose "Found pending request for Entra role: $($role.DisplayName) (ID: $($role.RoleDefinitionId))"
                                        }
                                    } else {
                                        Write-Verbose "Entra role $($role.DisplayName) is missing RoleDefinitionId property - skipping pending check"
                                    }
                                } elseif ($role.Type -eq 'Group') {
                                    # Check if the role has the required property
                                    if ($role.PSObject.Properties['GroupId']) {
                                        $pendingMatch = $pendingRequests | Where-Object { 
                                            $_.Type -eq 'Group' -and 
                                            $_.PSObject.Properties['GroupId'] -and 
                                            $_.GroupId -eq $role.GroupId 
                                        } | Select-Object -First 1
                                        $hasPendingRequest = [bool]$pendingMatch
                                        if ($pendingMatch) {
                                            Write-Verbose "Found pending request for Group role: $($role.DisplayName) (ID: $($role.GroupId))"
                                        }
                                    } else {
                                        Write-Verbose "Group role $($role.DisplayName) is missing GroupId property - skipping pending check"
                                    }
                                }
                                
                                $item.Text = "$typePrefix $($role.DisplayName)"
                                
                                # Column 1: Scope - show permission scope level (Directory, AU, or Group)
                                $scopeDisplay = "Directory"  # Default for Entra roles
                                if ($role.Type -eq 'Entra') {
                                    # For Entra roles, show the scope where permissions apply
                                    if ($role.PSObject.Properties['DirectoryScopeId'] -and $role.DirectoryScopeId) {
                                        # Use the DirectoryScopeId to determine scope
                                        if ($role.DirectoryScopeId -eq '/' -or [string]::IsNullOrEmpty($role.DirectoryScopeId)) {
                                            $scopeDisplay = "Directory"
                                        } elseif ($role.DirectoryScopeId -like '/administrativeUnits/*') {
                                            # Try to get AU display name from Scope property first
                                            if ($role.PSObject.Properties['Scope'] -and $role.Scope -and $role.Scope.StartsWith('AU: ')) {
                                                $scopeDisplay = $role.Scope
                                            } else {
                                                # Fallback to AU ID
                                                $auId = $role.DirectoryScopeId -replace '^/administrativeUnits/', ''
                                                $scopeDisplay = "AU: $auId"
                                            }
                                        } else {
                                            # Other scopes - use Scope property if available
                                            if ($role.PSObject.Properties['Scope'] -and $role.Scope -and $role.Scope -ne "Unknown Scope") {
                                                $scopeDisplay = $role.Scope
                                            } else {
                                                $scopeDisplay = "Directory"
                                            }
                                        }
                                    } else {
                                        # Fallback to Scope property
                                        if ($role.PSObject.Properties['Scope'] -and $role.Scope -and $role.Scope -ne "Unknown Scope") {
                                            $scopeDisplay = $role.Scope
                                        } else {
                                            $scopeDisplay = "Directory"
                                        }
                                    }
                                } elseif ($role.Type -eq 'Group') {
                                    # For groups, show the determined scope (Directory or AU)
                                    if ($role.PSObject.Properties['Scope'] -and $role.Scope -and $role.Scope -ne "Group") {
                                        $scopeDisplay = $role.Scope
                                    } else {
                                        $scopeDisplay = "Directory"  # Default for groups
                                    }
                                } elseif ($role.Type -eq 'AzureResource') {
                                    $scopeDisplay = if ($role.PSObject.Properties['ScopeDisplayName'] -and $role.ScopeDisplayName -and $role.ScopeDisplayName -ne "Unknown Scope") {
                                        $role.ScopeDisplayName
                                    } else {
                                        "Subscription/Resource"
                                    }
                                } else {
                                    # Fallback: use the scope if available and not "Unknown Scope"
                                    if ($role.PSObject.Properties['Scope'] -and $role.Scope -and $role.Scope -ne "Unknown Scope") {
                                        $scopeDisplay = $role.Scope
                                    }
                                }
                                $item.SubItems.Add($scopeDisplay) | Out-Null
                                
                                # Column 2: MemberType - show correct role type for each context
                                $memberType = "Direct"  # Default for Entra roles
                                if ($role.Type -eq 'Group') {
                                    # For groups, show Member or Owner based on AccessId
                                    if ($role.PSObject.Properties['AccessId'] -and $role.AccessId) {
                                        $memberType = switch ($role.AccessId.ToLower()) {
                                            'member' { 'Member' }
                                            'owner' { 'Owner' }
                                            default { $role.AccessId }
                                        }
                                    } elseif ($role.PSObject.Properties['MemberType'] -and $role.MemberType) {
                                        $memberType = $role.MemberType
                                    }
                                } elseif ($role.Type -eq 'Entra') {
                                    # For Entra roles, show assignment method (Direct or Group)
                                    if ($role.PSObject.Properties['MemberType'] -and $role.MemberType -and $role.MemberType -ne "Unknown") {
                                        $memberType = $role.MemberType
                                    }
                                } else {
                                    # For other types, use existing logic
                                    if ($role.PSObject.Properties['MemberType'] -and $role.MemberType -and $role.MemberType -ne "Unknown") {
                                        $memberType = $role.MemberType
                                    } elseif ($role.PSObject.Properties['AssignmentType'] -and $role.AssignmentType) {
                                        $memberType = $role.AssignmentType
                                    }
                                }
                                $item.SubItems.Add($memberType) | Out-Null

                                # Column 3: Maximum activation duration
                                $maxDuration = "8h"
                                if ($policyInfo -and $policyInfo.MaxDuration) {
                                    $hours = $policyInfo.MaxDuration
                                    $maxDuration = "${hours}h"
                                }
                                $item.SubItems.Add($maxDuration) | Out-Null
                                
                                # Column 4: MFA requirement
                                $mfaRequired = if ($policyInfo -and $policyInfo.RequiresMFA) { "Yes" } else { "No" }
                                $item.SubItems.Add($mfaRequired) | Out-Null
                                
                                # Column 5: Authentication context requirement
                                $authContext = if ($policyInfo -and $policyInfo.RequiresAuthenticationContext) { "Required" } else { "No" }
                                $item.SubItems.Add($authContext) | Out-Null
                                
                                # Column 6: Justification requirement
                                $justification = if ($policyInfo -and $policyInfo.RequiresJustification) { "Required" } else { "No" }
                                $item.SubItems.Add($justification) | Out-Null
                                
                                # Column 7: Ticket requirement
                                $ticket = if ($policyInfo -and $policyInfo.RequiresTicket) { "Yes" } else { "No" }
                                $item.SubItems.Add($ticket) | Out-Null
                                
                                # Column 8: Approval requirement
                                $approval = if ($policyInfo -and $policyInfo.RequiresApproval) { "Required" } else { "No" }
                                $item.SubItems.Add($approval) | Out-Null
                                
                                # Column 9: Pending Approval status
                                $pendingApproval = if ($hasPendingRequest) { "Yes" } else { "No" }
                                $item.SubItems.Add($pendingApproval) | Out-Null
                                
                                # Apply alternating row colors for better readability
                                if ($itemIndex % 2 -eq 1) {
                                    $item.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)  # Light gray
                                }
                                else {
                                    $item.BackColor = [System.Drawing.Color]::White
                                }
                                
                                # Apply color coding based on policy requirements
                                if ($policyInfo) {
                                    if ($policyInfo.RequiresApproval) {
                                        $item.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)  # Microsoft blue - requires approval
                                    }
                                    elseif ($policyInfo.RequiresAuthenticationContext) {
                                        $item.ForeColor = [System.Drawing.Color]::FromArgb(0, 78, 146)  # Dark blue - enhanced security
                                    }
                                    else {
                                        $item.ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)  # Default dark gray
                                    }
                                }
                                
                                # Store the role object with policy info for later retrieval
                                $roleWithPolicy = $role | Add-Member -NotePropertyName PolicyInfo -NotePropertyValue $policyInfo -PassThru -Force
                                $item.Tag = $roleWithPolicy
                                
                                # Add item to ListView
                                Write-Verbose "Adding ListView item for: $($role.DisplayName)"
                                $eligibleListView.Items.Add($item) | Out-Null
                                Write-Verbose "Successfully added item $($itemIndex + 1) for: $($role.DisplayName)"
                                $itemIndex++
                            }
                            catch {
                                Write-Warning "Failed to add eligible role '$($role.DisplayName)' to list: $_"
                                Write-Verbose "Exception details: $($_.Exception.Message)"
                                Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
                            }
                        }
                    }
                    finally {
                        # Resume UI updates
                        $eligibleListView.EndUpdate()
                    }
                    
                    Write-Verbose "Added $($eligibleListView.Items.Count) items to eligible roles list"
                    
                    # Auto-size columns to fit content
                    foreach ($column in $eligibleListView.Columns) {
                        $column.Width = -2  # Auto-size to content
                    }
                    
                    # Ensure column headers are fully visible
                    $graphics = $eligibleListView.CreateGraphics()
                    $font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
                    
                    for ($i = 0; $i -lt $eligibleListView.Columns.Count; $i++) {
                        $column = $eligibleListView.Columns[$i]
                        $headerText = $column.Text
                        
                        # Calculate minimum width based on header text
                        $textSize = $graphics.MeasureString($headerText, $font)
                        $minWidth = [int]$textSize.Width + 20  # Add padding
                        
                        # Ensure specific columns have adequate width
                        if ($headerText -eq "Max Duration" -or $headerText -eq "Justification" -or $headerText -eq "Pending Approval") {
                            $minWidth = [Math]::Max($minWidth, 110)
                        }
                        
                        if ($column.Width -lt $minWidth) {
                            $column.Width = $minWidth
                        }
                    }
                    
                    $font.Dispose()
                    $graphics.Dispose()
                    
                    # Update the eligible role count label
                    $eligiblePanel = $Form.Controls.Find('pnlEligible', $true)[0]
                    if ($eligiblePanel) {
                        # Find the header panel containing the count label
                        $headerPanel = $eligiblePanel.Controls | Where-Object { 
                            $_ -is [System.Windows.Forms.Panel] -and 
                            $_.BackColor.ToArgb() -eq [System.Drawing.Color]::FromArgb(91, 203, 255).ToArgb() 
                        }
                        if ($headerPanel) {
                            $lblCount = $headerPanel.Controls['lblEligibleCount']
                            if ($lblCount) {
                                $lblCount.Text = "$($eligibleRoles.Count) roles available"
                                Write-Verbose "Updated eligible role count: $($eligibleRoles.Count)"
                            }
                        }
                    }
                }
            }
            catch {
                Write-Error "Failed to update eligible roles list: $_"
                
                # Display error in the ListView
                $eligibleListView = $Form.Controls.Find("lstEligible", $true)[0]
                if ($eligibleListView) {
                    $eligibleListView.Items.Clear()
                    $errorItem = New-Object System.Windows.Forms.ListViewItem
                    $errorItem.Text = "Error loading eligible roles"
                    $errorItem.SubItems.Add($_.ToString()) | Out-Null
                    $errorItem.SubItems.Add("") | Out-Null  # Empty max duration
                    $errorItem.SubItems.Add("") | Out-Null  # Empty MFA
                    $errorItem.SubItems.Add("") | Out-Null  # Empty auth context
                    $errorItem.SubItems.Add("") | Out-Null  # Empty justification
                    $errorItem.SubItems.Add("") | Out-Null  # Empty ticket
                    $errorItem.SubItems.Add("") | Out-Null  # Empty approval
                    $errorItem.SubItems.Add("") | Out-Null  # Empty pending approval
                    $errorItem.ForeColor = [System.Drawing.Color]::Red
                    $eligibleListView.Items.Add($errorItem) | Out-Null
                }
            }
        }
        
        # Final splash screen update
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            if ($ownSplash) {
                $SplashForm.UpdateStatus("Role data loaded successfully!", 100)
                Start-Sleep -Milliseconds 500
            } else {
                Update-LoadingStatus -SplashForm $SplashForm -Status "Initialization complete!" -Progress 100
                Start-Sleep -Milliseconds 500
                Close-LoadingSplash -SplashForm $SplashForm
            }
        }
    }    
    finally {
        # Close splash if we created it
        if ($ownSplash -and $SplashForm -and -not $SplashForm.IsDisposed) {
            $SplashForm.Close()
        }
    }

    Write-Verbose "Update-PIMRolesList completed successfully"
}