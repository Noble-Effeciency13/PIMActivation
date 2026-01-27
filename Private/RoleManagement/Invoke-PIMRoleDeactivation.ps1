function Invoke-PIMRoleDeactivation {
    <#
    .SYNOPSIS
        Deactivates selected active PIM roles.
    
    .DESCRIPTION
        Handles the deactivation of active PIM roles including:
        - Both Entra ID directory roles and PIM-enabled groups
        - Progress tracking with splash screen
        - Comprehensive error handling
    
    .PARAMETER CheckedItems
        Array of checked ListView items representing the active roles to deactivate.
    
    .PARAMETER Form
        Reference to the main form for UI updates.
    
    .EXAMPLE
        Invoke-PIMRoleDeactivation -CheckedItems $selectedRoles -Form $mainForm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$CheckedItems,
        
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form
    )
    
    Write-Verbose "Starting deactivation process for $($CheckedItems.Count) role(s)"
    
    # Initialize splash variable
    $operationSplash = $null
    
    try {
        # Confirm deactivation first (before showing splash)
        $roleNames = @($CheckedItems | ForEach-Object { 
            if ($_.Tag.Scope -and $_.Tag.Scope -ne "Directory") {
                "$($_.Tag.DisplayName) [$($_.Tag.Scope)]"
            }
            else {
                $_.Tag.DisplayName
            }
        })
        $message = "Are you sure you want to deactivate the following role(s)?`n`n$($roleNames -join "`n")"
        
        $result = Show-TopMostMessageBox -Message $message -Title "Confirm Deactivation" -Buttons YesNo -Icon Question
        
        if ($result -ne 'Yes') {
            Write-Verbose "Deactivation cancelled by user"
            return
        }
        
        # Show operation splash AFTER user confirms
        $operationSplash = Show-OperationSplash -Title "Role Deactivation" -InitialMessage "Preparing role deactivation..." -ShowProgressBar $true
        
        # Process deactivations
        $deactivationErrors = @()
        $successCount = 0
        $totalRoles = @($CheckedItems).Count
        $currentRole = 0
        
        foreach ($item in $CheckedItems) {
            try {
                $currentRole++
                $roleData = $item.Tag
                $progressPercent = [int](($currentRole / $totalRoles) * 100)
                
                $scopeInfo = if ($roleData.Scope -and $roleData.Scope -ne "Directory") { " [$($roleData.Scope)]" } else { "" }
                $operationSplash.UpdateStatus("Deactivating $($roleData.DisplayName)$scopeInfo ... ($currentRole of $totalRoles)", $progressPercent)
                
                Write-Verbose "Deactivating role: $($roleData.DisplayName) [Type: $($roleData.Type)]"
                
                # Create cancellation request
                $requestBody = @{
                    principalId   = $script:CurrentUser.Id
                    action        = "selfDeactivate"
                    justification = "Deactivated via PowerShell"
                }
                
                switch ($roleData.Type) {
                    'Entra' {
                        # Find the active assignment schedule ID
                        if ($roleData.ScheduleId) {
                            $requestBody.roleAssignmentScheduleId = $roleData.ScheduleId
                        }
                        else {
                            # Query for the active schedule
                            Write-Verbose "Querying for active role assignment schedules for RoleDefinitionId: $($roleData.RoleDefinitionId)"
                            $activeSchedules = @(Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "principalId eq '$($script:CurrentUser.Id)' and roleDefinitionId eq '$($roleData.RoleDefinitionId)'" -ErrorAction SilentlyContinue)
                            
                            if ($activeSchedules -and $activeSchedules.Count -gt 0) {
                                Write-Verbose "Found $($activeSchedules.Count) active schedule(s), using first one: $($activeSchedules[0].Id)"
                                $requestBody.roleAssignmentScheduleId = $activeSchedules[0].Id
                            }
                            else {
                                throw "Could not find active assignment schedule for deactivation of role: $($roleData.DisplayName)"
                            }
                        }
                        
                        $requestBody.roleDefinitionId = $roleData.RoleDefinitionId
                        $requestBody.directoryScopeId = if ($roleData.DirectoryScopeId) { $roleData.DirectoryScopeId } else { "/" }
                        
                        $response = New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $requestBody -ErrorAction Stop
                        Write-Verbose "Entra role deactivated successfully"
                        $successCount++
                    }
                    
                    'Group' {
                        Write-Verbose "Processing group deactivation for GroupId: $($roleData.GroupId)"
                        
                        # Validate required group data
                        if (-not $roleData.GroupId) {
                            throw "Missing GroupId for group role deactivation: $($roleData.DisplayName)"
                        }
                        
                        $groupRequestBody = @{
                            principalId   = $script:CurrentUser.Id
                            groupId       = $roleData.GroupId
                            action        = "selfDeactivate"
                            justification = "Deactivated via PowerShell"
                            accessId      = "member"
                        }
                        
                        # Find the active assignment schedule ID
                        if ($roleData.ScheduleId) {
                            $groupRequestBody.assignmentScheduleId = $roleData.ScheduleId
                        }
                        else {
                            # Query for the active schedule
                            Write-Verbose "Querying for active group assignment schedules for GroupId: $($roleData.GroupId)"
                            $activeSchedules = @(Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule -Filter "principalId eq '$($script:CurrentUser.Id)' and groupId eq '$($roleData.GroupId)'" -ErrorAction SilentlyContinue)
                            
                            if ($activeSchedules -and $activeSchedules.Count -gt 0) {
                                Write-Verbose "Found $($activeSchedules.Count) active schedule(s), using first one: $($activeSchedules[0].Id)"
                                $groupRequestBody.assignmentScheduleId = $activeSchedules[0].Id
                            }
                            else {
                                throw "Could not find active assignment schedule for group deactivation: $($roleData.DisplayName). The group role may not be currently active or may have been assigned through a different mechanism."
                            }
                        }
                        
                        $response = New-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleRequest -BodyParameter $groupRequestBody -ErrorAction Stop
                        Write-Verbose "Group role deactivated successfully"
                        $successCount++
                    }
                    
                    'AzureResource' {
                        Write-Verbose "Processing Azure Resource role deactivation for scope: $($roleData.FullScope)"

                        # Validate required Azure role data
                        if (-not $roleData.RoleDefinitionId -or -not $roleData.FullScope) {
                            throw "Missing Azure Resource role details (RoleDefinitionId/Scope) for deactivation: $($roleData.DisplayName)"
                        }

                        $roleDefId = if ($roleData.RoleDefinitionId.StartsWith('/')) {
                            $roleData.RoleDefinitionId
                        }
                        else {
                            "$($roleData.FullScope)/providers/Microsoft.Authorization/roleDefinitions/$($roleData.RoleDefinitionId)"
                        }

                        # Build deactivation parameters for Az.Resources
                        $deactivateParams = @{
                            Name             = ([System.Guid]::NewGuid().ToString())
                            Scope            = $roleData.FullScope
                            RoleDefinitionId = $roleDefId
                            PrincipalId      = $script:CurrentUser.Id
                            RequestType      = 'SelfDeactivate'
                            Justification    = 'Deactivated via PowerShell'
                        }

                        try {
                            Write-Verbose "Submitting Azure Resource deactivation using New-AzRoleAssignmentScheduleRequest"
                            $response = New-AzRoleAssignmentScheduleRequest @deactivateParams -ErrorAction Stop
                            Write-Verbose "Azure Resource role deactivation request submitted successfully"
                            $successCount++
                        }
                        catch {
                            throw "Azure Resource deactivation failed: $($_.Exception.Message)"
                        }
                    }
                    
                    default {
                        throw "Unsupported role type: $($roleData.Type)"
                    }
                }
            }
            catch {
                $errorMessage = if ($_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
                
                # Check for  minimum active duration error
                if ($errorMessage -match "Active.*duration.*too short") {
                    # Calculate remaining time if StartDateTime is available
                    $remainingTimeMsg = ""
                    if ($roleData.PSObject.Properties['StartDateTime'] -and $roleData.StartDateTime) {
                        try {
                            $startTime = [DateTime]$roleData.StartDateTime
                            $minimumEndTime = $startTime.AddMinutes(5)
                            $now = [DateTime]::UtcNow
                            if ($minimumEndTime -gt $now) {
                                $remainingSeconds = [int]($minimumEndTime - $now).TotalSeconds
                                $remainingMinutes = [math]::Ceiling($remainingSeconds / 60)
                                $remainingTimeMsg = " Please wait approximately $remainingMinutes minute(s)."
                            }
                        }
                        catch {
                            Write-Verbose "Could not calculate remaining time: $_"
                        }
                    }
                    $errorMessage = "Role must be active for at least 5 minutes before it can be deactivated.$remainingTimeMsg"
                }
                
                # Build role identifier with resource info (same logic as Active Roles list)
                $roleIdentifier = $roleData.DisplayName
                if ($roleData.Scope -and $roleData.Scope -ne "Directory") {
                    $roleIdentifier = "$($roleData.DisplayName) [$($roleData.Scope)]"
                }

                $deactivationErrors += "$roleIdentifier`: $errorMessage"
                Write-Warning "Failed to deactivate $roleIdentifier`: $errorMessage"
            }
        }
        
        $operationSplash.UpdateStatus("Completing deactivation process...", 95)
        
        # Close splash before showing results dialog
        if ($operationSplash -and -not $operationSplash.IsDisposed) {
            $operationSplash.Close()
            $operationSplash = $null
        }
        
        # Display results - always show a message to the user
        $errorCount = @($deactivationErrors).Count
        Write-Verbose "Deactivation complete. Success: $successCount, Errors: $errorCount"
        
        if ($errorCount -gt 0) {
            $message = "Successfully deactivated: $successCount of $totalRoles role(s)`n`nErrors ($errorCount):`n`n$($deactivationErrors -join "`n`n")"
            Show-TopMostMessageBox -Message $message -Title "Deactivation Results" -Icon Warning
        }
        elseif ($successCount -gt 0) {
            Show-TopMostMessageBox -Message "Successfully deactivated all $successCount role(s)!" -Title "Success" -Icon Information
        }
        
        # Clear role cache to ensure fresh data is fetched after deactivation
        if ($successCount -gt 0) {
            Write-Verbose "Waiting for Microsoft Graph to process deactivation changes..."
            Start-Sleep -Seconds 3  # Add delay for Graph propagation
            
            Write-Verbose "Clearing role cache to force fresh data retrieval after deactivation"
            $script:CachedEligibleRoles = $null
            $script:CachedActiveRoles = $null
            $script:LastRoleFetchTime = $null

            # Mark affected Azure subscriptions as dirty for delta refresh and clear any override expirations
            try {
                foreach ($item in $CheckedItems) {
                    $roleData = $item.Tag
                    if ($roleData -and $roleData.Type -eq 'AzureResource') {
                        if (-not (Get-Variable -Name 'DirtyAzureSubscriptions' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyAzureSubscriptions = @() }
                        if ($roleData.PSObject.Properties['SubscriptionId'] -and $roleData.SubscriptionId) {
                            $script:DirtyAzureSubscriptions += $roleData.SubscriptionId
                            $script:DirtyAzureSubscriptions = @($script:DirtyAzureSubscriptions | Select-Object -Unique)
                            Write-Verbose "Marked subscription $($roleData.SubscriptionId) as dirty after deactivation"
                        }
                        # If management group scope, mark MG dirty for delta refresh
                        if ($roleData.PSObject.Properties['FullScope'] -and $roleData.FullScope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                            if (-not (Get-Variable -Name 'DirtyManagementGroups' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyManagementGroups = @() }
                            $mgName = $matches[1]
                            $script:DirtyManagementGroups += $mgName
                            $script:DirtyManagementGroups = @($script:DirtyManagementGroups | Select-Object -Unique)
                            Write-Verbose "Marked management group ${mgName} as dirty after deactivation"
                        }

                        # Remove any Azure active override expiration for this role/scope
                        if (Get-Variable -Name 'AzureActiveOverrides' -Scope Script -ErrorAction SilentlyContinue) {
                            $roleDefKey = $roleData.RoleDefinitionId
                            if ($roleDefKey -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefKey = $matches[1] }
                            $overrideKey = "$($roleData.FullScope)|$($roleDefKey)"
                            if ($script:AzureActiveOverrides.ContainsKey($overrideKey)) {
                                $null = $script:AzureActiveOverrides.Remove($overrideKey)
                                Write-Verbose "Cleared Azure active override for $overrideKey after deactivation"
                            }
                            # Also remove from AzureRolesCache if present
                            if (Get-Variable -Name 'AzureRolesCache' -Scope Script -ErrorAction SilentlyContinue) {
                                $script:AzureRolesCache = @($script:AzureRolesCache | Where-Object {
                                        if ($_.PSObject.Properties['RoleDefinitionId'] -and $_.PSObject.Properties['FullScope']) {
                                            $rd = $_.RoleDefinitionId; if ($rd -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $rd = $matches[1] }
                                            -not ($rd -eq $roleDefKey -and $_.FullScope -eq $roleData.FullScope)
                                        }
                                        else { $true }
                                    })
                                Write-Verbose "Pruned deactivated Azure role from AzureRolesCache for key $overrideKey"
                            }
                        }
                    }
                }
            }
            catch { Write-Verbose "Post-deactivation delta marking failed: $($_.Exception.Message)" }
        }
        
        try {
            # Per refresh semantics: only refresh ACTIVE roles after deactivation
            Update-PIMRolesList -Form $Form -RefreshActive
        }
        catch {
            Write-Warning "Failed to refresh role lists: $_"
        }
        
    }
    finally {
        # Ensure splash is closed
        if ($operationSplash -and -not $operationSplash.IsDisposed) {
            $operationSplash.Close()
        }
    }
    
    Write-Verbose "Deactivation process completed - Success: $successCount, Errors: $errorCount"
}