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
    
    # Show operation splash
    $operationSplash = Show-OperationSplash -Title "Role Deactivation" -InitialMessage "Preparing role deactivation..." -ShowProgressBar $true
    
    try {
        # Confirm deactivation
        $roleNames = @($CheckedItems | ForEach-Object { $_.Tag.DisplayName })
        $message = "Are you sure you want to deactivate the following role(s)?`n`n$($roleNames -join "`n")"
        
        $result = Show-TopMostMessageBox -Message $message -Title "Confirm Deactivation" -Buttons YesNo -Icon Question
        
        if ($result -ne 'Yes') {
            Write-Verbose "Deactivation cancelled by user"
            $operationSplash.Close()
            return
        }
        
        # Process deactivations
        $deactivationErrors = @()
        $successCount = 0
        $totalRoles = $CheckedItems.Count
        $currentRole = 0
        
        foreach ($item in $CheckedItems) {
            try {
                $currentRole++
                $roleData = $item.Tag
                $progressPercent = [int](($currentRole / $totalRoles) * 100)
                
                $operationSplash.UpdateStatus("Deactivating $($roleData.DisplayName)... ($currentRole of $totalRoles)", $progressPercent)
                
                Write-Verbose "Deactivating role: $($roleData.DisplayName) [Type: $($roleData.Type)]"
                
                # Create cancellation request
                $requestBody = @{
                    principalId = $script:CurrentUser.Id
                    action = "selfDeactivate"
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
                            $activeSchedules = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "principalId eq '$($script:CurrentUser.Id)' and roleDefinitionId eq '$($roleData.RoleDefinitionId)'"
                            if ($activeSchedules) {
                                $requestBody.roleAssignmentScheduleId = $activeSchedules[0].Id
                            }
                            else {
                                throw "Could not find active assignment schedule for deactivation"
                            }
                        }
                        
                        $requestBody.roleDefinitionId = $roleData.RoleDefinitionId
                        $requestBody.directoryScopeId = if ($roleData.DirectoryScopeId) { $roleData.DirectoryScopeId } else { "/" }
                        
                        $response = New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $requestBody
                        Write-Verbose "Entra role deactivated successfully"
                    }
                    
                    'Group' {
                        $groupRequestBody = @{
                            principalId = $script:CurrentUser.Id
                            groupId = $roleData.GroupId
                            action = "selfDeactivate"
                            justification = "Deactivated via PowerShell"
                            accessId = "member"
                        }
                        
                        # Find the active assignment schedule ID
                        if ($roleData.ScheduleId) {
                            $groupRequestBody.assignmentScheduleId = $roleData.ScheduleId
                        }
                        else {
                            # Query for the active schedule
                            $activeSchedules = Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule -Filter "principalId eq '$($script:CurrentUser.Id)' and groupId eq '$($roleData.GroupId)'"
                            if ($activeSchedules) {
                                $groupRequestBody.assignmentScheduleId = $activeSchedules[0].Id
                            }
                        }
                        
                        $response = New-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleRequest -BodyParameter $groupRequestBody
                        Write-Verbose "Group role deactivated successfully"
                    }
                    
                    default {
                        throw "Unsupported role type: $($roleData.Type)"
                    }
                }
                
                $successCount++
            }
            catch {
                $errorMessage = if ($_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
                $deactivationErrors += "$($roleData.DisplayName): $errorMessage"
                Write-Warning "Failed to deactivate $($roleData.DisplayName): $errorMessage"
            }
        }
        
        $operationSplash.UpdateStatus("Completing deactivation process...", 95)
        
        # Display results
        if ($deactivationErrors.Count -gt 0) {
            $message = "Successfully deactivated $successCount of $totalRoles role(s).`n`nErrors:`n$($deactivationErrors -join "`n")"
            Show-TopMostMessageBox -Message $message -Title "Deactivation Results" -Icon Warning
        }
        else {
            Show-TopMostMessageBox -Message "Successfully deactivated all $successCount role(s)!" -Title "Success" -Icon Information
        }
        
        # Refresh role lists
        $operationSplash.UpdateStatus("Refreshing role lists...", 98)
        
        try {
            Update-PIMRolesList -Form $Form -RefreshActive -RefreshEligible
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
    
    Write-Verbose "Deactivation process completed - Success: $successCount, Errors: $($deactivationErrors.Count)"
}