function Get-PIMPendingRequests {
    <#
    .SYNOPSIS
        Retrieves pending PIM role assignment requests for the current user.
    
    .DESCRIPTION
        Queries Microsoft Graph to get all pending role assignment schedule requests
        for the current user, including both Entra ID roles and PIM group memberships.
    
    .PARAMETER UserId
        The user ID to check for pending requests. If not provided, uses the current user.
    
    .OUTPUTS
        Array of pending request objects with role information and request details.
    #>
    [CmdletBinding()]
    param(
        [string]$UserId
    )
    
    if (-not $UserId) {
        if (-not $script:CurrentUser -or -not $script:CurrentUser.Id) {
            Write-Warning "No user ID provided and no current user context available"
            return @()
        }
        $UserId = $script:CurrentUser.Id
    }
    
    $pendingRequests = @()
    
    try {
        Write-Verbose "Retrieving pending Entra ID role requests for user: $UserId"
        
        # Verify Graph connection is available
        $graphContext = Get-MgContext
        if (-not $graphContext) {
            Write-Warning "No Microsoft Graph connection available for pending request lookup"
            return @()
        }
        
        Write-Verbose "Graph connection confirmed for account: $($graphContext.Account)"
        
        # Get pending Entra ID role assignment requests
        $entraRequests = @(Get-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -Filter "principalId eq '$UserId' and status eq 'PendingApproval'" -ErrorAction SilentlyContinue)
        
        Write-Verbose "Found $($entraRequests.Count) pending Entra ID role requests"
        
        foreach ($request in $entraRequests) {
            if ($request.RoleDefinitionId) {
                try {
                    # Get role definition separately
                    $roleDefinition = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $request.RoleDefinitionId -ErrorAction SilentlyContinue
                    
                    $pendingRequest = [PSCustomObject]@{
                        Type = 'Entra'
                        RoleDefinitionId = $request.RoleDefinitionId
                        RoleName = if ($roleDefinition) { $roleDefinition.DisplayName } else { "Unknown Role" }
                        RequestId = $request.Id
                        Action = $request.Action
                        Status = $request.Status
                        CreatedDateTime = $request.CreatedDateTime
                        DirectoryScopeId = $request.DirectoryScopeId
                        Justification = $request.Justification
                        ScheduleInfo = $request.ScheduleInfo
                    }
                    
                    $pendingRequests += $pendingRequest
                    Write-Verbose "Added pending Entra request: $($pendingRequest.RoleName) (ID: $($pendingRequest.RoleDefinitionId))"
                }
                catch {
                    Write-Warning "Failed to get role definition for ID $($request.RoleDefinitionId): $_"
                    # Add the request even without role definition
                    $pendingRequest = [PSCustomObject]@{
                        Type = 'Entra'
                        RoleDefinitionId = $request.RoleDefinitionId
                        RoleName = "Unknown Role"
                        RequestId = $request.Id
                        Action = $request.Action
                        Status = $request.Status
                        CreatedDateTime = $request.CreatedDateTime
                        DirectoryScopeId = $request.DirectoryScopeId
                        Justification = $request.Justification
                        ScheduleInfo = $request.ScheduleInfo
                    }
                    $pendingRequests += $pendingRequest
                    Write-Verbose "Added pending Entra request with unknown role name: $($pendingRequest.RoleDefinitionId)"
                }
            }
        }
        
        Write-Verbose "Found $($entraRequests.Count) pending Entra ID role requests"
        
        # Get pending PIM group assignment requests
        Write-Verbose "Retrieving pending PIM group requests for user: $UserId"
        
        try {
            $groupRequests = @(Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleRequest -Filter "principalId eq '$UserId' and status eq 'PendingProvisioning'" -ExpandProperty Group -ErrorAction SilentlyContinue)
            
            foreach ($request in $groupRequests) {
                if ($request.Group) {
                    $groupRequest = [PSCustomObject]@{
                        Type = 'Group'
                        GroupId = $request.GroupId
                        RoleName = $request.Group.DisplayName
                        RequestId = $request.Id
                        Action = $request.Action
                        Status = $request.Status
                        CreatedDateTime = $request.CreatedDateTime
                        AccessId = $request.AccessId
                        Justification = $request.Justification
                        ScheduleInfo = $request.ScheduleInfo
                    }
                    $pendingRequests += $groupRequest
                    Write-Verbose "Added pending Group request: $($groupRequest.RoleName) (ID: $($groupRequest.GroupId))"
                }
            }
            
            Write-Verbose "Found $($groupRequests.Count) pending PIM group requests"
        }
        catch {
            Write-Warning "Failed to retrieve PIM group requests: $($_.Exception.Message)"
            # Continue processing even if group requests fail
        }
        Write-Verbose "Total pending requests found: $($pendingRequests.Count)"
        
        return $pendingRequests
    }
    catch {
        Write-Warning "Failed to retrieve pending PIM requests: $($_.Exception.Message)"
        Write-Verbose "Exception details: $($_.Exception.ToString())"
        Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
        return @()
    }
}
