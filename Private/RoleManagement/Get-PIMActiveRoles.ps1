function Get-PIMActiveRoles {
    <#
    .SYNOPSIS
        Retrieves currently active PIM role assignments for the authenticated user.
    
    .DESCRIPTION
        Gets all active Privileged Identity Management (PIM) role assignments across:
        - Entra ID directory roles
        - PIM-enabled groups
        - Azure resource roles (if configured)
        
        The function filters all available roles to return only those currently active,
        with formatted output suitable for display or further processing.
    
    .PARAMETER None
        This function uses module-level configuration variables for scope determination.
    
    .OUTPUTS
        System.Object[]
        Array of custom objects representing active PIM roles with the following properties:
        - Type: Role type (EntraRole, Group, AzureResource)
        - DisplayName: Human-readable role name
        - EndDateTime: When the activation expires
        - ResourceName: Name of the resource the role applies to
        - Scope: Formatted scope display (Directory, AU name, or resource path)
        - MemberType: How the role was assigned (Direct, Eligible, etc.)
        - ScheduleId: Internal ID for deactivation operations
    
    .EXAMPLE
        PS> Get-PIMActiveRoles
        Returns all currently active PIM roles for the authenticated user.
    
    .NOTES
        Requires an authenticated user context and appropriate permissions to query PIM assignments.
        Uses module-scope variables: $script:CurrentUser, $script:IncludeEntraRoles, 
        $script:IncludeGroups, $script:IncludeAzureResources
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Starting active PIM role retrieval"
    
    # Validate user context
    if (-not $script:CurrentUser -or -not $script:CurrentUser.Id) {
        Write-Warning "No authenticated user context available. Please ensure you're logged in."
        return @()
    }
    
    Write-Verbose "Retrieving roles for user: $($script:CurrentUser.Id)"
    
    # Prepare parameters for role query
    $roleParams = @{
        UserId = $script:CurrentUser.Id
        IncludeEntraRoles = $script:IncludeEntraRoles
        IncludeGroups = $script:IncludeGroups
        IncludeAzureResources = $script:IncludeAzureResources
    }
    
    try {
        # Get all roles for the user
        $allRoles = Get-PIMRoles @roleParams
        Write-Verbose "Retrieved $($allRoles.Count) total role assignments"
        
        # Filter to active roles only
        $activeRoles = @($allRoles | Where-Object { $_.Status -eq 'Active' })
        Write-Verbose "Found $($activeRoles.Count) active role assignments"
        
        if ($activeRoles.Count -eq 0) {
            Write-Verbose "No active PIM roles found"
            return @()
        }
        
        # Process and format each active role
        $formattedRoles = foreach ($role in $activeRoles) {
            Write-Verbose "Processing role: $($role.Name) ($($role.Type))"
            
            # Determine assignment type
            $memberType = if ($role.MemberType) { 
                $role.MemberType 
            } else {
                Get-MembershipType -Assignment $role.Assignment -RoleType $role.Type
            }
            
            # Extract schedule identifier for deactivation
            $scheduleId = $null
            if ($role.Assignment) {
                $scheduleId = if ($role.Assignment.Id) { $role.Assignment.Id } else { $role.Assignment.RoleAssignmentScheduleId }
            }
            
            # Format scope display
            $scopeDisplay = Get-FormattedScope -DirectoryScopeId $role.DirectoryScopeId -RoleType $role.Type
            
            # Create formatted role object
            $formattedRole = [PSCustomObject]@{
                Type = $role.Type
                DisplayName = $role.Name
                EndDateTime = $role.EndDateTime
                ResourceName = $role.ResourceName
                DirectoryScopeId = $role.DirectoryScopeId
                Assignment = $role.Assignment
                ResourceId = $role.ResourceId
                GroupId = ($role.Type -eq 'Group') ? $role.ResourceId : $null
                Scope = $scopeDisplay
                MemberType = $memberType
                ScheduleId = $scheduleId
                RoleDefinitionId = $role.Id
            }
            
            # Adjust scope for Azure resources
            if ($role.Type -eq 'AzureResource') {
                $formattedRole.Scope = $role.DirectoryScopeId
            }
            
            $formattedRole
        }
        
        Write-Verbose "Successfully processed $($formattedRoles.Count) active roles"
        return $formattedRoles
    }
    catch {
        $errorMessage = "Failed to retrieve active PIM roles: $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Verbose "Exception details: $($_.Exception.GetType().Name) - $($_.ScriptStackTrace)"
        return @()
    }
}

# Helper function to format scope display
function Get-FormattedScope {
    param(
        [string]$DirectoryScopeId,
        [string]$RoleType
    )
    
    if (-not $DirectoryScopeId -or $DirectoryScopeId -eq "/" -or $DirectoryScopeId -eq "Directory") {
        return "Directory"
    }
    
    # Handle Administrative Unit scopes
    if ($DirectoryScopeId -match "^/administrativeUnits/(.+)$") {
        $auId = $Matches[1]
        try {
            $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $auId -ErrorAction Stop
            return "AU: $($au.DisplayName)"
        }
        catch {
            Write-Verbose "Unable to resolve Administrative Unit name for ID: $auId"
            return "AU: $auId"
        }
    }
    
    return $DirectoryScopeId
}