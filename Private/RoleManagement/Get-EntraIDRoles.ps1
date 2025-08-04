function Get-EntraIDRoles {
    <#
    .SYNOPSIS
        Retrieves Entra ID directory roles for a specified user.
    
    .DESCRIPTION
        Gets both active and eligible Entra ID directory roles from Microsoft Graph API.
        Returns a collection of role objects with standardized properties for PIM management.
    
    .PARAMETER UserId
        The user ID (Object ID) to retrieve roles for. This should be the user's Azure AD Object ID.
    
    .EXAMPLE
        Get-EntraIDRoles -UserId "12345678-1234-1234-1234-123456789012"
        
        Retrieves all active and eligible Entra ID roles for the specified user.
    
    .EXAMPLE
        Get-EntraIDRoles -UserId $env:USER_OBJECT_ID -Verbose
        
        Retrieves roles with verbose output showing the operation progress.
    
    .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        Returns an array of role objects with properties: Id, Name, Type, Status, Source, 
        ResourceId, ResourceName, StartDateTime, EndDateTime, MemberType, DirectoryScopeId, 
        PrincipalId, and Assignment.
    
    .NOTES
        Requires Microsoft Graph PowerShell SDK with appropriate permissions:
        - RoleManagement.Read.Directory
        - Directory.Read.All
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )
    
    $roles = @()
    
    try {
        Write-Verbose "Retrieving Entra ID roles for user: $UserId"
        
        # Get eligible role assignments
        Write-Verbose "Querying eligible role assignments..."
        $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance `
            -Filter "principalId eq '$UserId'" `
            -ExpandProperty roleDefinition -ErrorAction Stop
        
        $eligibleAssignments = @($eligibleAssignments)

        Write-Verbose "Processing $($eligibleAssignments.Count) eligible role assignment(s)"
        foreach ($assignment in $eligibleAssignments) {
            # Determine member type
            $memberType = if ($assignment.MemberType) {
                switch ($assignment.MemberType) {
                    "Direct" { "Direct" }
                    "Group" { "Group" }
                    "Inherited" { "Inherited" }
                    default { $assignment.MemberType }
                }
            } else {
                "Direct"
            }
            
            $roles += [PSCustomObject]@{
                Id = $assignment.RoleDefinitionId
                Name = $assignment.RoleDefinition.DisplayName
                Type = 'Entra'
                Status = 'Eligible'
                Source = 'EntraID'
                ResourceId = $null
                ResourceName = 'Entra ID Directory'
                StartDateTime = $assignment.StartDateTime
                EndDateTime = $assignment.EndDateTime
                MemberType = $memberType  # ADD THIS
                DirectoryScopeId = $assignment.DirectoryScopeId
                PrincipalId = $assignment.PrincipalId
                Assignment = $assignment
            }
        }
        
        # Get active role assignments
        Write-Verbose "Querying active role assignments..."
        $activeAssignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance `
            -Filter "principalId eq '$UserId'" `
            -ErrorAction Stop
        
        $activeAssignments = @($activeAssignments)
        
        Write-Verbose "Processing $($activeAssignments.Count) active role assignment(s)"
        foreach ($assignment in $activeAssignments) {
            try {
                # Retrieve role definition details
                $roleDetails = Get-MgDirectoryRole -Filter "roleTemplateId eq '$($assignment.RoleDefinitionId)'" -ErrorAction Stop
                
                if ($roleDetails) {
                    $roles += [PSCustomObject]@{
                        Id = $assignment.RoleDefinitionId
                        Name = $roleDetails.DisplayName
                        Type = 'Entra'
                        Status = 'Active'
                        Source = 'EntraID'
                        ResourceId = $null
                        ResourceName = 'Entra ID Directory'
                        StartDateTime = $assignment.StartDateTime
                        EndDateTime = $assignment.EndDateTime
                        MemberType = $assignment.MemberType
                        DirectoryScopeId = $assignment.DirectoryScopeId
                        PrincipalId = $assignment.PrincipalId
                        Assignment = $assignment
                    }
                }
            }
            catch {
                Write-Verbose "Unable to retrieve role definition for ID '$($assignment.RoleDefinitionId)': $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "Successfully retrieved $($roles.Count) total Entra ID role(s)"
    }
    catch {
        Write-Warning "Error retrieving Entra ID roles: $($_.Exception.Message)"
        Write-Verbose "Entra ID role retrieval error: $($_.Exception.GetType().Name) - $($_.ScriptStackTrace)"
    }
    
    # Ensure we always return an array
    if ($null -eq $roles) {
        return @()
    }
    elseif ($roles -isnot [array]) {
        return @($roles)
    }
    
    return $roles
}