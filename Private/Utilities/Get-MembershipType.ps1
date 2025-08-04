function Get-MembershipType {
    <#
    .SYNOPSIS
        Determines whether a role assignment is direct or through group membership.
    
    .DESCRIPTION
        Analyzes a role assignment object to determine if the role is assigned directly 
        to the principal or inherited through group membership. Supports Entra roles, 
        Group roles, and Azure Resource roles.
    
    .PARAMETER Assignment
        The role assignment object to analyze. This can be from various sources including
        active assignments, eligible assignments, or role eligibility schedule instances.
    
    .PARAMETER RoleType
        Specifies the type of role assignment to analyze.
        Valid values: 'Entra', 'Group', 'AzureResource'
    
    .EXAMPLE
        Get-MembershipType -Assignment $entraAssignment -RoleType 'Entra'
        Returns 'Direct' or 'Group' based on the assignment's membership type.
    
    .EXAMPLE
        Get-MembershipType -Assignment $groupAssignment -RoleType 'Group'
        Returns 'Direct' since group assignments are always direct.
    
    .OUTPUTS
        System.String
        Returns either 'Direct' or 'Group' indicating the membership type.
    
    .NOTES
        For Entra roles, the function checks multiple properties including MemberType,
        IsTransitive, AssignmentType, and may query eligible assignments to determine
        the true membership source.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Assignment,
        
        [Parameter(Mandatory)]
        [ValidateSet('Entra', 'Group', 'AzureResource')]
        [string]$RoleType
    )
    
    try {
        Write-Verbose "Analyzing membership type for $RoleType role assignment"
        
        # For Entra roles, check multiple properties to determine group membership
        if ($RoleType -eq 'Entra' -and $Assignment) {
            
            # Check MemberType property first (most reliable indicator)
            if ($Assignment.PSObject.Properties['MemberType']) {
                $memberType = $Assignment.MemberType
                Write-Verbose "Found MemberType property: $memberType"
                
                switch ($memberType) {
                    'Direct' { 
                        Write-Verbose "Assignment is direct based on MemberType"
                        return 'Direct' 
                    }
                    { $_ -in @('Group', 'Inherited', 'Transitive') } { 
                        Write-Verbose "Assignment is through group membership based on MemberType: $memberType"
                        return 'Group' 
                    }
                    default { 
                        Write-Verbose "Unknown MemberType '$memberType', defaulting to Direct"
                        return 'Direct' 
                    }
                }
            }
            
            # Check for transitive membership indication
            if ($Assignment.PSObject.Properties['IsTransitive'] -and $Assignment.IsTransitive) {
                Write-Verbose "Assignment marked as transitive, indicating group membership"
                return 'Group'
            }
            
            # Check AssignmentType for additional context
            if ($Assignment.PSObject.Properties['AssignmentType']) {
                Write-Verbose "Checking AssignmentType: $($Assignment.AssignmentType)"
                
                if ($Assignment.AssignmentType -eq 'Assigned') {
                    Write-Verbose "Assignment type is 'Assigned', likely direct"
                    return 'Direct'
                }
                elseif ($Assignment.AssignmentType -eq 'Activated') {
                    Write-Verbose "Assignment is activated, checking origin"
                    
                    if ($Assignment.PSObject.Properties['AssignmentOrigin'] -and 
                        $Assignment.AssignmentOrigin -like '*Group*') {
                        Write-Verbose "Assignment origin indicates group membership"
                        return 'Group'
                    }
                }
            }
            
            # Check for explicit group identifiers
            if ($Assignment.PSObject.Properties['GroupId'] -and $Assignment.GroupId) {
                Write-Verbose "GroupId found, assignment is through group membership"
                return 'Group'
            }
            
            # Check principal type
            if ($Assignment.PSObject.Properties['PrincipalType'] -and 
                $Assignment.PrincipalType -eq 'Group') {
                Write-Verbose "PrincipalType is Group"
                return 'Group'
            }
            
            # For active assignments, check underlying eligible assignments
            if ($Assignment.PSObject.Properties['RoleDefinitionId'] -and 
                $Assignment.PSObject.Properties['PrincipalId']) {
                
                Write-Verbose "Checking eligible assignments for group membership indicators"
                
                try {
                    $eligibleParams = @{
                        Filter = "principalId eq '$($Assignment.PrincipalId)' and roleDefinitionId eq '$($Assignment.RoleDefinitionId)'"
                        ErrorAction = 'SilentlyContinue'
                    }
                    $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance @eligibleParams
                    
                    foreach ($eligible in $eligibleAssignments) {
                        if ($eligible.MemberType -in @('Group', 'Inherited', 'Transitive')) {
                            Write-Verbose "Found group-based eligible assignment with MemberType: $($eligible.MemberType)"
                            return 'Group'
                        }
                    }
                    
                    Write-Verbose "No group-based eligible assignments found"
                }
                catch {
                    Write-Verbose "Unable to query eligible assignments: $($_.Exception.Message)"
                }
            }
        }
        
        # Group roles are always considered direct assignments
        if ($RoleType -eq 'Group') {
            Write-Verbose "Group role type - returning Direct (you are directly assigned to the group)"
            return 'Direct'
        }
        
        # Azure Resource roles - for future implementation
        if ($RoleType -eq 'AzureResource') {
            Write-Verbose "Azure Resource role analysis not yet implemented, defaulting to Direct"
            return 'Direct'
        }
        
        # Default fallback
        Write-Verbose "No clear membership indicators found, defaulting to Direct"
        return 'Direct'
    }
    catch {
        Write-Verbose "Error analyzing membership type: $($_.Exception.Message)"
        return 'Direct'
    }
}