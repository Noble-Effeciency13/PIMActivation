function Get-PIMEligibleRoles {
    <#
    .SYNOPSIS
        Retrieves eligible PIM roles for the current user.
    
    .DESCRIPTION
        Gets all eligible Privileged Identity Management (PIM) roles across Entra ID and Groups
        for the current user context. Azure resource roles are not included in this version.
        
        The function filters roles to only return those with 'Eligible' status and formats
        them with policy information and scope details for easy consumption.
    
    .EXAMPLE
        Get-PIMEligibleRoles
        
        Returns all eligible PIM roles for the current user across Entra ID and Groups.
    
    .OUTPUTS
        System.Object[]
        Array of role objects containing:
        - Type: Role type (Entra, Group)
        - DisplayName: Human-readable role name
        - PolicyInfo: Role policy configuration
        - ResourceName: Name of the resource the role applies to
        - RoleDefinitionId: Unique identifier for the role definition
        - Assignment: Assignment details
        - Scope: Formatted scope display (Directory, AU name, etc.)
        - Type-specific properties (DirectoryScopeId, GroupId, SubscriptionId)
    
    .NOTES
        Requires an active user context ($script:CurrentUser) to be available.
        Uses module-level flags $script:IncludeEntraRoles and $script:IncludeGroups
        to determine which role types to retrieve.
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Starting PIM eligible roles retrieval"
    
    # Validate current user context
    if (-not $script:CurrentUser -or -not $script:CurrentUser.Id) {
        Write-Verbose "No current user context available - returning empty array"
        return @()
    }
    
    Write-Verbose "Processing roles for user: $($script:CurrentUser.UserPrincipalName) ($($script:CurrentUser.Id))"
    
    # Prepare parameters for role retrieval
    $params = @{
        UserId = $script:CurrentUser.Id
        IncludeEntraRoles = $script:IncludeEntraRoles
        IncludeGroups = $script:IncludeGroups
        IncludeAzureResources = $false  # Explicitly disabled for this version
    }
    
    try {
        # Get all roles for the user
        $allRoles = Get-PIMRoles @params
        
        # Ensure we have an array to work with
        if ($null -eq $allRoles) {
            $allRoles = @()
        }
        elseif ($allRoles -isnot [array]) {
            $allRoles = @($allRoles)
        }
        
        Write-Verbose "Retrieved $($allRoles.Count) total roles"
        
        # Filter to eligible roles only
        $eligibleRoles = @($allRoles | Where-Object { $_.Status -eq 'Eligible' })
        
        # Ensure we have an array after filtering
        if ($null -eq $eligibleRoles) {
            $eligibleRoles = @()
        }
        
        Write-Verbose "Found $($eligibleRoles.Count) eligible roles"
        
        # Process and format each eligible role
        $formattedRoles = foreach ($role in $eligibleRoles) {
            Write-Verbose "Processing role: $($role.Name) (Type: $($role.Type))"
            
            # Format scope display for better readability
            $scopeDisplay = Get-FormattedScopeDisplay -Role $role
            
            # Get policy information for the role
            $policyInfo = Get-PIMRolePolicy -Role $role
            
            # Create base formatted role object
            $formattedRole = [PSCustomObject]@{
                Type = $role.Type
                DisplayName = $role.Name
                PolicyInfo = $policyInfo
                ResourceName = $role.ResourceName
                Assignment = $role.Assignment
                Scope = $scopeDisplay
                MemberType = $role.MemberType  # This line ensures MemberType is passed through
            }
            
            # Add the appropriate ID property based on role type
            if ($role.Type -eq 'Group') {
                $formattedRole | Add-Member -NotePropertyName 'GroupId' -NotePropertyValue $role.ResourceId
                $formattedRole | Add-Member -NotePropertyName 'RoleDefinitionId' -NotePropertyValue $null
            } else {
                # Entra roles and others use RoleDefinitionId
                $formattedRole | Add-Member -NotePropertyName 'RoleDefinitionId' -NotePropertyValue $role.Id
                $formattedRole | Add-Member -NotePropertyName 'GroupId' -NotePropertyValue $null
            }
            
            # Add type-specific properties
            Add-TypeSpecificProperties -FormattedRole $formattedRole -SourceRole $role
            
            $formattedRole
        }
        
        # Ensure we return an array
        if ($null -eq $formattedRoles) {
            $formattedRoles = @()
        }
        elseif ($formattedRoles -isnot [array]) {
            $formattedRoles = @($formattedRoles)
        }
        
        Write-Verbose "Successfully processed $($formattedRoles.Count) eligible roles"
        return $formattedRoles
    }
    catch {
        Write-Error "Failed to retrieve eligible roles: $($_.Exception.Message)"
        Write-Verbose "Exception details: $($_.Exception.GetType().Name) - $($_.ScriptStackTrace)"
        return @()
    }
}

# Helper function to format scope display
function Get-FormattedScopeDisplay {
    param($Role)
    
    $scopeDisplay = "Directory"
    
    if ($Role.DirectoryScopeId -and $Role.DirectoryScopeId -ne "/" -and $Role.DirectoryScopeId -ne "Directory") {
        if ($Role.DirectoryScopeId -match "^/administrativeUnits/(.+)$") {
            $auId = $Matches[1]
            try {
                $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $auId -ErrorAction Stop
                $scopeDisplay = "AU: $($au.DisplayName)"
            }
            catch {
                Write-Verbose "Could not retrieve Administrative Unit name for ID: $auId"
                $scopeDisplay = "AU: $auId"
            }
        }
        else {
            $scopeDisplay = $Role.DirectoryScopeId
        }
    }
    
    return $scopeDisplay
}

# Helper function to add type-specific properties
function Add-TypeSpecificProperties {
    param($FormattedRole, $SourceRole)
    
    switch ($SourceRole.Type) {
        'Entra' {
            $FormattedRole | Add-Member -NotePropertyName 'DirectoryScopeId' -NotePropertyValue $SourceRole.DirectoryScopeId
            $FormattedRole | Add-Member -NotePropertyName 'SubscriptionId' -NotePropertyValue $null
        }
        'Group' {
            $FormattedRole | Add-Member -NotePropertyName 'DirectoryScopeId' -NotePropertyValue $null
            $FormattedRole | Add-Member -NotePropertyName 'SubscriptionId' -NotePropertyValue $null
        }
        default {
            # Future Azure resource roles
            $FormattedRole | Add-Member -NotePropertyName 'DirectoryScopeId' -NotePropertyValue $null
            $FormattedRole | Add-Member -NotePropertyName 'SubscriptionId' -NotePropertyValue $null
        }
    }
}