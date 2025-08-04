function Get-GroupRoles {
    <#
    .SYNOPSIS
        Retrieves PIM-enabled group memberships for a user.
    
    .DESCRIPTION
        Gets both active and eligible PIM group memberships from Microsoft Graph API.
        Returns standardized role objects that include group details, membership status,
        and access type (member or owner).
    
    .PARAMETER UserId
        The Azure AD user ID (GUID) to retrieve group memberships for.
    
    .EXAMPLE
        Get-GroupRoles -UserId "12345678-1234-1234-1234-123456789012"
        Retrieves all PIM group memberships for the specified user.
    
    .EXAMPLE
        Get-GroupRoles -UserId $env:UserPrincipalName -Verbose
        Retrieves group memberships with detailed verbose output.
    
    .OUTPUTS
        PSCustomObject[]
        Returns an array of objects containing group membership details including:
        - Id: Group ID
        - Name: Group display name
        - Type: Always 'Group'
        - Status: 'Eligible' or 'Active'
        - MemberType: 'member' or 'owner'
        - StartDateTime/EndDateTime: Assignment validity period
    
    .NOTES
        Requires Microsoft Graph PowerShell SDK with appropriate permissions:
        - PrivilegedAccess.Read.AzureADGroup
        - Group.Read.All (for group details)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$UserId
    )
    
    begin {
        $roles = [System.Collections.ArrayList]::new()
    }
    
    process {
        try {
            Write-Verbose "Retrieving PIM group memberships for user: $UserId"
            
            # Get eligible group memberships
            Write-Verbose "Querying eligible group memberships..."
            $eligibleGroups = $null
            try {
                $eligibleParams = @{
                    Filter = "principalId eq '$UserId'"
                    ExpandProperty = 'group'
                    ErrorAction = 'Stop'
                }
                $eligibleGroups = Get-MgIdentityGovernancePrivilegedAccessGroupEligibilityScheduleInstance @eligibleParams
            }
            catch {
                Write-Verbose "No eligible groups found or access denied: $($_.Exception.Message)"
                $eligibleGroups = @()
            }
            
            # Normalize to array
            $eligibleGroups = @($eligibleGroups)
            Write-Verbose "Found $($eligibleGroups.Count) eligible group membership(s)"
            
            # Process eligible memberships
            foreach ($membership in $eligibleGroups) {
                if ($membership.Group) {
                    Write-Verbose "Processing eligible group: $($membership.Group.DisplayName) (Access: $($membership.AccessId))"
                    
                    $null = $roles.Add([PSCustomObject]@{
                        Id = $membership.GroupId
                        Name = $membership.Group.DisplayName
                        Type = 'Group'
                        Status = 'Eligible'
                        Source = 'PIMGroup'
                        ResourceId = $membership.GroupId
                        ResourceName = $membership.Group.DisplayName
                        StartDateTime = $membership.StartDateTime
                        EndDateTime = $membership.EndDateTime
                        MemberType = $membership.AccessId
                        DirectoryScopeId = $null
                        PrincipalId = $membership.PrincipalId
                        Assignment = $membership
                    })
                }
            }
            
            # Get active group memberships
            Write-Verbose "Querying active group memberships..."
            $activeGroups = $null
            try {
                $activeParams = @{
                    Filter = "principalId eq '$UserId'"
                    ExpandProperty = 'group'
                    ErrorAction = 'Stop'
                }
                $activeGroups = Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleInstance @activeParams
            }
            catch {
                Write-Verbose "No active groups found or access denied: $($_.Exception.Message)"
                $activeGroups = @()
            }
            
            # Normalize to array
            $activeGroups = @($activeGroups)
            Write-Verbose "Found $($activeGroups.Count) active group membership(s)"
            
            # Process active memberships
            foreach ($membership in $activeGroups) {
                if ($membership.Group) {
                    Write-Verbose "Processing active group: $($membership.Group.DisplayName) (Access: $($membership.AccessId))"
                    
                    $null = $roles.Add([PSCustomObject]@{
                        Id = $membership.GroupId
                        Name = $membership.Group.DisplayName
                        Type = 'Group'
                        Status = 'Active'
                        Source = 'PIMGroup'
                        ResourceId = $membership.GroupId
                        ResourceName = $membership.Group.DisplayName
                        StartDateTime = $membership.StartDateTime
                        EndDateTime = $membership.EndDateTime
                        MemberType = $membership.AccessId
                        DirectoryScopeId = $null
                        PrincipalId = $membership.PrincipalId
                        Assignment = $membership
                    })
                }
            }
            
            Write-Verbose "Retrieved $($roles.Count) total PIM group membership(s)"
        }
        catch {
            Write-Warning "Failed to retrieve group memberships: $($_.Exception.Message)"
            Write-Verbose "Full error details: $($_.Exception.ToString())"
        }
    }
    
    end {
        # Always return an array
        return @($roles)
    }
}