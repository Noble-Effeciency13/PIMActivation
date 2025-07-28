function Test-PIMRoleEligibility {
    <#
    .SYNOPSIS
        Tests if a user has an eligible assignment for a specific role in Azure AD Privileged Identity Management.
    
    .DESCRIPTION
        Verifies that an eligible role assignment exists for the specified user and role before attempting activation.
        This function queries the Microsoft Graph API to check for active eligibility schedule instances.
    
    .PARAMETER UserId
        The Azure AD user's object ID (GUID) to check for role eligibility.
    
    .PARAMETER RoleDefinitionId
        The Azure AD role definition ID (GUID) to verify eligibility for.
    
    .EXAMPLE
        Test-PIMRoleEligibility -UserId "12345678-1234-1234-1234-123456789012" -RoleDefinitionId "62e90394-69f5-4237-9190-012177145e10"
        
        Tests if the specified user has an eligible assignment for the Global Administrator role.
    
    .EXAMPLE
        $eligibility = Test-PIMRoleEligibility -UserId $currentUser.Id -RoleDefinitionId $globalAdminRole.Id -Verbose
        if ($eligibility.IsEligible) {
            Write-Host "User is eligible for role activation"
        }
        
        Checks eligibility with verbose output and processes the result.
    
    .OUTPUTS
        PSCustomObject
        Returns an object with the following properties:
        - IsEligible: Boolean indicating if the user has an eligible assignment
        - EligibilityId: The ID of the eligibility schedule instance (if found)
        - Details: Full details of the eligibility schedule instance
        - Error: Error message if the check failed
    
    .NOTES
        Requires the Microsoft.Graph.Identity.Governance module.
        The calling context must have appropriate permissions to read role eligibility schedules.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$RoleDefinitionId
    )
    
    # Initialize result object
    $result = [PSCustomObject]@{
        IsEligible = $false
        EligibilityId = $null
        Details = $null
        Error = $null
    }
    
    try {
        Write-Verbose "Checking PIM role eligibility for user ID: $UserId"
        Write-Verbose "Target role definition ID: $RoleDefinitionId"
        
        # Query for eligibility schedule instances matching the user and role
        $filter = "principalId eq '$UserId' and roleDefinitionId eq '$RoleDefinitionId'"
        Write-Verbose "Applying filter: $filter"
        
        $eligibilityInstances = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance `
            -Filter $filter `
            -ErrorAction Stop
        
        # Ensure we have an array for consistent processing
        $eligibilityInstances = @($eligibilityInstances)
        Write-Verbose "Found $($eligibilityInstances.Count) eligibility instance(s)"

        if ($eligibilityInstances.Count -gt 0) {
            $instance = $eligibilityInstances[0]
            
            $result.IsEligible = $true
            $result.EligibilityId = $instance.Id
            $result.Details = $instance
            
            Write-Verbose "Eligible assignment found with ID: $($instance.Id)"
            
            # Log additional properties if available
            if ($instance.PSObject.Properties.Name -contains 'AssignmentType') {
                Write-Verbose "Assignment type: $($instance.AssignmentType)"
            }
            
            if ($instance.PSObject.Properties.Name -contains 'MemberType') {
                Write-Verbose "Member type: $($instance.MemberType)"
            }
            
            if ($instance.PSObject.Properties.Name -contains 'StartDateTime') {
                Write-Verbose "Eligibility start: $($instance.StartDateTime)"
            }
            
            if ($instance.PSObject.Properties.Name -contains 'EndDateTime') {
                Write-Verbose "Eligibility end: $($instance.EndDateTime)"
            }
        }
        else {
            $result.Error = "No eligible assignments found for the specified user and role"
            Write-Verbose "No eligible assignments found - user may not have PIM eligibility for this role"
        }
    }
    catch {
        $errorMessage = "Failed to check role eligibility: $($_.Exception.Message)"
        $result.Error = $errorMessage
        Write-Verbose $errorMessage
        Write-Debug "Full exception details: $($_ | Out-String)"
    }
    
    Write-Verbose "Eligibility check completed. IsEligible: $($result.IsEligible)"
    return $result
}