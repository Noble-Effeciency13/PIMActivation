function Get-PIMRoles {
    <#
    .SYNOPSIS
        Retrieves Privileged Identity Management (PIM) roles for a specified user.
    
    .DESCRIPTION
        Modular function to retrieve different types of PIM roles including Entra ID directory roles,
        PIM-enabled groups, and Azure resource roles. Each role type is handled by a separate 
        private function for better modularity and error handling.
    
    .PARAMETER IncludeEntraRoles
        Switch to include Entra ID directory roles in the results.
    
    .PARAMETER IncludeGroups
        Switch to include PIM-enabled group memberships in the results.
    
    .PARAMETER IncludeAzureResources
        Switch to include Azure resource roles (subscriptions, resource groups, etc.) in the results.
    
    .PARAMETER UserId
        The user ID (object ID) to retrieve PIM roles for. If not specified, retrieves roles for the current user.
    
    .EXAMPLE
        Get-PIMRoles -IncludeEntraRoles -UserId "12345678-1234-1234-1234-123456789012"
        Retrieves only Entra ID directory roles for the specified user.
    
    .EXAMPLE
        Get-PIMRoles -IncludeEntraRoles -IncludeGroups -IncludeAzureResources
        Retrieves all types of PIM roles for the current user.
    
    .OUTPUTS
        System.Array
        Returns an array of PIM role objects containing role details from the specified sources.
    
    .NOTES
        Requires appropriate permissions to read PIM role assignments.
        - For Entra roles: Privileged Role Administrator or similar
        - For Groups: Groups Administrator or similar  
        - For Azure resources: Reader access to Azure subscriptions
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeEntraRoles,
        [switch]$IncludeGroups,
        [switch]$IncludeAzureResources,
        [string]$UserId
    )
    
    Write-Verbose "Starting PIM role retrieval for user: $(if ($UserId) { $UserId } else { 'current user' })"
    Write-Verbose "Role types requested - Entra: $IncludeEntraRoles, Groups: $IncludeGroups, Azure: $IncludeAzureResources"
    
    $allRoles = @()
    
    # Get Entra ID directory roles
    if ($IncludeEntraRoles) {
        Write-Verbose "Retrieving Entra ID directory roles..."
        try {
            $entraRoles = Get-EntraIDRoles -UserId $UserId
            
            # Ensure we always have an array
            if ($null -eq $entraRoles) {
                $entraRoles = @()
            }
            elseif ($entraRoles -isnot [array]) {
                $entraRoles = @($entraRoles)
            }
            
            $roleCount = $entraRoles.Count
            Write-Verbose "Successfully retrieved $roleCount Entra ID role(s)"
            
            if ($roleCount -gt 0) {
                $allRoles += $entraRoles
            }
        }
        catch {
            Write-Warning "Failed to retrieve Entra ID roles: $($_.Exception.Message)"
            Write-Verbose "Entra ID roles error details - Type: $($_.Exception.GetType().Name), Stack: $($_.ScriptStackTrace)"
        }
    }
    
    # Get PIM-enabled group memberships
    if ($IncludeGroups) {
        Write-Verbose "Retrieving PIM-enabled group roles..."
        try {
            $groupRoles = Get-GroupRoles -UserId $UserId
            
            # Ensure we always have an array
            if ($null -eq $groupRoles) {
                $groupRoles = @()
            }
            elseif ($groupRoles -isnot [array]) {
                $groupRoles = @($groupRoles)
            }
            
            $roleCount = $groupRoles.Count
            Write-Verbose "Successfully retrieved $roleCount group role(s)"
            
            if ($roleCount -gt 0) {
                $allRoles += $groupRoles
            }
        }
        catch {
            Write-Warning "Failed to retrieve group roles: $($_.Exception.Message)"
            Write-Verbose "Group roles error details - Type: $($_.Exception.GetType().Name), Stack: $($_.ScriptStackTrace)"
        }
    }
    
    # Get Azure resource roles
    if ($IncludeAzureResources) {
        Write-Verbose "Retrieving Azure resource roles..."
        try {
            $azureRoles = Get-AzureResourceRoles -UserId $UserId
            $roleCount = $azureRoles.Count
            Write-Verbose "Successfully retrieved $roleCount Azure resource role(s)"
            
            if ($azureRoles) {
                $allRoles += $azureRoles
            }
        }
        catch {
            Write-Warning "Failed to retrieve Azure resource roles: $($_.Exception.Message)"
            Write-Verbose "Azure roles error details - Type: $($_.Exception.GetType().Name), Stack: $($_.ScriptStackTrace)"
        }
    }
    
    $totalCount = $allRoles.Count
    Write-Verbose "PIM role retrieval completed. Total roles found: $totalCount"
    
    return $allRoles
}