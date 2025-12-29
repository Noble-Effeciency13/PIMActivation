function Get-AzureScopeInfo {
    <#
    .SYNOPSIS
        Converts Azure scope to display information matching Entra portal format.
    
    .DESCRIPTION
        Parses Azure ARM scope strings and returns formatted display information
        that aligns with how roles are displayed in the Entra portal.
    
    .PARAMETER Scope
        The Azure ARM scope string (e.g., /subscriptions/xxx/resourceGroups/yyy).
    
    .EXAMPLE
        Get-AzureScopeInfo -Scope "/subscriptions/12345/resourceGroups/myRG"
        
        Returns scope information formatted for display.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Scope
    )
    
    if ($Scope -eq "/" -or $Scope -eq "") {
        return @{
            ResourceDisplay = "/"
            ScopeType       = "Tenant"
        }
    }
    
    # Parse management group scope
    if ($Scope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
        $mgName = $matches[1]
        
        # Check if this is the tenant root group (often matches tenant ID format)
        if ($mgName -match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$") {
            $displayName = "Tenant Root Group"
        }
        else {
            $displayName = $mgName
        }
        
        return @{
            ResourceDisplay = $displayName
            ScopeType       = "Management group"
        }
    }
    
    # Parse subscription-level scope
    if ($Scope -match "^/subscriptions/([^/]+)$") {
        $subscriptionId = $matches[1]
        try {
            $subscription = Get-AzSubscription -SubscriptionId $subscriptionId -ErrorAction SilentlyContinue
            $subscriptionName = if ($subscription) { $subscription.Name } else { $subscriptionId }
        }
        catch {
            $subscriptionName = $subscriptionId
        }
        
        return @{
            ResourceDisplay = $subscriptionName
            ScopeType       = "Subscription"
        }
    }
    
    # Parse resource group scope
    if ($Scope -match "^/subscriptions/([^/]+)/resourceGroups/([^/]+)$") {
        return @{
            ResourceDisplay = $matches[2]
            ScopeType       = "Resource Group"
        }
    }
    
    # Parse specific resource scope - handle complex resource paths
    if ($Scope -match "^/subscriptions/([^/]+)/resourceGroups/([^/]+)/providers/([^/]+)/([^/]+)/(.+)$") {
        $subscriptionId = $matches[1]
        $resourceGroupName = $matches[2]
        $resourceProviderNamespace = $matches[3]
        $resourceType = $matches[4]
        $resourcePath = $matches[5]
        
        # Handle nested resources (like storage account file shares)
        $pathSegments = $resourcePath.Split('/')
        
        # For storage accounts with nested resources like file shares
        if ($resourceProviderNamespace -eq 'Microsoft.Storage' -and $pathSegments.Count -gt 1) {
            $storageAccountName = $pathSegments[0]
            $nestedResourceType = $pathSegments[1] # e.g., "fileServices", "default", etc.
            $nestedResourceName = if ($pathSegments.Count -gt 2) { $pathSegments[-1] } else { $pathSegments[1] }
            
            # Special handling for file shares
            if ($nestedResourceType -eq 'fileServices' -and $pathSegments.Count -gt 3) {
                $fileShareName = $pathSegments[-1]
                return @{
                    ResourceDisplay = "$fileShareName ($storageAccountName)"
                    ScopeType       = "Fileshare"
                }
            }
            # General nested storage resource
            return @{
                ResourceDisplay = "$nestedResourceName ($storageAccountName)"
                ScopeType       = "Storage resource"
            }
        }
        
        # Extract the final resource name from potentially nested path
        $resourceName = $pathSegments[-1]
        
        # Format the resource type for display
        $displayResourceType = switch ($resourceProviderNamespace) {
            'Microsoft.Storage' { 
                switch ($resourceType) {
                    'storageAccounts' { 'Storage account' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.Compute' {
                switch ($resourceType) {
                    'virtualMachines' { 'Virtual machine' }
                    'disks' { 'Disk' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.Web' {
                switch ($resourceType) {
                    'sites' { 'App Service' }
                    'serverfarms' { 'App Service plan' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.KeyVault' {
                switch ($resourceType) {
                    'vaults' { 'Key vault' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.Network' {
                switch ($resourceType) {
                    'virtualNetworks' { 'Virtual network' }
                    'networkSecurityGroups' { 'Network security group' }
                    'publicIPAddresses' { 'Public IP address' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.Sql' {
                switch ($resourceType) {
                    'servers' { 'SQL server' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.Authorization' {
                switch ($resourceType) {
                    'roleDefinitions' { 'Role definition' }
                    'policyDefinitions' { 'Policy definition' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.Resources' {
                switch ($resourceType) {
                    'resourceGroups' { 'Resource group' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            'Microsoft.DesktopVirtualization' {
                switch ($resourceType) {
                    'applicationgroups' { 'Application group' }
                    'hostpools' { 'Host pool' }
                    'workspaces' { 'Workspace' }
                    default { "$resourceProviderNamespace/$resourceType" }
                }
            }
            default { 
                # Format as readable text for common patterns
                if ($resourceType -eq 'clusters') { 'Cluster' }
                elseif ($resourceType -eq 'databases') { 'Database' }
                elseif ($resourceType -eq 'accounts') { 'Account' }
                else { "$resourceProviderNamespace/$resourceType" }
            }
        }
        
        return @{
            ResourceDisplay = $resourceName
            ScopeType       = $displayResourceType
        }
    }
    
    # Handle nested resource scopes (like SQL databases)
    if ($Scope -match "^/subscriptions/([^/]+)/resourceGroups/([^/]+)/providers/([^/]+)/.+$") {
        $resourceProviderNamespace = $matches[3]
        
        # Extract the last segment as resource name
        $segments = $Scope.Split('/')
        $resourceName = $segments[-1]
        
        return @{
            ResourceDisplay = $resourceName
            ScopeType       = $resourceProviderNamespace
        }
    }
    
    # Fallback for unknown scope format
    return @{
        ResourceDisplay = $Scope
        ScopeType       = "Unknown"
    }
}