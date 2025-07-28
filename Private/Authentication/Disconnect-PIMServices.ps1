function Disconnect-PIMServices {
    <#
    .SYNOPSIS
        Disconnects from Microsoft Graph and Azure services used by PIM operations.
    
    .DESCRIPTION
        Cleanly disconnects from all connected Microsoft Graph and Azure services,
        clears policy caches, and performs cleanup operations. This function should
        be called when finishing PIM-related tasks to ensure proper session cleanup.
    
    .EXAMPLE
        Disconnect-PIMServices
        
        Disconnects from all PIM-related services and clears caches.
    
    .EXAMPLE
        Disconnect-PIMServices -Verbose
        
        Disconnects from services with detailed verbose output showing each step.
    
    .NOTES
        - Clears the PIM policy cache before disconnecting
        - Safely handles disconnection even if services are not connected
        - Uses SilentlyContinue to prevent errors for already disconnected services
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Starting PIM services disconnection process"
    
    try {
        # Clear policy cache when disconnecting
        Write-Verbose "Clearing PIM policy cache"
        Clear-PIMPolicyCache
        
        # Clear authentication context tokens
        Write-Verbose "Clearing authentication context session state"
        $script:CurrentAuthContextToken = $null
        $script:CurrentAuthContextRefreshToken = $null
        $script:AuthContextTokens = @{}
        $script:JustCompletedAuthContext = $false
        $script:AuthContextCompletionTime = $null
        
        # Disconnect from Microsoft Graph
        Write-Verbose "Attempting to disconnect from Microsoft Graph"
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Verbose "Successfully disconnected from Microsoft Graph"
        
        # Disconnect from Azure if connected
        if ($script:IncludeAzureResources -and (Get-Module -Name Az.Accounts -ErrorAction SilentlyContinue)) {
            Write-Verbose "Azure resources are included and Az.Accounts module is available"
            Write-Verbose "Attempting to disconnect from Azure"
            $null = Disconnect-AzAccount -ErrorAction SilentlyContinue
            Write-Verbose "Successfully disconnected from Azure"
        }
        else {
            Write-Verbose "Azure disconnection skipped (not connected or module not available)"
        }
        
        Write-Verbose "PIM services disconnection completed successfully"
    }
    catch {
        Write-Warning "Error occurred during PIM services disconnection: $($_.Exception.Message)"
        Write-Verbose "Full error details: $_"
    }
}