function Clear-PIMPolicyCache {
    <#
    .SYNOPSIS
        Clears the PIM policy cache and authentication context cache.
    
    .DESCRIPTION
        Clears all cached policy information and authentication contexts. This function is typically 
        used when switching between different Azure AD accounts or when policy information needs to 
        be refreshed from the source.
        
        The function resets:
        - PolicyCache: Stores cached PIM policy configurations
        - AuthenticationContextCache: Stores cached authentication context information
        - EntraPoliciesLoaded: Flag indicating whether Entra ID policies have been loaded
    
    .EXAMPLE
        Clear-PIMPolicyCache
        Clears all PIM-related caches and resets the policy loaded flag.
    
    .NOTES
        This function affects script-scoped variables and should be called when you need to ensure
        fresh policy data is retrieved on the next PIM operation.
    #>
    [CmdletBinding()]
    param()
    
    # Clear policy cache
    $script:PolicyCache = @{}
    
    # Clear authentication context cache
    $script:AuthenticationContextCache = @{}
    
    # Reset Entra policies loaded flag
    $script:EntraPoliciesLoaded = $false
    
    Write-Verbose "PIM caches cleared: PolicyCache, AuthenticationContextCache, and EntraPoliciesLoaded flag reset"
}