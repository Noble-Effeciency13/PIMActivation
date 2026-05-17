function Clear-PIMPolicyCache {
    <#
    .SYNOPSIS
        Clears the PIM policy cache and authentication context cache.

    .DESCRIPTION
        Clears all cached policy information and authentication contexts. This function is typically 
        used when switching between different Azure AD accounts or when policy information needs to 
        be refreshed from the source. By default only in-memory caches are cleared. Use IncludePersistent
        when a full refresh should also remove the tenant-scoped cache stored under LOCALAPPDATA.
        
        The function resets:
        - PolicyCache: Stores cached PIM policy configurations
        - AuthenticationContextCache: Stores cached authentication context information
        - EntraPoliciesLoaded: Flag indicating whether Entra ID policies have been loaded

    .PARAMETER IncludePersistent
        Removes the tenant-scoped persistent policy-cache folder in addition to clearing the in-memory caches.
        This is used by Full Refresh so policies are fetched from source and the local cache is rebuilt.

    .EXAMPLE
        Clear-PIMPolicyCache
        Clears all PIM-related caches and resets the policy loaded flag.
    
    .NOTES
        This function affects script-scoped variables and should be called when you need to ensure
        fresh policy data is retrieved on the next PIM operation.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludePersistent
    )
    
    # Clear policy cache
    $script:PolicyCache = @{}
    $script:PIMPolicyCacheLoadedForScope = $null
    
    # Clear authentication context cache
    $script:AuthenticationContextCache = @{}

    if ($IncludePersistent) {
        try {
            $policyCachePath = Get-PIMPolicyCacheStorePath
            if ($policyCachePath -and (Test-Path -Path $policyCachePath)) {
                Remove-Item -Path $policyCachePath -Recurse -Force -ErrorAction Stop
                Write-Verbose "Persistent PIM policy cache removed from $policyCachePath"
            }
        }
        catch {
            Write-Verbose "Failed to remove persistent PIM policy cache: $($_.Exception.Message)"
        }
    }
    
    # Clear role caches
    $script:CachedEligibleRoles = $null
    $script:CachedActiveRoles = $null
    $script:LastRoleFetchTime = $null
    
    # Reset Entra policies loaded flag
    $script:EntraPoliciesLoaded = $false
    
    Write-Verbose "PIM caches cleared: PolicyCache, AuthenticationContextCache, RoleCaches, and EntraPoliciesLoaded flag reset"
}