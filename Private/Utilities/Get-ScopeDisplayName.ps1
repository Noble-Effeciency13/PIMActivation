function Get-ScopeDisplayName {
    <#
    .SYNOPSIS
        Converts role scope IDs to user-friendly display names.
    
    .DESCRIPTION
        Transforms Azure AD directory scope identifiers into readable names.
        Handles directory root scope and administrative unit scopes.
    
    .PARAMETER Scope
        The scope identifier to convert. Can be '/', '/administrativeUnits/{id}', or other scope patterns.
    
    .EXAMPLE
        Get-ScopeDisplayName -Scope '/'
        Returns 'Directory'
    
    .EXAMPLE
        Get-ScopeDisplayName -Scope '/administrativeUnits/12345678-1234-1234-1234-123456789012'
        Returns 'AU: Marketing Department' (or the AU ID if name lookup fails)
    
    .OUTPUTS
        System.String
        Returns a human-readable scope name.
    
    .NOTES
        Requires Microsoft Graph PowerShell SDK for administrative unit name resolution.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowEmptyString()]
        [string]$Scope
    )
    
    # Initialize simple caches to avoid repeated lookups
    if (-not (Test-Path Variable:script:ScopeNameCache) -or -not $script:ScopeNameCache) {
        $script:ScopeNameCache = @{}
    }
    if (-not (Test-Path Variable:script:AuNameCache) -or -not $script:AuNameCache) {
        $script:AuNameCache = @{}
    }
    
    # Return 'Directory' for null, empty, or root scope
    if ([string]::IsNullOrEmpty($Scope) -or $Scope -eq '/') {
        return 'Directory'
    }

    # Return cached scope display name if available
    if ($script:ScopeNameCache.ContainsKey($Scope)) {
        return $script:ScopeNameCache[$Scope]
    }
    
    # Parse administrative unit scopes
    if ($Scope -match '^/administrativeUnits/(.+)$') {
        $auId = $Matches[1]
        try {
            if (-not $script:AuNameCache.ContainsKey($auId)) {
                $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $auId -ErrorAction Stop
                if ($au) {
                    $script:AuNameCache[$auId] = $au.DisplayName
                } else {
                    $script:AuNameCache[$auId] = $null
                }
            }
            $resolved = $script:AuNameCache[$auId]
            $display = if ($resolved) { "AU: $resolved" } else { "AU: $auId" }
            $script:ScopeNameCache[$Scope] = $display
            return $display
        }
        catch {
            Write-Verbose "Failed to resolve AU name for ID: $auId"
            $display = "AU: $auId"
            $script:ScopeNameCache[$Scope] = $display
            return $display
        }
    }
    
    # Return original scope for unrecognized patterns
    $script:ScopeNameCache[$Scope] = $Scope
    return $Scope
}