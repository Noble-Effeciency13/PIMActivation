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
    
    # Return 'Directory' for null, empty, or root scope
    if ([string]::IsNullOrEmpty($Scope) -or $Scope -eq '/') {
        return 'Directory'
    }
    
    # Parse administrative unit scopes
    if ($Scope -match '^/administrativeUnits/(.+)$') {
        $auId = $Matches[1]
        try {
            $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $auId -ErrorAction Stop
            return "AU: $($au.DisplayName)"
        }
        catch {
            Write-Verbose "Failed to resolve AU name for ID: $auId"
            return "AU: $auId"
        }
    }
    
    # Return original scope for unrecognized patterns
    return $Scope
}