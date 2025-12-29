function Get-FormattedScope {
    param(
        [string]$DirectoryScopeId,
        [string]$RoleType
    )
    
    if (-not $DirectoryScopeId -or $DirectoryScopeId -eq "/" -or $DirectoryScopeId -eq "Directory") {
        return "Directory"
    }
    
    # Handle Administrative Unit scopes
    if ($DirectoryScopeId -match "^/administrativeUnits/(.+)$") {
        $auId = $Matches[1]
        try {
            $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $auId -ErrorAction Stop
            return "AU: $($au.DisplayName)"
        }
        catch {
            Write-Verbose "Unable to resolve Administrative Unit name for ID: $auId"
            return "AU: $auId"
        }
    }
    
    return $DirectoryScopeId
}