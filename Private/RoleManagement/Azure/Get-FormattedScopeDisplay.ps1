function Get-FormattedScopeDisplay {
    param($Role)
    
    $scopeDisplay = "Directory"
    
    if ($Role.DirectoryScopeId -and $Role.DirectoryScopeId -ne "/" -and $Role.DirectoryScopeId -ne "Directory") {
        if ($Role.DirectoryScopeId -match "^/administrativeUnits/(.+)$") {
            $auId = $Matches[1]
            try {
                $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $auId -ErrorAction Stop
                $scopeDisplay = "AU: $($au.DisplayName)"
            }
            catch {
                Write-Verbose "Could not retrieve Administrative Unit name for ID: $auId"
                $scopeDisplay = "AU: $auId"
            }
        }
        else {
            $scopeDisplay = $Role.DirectoryScopeId
        }
    }
    
    return $scopeDisplay
}