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
        elseif ($Role.DirectoryScopeId -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
            $mgId = $Matches[1]
            try {
                $mgInfo = Get-AzManagementGroup -GroupId $mgId -ErrorAction SilentlyContinue
                if ($mgInfo -and $mgInfo.DisplayName) {
                    if ($mgId -eq 'root' -or $mgInfo.Name -eq 'root' -or $mgInfo.DisplayName -match "(?i)tenant root") {
                        $scopeDisplay = "/"
                    }
                    else {
                        $scopeDisplay = "MG: $($mgInfo.DisplayName)"
                    }
                }
                else {
                    if ($mgId -eq 'root') { $scopeDisplay = "/" } else { $scopeDisplay = "MG: $mgId" }
                }
            }
            catch {
                Write-Verbose "Could not retrieve Management Group info for ID: $mgId"
                if ($mgId -eq 'root') { $scopeDisplay = "/" } else { $scopeDisplay = "MG: $mgId" }
            }
        }
        else {
            $scopeDisplay = $Role.DirectoryScopeId
        }
    }
    
    return $scopeDisplay
}