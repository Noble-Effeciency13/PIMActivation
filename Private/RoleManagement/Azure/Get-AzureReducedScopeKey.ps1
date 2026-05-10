function Get-AzureReducedScopeKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$RoleData
    )

    if ($RoleData.PSObject.Properties['EligibilityScheduleName'] -and $RoleData.EligibilityScheduleName) {
        return "EligibilityName|$($RoleData.EligibilityScheduleName)"
    }

    if ($RoleData.PSObject.Properties['EligibilityScheduleId'] -and $RoleData.EligibilityScheduleId) {
        return "EligibilityId|$($RoleData.EligibilityScheduleId)"
    }

    $scope = if ($RoleData.PSObject.Properties['FullScope'] -and $RoleData.FullScope) { $RoleData.FullScope } else { '' }
    $roleDefinitionId = if ($RoleData.PSObject.Properties['RoleDefinitionId'] -and $RoleData.RoleDefinitionId) { $RoleData.RoleDefinitionId } else { '' }
    $displayName = if ($RoleData.PSObject.Properties['DisplayName'] -and $RoleData.DisplayName) { $RoleData.DisplayName } else { '' }

    return "Role|$scope|$roleDefinitionId|$displayName"
}
