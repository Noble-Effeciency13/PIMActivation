function Get-PIMActivationProfileRoleKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$RoleData
    )

    if ($RoleData -is [string]) {
        return "Name|$RoleData"
    }

    $type = if ($RoleData.PSObject.Properties['Type'] -and $RoleData.Type) { [string]$RoleData.Type } else { 'Unknown' }
    switch ($type) {
        'Azure' { $type = 'AzureResource' }
    }

    switch ($type) {
        'AzureResource' {
            $roleDefinitionId = if ($RoleData.PSObject.Properties['RoleDefinitionId'] -and $RoleData.RoleDefinitionId) { $RoleData.RoleDefinitionId } elseif ($RoleData.PSObject.Properties['Id']) { $RoleData.Id } else { '' }
            $scope = if ($RoleData.PSObject.Properties['FullScope'] -and $RoleData.FullScope) { $RoleData.FullScope } elseif ($RoleData.PSObject.Properties['DirectoryScopeId']) { $RoleData.DirectoryScopeId } else { '' }
            return "AzureResource|$roleDefinitionId|$scope"
        }
        'Group' {
            $groupId = if ($RoleData.PSObject.Properties['GroupId'] -and $RoleData.GroupId) { $RoleData.GroupId } elseif ($RoleData.PSObject.Properties['ResourceId']) { $RoleData.ResourceId } else { '' }
            $accessId = if ($RoleData.PSObject.Properties['AccessId'] -and $RoleData.AccessId) {
                $RoleData.AccessId
            }
            elseif ($RoleData.PSObject.Properties['Assignment'] -and $RoleData.Assignment -and $RoleData.Assignment.PSObject.Properties['AccessId'] -and $RoleData.Assignment.AccessId) {
                $RoleData.Assignment.AccessId
            }
            elseif ($RoleData.PSObject.Properties['MemberType'] -and $RoleData.MemberType) {
                $RoleData.MemberType
            }
            elseif ($RoleData.PSObject.Properties['MembershipType'] -and $RoleData.MembershipType) {
                $RoleData.MembershipType
            }
            else {
                'member'
            }
            $accessId = ([string]$accessId).Trim().ToLowerInvariant()
            if ($accessId -notin @('member', 'owner')) { $accessId = 'member' }
            return "Group|$groupId|$accessId"
        }
        'Entra' {
            $roleDefinitionId = if ($RoleData.PSObject.Properties['RoleDefinitionId'] -and $RoleData.RoleDefinitionId) { $RoleData.RoleDefinitionId } elseif ($RoleData.PSObject.Properties['Id']) { $RoleData.Id } else { '' }
            $directoryScopeId = if ($RoleData.PSObject.Properties['DirectoryScopeId'] -and $RoleData.DirectoryScopeId) { $RoleData.DirectoryScopeId } else { '/' }
            return "Entra|$roleDefinitionId|$directoryScopeId"
        }
        default {
            $displayName = if ($RoleData.PSObject.Properties['DisplayName'] -and $RoleData.DisplayName) { $RoleData.DisplayName } elseif ($RoleData.PSObject.Properties['Name']) { $RoleData.Name } else { '' }
            return "$type|$displayName"
        }
    }
}
