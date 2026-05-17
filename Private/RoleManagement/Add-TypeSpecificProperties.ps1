function Add-TypeSpecificProperties {
    param($FormattedRole, $SourceRole)
    
    switch ($SourceRole.Type) {
        'Entra' {
            $FormattedRole | Add-Member -NotePropertyName 'DirectoryScopeId' -NotePropertyValue $SourceRole.DirectoryScopeId
            $FormattedRole | Add-Member -NotePropertyName 'SubscriptionId' -NotePropertyValue $null
        }
        'Group' {
            $FormattedRole | Add-Member -NotePropertyName 'DirectoryScopeId' -NotePropertyValue $null
            $FormattedRole | Add-Member -NotePropertyName 'SubscriptionId' -NotePropertyValue $null
            $accessId = if ($SourceRole.PSObject.Properties['AccessId'] -and $SourceRole.AccessId) {
                $SourceRole.AccessId
            }
            elseif ($SourceRole.PSObject.Properties['Assignment'] -and $SourceRole.Assignment -and $SourceRole.Assignment.PSObject.Properties['AccessId'] -and $SourceRole.Assignment.AccessId) {
                $SourceRole.Assignment.AccessId
            }
            elseif ($SourceRole.PSObject.Properties['MemberType'] -and $SourceRole.MemberType -in @('member', 'owner')) {
                $SourceRole.MemberType
            }
            else {
                $null
            }
            $FormattedRole | Add-Member -NotePropertyName 'AccessId' -NotePropertyValue $accessId
        }
        default {
            # Future Azure resource roles
            $FormattedRole | Add-Member -NotePropertyName 'DirectoryScopeId' -NotePropertyValue $null
            $FormattedRole | Add-Member -NotePropertyName 'SubscriptionId' -NotePropertyValue $null
        }
    }
}