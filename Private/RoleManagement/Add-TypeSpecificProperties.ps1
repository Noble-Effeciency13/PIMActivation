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
        }
        default {
            # Future Azure resource roles
            $FormattedRole | Add-Member -NotePropertyName 'DirectoryScopeId' -NotePropertyValue $null
            $FormattedRole | Add-Member -NotePropertyName 'SubscriptionId' -NotePropertyValue $null
        }
    }
}