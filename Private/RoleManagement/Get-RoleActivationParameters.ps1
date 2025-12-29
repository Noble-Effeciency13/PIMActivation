function Get-RoleActivationParameters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$RoleData,
        
        [Parameter(Mandatory)]
        [string]$Justification,
        
        [Parameter(Mandatory)]
        [hashtable]$EffectiveDuration,
        
        [hashtable]$TicketInfo
    )
    
    $activationParams = @{
        action        = "selfActivate"
        justification = $Justification
        principalId   = $script:CurrentUser.Id
        scheduleInfo  = @{
            startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
            expiration    = @{
                duration = "PT$($EffectiveDuration.Hours)H$($EffectiveDuration.Minutes)M"
                type     = "afterDuration"
            }
        }
    }
    
    # Add role-specific parameters
    switch ($RoleData.Type) {
        'Entra' {
            $activationParams.roleDefinitionId = $RoleData.RoleDefinitionId
            $activationParams.directoryScopeId = if ($RoleData.DirectoryScopeId) { $RoleData.DirectoryScopeId } else { "/" }
        }
        
        'Group' {
            $activationParams.groupId = $RoleData.GroupId
            $activationParams.accessId = "member"
        }
        
        'AzureResource' {
            # Azure Resource roles use different parameter structure
            return @{
                Scope            = $RoleData.FullScope
                RoleDefinitionId = $RoleData.RoleDefinitionId  
                PrincipalId      = $script:CurrentUser.Id
                RequestType      = 'SelfActivate'
                Justification    = $Justification
                ScheduleInfo     = @{
                    StartDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                    Expiration    = @{
                        Type     = 'AfterDuration'
                        Duration = "PT$($EffectiveDuration.Hours)H$($EffectiveDuration.Minutes)M"
                    }
                }
                TicketInfo       = if ($TicketInfo -and $TicketInfo.ticketNumber) { $TicketInfo } else { $null }
            }
        }
    }
    
    # Add ticket info for Entra/Group roles if present
    if ($TicketInfo -and $TicketInfo.ticketNumber) {
        $activationParams.ticketInfo = $TicketInfo
    }
    
    return $activationParams
}