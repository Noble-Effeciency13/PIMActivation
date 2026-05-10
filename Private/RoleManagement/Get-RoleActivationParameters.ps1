function Get-RoleActivationParameters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$RoleData,
        
        [Parameter(Mandatory)]
        [string]$Justification,
        
        [Parameter(Mandatory)]
        [hashtable]$EffectiveDuration,
        
        [hashtable]$TicketInfo,

        [string]$AzureTargetScope,

        [string]$LinkedRoleEligibilityScheduleId
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
            $originalScope = if ($RoleData.PSObject.Properties['FullScope'] -and $RoleData.FullScope) {
                $RoleData.FullScope
            }
            elseif ($RoleData.PSObject.Properties['Scope'] -and $RoleData.Scope) {
                $RoleData.Scope
            }
            elseif ($RoleData.PSObject.Properties['DirectoryScopeId'] -and $RoleData.DirectoryScopeId) {
                $RoleData.DirectoryScopeId
            }
            else {
                throw "Azure Resource role '$($RoleData.DisplayName)' is missing an activation scope."
            }
            $targetScope = if (-not [string]::IsNullOrWhiteSpace($AzureTargetScope)) { $AzureTargetScope } else { $originalScope }
            $scopeValidation = Test-AzureReducedScope -OriginalScope $originalScope -TargetScope $targetScope

            if (-not $scopeValidation.IsValid) {
                throw $scopeValidation.ErrorMessage
            }

            $effectiveScope = $scopeValidation.TargetScope
            $originalScope = $scopeValidation.OriginalScope
            $isReducedScope = [bool]$scopeValidation.IsReducedScope
            $linkedEligibilityId = if ($LinkedRoleEligibilityScheduleId) {
                $LinkedRoleEligibilityScheduleId
            }
            elseif ($isReducedScope -and $RoleData.PSObject.Properties['EligibilityScheduleName'] -and $RoleData.EligibilityScheduleName) {
                $RoleData.EligibilityScheduleName
            }
            elseif ($isReducedScope -and $RoleData.PSObject.Properties['EligibilityScheduleId'] -and $RoleData.EligibilityScheduleId) {
                $RoleData.EligibilityScheduleId
            }
            else {
                $null
            }

            $roleDefinitionId = $RoleData.RoleDefinitionId
            if ($roleDefinitionId -and -not $roleDefinitionId.StartsWith('/')) {
                if ($originalScope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                    $roleDefinitionId = "/subscriptions/$($matches[1])/providers/Microsoft.Authorization/roleDefinitions/$roleDefinitionId"
                }
                elseif ($effectiveScope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                    $roleDefinitionId = "/subscriptions/$($matches[1])/providers/Microsoft.Authorization/roleDefinitions/$roleDefinitionId"
                }
                else {
                    $roleDefinitionId = "/providers/Microsoft.Authorization/roleDefinitions/$roleDefinitionId"
                }
            }

            # Azure Resource roles use different parameter structure
            return @{
                Scope                           = $effectiveScope
                OriginalScope                   = $originalScope
                IsReducedScope                  = $isReducedScope
                LinkedRoleEligibilityScheduleId = if ($isReducedScope) { $linkedEligibilityId } else { $null }
                RoleDefinitionId                = $roleDefinitionId
                PrincipalId                     = $script:CurrentUser.Id
                RequestType                     = 'SelfActivate'
                Justification                   = $Justification
                ScheduleInfo                    = @{
                    StartDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                    Expiration    = @{
                        Type     = 'AfterDuration'
                        Duration = "PT$($EffectiveDuration.Hours)H$($EffectiveDuration.Minutes)M"
                    }
                }
                TicketInfo                      = if ($TicketInfo -and $TicketInfo.ticketNumber) { $TicketInfo } else { $null }
            }
        }
    }
    
    # Add ticket info for Entra/Group roles if present
    if ($TicketInfo -and $TicketInfo.ticketNumber) {
        $activationParams.ticketInfo = $TicketInfo
    }
    
    return $activationParams
}