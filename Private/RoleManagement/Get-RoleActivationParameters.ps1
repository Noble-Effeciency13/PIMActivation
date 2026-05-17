function Get-RoleActivationParameters {
    <#
    .SYNOPSIS
        Builds role activation request parameters for Microsoft Graph and Azure Resource PIM.

    .DESCRIPTION
        Creates the scheduleInfo and role-specific payload fields needed to submit PIM activation
        requests. The function supports immediate and scheduled activations by accepting an optional
        local start time and converting it to the UTC timestamp format expected by Graph and ARM.

    .PARAMETER RoleData
        Role metadata for the selected Entra, Group, or Azure Resource role.

    .PARAMETER Justification
        Justification text to include with the activation request.

    .PARAMETER EffectiveDuration
        Hashtable containing the duration after policy maximums have been applied.

    .PARAMETER TicketInfo
        Optional ticket metadata when the role policy requires ticketing.

    .PARAMETER ScheduleStartTime
        Optional local date and time when the activation should start. When omitted, the request
        starts immediately.

    .PARAMETER AzureTargetScope
        Optional reduced Azure Resource scope to use for Azure role activation.

    .PARAMETER LinkedRoleEligibilityScheduleId
        Optional Azure eligibility schedule ID required when activating at a reduced scope.

    .OUTPUTS
        Hashtable
        Returns a Graph-compatible activation hashtable, or an Azure Resource activation hashtable
        for Azure Resource roles.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$RoleData,
        
        [Parameter(Mandatory)]
        [string]$Justification,
        
        [Parameter(Mandatory)]
        [hashtable]$EffectiveDuration,
        
        [hashtable]$TicketInfo,

        [datetime]$ScheduleStartTime,

        [string]$AzureTargetScope,

        [string]$LinkedRoleEligibilityScheduleId
    )

    $activationStartDateTime = if ($PSBoundParameters.ContainsKey('ScheduleStartTime')) { [datetime]$ScheduleStartTime } else { Get-Date }
    if ($activationStartDateTime.Kind -eq [System.DateTimeKind]::Unspecified) {
        $activationStartDateTime = [datetime]::SpecifyKind($activationStartDateTime, [System.DateTimeKind]::Local)
    }
    else {
        $activationStartDateTime = $activationStartDateTime.ToLocalTime()
    }
    $activationStartDateTimeUtc = $activationStartDateTime.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
    
    $activationParams = @{
        action        = "selfActivate"
        justification = $Justification
        principalId   = $script:CurrentUser.Id
        scheduleInfo  = @{
            startDateTime = $activationStartDateTimeUtc
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
            $groupAccessCandidates = @()
            if ($RoleData.PSObject.Properties['AccessId'] -and $RoleData.AccessId) {
                $groupAccessCandidates += $RoleData.AccessId
            }
            if ($RoleData.PSObject.Properties['Assignment'] -and $RoleData.Assignment -and $RoleData.Assignment.PSObject.Properties['AccessId'] -and $RoleData.Assignment.AccessId) {
                $groupAccessCandidates += $RoleData.Assignment.AccessId
            }
            if ($RoleData.PSObject.Properties['MemberType'] -and $RoleData.MemberType) {
                $groupAccessCandidates += $RoleData.MemberType
            }
            if ($RoleData.PSObject.Properties['MembershipType'] -and $RoleData.MembershipType) {
                $groupAccessCandidates += $RoleData.MembershipType
            }

            $groupAccessId = 'member'
            foreach ($candidate in $groupAccessCandidates) {
                $normalizedAccessId = ([string]$candidate).Trim().ToLowerInvariant()
                if ($normalizedAccessId -in @('member', 'owner')) {
                    $groupAccessId = $normalizedAccessId
                    break
                }
            }

            $activationParams.groupId = $RoleData.GroupId
            $activationParams.accessId = $groupAccessId
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
                    StartDateTime = $activationStartDateTimeUtc
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