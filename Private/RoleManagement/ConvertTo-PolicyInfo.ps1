function ConvertTo-PolicyInfo {
    <#
    .SYNOPSIS
        Converts a Graph API policy object to a standardized policy info object.
    
    .PARAMETER Policy
        The policy object returned from the Graph API.
    
    .OUTPUTS
        PSCustomObject with standardized policy information.
    #>
    param(
        [Parameter(Mandatory)]
        $Policy
    )
    
    $policyInfo = [PSCustomObject]@{
        MaxDuration = 8
        RequiresMfa = $false
        RequiresJustification = $false
        RequiresTicket = $false
        RequiresApproval = $false
        RequiresAuthenticationContext = $false
        AuthenticationContextId = $null
        AuthenticationContextDisplayName = $null
        AuthenticationContextDescription = $null
        AuthenticationContextDetails = $null
    }
    
    if (-not $Policy.Rules) {
        Write-Verbose "Policy has no rules, returning defaults"
        return $policyInfo
    }
    
    foreach ($rule in $Policy.Rules) {
        $ruleType = $rule.AdditionalProperties['@odata.type'] ?? $rule.'@odata.type'
        
        switch ($ruleType) {
            '#microsoft.graph.unifiedRoleManagementPolicyExpirationRule' {
                if ($rule.AdditionalProperties.maximumDuration -or $rule.maximumDuration) {
                    $duration = $rule.AdditionalProperties.maximumDuration ?? $rule.maximumDuration
                    try {
                        $timespan = [System.Xml.XmlConvert]::ToTimeSpan($duration)
                        $policyInfo.MaxDuration = [int]$timespan.TotalHours
                        Write-Verbose "Set max duration to $($policyInfo.MaxDuration) hours"
                    }
                    catch {
                        Write-Verbose "Could not parse duration: $duration"
                    }
                }
            }
            '#microsoft.graph.unifiedRoleManagementPolicyEnablementRule' {
                $enabledRules = @($rule.AdditionalProperties.enabledRules ?? $rule.enabledRules ?? @())
                $policyInfo.RequiresJustification = 'Justification' -in $enabledRules
                $policyInfo.RequiresTicket = 'Ticketing' -in $enabledRules
                $policyInfo.RequiresMfa = 'MultiFactorAuthentication' -in $enabledRules
                $policyInfo.RequiresAuthenticationContext = 'AuthenticationContext' -in $enabledRules
                Write-Verbose "Enablement rules: MFA=$($policyInfo.RequiresMfa), Justification=$($policyInfo.RequiresJustification), Ticket=$($policyInfo.RequiresTicket), AuthContext=$($policyInfo.RequiresAuthenticationContext)"
            }
            '#microsoft.graph.unifiedRoleManagementPolicyApprovalRule' {
                $setting = $rule.AdditionalProperties.setting ?? $rule.setting
                if ($setting -and $setting.isApprovalRequired) {
                    $policyInfo.RequiresApproval = $true
                    Write-Verbose "Approval required: true"
                }
            }
            '#microsoft.graph.unifiedRoleManagementPolicyAuthenticationContextRule' {
                if (($rule.AdditionalProperties.isEnabled ?? $rule.isEnabled) -and 
                    ($rule.AdditionalProperties.claimValue ?? $rule.claimValue)) {
                    $policyInfo.RequiresAuthenticationContext = $true
                    $policyInfo.AuthenticationContextId = $rule.AdditionalProperties.claimValue ?? $rule.claimValue
                    Write-Verbose "Authentication context required: $($policyInfo.AuthenticationContextId)"
                }
            }
        }
    }
    
    return $policyInfo
}