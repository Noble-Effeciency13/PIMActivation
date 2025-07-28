function Get-PIMRolePolicy {
    <#
    .SYNOPSIS
        Retrieves policy information for a PIM role.
    
    .DESCRIPTION
        Gets the policy requirements for activating a specific PIM role including maximum duration,
        MFA requirements, justification requirements, approval requirements, and authentication context.
        
        Supports Entra ID roles, PIM for Groups, and Azure Resource roles. Uses intelligent caching
        to reduce repeated API calls and improve performance.
    
    .PARAMETER Role
        The role object to get policy information for. Must contain Type, Id, Name properties.
        For Group roles, must also contain ResourceId property.
    
    .EXAMPLE
        Get-PIMRolePolicy -Role $role
        Returns policy information for the specified role.
    
    .OUTPUTS
        PSCustomObject
        Returns an object with the following properties:
        - MaxDuration: Maximum activation duration in hours
        - RequiresMfa: Whether MFA is required for activation
        - RequiresJustification: Whether justification text is required
        - RequiresTicket: Whether ticket/tracking number is required
        - RequiresApproval: Whether approval workflow is required
        - RequiresAuthenticationContext: Whether specific authentication context is required
        - AuthenticationContextId: The ID of the required authentication context
        - AuthenticationContextDisplayName: Display name of the authentication context
        - AuthenticationContextDescription: Description of the authentication context
        - AuthenticationContextDetails: Full authentication context object
    
    .NOTES
        Uses module-level caching for both policies and authentication contexts to improve performance.
        Gracefully handles API failures by returning sensible defaults.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Role
    )
    
    # Initialize module-level cache for policies
    if (-not (Test-Path Variable:script:PolicyCache) -or -not $script:PolicyCache) {
        $script:PolicyCache = @{}
        Write-Verbose "Initialized policy cache"
    }
    
    # Create cache key based on role type and ID
    $cacheKey = "$($Role.Type)_$($Role.Id)"
    if ($Role.Type -eq 'Group') {
        $cacheKey = "Group_$($Role.ResourceId)"
    }
    
    # Return cached result if available
    if ($script:PolicyCache.ContainsKey($cacheKey)) {
        Write-Verbose "Retrieved cached policy for role: $($Role.Name)"
        return $script:PolicyCache[$cacheKey]
    }
    
    Write-Verbose "Retrieving policy for role: $($Role.Name) [Type: $($Role.Type)]"
    
    # Initialize policy object with default values
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
    
    # Initialize authentication context cache if needed
    if (-not $script:AuthenticationContextCache) {
        $script:AuthenticationContextCache = @{}
        Write-Verbose "Caching available authentication contexts..."
        
        try {
            $availableContexts = Get-AuthenticationContexts
            if ($availableContexts) {
                foreach ($context in $availableContexts) {
                    $script:AuthenticationContextCache[$context.Id] = $context
                }
                Write-Verbose "Cached $($availableContexts.Count) authentication contexts"
            }
        }
        catch {
            Write-Verbose "Authentication contexts not available: $($_.Exception.Message)"
        }
    }
    
    try {
        switch ($Role.Type) {
            'Entra' {
                Write-Verbose "Processing Entra ID role policy [RoleId: $($Role.Id)]"
                try {
                    # Get policy assignments for this role
                    $policyAssignments = Get-MgPolicyRoleManagementPolicyAssignment -Filter "scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '$($Role.Id)'" -ErrorAction Stop
                    
                    $policyAssignmentsArray = @($policyAssignments)
                    if ($policyAssignmentsArray.Count -eq 0) {
                        Write-Verbose "No policy assignments found for Entra role"
                        break
                    }
                    
                    $assignment = $policyAssignmentsArray[0]
                    Write-Verbose "Found policy assignment [PolicyId: $($assignment.PolicyId)]"
                    
                    # Get the policy with expanded rules
                    $policy = Get-MgPolicyRoleManagementPolicy -UnifiedRoleManagementPolicyId $assignment.PolicyId -ExpandProperty "rules" -ErrorAction Stop
                    
                    if ($policy -and $policy.Rules) {
                        $rulesArray = @($policy.Rules)
                        Write-Verbose "Processing $($rulesArray.Count) policy rules"
                        
                        foreach ($rule in $rulesArray) {
                            if ($rule.AdditionalProperties) {
                                $ruleType = $rule.AdditionalProperties['@odata.type']
                                
                                switch ($ruleType) {
                                    '#microsoft.graph.unifiedRoleManagementPolicyExpirationRule' {
                                        if ($rule.AdditionalProperties.maximumDuration) {
                                            try {
                                                $duration = [System.Xml.XmlConvert]::ToTimeSpan($rule.AdditionalProperties.maximumDuration)
                                                $policyInfo.MaxDuration = [int]$duration.TotalHours
                                            }
                                            catch {
                                                Write-Verbose "Could not parse duration: $($rule.AdditionalProperties.maximumDuration)"
                                            }
                                        }
                                    }
                                    '#microsoft.graph.unifiedRoleManagementPolicyEnablementRule' {
                                        if ($rule.AdditionalProperties.enabledRules) {
                                            $enabledRulesArray = @($rule.AdditionalProperties.enabledRules)
                                            $policyInfo.RequiresJustification = 'Justification' -in $enabledRulesArray
                                            $policyInfo.RequiresTicket = 'Ticketing' -in $enabledRulesArray
                                            $policyInfo.RequiresMfa = 'MultiFactorAuthentication' -in $enabledRulesArray
                                            $policyInfo.RequiresAuthenticationContext = 'AuthenticationContext' -in $enabledRulesArray
                                        }
                                    }
                                    '#microsoft.graph.unifiedRoleManagementPolicyApprovalRule' {
                                        if ($rule.AdditionalProperties.setting -and $rule.AdditionalProperties.setting.isApprovalRequired) {
                                            $policyInfo.RequiresApproval = $true
                                        }
                                    }
                                    '#microsoft.graph.unifiedRoleManagementPolicyAuthenticationContextRule' {
                                        if ($rule.AdditionalProperties.isEnabled -and $rule.AdditionalProperties.claimValue) {
                                            $policyInfo.RequiresAuthenticationContext = $true
                                            $contextId = $rule.AdditionalProperties.claimValue
                                            $policyInfo.AuthenticationContextId = $contextId
                                            
                                            # Enhance with cached context information
                                            if ($script:AuthenticationContextCache.ContainsKey($contextId)) {
                                                $contextInfo = $script:AuthenticationContextCache[$contextId]
                                                $policyInfo.AuthenticationContextDisplayName = $contextInfo.DisplayName
                                                $policyInfo.AuthenticationContextDescription = $contextInfo.Description
                                                $policyInfo.AuthenticationContextDetails = $contextInfo
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to retrieve Entra role policy for $($Role.Name): $($_.Exception.Message)"
                }
            }
            
            'Group' {
                Write-Verbose "Processing PIM for Groups policy [GroupId: $($Role.ResourceId)]"
                if ($Role.ResourceId) {
                    try {
                        # Get policy assignments for the group
                        $uri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=scopeId eq '$($Role.ResourceId)' and scopeType eq 'Group'"
                        $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
                        
                        if ($response.value -and @($response.value).Count -gt 0) {
                            $assignmentsArray = @($response.value)
                            
                            # Look for member role assignment first, then owner
                            $assignment = $assignmentsArray | Where-Object { $_.roleDefinitionId -eq 'member' } | Select-Object -First 1
                            if (-not $assignment) {
                                $assignment = $assignmentsArray | Where-Object { $_.roleDefinitionId -eq 'owner' } | Select-Object -First 1
                            }
                            
                            if ($assignment) {
                                Write-Verbose "Found group policy assignment [Role: $($assignment.roleDefinitionId), PolicyId: $($assignment.policyId)]"
                                
                                # Get the policy with expanded rules
                                $policyUri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies/$($assignment.policyId)?`$expand=rules"
                                $policyResponse = Invoke-MgGraphRequest -Uri $policyUri -Method GET -ErrorAction Stop
                                
                                if ($policyResponse.rules) {
                                    $rulesArray = @($policyResponse.rules)
                                    Write-Verbose "Processing $($rulesArray.Count) group policy rules"
                                    
                                    foreach ($rule in $rulesArray) {
                                        if ($rule.id -like "*Expiration_EndUser_Assignment" -and $rule.maximumDuration) {
                                            try {
                                                $duration = [System.Xml.XmlConvert]::ToTimeSpan($rule.maximumDuration)
                                                $policyInfo.MaxDuration = [int]$duration.TotalHours
                                            }
                                            catch {
                                                Write-Verbose "Could not parse group duration: $($rule.maximumDuration)"
                                            }
                                        }
                                        elseif ($rule.id -like "*Enablement_EndUser_Assignment" -and $rule.enabledRules) {
                                            $enabledRulesArray = @($rule.enabledRules)
                                            $policyInfo.RequiresJustification = 'Justification' -in $enabledRulesArray
                                            $policyInfo.RequiresTicket = 'Ticketing' -in $enabledRulesArray
                                            $policyInfo.RequiresMfa = 'MultiFactorAuthentication' -in $enabledRulesArray
                                            $policyInfo.RequiresAuthenticationContext = 'AuthenticationContext' -in $enabledRulesArray
                                        }
                                        elseif ($rule.id -like "*Approval_EndUser_Assignment" -and $rule.setting.isApprovalRequired) {
                                            $policyInfo.RequiresApproval = $true
                                        }
                                        elseif ($rule.id -like "*AuthenticationContext_EndUser_Assignment" -and $rule.isEnabled -and $rule.claimValue) {
                                            $policyInfo.RequiresAuthenticationContext = $true
                                            $contextId = $rule.claimValue
                                            $policyInfo.AuthenticationContextId = $contextId
                                            
                                            # Enhance with cached context information
                                            if ($script:AuthenticationContextCache.ContainsKey($contextId)) {
                                                $contextInfo = $script:AuthenticationContextCache[$contextId]
                                                $policyInfo.AuthenticationContextDisplayName = $contextInfo.DisplayName
                                                $policyInfo.AuthenticationContextDescription = $contextInfo.Description
                                                $policyInfo.AuthenticationContextDetails = $contextInfo
                                            }
                                        }
                                    }
                                }
                            }
                            else {
                                Write-Verbose "No suitable role assignment found for group"
                            }
                        }
                        else {
                            Write-Verbose "No policy assignments found for group"
                        }
                    }
                    catch {
                        Write-Warning "Failed to retrieve group policy for $($Role.Name): $($_.Exception.Message)"
                        # Set sensible defaults for groups
                        $policyInfo.RequiresJustification = $true
                    }
                }
            }
            
            'AzureResource' {
                Write-Verbose "Using default policy for Azure Resource role"
                $policyInfo.RequiresJustification = $true
            }
        }
        
        # Cache the result
        $script:PolicyCache[$cacheKey] = $policyInfo
        
        # Create summary for verbose output
        $requirements = @()
        if ($policyInfo.RequiresMfa) { $requirements += "MFA" }
        if ($policyInfo.RequiresJustification) { $requirements += "Justification" }
        if ($policyInfo.RequiresTicket) { $requirements += "Ticket" }
        if ($policyInfo.RequiresApproval) { $requirements += "Approval" }
        if ($policyInfo.RequiresAuthenticationContext) { 
            if ($policyInfo.AuthenticationContextDisplayName) {
                $requirements += "AuthContext ($($policyInfo.AuthenticationContextDisplayName))"
            } else {
                $requirements += "AuthContext ($($policyInfo.AuthenticationContextId))"
            }
        }
        
        $requirementsSummary = if ($requirements.Count -gt 0) { $requirements -join ", " } else { "None" }
        Write-Verbose "Policy cached for $($Role.Name): Duration=$($policyInfo.MaxDuration)h, Requirements=[$requirementsSummary]"
    }
    catch {
        Write-Warning "Failed to retrieve policy for role $($Role.Name): $($_.Exception.Message)"
        # Cache the default to avoid repeated failures
        $script:PolicyCache[$cacheKey] = $policyInfo
    }
    
    return $policyInfo
}
