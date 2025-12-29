function Get-AzureResourcePIMPolicy {
    <#
    .SYNOPSIS
        Retrieves Azure Resource PIM policy settings for a specific role.
    
    .DESCRIPTION
        Gets the actual PIM policy configuration for Azure Resource roles including
        activation requirements, maximum duration, approval settings, etc.
    
    .PARAMETER RoleDefinitionId
        The Azure role definition ID.
    
    .PARAMETER SubscriptionId
        The subscription ID where the role is assigned.
    
    .PARAMETER Scope
        The specific scope of the role assignment.
    
    .OUTPUTS
        PSCustomObject containing policy information
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RoleDefinitionId,
        
        [Parameter(Mandatory)]
        [string]$SubscriptionId,
        
        [Parameter()]
        [string]$Scope
    )
    
    Write-Verbose "Fetching Azure Resource PIM policy for role: $RoleDefinitionId in subscription: $SubscriptionId"
    
    try {
        # Azure Resource PIM policies are retrieved differently than Entra ID
        # They use the Azure Management API directly
        
        $context = Get-AzContext
        if (-not $context) {
            Write-Warning "No Azure context available for policy retrieval"
            return $null
        }
        
        # Get access token for Azure Management API
        $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id, $null, "Never", $null, "https://management.azure.com/").AccessToken
        
        # Azure Resource PIM policies are scoped to the resource scope
        $policyScope = if ($Scope) { $Scope } else { "/subscriptions/$SubscriptionId" }

        # Normalize role definition ID to GUID and construct correct path based on scope type
        $roleDefGuid = $RoleDefinitionId
        if ($roleDefGuid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
            $roleDefGuid = $matches[1]
        }

        $isManagementGroupScope = ($policyScope -match "^/providers/Microsoft\.Management/managementGroups/")
        $roleDefPath = if ($isManagementGroupScope) {
            "/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"
        } else {
            "/subscriptions/$SubscriptionId/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"
        }
        
        # Call Azure REST API to get PIM role settings
        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type'  = 'application/json'
        }
        
        # Try to get the role setting for this specific role definition at this scope
        $uri = "https://management.azure.com$policyScope/providers/Microsoft.Authorization/roleManagementPolicyAssignments?api-version=2020-10-01&`$filter=roleDefinitionId eq '$roleDefPath'"

        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction SilentlyContinue
        
        if ($response -and $response.value -and $response.value.Count -gt 0) {
            $policyAssignment = $response.value[0]
            $policyId = $policyAssignment.properties.policyId
            
            # Get the actual policy definition
            $policyUri = "https://management.azure.com$policyScope/providers/Microsoft.Authorization/roleManagementPolicies/${policyId}?api-version=2020-10-01"
            $policyResponse = Invoke-RestMethod -Uri $policyUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
            
            if ($policyResponse -and $policyResponse.properties) {
                $policy = $policyResponse.properties
                
                # Parse the policy rules to extract relevant settings
                $maxDuration = 8  # Default
                $requiresMfa = $false  # Azure Resource roles typically don't require MFA activation
                $requiresJustification = $true  # Default
                $requiresApproval = $false
                
                if ($policy.rules) {
                    foreach ($rule in $policy.rules) {
                        switch ($rule.id) {
                            'Activation_Admin_Duration' {
                                if ($rule.maximumDuration) {
                                    # Parse duration (format: PT8H for 8 hours)
                                    if ($rule.maximumDuration -match "PT(\d+)H") {
                                        $maxDuration = [int]$matches[1]
                                    }
                                }
                            }
                            'Activation_Admin_MFA' {
                                $requiresMfa = $rule.setting.mfaRequired -eq $true
                            }
                            'Activation_Admin_Justification' {
                                $requiresJustification = $rule.setting.justificationRequired -eq $true
                            }
                            'Activation_Admin_Approval' {
                                $requiresApproval = $rule.setting.approvalRequired -eq $true
                            }
                        }
                    }
                }
                
                Write-Verbose "Retrieved Azure Resource PIM policy: MaxDuration=$maxDuration, MFA=$requiresMfa, Justification=$requiresJustification, Approval=$requiresApproval"
                
                return [PSCustomObject]@{
                    MaxDuration                      = $maxDuration
                    RequiresMfa                      = $requiresMfa
                    RequiresJustification            = $requiresJustification
                    RequiresTicket                   = $false  # Azure Resource roles don't typically use tickets
                    RequiresApproval                 = $requiresApproval
                    RequiresAuthenticationContext    = $false  # Not used for Azure Resource roles
                    AuthenticationContextId          = $null
                    AuthenticationContextDisplayName = $null
                    AuthenticationContextDescription = $null
                    AuthenticationContextDetails     = $null
                    NotificationSettings             = $policy.notificationSettings
                    ApprovalSettings                 = if ($requiresApproval) { $policy.approvalSettings } else { $null }
                }
            }
        }
        
        # Fallback: enumerate all policy assignments at scope and match locally when server-side filter fails
        try {
            $listUri = "https://management.azure.com$policyScope/providers/Microsoft.Authorization/roleManagementPolicyAssignments?api-version=2020-10-01"
            $listResponse = Invoke-RestMethod -Uri $listUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
            if ($listResponse -and $listResponse.value -and $listResponse.value.Count -gt 0) {
                $matched = $listResponse.value | Where-Object {
                    $_.properties.roleDefinitionId -eq $roleDefPath -or
                    ($_.properties.roleDefinitionId -match "([a-fA-F0-9\-]{36})" -and $matches[1] -eq $roleDefGuid)
                } | Select-Object -First 1

                if ($matched) {
                    $policyId = $matched.properties.policyId
                    $policyUri = "https://management.azure.com$policyScope/providers/Microsoft.Authorization/roleManagementPolicies/${policyId}?api-version=2020-10-01"
                    $policyResponse = Invoke-RestMethod -Uri $policyUri -Headers $headers -Method Get -ErrorAction SilentlyContinue
                    if ($policyResponse -and $policyResponse.properties) {
                        $policy = $policyResponse.properties

                        # Parse the policy rules to extract relevant settings
                        $maxDuration = 8  # Default
                        $requiresMfa = $false
                        $requiresJustification = $true
                        $requiresApproval = $false

                        if ($policy.rules) {
                            foreach ($rule in $policy.rules) {
                                switch ($rule.id) {
                                    'Activation_Admin_Duration' {
                                        if ($rule.maximumDuration -and ($rule.maximumDuration -match "PT(\d+)H")) { $maxDuration = [int]$matches[1] }
                                    }
                                    'Activation_Admin_MFA' { $requiresMfa = $rule.setting.mfaRequired -eq $true }
                                    'Activation_Admin_Justification' { $requiresJustification = $rule.setting.justificationRequired -eq $true }
                                    'Activation_Admin_Approval' { $requiresApproval = $rule.setting.approvalRequired -eq $true }
                                }
                            }
                        }

                        Write-Verbose "Retrieved Azure Resource PIM policy via fallback: MaxDuration=$maxDuration, MFA=$requiresMfa, Justification=$requiresJustification, Approval=$requiresApproval"

                        return [PSCustomObject]@{
                            MaxDuration                      = $maxDuration
                            RequiresMfa                      = $requiresMfa
                            RequiresJustification            = $requiresJustification
                            RequiresTicket                   = $false
                            RequiresApproval                 = $requiresApproval
                            RequiresAuthenticationContext    = $false
                            AuthenticationContextId          = $null
                            AuthenticationContextDisplayName = $null
                            AuthenticationContextDescription = $null
                            AuthenticationContextDetails     = $null
                            NotificationSettings             = $policy.notificationSettings
                            ApprovalSettings                 = if ($requiresApproval) { $policy.approvalSettings } else { $null }
                        }
                    }
                }
            }
        } catch { Write-Verbose "Fallback Azure Resource PIM policy enumeration failed: $($_.Exception.Message)" }

        Write-Verbose "Could not retrieve Azure Resource PIM policy; returning null to use defaults upstream"
        return $null
    }
    catch {
        Write-Verbose "Failed to retrieve Azure Resource PIM policy: $($_.Exception.Message)"
        return $null
    }
}