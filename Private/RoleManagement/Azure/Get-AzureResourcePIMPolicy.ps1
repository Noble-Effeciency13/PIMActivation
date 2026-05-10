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
        
        [Parameter()]
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
        $tokenObj = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop
        $token = if ($tokenObj.Token -is [System.Security.SecureString]) {
            [System.Net.NetworkCredential]::new('', $tokenObj.Token).Password
        } else { $tokenObj.Token }
        
        $policyApiVersion = '2020-10-01-preview'

        # Azure Resource PIM policy assignments are queried in the context of a scope,
        # but policy identity is role-based for this module's cache and UI behavior.
        $policyScope = if ($Scope) { $Scope } else { "/subscriptions/$SubscriptionId" }

        # Normalize role definition ID to GUID and construct correct path based on scope type
        $roleDefGuid = $RoleDefinitionId
        if ($roleDefGuid -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
            $roleDefGuid = $matches[1]
        }

        $isManagementGroupScope = ($policyScope -match "^/providers/Microsoft\.Management/managementGroups/")
        # Prefer the original full ARM path from the schedule; only reconstruct when it is a bare GUID
        $roleDefPath = if ($RoleDefinitionId -match "^/") {
            $RoleDefinitionId
        } elseif ($isManagementGroupScope) {
            "/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"
        } elseif ($SubscriptionId) {
            "/subscriptions/$SubscriptionId/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"
        } else {
            "/providers/Microsoft.Authorization/roleDefinitions/$roleDefGuid"
        }
        
        # Call Azure REST API to get PIM role settings
        $headers = @{
            'Authorization' = "Bearer $token"
            'Content-Type'  = 'application/json'
        }

        # Helper: build a "policy unavailable" object so callers can distinguish missing data from defaults
        $unavailablePolicy = [PSCustomObject]@{
            MaxDuration                      = 8
            RequiresMfa                      = $false
            RequiresJustification            = $true
            RequiresTicket                   = $false
            RequiresApproval                 = $false
            RequiresAuthenticationContext    = $false
            AuthenticationContextId          = $null
            AuthenticationContextDisplayName = $null
            AuthenticationContextDescription = $null
            AuthenticationContextDetails     = $null
            NotificationSettings             = $null
            ApprovalSettings                 = $null
            PolicyUnavailable                = $true
        }

        $policyLookupScopes = [System.Collections.ArrayList]::new()
        $seenLookupScopes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $AddPolicyLookupScope = {
            param([string]$CandidateScope)
            if ($CandidateScope -and $seenLookupScopes.Add($CandidateScope)) {
                $null = $policyLookupScopes.Add($CandidateScope)
            }
        }

        & $AddPolicyLookupScope $policyScope
        if ($policyScope -match '^(\/providers/Microsoft\.Management/managementGroups/[^\/]+)') {
            & $AddPolicyLookupScope $matches[1]
        }
        elseif ($policyScope -match '^(\/subscriptions\/[^\/]+)') {
            & $AddPolicyLookupScope $matches[1]
        }
        elseif ($SubscriptionId) {
            & $AddPolicyLookupScope "/subscriptions/$SubscriptionId"
        }

        if ($policyLookupScopes.Count -eq 0) {
            Write-Verbose "Get-AzureResourcePIMPolicy: no policy lookup scope available for '$roleDefGuid'; returning unavailable-policy marker."
            return $unavailablePolicy
        }

        # Helper: extract HTTP status code from a caught Invoke-RestMethod exception
        $GetHttpStatus = {
            param($err)
            try {
                if ($err.Exception.Response) { return [int]$err.Exception.Response.StatusCode }
            } catch {}
            return 0
        }

        # Helper: parse policy rules into the standard policy object
        $ParseRules = {
            param($rules)
            $maxDuration          = 8
            $requiresMfa          = $false
            $requiresJustification = $true
            $requiresTicket        = $false
            $requiresApproval      = $false
            $requiresAuthenticationContext = $false
            $authenticationContextId = $null
            foreach ($rule in $rules) {
                switch ($rule.id) {
                    'Expiration_EndUser_Assignment' {
                        if ($rule.maximumDuration) {
                            try {
                                $ts = [System.Xml.XmlConvert]::ToTimeSpan($rule.maximumDuration)
                                $maxDuration = [int]$ts.TotalHours
                            } catch {
                                if ($rule.maximumDuration -match 'PT(\d+)H') { $maxDuration = [int]$matches[1] }
                                elseif ($rule.maximumDuration -match 'PT(\d+)M') { $maxDuration = [Math]::Max(1, [int]([int]$matches[1] / 60)) }
                            }
                        }
                    }
                    'Enablement_EndUser_Assignment' {
                        if ($rule.enabledRules) {
                            $enabledArr = @($rule.enabledRules)
                            $requiresJustification = 'Justification' -in $enabledArr
                            $requiresMfa           = 'MultiFactorAuthentication' -in $enabledArr
                            $requiresTicket        = 'Ticketing' -in $enabledArr
                        }
                    }
                    'Approval_EndUser_Assignment' {
                        if ($rule.setting -and $null -ne $rule.setting.isApprovalRequired) {
                            $requiresApproval = [bool]$rule.setting.isApprovalRequired
                        }
                    }
                    'AuthenticationContext_EndUser_Assignment' {
                        if ($rule.isEnabled -eq $true) {
                            $requiresAuthenticationContext = $true
                            if ($rule.claimValue) { $authenticationContextId = $rule.claimValue }
                        }
                    }
                }
            }
            return [PSCustomObject]@{
                MaxDuration                      = $maxDuration
                RequiresMfa                      = $requiresMfa
                RequiresJustification            = $requiresJustification
                RequiresTicket                   = $requiresTicket
                RequiresApproval                 = $requiresApproval
                RequiresAuthenticationContext    = $requiresAuthenticationContext
                AuthenticationContextId          = $authenticationContextId
                AuthenticationContextDisplayName = $null
                AuthenticationContextDescription = $null
                AuthenticationContextDetails     = $null
                PolicyUnavailable                = $false
            }
        }

        $encodedRoleDefinitionId = [System.Uri]::EscapeDataString($roleDefPath)
        $filter = "roleDefinitionId%20eq%20'$encodedRoleDefinitionId'"

        foreach ($policyLookupScope in $policyLookupScopes) {
            $nextUri = "https://management.azure.com$policyLookupScope/providers/Microsoft.Authorization/roleManagementPolicyAssignments?api-version=$policyApiVersion&`$filter=$filter"
            $policyId = $null

            while ($nextUri -and -not $policyId) {
                $listResponse = $null
                try {
                    $listResponse = Invoke-RestMethod -Uri $nextUri -Headers $headers -Method Get -ErrorAction Stop
                } catch {
                    $httpStatus = & $GetHttpStatus $_
                    if ($httpStatus -in @(401, 403)) {
                        Write-Verbose "Get-AzureResourcePIMPolicy: HTTP $httpStatus at '$policyLookupScope' - Microsoft.Authorization/roleManagementPolicies/read required (Reader role). Returning unavailable policy."
                        return $unavailablePolicy
                    }
                    Write-Verbose "Get-AzureResourcePIMPolicy: policy assignment list failed at '$policyLookupScope' (HTTP $httpStatus): $($_.Exception.Message)"
                    break
                }

                $matched = @($listResponse.value) | Where-Object {
                    $rdId = $_.properties.roleDefinitionId
                    ($rdId -eq $roleDefPath) -or
                    ($rdId -match '/roleDefinitions/([a-fA-F0-9\-]{36})$' -and $matches[1] -eq $roleDefGuid) -or
                    ($rdId -eq $roleDefGuid)
                } | Select-Object -First 1

                if ($matched) {
                    $policyId = $matched.properties.policyId
                } else {
                    $nextUri = if ($listResponse.nextLink) { $listResponse.nextLink } else { $null }
                }
            }

            if (-not $policyId) {
                Write-Verbose "Get-AzureResourcePIMPolicy: no policy assignment matched '$roleDefGuid' at '$policyLookupScope'"
                continue
            }

            # policyId may be a full ARM path (/subscriptions/.../roleManagementPolicies/{name})
            # or just the policy name. Handle both.
            $policyUri = if ($policyId -match '^/') {
                "https://management.azure.com${policyId}?api-version=$policyApiVersion"
            } else {
                "https://management.azure.com$policyLookupScope/providers/Microsoft.Authorization/roleManagementPolicies/${policyId}?api-version=$policyApiVersion"
            }

            $policyResponse = $null
            try {
                $policyResponse = Invoke-RestMethod -Uri $policyUri -Headers $headers -Method Get -ErrorAction Stop
            } catch {
                $httpStatus = & $GetHttpStatus $_
                Write-Verbose "Get-AzureResourcePIMPolicy: policy fetch failed at '$policyLookupScope' (HTTP $httpStatus): $($_.Exception.Message)"
                continue
            }

            if ($policyResponse -and $policyResponse.properties -and $policyResponse.properties.rules) {
                Write-Verbose "Get-AzureResourcePIMPolicy: successfully parsed policy for '$roleDefGuid' at '$policyLookupScope'"
                return & $ParseRules -rules $policyResponse.properties.rules
            }
        }

        Write-Verbose "Get-AzureResourcePIMPolicy: could not retrieve policy for '$roleDefGuid' at scopes '$($policyLookupScopes -join ', ')'; returning unavailable-policy marker."
        return $unavailablePolicy
    }
    catch {
        Write-Verbose "Failed to retrieve Azure Resource PIM policy: $($_.Exception.Message)"
        return $unavailablePolicy
    }
}