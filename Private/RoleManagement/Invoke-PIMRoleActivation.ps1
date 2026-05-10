function Invoke-PIMRoleActivation {
    <#
    .SYNOPSIS
        Activates selected PIM (Privileged Identity Management) roles with enhanced error handling and policy compliance.
    
    .DESCRIPTION
        Handles the complete PIM role activation process including:
        - Policy requirement validation (justification, tickets, MFA, authentication context)
        - Duration calculations based on role policies
        - Authentication context challenges for conditional access policies
        - Both Entra ID directory roles and PIM-enabled groups
        - Comprehensive error handling with user-friendly messages
        
        The function supports both standard Microsoft Graph SDK calls and direct REST API calls
        for roles requiring authentication context tokens.
    
    .PARAMETER CheckedItems
        Array of checked ListView items representing the roles to activate.
        Each item must have a Tag property containing role metadata.
    
    .PARAMETER Form
        Reference to the main Windows Forms object for UI updates and refresh operations.
    
    .EXAMPLE
        Invoke-PIMRoleActivation -CheckedItems $selectedRoles -Form $mainForm
        
        Activates the selected PIM roles with appropriate policy validation.
    
    .NOTES
        - Requires Microsoft Graph PowerShell SDK
        - Supports authentication context challenges for conditional access
        - Handles both directory roles and group memberships
        - Duration is automatically adjusted based on role policy limits
        - Uses script-scoped variables for authentication state management
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$CheckedItems,
        
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form
    )
    
    Write-Verbose "Starting activation process for $($CheckedItems.Count) role(s)"
    
    # Initialize the splash form variable
    $operationSplash = $null

    $GetActivationResponseId = {
        param($ActivationResult)

        if (-not $ActivationResult) { return $null }

        $response = $null
        if ($ActivationResult -is [System.Collections.IDictionary]) {
            if ($ActivationResult.Contains('Response')) { $response = $ActivationResult['Response'] }
        }
        elseif ($ActivationResult.PSObject.Properties['Response']) {
            $response = $ActivationResult.Response
        }

        if (-not $response) { return $null }

        if ($response -is [System.Collections.IDictionary]) {
            if ($response.Contains('id')) { return [string]$response['id'] }
            if ($response.Contains('Id')) { return [string]$response['Id'] }
        }
        else {
            if ($response.PSObject.Properties['id']) { return [string]$response.id }
            if ($response.PSObject.Properties['Id']) { return [string]$response.Id }
        }

        return $null
    }

    $GetStatusSafeRoleList = {
        param([string[]]$RoleNames)

        $roleList = (@($RoleNames | Where-Object { $_ }) -join ', ')
        if ($roleList.Length -gt 90) {
            return "$($roleList.Substring(0, 87))..."
        }
        return $roleList
    }
    
    try {
        # Initialize duration from script variable or use default
        $requestedHours = 8
        $requestedMinutes = 0
        
        if ($script:RequestedDuration) {
            $requestedHours = $script:RequestedDuration.Hours
            $requestedMinutes = $script:RequestedDuration.Minutes
        }
        else {
            # Get from form controls if available
            $cmbHours = $Form.Controls.Find("cmbHours", $true)[0]
            $cmbMinutes = $Form.Controls.Find("cmbMinutes", $true)[0]
            
            if ($cmbHours -and $cmbMinutes) {
                $requestedHours = [int]$cmbHours.SelectedItem
                $requestedMinutes = [int]$cmbMinutes.SelectedItem
            }
        }
        
        $requestedTotalMinutes = ($requestedHours * 60) + $requestedMinutes
        Write-Verbose "Using requested duration: $requestedHours hours, $requestedMinutes minutes"

        # Analyze policy requirements across all selected roles
        $policyRequirements = @{
            RequiresJustification = $false
            RequiresTicket        = $false
            RequiresMfa           = $false
            RequiresAuthContext   = $false
            AuthContextIds        = @()
        }
        
        foreach ($item in $CheckedItems) {
            $roleData = $item.Tag
            if ($roleData.PolicyInfo) {
                if ($roleData.PolicyInfo.RequiresJustification) { $policyRequirements.RequiresJustification = $true }
                if ($roleData.PolicyInfo.RequiresTicket) { $policyRequirements.RequiresTicket = $true }
                if ($roleData.PolicyInfo.RequiresMfa) { $policyRequirements.RequiresMfa = $true }
                if ($roleData.PolicyInfo.RequiresAuthenticationContext -and $roleData.PolicyInfo.AuthenticationContextId) {
                    $policyRequirements.RequiresAuthContext = $true
                    $policyRequirements.AuthContextIds += $roleData.PolicyInfo.AuthenticationContextId
                }
            }
        }
        
        # Remove duplicate authentication contexts
        $policyRequirements.AuthContextIds = @($policyRequirements.AuthContextIds | Select-Object -Unique)
        
        Write-Verbose "Policy analysis complete - Justification: $($policyRequirements.RequiresJustification), Ticket: $($policyRequirements.RequiresTicket), MFA: $($policyRequirements.RequiresMfa), Auth Context: $($policyRequirements.RequiresAuthContext)"
        
        # Collect justification and ticket information
        $justification = "PowerShell activation"
        $ticketInfo = $null  # Initialize as null instead of empty hashtable
        $selectedAzureRoleItems = @($CheckedItems | Where-Object { $_.Tag -and $_.Tag.Type -eq 'AzureResource' -and $_.Tag.Status -eq 'Eligible' })
        $showAzureReducedScope = $selectedAzureRoleItems.Count -gt 0
        $result = $null
        
        # Show activation dialog for required or optional information
        if ($policyRequirements.RequiresJustification -or $policyRequirements.RequiresTicket -or $CheckedItems.Count -gt 0) {
            Write-Verbose "Showing activation dialog for justification/ticket requirements"
            $result = Show-PIMActivationDialog -RequiresJustification:$policyRequirements.RequiresJustification `
                -RequiresTicket:$policyRequirements.RequiresTicket `
                -OptionalJustification:$(-not $policyRequirements.RequiresJustification) `
                -ShowAzureReducedScope:$showAzureReducedScope
            
            if ($result.Cancelled) {
                Write-Verbose "User cancelled activation"
                return
            }
            
            $justification = $result.Justification
            if ($result.TicketNumber) {
                $ticketInfo = @{
                    ticketNumber = $result.TicketNumber
                    ticketSystem = $result.TicketSystem
                }
            }
        }

        $azureReducedScopeOverrides = @{}
        $azureTargetScope = if ($result -and $result.PSObject.Properties['AzureReducedScope'] -and -not [string]::IsNullOrWhiteSpace($result.AzureReducedScope)) {
            $result.AzureReducedScope.Trim()
        }
        else {
            $null
        }

        if ($azureTargetScope) {
            Write-Verbose "Azure reduced scope entered in activation dialog: $azureTargetScope"

            foreach ($item in $selectedAzureRoleItems) {
                $roleData = $item.Tag
                $displayName = if ($roleData.PSObject.Properties['DisplayName'] -and $roleData.DisplayName) { $roleData.DisplayName } else { 'Azure Resource role' }
                $originalScope = if ($roleData.PSObject.Properties['FullScope'] -and $roleData.FullScope) {
                    $roleData.FullScope
                }
                elseif ($roleData.PSObject.Properties['Scope'] -and $roleData.Scope -match '^/') {
                    $roleData.Scope
                }
                elseif ($roleData.PSObject.Properties['DirectoryScopeId'] -and $roleData.DirectoryScopeId) {
                    $roleData.DirectoryScopeId
                }
                else {
                    Show-TopMostMessageBox -Message "Cannot use reduced scope for '$displayName' because the original Azure scope was not found." -Title 'Reduced Scope Validation' -Icon Warning
                    return
                }

                $validation = Test-AzureReducedScope -OriginalScope $originalScope -TargetScope $azureTargetScope
                if (-not $validation.IsValid) {
                    Show-TopMostMessageBox -Message "Reduced scope is not valid for '$displayName'.`r`n$($validation.ErrorMessage)" -Title 'Reduced Scope Validation' -Icon Warning
                    return
                }

                if ($validation.IsReducedScope) {
                    $linkedEligibilityId = if ($roleData.PSObject.Properties['EligibilityScheduleName'] -and $roleData.EligibilityScheduleName) {
                        $roleData.EligibilityScheduleName
                    }
                    elseif ($roleData.PSObject.Properties['EligibilityScheduleId'] -and $roleData.EligibilityScheduleId) {
                        $roleData.EligibilityScheduleId
                    }
                    else {
                        $null
                    }

                    if (-not $linkedEligibilityId) {
                        Show-TopMostMessageBox -Message "Cannot use reduced scope for '$displayName' because the eligibility schedule ID was not found." -Title 'Reduced Scope Validation' -Icon Warning
                        return
                    }

                    $scopeKey = Get-AzureReducedScopeKey -RoleData $roleData
                    $azureReducedScopeOverrides[$scopeKey] = [PSCustomObject]@{
                        TargetScope                     = $validation.TargetScope
                        OriginalScope                   = $validation.OriginalScope
                        LinkedRoleEligibilityScheduleId = $linkedEligibilityId
                    }
                }
            }

            Write-Verbose "Azure reduced scope override count: $($azureReducedScopeOverrides.Count)"
        }

        $GetAzureScopeOverride = {
            param($RoleData)

            if (-not $azureReducedScopeOverrides -or $azureReducedScopeOverrides.Count -eq 0) { return $null }

            $overrideKey = Get-AzureReducedScopeKey -RoleData $RoleData
            if ($azureReducedScopeOverrides.ContainsKey($overrideKey)) {
                return $azureReducedScopeOverrides[$overrideKey]
            }

            return $null
        }
        
        # Split roles by whether they require authentication context step-up.
        $authContextRoles = @()
        $noContextRoles = @()
        
        foreach ($item in $CheckedItems) {
            $roleData = $item.Tag
            
            if ($roleData.PolicyInfo -and $roleData.PolicyInfo.RequiresAuthenticationContext -and $roleData.PolicyInfo.AuthenticationContextId) {
                $authContextRoles += $item
            }
            else {
                $noContextRoles += $item
            }
        }

        $authContextIds = @($authContextRoles | ForEach-Object { $_.Tag.PolicyInfo.AuthenticationContextId } | Where-Object { $_ } | Sort-Object -Unique)
        $authContextLabel = $authContextIds -join ', '
        
        Write-Verbose "Authentication-context roles: $($authContextRoles.Count) across $($authContextIds.Count) context(s) [$authContextLabel], $($noContextRoles.Count) without context"

        # NOW show the splash form after all user input has been collected
        $operationSplash = Show-OperationSplash -Title "Role Activation" -InitialMessage "Processing role activations..." -ShowProgressBar $true
        $activationErrors = @()
        $successCount = 0
        $totalRoles = $CheckedItems.Count
        $currentRole = 0
        
        # Process roles that require authentication context with one shared step-up per resource.
        if ($authContextRoles.Count -gt 0) {
            $authContextGraphItems = @($authContextRoles | Where-Object { $_.Tag.Type -in @('Entra', 'Group') })
            $authContextAzureItems = @($authContextRoles | Where-Object { $_.Tag.Type -eq 'AzureResource' })
            $graphAuthContextIds = @($authContextGraphItems | ForEach-Object { $_.Tag.PolicyInfo.AuthenticationContextId } | Where-Object { $_ } | Sort-Object -Unique)
            $azureAuthContextIds = @($authContextAzureItems | ForEach-Object { $_.Tag.PolicyInfo.AuthenticationContextId } | Where-Object { $_ } | Sort-Object -Unique)
            $graphAuthContextLabel = $graphAuthContextIds -join ', '

            Write-Verbose "Authentication-context roles to activate: $($authContextGraphItems.Count) Graph, $($authContextAzureItems.Count) Azure Resource"

            if ($operationSplash -and -not $operationSplash.IsDisposed) {
                $operationSplash.UpdateStatus("Completing authentication requirements: $authContextLabel", 15)
            }

            if ($authContextGraphItems.Count -gt 0) {
                $graphAuthContextToken = Get-AuthenticationContextToken -ContextIds $graphAuthContextIds -Scopes @('https://graph.microsoft.com/.default') -CacheNamespace 'Graph'

                if (-not $graphAuthContextToken) {
                    Write-Warning "Failed to obtain a shared Graph authentication context token. Falling back to sequential activation for Graph auth-context roles."

                    foreach ($item in $authContextGraphItems) {
                        $currentRole++
                        $roleData = $item.Tag
                        $contextId = $roleData.PolicyInfo.AuthenticationContextId
                        $progressPercent = [int](($currentRole / $totalRoles) * 100)

                        if ($operationSplash -and -not $operationSplash.IsDisposed) {
                            $scopeInfo = if ($roleData.Scope -and $roleData.Scope -ne "Directory") { " [$($roleData.Scope)]" } else { "" }
                            $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)$scopeInfo... ($currentRole of $totalRoles)", $progressPercent)
                        }

                        $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration

                        try {
                            $result = Invoke-SingleRoleActivation -RoleData $roleData -Justification $justification -EffectiveDuration $effectiveDuration -TicketInfo $ticketInfo -AuthenticationContextId $contextId -UseFallbackMethod
                            if ($result.Success) {
                                $successCount++
                                Write-Verbose "$($roleData.Type) role activated with authentication context fallback"
                            }
                            else {
                                $friendlyError = if ($result.Error -and $result.Error.Exception) { Get-FriendlyErrorMessage -Exception $result.Error.Exception -ErrorDetails $result.ErrorDetails } elseif ($result.ErrorMessage) { $result.ErrorMessage } else { "Activation failed" }
                                $activationErrors += "$($roleData.DisplayName): $friendlyError"
                                Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
                            }
                        }
                        catch {
                            $activationErrors += "$($roleData.DisplayName): $($_.Exception.Message)"
                            Write-Warning "Failed to activate $($roleData.DisplayName): $($_.Exception.Message)"
                        }
                    }
                }
                else {
                    Write-Verbose "Successfully obtained shared Graph authentication context token for context(s): $graphAuthContextLabel"

                    $graphAuthBatchRequestBodies = [System.Collections.ArrayList]::new()
                    $graphAuthBatchMeta = [System.Collections.ArrayList]::new()

                    foreach ($item in $authContextGraphItems) {
                        $roleData = $item.Tag
                        $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                        $activationParams = Get-RoleActivationParameters -RoleData $roleData -Justification $justification -EffectiveDuration $effectiveDuration -TicketInfo $ticketInfo
                        $batchId = "$($graphAuthBatchRequestBodies.Count + 1)"

                        $reqUrl = switch ($roleData.Type) {
                            'Entra' { '/roleManagement/directory/roleAssignmentScheduleRequests' }
                            'Group' { '/identityGovernance/privilegedAccess/group/assignmentScheduleRequests' }
                        }

                        $reqBody = if ($roleData.Type -eq 'Entra') {
                            $requestBodyData = @{
                                action           = $activationParams.action
                                principalId      = $activationParams.principalId
                                roleDefinitionId = $activationParams.roleDefinitionId
                                directoryScopeId = $activationParams.directoryScopeId
                                justification    = $activationParams.justification
                                scheduleInfo     = $activationParams.scheduleInfo
                            }
                            if ($activationParams.ContainsKey('ticketInfo') -and $activationParams.ticketInfo) { $requestBodyData.ticketInfo = $activationParams.ticketInfo }
                            $requestBodyData
                        }
                        else {
                            $requestBodyData = @{
                                action        = $activationParams.action
                                principalId   = $activationParams.principalId
                                groupId       = $activationParams.groupId
                                accessId      = $activationParams.accessId
                                justification = $activationParams.justification
                                scheduleInfo  = $activationParams.scheduleInfo
                            }
                            if ($activationParams.ContainsKey('ticketInfo') -and $activationParams.ticketInfo) { $requestBodyData.ticketInfo = $activationParams.ticketInfo }
                            $requestBodyData
                        }

                        $null = $graphAuthBatchRequestBodies.Add(@{
                            id      = $batchId
                            method  = 'POST'
                            url     = $reqUrl
                            headers = @{ 'Content-Type' = 'application/json' }
                            body    = $reqBody
                        })
                        $null = $graphAuthBatchMeta.Add([PSCustomObject]@{
                            BatchId           = $batchId
                            Item              = $item
                            RoleData          = $roleData
                            EffectiveDuration = $effectiveDuration
                        })
                    }

                    $authContextBatchSize = 20
                    for ($bi = 0; $bi -lt $graphAuthBatchRequestBodies.Count; $bi += $authContextBatchSize) {
                        $chunkEnd = [Math]::Min($bi + $authContextBatchSize - 1, $graphAuthBatchRequestBodies.Count - 1)
                        $chunk = @($graphAuthBatchRequestBodies[$bi..$chunkEnd])
                        $batchBody = @{ requests = $chunk }
                        $batchNumber = [int]($bi / $authContextBatchSize) + 1
                        $chunkRoleNames = @(
                            foreach ($reqItem in $chunk) {
                                $meta = $graphAuthBatchMeta | Where-Object { $_.BatchId -eq $reqItem.id } | Select-Object -First 1
                                if ($meta -and $meta.RoleData) { $meta.RoleData.DisplayName }
                            }
                        )
                        $chunkRoleList = @($chunkRoleNames | Where-Object { $_ }) -join ', '
                        $statusRoleList = & $GetStatusSafeRoleList $chunkRoleNames

                        try {
                            Write-Verbose "Submitting Graph authentication-context batch $batchNumber with $($chunk.Count) role(s): $chunkRoleList"
                            if ($operationSplash -and -not $operationSplash.IsDisposed) {
                                $operationSplash.UpdateStatus("Submitting auth-context Graph batch $batchNumber`: $statusRoleList", 35)
                            }

                            $graphBatchHeaders = @{ 'Authorization' = "Bearer $graphAuthContextToken"; 'Content-Type' = 'application/json' }
                            $batchJson = $batchBody | ConvertTo-Json -Depth 12 -Compress
                            $batchResponse = Invoke-RestMethod -Method Post -Uri 'https://graph.microsoft.com/v1.0/$batch' -Headers $graphBatchHeaders -Body $batchJson -ErrorAction Stop

                            foreach ($resp in @($batchResponse.responses)) {
                                $meta = $graphAuthBatchMeta | Where-Object { $_.BatchId -eq $resp.id } | Select-Object -First 1
                                if (-not $meta) { continue }
                                $currentRole++
                                $roleData = $meta.RoleData

                                if ($resp.status -ge 200 -and $resp.status -lt 300) {
                                    Write-Verbose "$($roleData.Type) role activated via auth-context Graph batch - id: $($resp.body.id)"
                                    $successCount++
                                }
                                else {
                                    $errMsg = if ($resp.body -and $resp.body.error -and $resp.body.error.message) { $resp.body.error.message } else { "HTTP $($resp.status)" }
                                    $activationErrors += "$($roleData.DisplayName): $errMsg"
                                    Write-Warning "Auth-context Graph batch activation failed for $($roleData.DisplayName): $errMsg"
                                }
                            }
                        }
                        catch {
                            Write-Warning "Auth-context Graph batch request failed, falling back to sequential for this chunk: $($_.Exception.Message)"
                            Write-Verbose "Auth-context Graph fallback batch $batchNumber roles: $chunkRoleList"

                            foreach ($reqItem in $chunk) {
                                $meta = $graphAuthBatchMeta | Where-Object { $_.BatchId -eq $reqItem.id } | Select-Object -First 1
                                if (-not $meta) { continue }
                                $currentRole++
                                $roleData = $meta.RoleData
                                try {
                                    $result = Invoke-SingleRoleActivation -RoleData $roleData -Justification $justification -EffectiveDuration $meta.EffectiveDuration -TicketInfo $ticketInfo -AuthContextToken $graphAuthContextToken
                                    if ($result.Success) {
                                        $successCount++
                                        Write-Verbose "$($roleData.Type) role activated with shared auth-context token fallback"
                                    }
                                    else {
                                        $friendlyError = if ($result.Error -and $result.Error.Exception) { Get-FriendlyErrorMessage -Exception $result.Error.Exception -ErrorDetails $result.ErrorDetails } elseif ($result.ErrorMessage) { $result.ErrorMessage } else { "Activation failed" }
                                        $activationErrors += "$($roleData.DisplayName): $friendlyError"
                                        Write-Warning "Auth-context sequential fallback failed for $($roleData.DisplayName): $friendlyError"
                                    }
                                }
                                catch {
                                    $activationErrors += "$($roleData.DisplayName): $($_.Exception.Message)"
                                }
                            }
                        }
                    }
                }
            }

            if ($authContextAzureItems.Count -gt 0) {
                $armAuthContextToken = Get-AuthenticationContextToken -ContextIds $azureAuthContextIds -Scopes @('https://management.azure.com/.default') -CacheNamespace 'ARM'

                if (-not $armAuthContextToken) {
                    Write-Warning "Failed to obtain a shared ARM authentication context token. Falling back to sequential activation for Azure auth-context roles."

                    foreach ($item in $authContextAzureItems) {
                        $currentRole++
                        $roleData = $item.Tag
                        $contextId = $roleData.PolicyInfo.AuthenticationContextId
                        $progressPercent = [int](($currentRole / $totalRoles) * 100)

                        if ($operationSplash -and -not $operationSplash.IsDisposed) {
                            $scopeInfo = if ($roleData.Scope -and $roleData.Scope -ne "Directory") { " [$($roleData.Scope)]" } else { "" }
                            $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)$scopeInfo... ($currentRole of $totalRoles)", $progressPercent)
                        }

                        $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                        $scopeOverride = & $GetAzureScopeOverride $roleData

                        try {
                            $singleActivationArgs = @{
                                RoleData                 = $roleData
                                Justification            = $justification
                                EffectiveDuration        = $effectiveDuration
                                TicketInfo               = $ticketInfo
                                AuthenticationContextId  = $contextId
                                UseFallbackMethod        = $true
                            }
                            if ($scopeOverride) {
                                $singleActivationArgs.AzureTargetScope = $scopeOverride.TargetScope
                                $singleActivationArgs.LinkedRoleEligibilityScheduleId = $scopeOverride.LinkedRoleEligibilityScheduleId
                            }
                            $result = Invoke-SingleRoleActivation @singleActivationArgs
                            if ($result.Success) {
                                $successCount++
                                Write-Verbose "Azure Resource role activated with authentication context fallback"
                            }
                            else {
                                $friendlyError = if ($result.ErrorMessage) { $result.ErrorMessage } elseif ($result.Error -and $result.Error.Exception) { Get-FriendlyErrorMessage -Exception $result.Error.Exception -ErrorDetails $result.ErrorDetails } else { "Activation failed" }
                                $activationErrors += "$($roleData.DisplayName): $friendlyError"
                                Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
                            }
                        }
                        catch {
                            $activationErrors += "$($roleData.DisplayName): $($_.Exception.Message)"
                            Write-Warning "Failed to activate $($roleData.DisplayName): $($_.Exception.Message)"
                        }
                    }
                }
                else {
                    $azureAuthJobList = [System.Collections.ArrayList]::new()
                    foreach ($item in $authContextAzureItems) {
                        $roleData = $item.Tag
                        $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                        $scopeOverride = & $GetAzureScopeOverride $roleData
                        $azureParamArgs = @{
                            RoleData          = $roleData
                            Justification     = $justification
                            EffectiveDuration = $effectiveDuration
                            TicketInfo        = $ticketInfo
                        }
                        if ($scopeOverride) {
                            $azureParamArgs.AzureTargetScope = $scopeOverride.TargetScope
                            $azureParamArgs.LinkedRoleEligibilityScheduleId = $scopeOverride.LinkedRoleEligibilityScheduleId
                        }
                        $azureParams = Get-RoleActivationParameters @azureParamArgs
                        if ($azureParams.IsReducedScope) {
                            Write-Verbose "Using reduced Azure scope for $($roleData.DisplayName): $($azureParams.OriginalScope) -> $($azureParams.Scope)"
                        }
                        $null = $azureAuthJobList.Add([PSCustomObject]@{
                            RoleData          = $roleData
                            EffectiveDuration = $effectiveDuration
                            Params            = $azureParams
                        })
                    }

                    $azureAuthRoleNames = @($azureAuthJobList | ForEach-Object { $_.RoleData.DisplayName } | Where-Object { $_ })
                    $azureAuthRoleList = $azureAuthRoleNames -join ', '
                    Write-Verbose "Submitting Azure Resource authentication-context batch with $($azureAuthJobList.Count) role(s): $azureAuthRoleList"
                    if ($operationSplash -and -not $operationSplash.IsDisposed) {
                        $statusRoleList = & $GetStatusSafeRoleList $azureAuthRoleNames
                        $operationSplash.UpdateStatus("Submitting auth-context Azure batch: $statusRoleList", 45)
                    }

                    $azureAuthResults = $azureAuthJobList | ForEach-Object -Parallel {
                        $job = $_
                        $params = $job.Params
                        $tok = $using:armAuthContextToken
                        try {
                            $requestName = [System.Guid]::NewGuid().ToString()
                            $roleDefId = if ($params.RoleDefinitionId.StartsWith('/')) {
                                $params.RoleDefinitionId
                            }
                            elseif ($params.OriginalScope -and $params.OriginalScope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                                "/subscriptions/$($matches[1])/providers/Microsoft.Authorization/roleDefinitions/$($params.RoleDefinitionId)"
                            }
                            elseif ($params.Scope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                                "/subscriptions/$($matches[1])/providers/Microsoft.Authorization/roleDefinitions/$($params.RoleDefinitionId)"
                            }
                            else {
                                "/providers/Microsoft.Authorization/roleDefinitions/$($params.RoleDefinitionId)"
                            }
                            $bodyObj = @{
                                properties = @{
                                    roleDefinitionId = $roleDefId
                                    principalId      = $params.PrincipalId
                                    requestType      = $params.RequestType
                                    justification    = $params.Justification
                                    scheduleInfo     = @{
                                        startDateTime = $params.ScheduleInfo.StartDateTime
                                        expiration    = @{
                                            type     = $params.ScheduleInfo.Expiration.Type
                                            duration = $params.ScheduleInfo.Expiration.Duration
                                        }
                                    }
                                }
                            }
                            if ($params.TicketInfo -and $params.TicketInfo.ticketNumber) {
                                $bodyObj.properties.ticketInfo = @{
                                    ticketNumber = $params.TicketInfo.ticketNumber
                                    ticketSystem = $params.TicketInfo.ticketSystem
                                }
                            }
                            if ($params.LinkedRoleEligibilityScheduleId) {
                                $bodyObj.properties.linkedRoleEligibilityScheduleId = $params.LinkedRoleEligibilityScheduleId
                            }
                            $hdrs = @{ 'Authorization' = "Bearer $tok"; 'Content-Type' = 'application/json' }
                            $uri = "https://management.azure.com$($params.Scope)/providers/Microsoft.Authorization/roleAssignmentScheduleRequests/$($requestName)?api-version=2020-10-01"
                            $bodyJson = $bodyObj | ConvertTo-Json -Depth 10 -Compress
                            $resp = Invoke-RestMethod -Uri $uri -Headers $hdrs -Method Put -Body $bodyJson -ErrorAction Stop
                            [PSCustomObject]@{ Success = $true; RoleData = $job.RoleData; Response = $resp; EffectiveDuration = $job.EffectiveDuration; ActivationScope = $params.Scope; OriginalScope = $params.OriginalScope; IsReducedScope = $params.IsReducedScope }
                        }
                        catch {
                            [PSCustomObject]@{ Success = $false; RoleData = $job.RoleData; ErrorMessage = $_.Exception.Message; EffectiveDuration = $job.EffectiveDuration; ActivationScope = $params.Scope; OriginalScope = $params.OriginalScope; IsReducedScope = $params.IsReducedScope }
                        }
                    } -ThrottleLimit 5

                    foreach ($azResult in @($azureAuthResults)) {
                        $currentRole++
                        $roleData = $azResult.RoleData
                        $effectiveDuration = $azResult.EffectiveDuration

                        if ($azResult.Success) {
                            $successCount++
                            Write-Verbose "Azure Resource role activated in auth-context parallel batch: $($roleData.DisplayName)"
                            if ($azResult.IsReducedScope) {
                                Write-Verbose "Azure Resource role used reduced scope: $($azResult.OriginalScope) -> $($azResult.ActivationScope)"
                            }
                            try {
                                if (-not (Get-Variable -Name 'AzureActiveOverrides' -Scope Script -ErrorAction SilentlyContinue)) { $script:AzureActiveOverrides = @{} }
                                $endUtc = (Get-Date).ToUniversalTime().AddHours($effectiveDuration.Hours).AddMinutes($effectiveDuration.Minutes)
                                $roleDefKey = $roleData.RoleDefinitionId
                                if ($roleDefKey -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefKey = $matches[1] }
                                $fullScope = if ($azResult.ActivationScope) { $azResult.ActivationScope } elseif ($roleData.FullScope) { $roleData.FullScope } else { $roleData.DirectoryScopeId }
                                $overrideKey = "$fullScope|$roleDefKey"
                                $script:AzureActiveOverrides[$overrideKey] = [PSCustomObject]@{ EndDateTime = $endUtc }
                                Write-Verbose "Recorded Azure active override for $overrideKey"
                                if (-not (Get-Variable -Name 'DirtyAzureSubscriptions' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyAzureSubscriptions = @() }
                                if (-not (Get-Variable -Name 'DirtyManagementGroups' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyManagementGroups = @() }
                                if ($fullScope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                                    $script:DirtyAzureSubscriptions += $matches[1]
                                    $script:DirtyAzureSubscriptions = @($script:DirtyAzureSubscriptions | Select-Object -Unique)
                                }
                                if ($fullScope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                                    $mgName = $matches[1]
                                    $script:DirtyManagementGroups += $mgName
                                    $script:DirtyManagementGroups = @($script:DirtyManagementGroups | Select-Object -Unique)
                                }
                            }
                            catch { Write-Verbose "Failed to record Azure active override: $($_.Exception.Message)" }
                        }
                        else {
                            $friendlyError = Get-FriendlyErrorMessage -Exception ([Exception]::new($azResult.ErrorMessage)) -ErrorDetails $null
                            $activationErrors += "$($roleData.DisplayName): $friendlyError"
                            Write-Warning "Azure Resource auth-context parallel activation failed for $($roleData.DisplayName): $($azResult.ErrorMessage)"
                        }
                    }
                }
            }
        }

        # ── Batch / parallel submit for non-auth-context roles ───────────────────
        # Split by role type: Graph API roles vs Azure Resource roles
        $graphNoCtxItems = @($noContextRoles | Where-Object { $_.Tag.Type -in @('Entra', 'Group') })
        $azureNoCtxItems = @($noContextRoles | Where-Object { $_.Tag.Type -eq 'AzureResource' })

        Write-Verbose "Non-context roles to activate: $($graphNoCtxItems.Count) Graph, $($azureNoCtxItems.Count) Azure Resource"

        # ── Graph roles: Microsoft Graph $batch API (up to 20 per batch call) ─
        if ($graphNoCtxItems.Count -gt 0) {
            if ($operationSplash -and -not $operationSplash.IsDisposed) {
                $operationSplash.UpdateStatus("Submitting $($graphNoCtxItems.Count) Graph role activation(s) in batch...", 60)
            }

            $batchRequestBodies = [System.Collections.ArrayList]::new()
            $graphBatchMeta     = [System.Collections.ArrayList]::new()

            foreach ($item in $graphNoCtxItems) {
                $roleData          = $item.Tag
                $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                $activationParams  = Get-RoleActivationParameters -RoleData $roleData -Justification $justification -EffectiveDuration $effectiveDuration -TicketInfo $ticketInfo
                $batchId           = "$($batchRequestBodies.Count + 1)"

                $reqUrl = switch ($roleData.Type) {
                    'Entra' { '/roleManagement/directory/roleAssignmentScheduleRequests' }
                    'Group' { '/identityGovernance/privilegedAccess/group/assignmentScheduleRequests' }
                }

                $reqBody = if ($roleData.Type -eq 'Entra') {
                    $b = @{
                        action           = $activationParams.action
                        principalId      = $activationParams.principalId
                        roleDefinitionId = $activationParams.roleDefinitionId
                        directoryScopeId = $activationParams.directoryScopeId
                        justification    = $activationParams.justification
                        scheduleInfo     = $activationParams.scheduleInfo
                    }
                    if ($activationParams.ContainsKey('ticketInfo') -and $activationParams.ticketInfo) { $b.ticketInfo = $activationParams.ticketInfo }
                    $b
                } else {
                    $b = @{
                        action        = $activationParams.action
                        principalId   = $activationParams.principalId
                        groupId       = $activationParams.groupId
                        accessId      = $activationParams.accessId
                        justification = $activationParams.justification
                        scheduleInfo  = $activationParams.scheduleInfo
                    }
                    if ($activationParams.ContainsKey('ticketInfo') -and $activationParams.ticketInfo) { $b.ticketInfo = $activationParams.ticketInfo }
                    $b
                }

                $null = $batchRequestBodies.Add(@{
                    id      = $batchId
                    method  = 'POST'
                    url     = $reqUrl
                    headers = @{ 'Content-Type' = 'application/json' }
                    body    = $reqBody
                })
                $null = $graphBatchMeta.Add([PSCustomObject]@{
                    BatchId           = $batchId
                    Item              = $item
                    RoleData          = $roleData
                    EffectiveDuration = $effectiveDuration
                })
            }

            # Submit in chunks of 20 (Graph batch limit)
            $BATCH_SIZE = 20
            for ($bi = 0; $bi -lt $batchRequestBodies.Count; $bi += $BATCH_SIZE) {
                $chunkEnd   = [Math]::Min($bi + $BATCH_SIZE - 1, $batchRequestBodies.Count - 1)
                $chunk      = @($batchRequestBodies[$bi..$chunkEnd])
                $batchBody  = @{ requests = $chunk }
                $batchNumber = [int]($bi / $BATCH_SIZE) + 1
                $chunkRoleNames = @(
                    foreach ($reqItem in $chunk) {
                        $meta = $graphBatchMeta | Where-Object { $_.BatchId -eq $reqItem.id } | Select-Object -First 1
                        if ($meta -and $meta.RoleData) { $meta.RoleData.DisplayName }
                    }
                )
                $chunkRoleList = @($chunkRoleNames | Where-Object { $_ }) -join ', '
                $statusRoleList = & $GetStatusSafeRoleList $chunkRoleNames

                try {
                    Write-Verbose "Submitting Graph activation batch $batchNumber with $($chunk.Count) role(s): $chunkRoleList"
                    if ($operationSplash -and -not $operationSplash.IsDisposed) {
                        $operationSplash.UpdateStatus("Submitting Graph batch $batchNumber`: $statusRoleList", 62)
                    }
                    $batchResponse = Invoke-MgGraphRequest -Method POST -Uri '$batch' -Body $batchBody -ErrorAction Stop

                    foreach ($resp in @($batchResponse.responses)) {
                        $meta     = $graphBatchMeta | Where-Object { $_.BatchId -eq $resp.id } | Select-Object -First 1
                        if (-not $meta) { continue }
                        $currentRole++
                        $roleData = $meta.RoleData

                        if ($resp.status -ge 200 -and $resp.status -lt 300) {
                            Write-Verbose "$($roleData.Type) role activated via Graph batch - id: $($resp.body.id)"
                            $successCount++
                        }
                        else {
                            $errMsg = if ($resp.body -and $resp.body.error -and $resp.body.error.message) {
                                $resp.body.error.message
                            } else { "HTTP $($resp.status)" }
                            $activationErrors += "$($roleData.DisplayName): $errMsg"
                            Write-Warning "Graph batch activation failed for $($roleData.DisplayName): $errMsg"
                        }
                    }
                }
                catch {
                    Write-Warning "Graph batch request failed, falling back to sequential for this chunk: $($_.Exception.Message)"
                    Write-Verbose "Graph fallback batch $batchNumber roles: $chunkRoleList"
                    # Fallback: activate each in this chunk individually
                    foreach ($reqItem in $chunk) {
                        $meta = $graphBatchMeta | Where-Object { $_.BatchId -eq $reqItem.id } | Select-Object -First 1
                        if (-not $meta) { continue }
                        $currentRole++
                        $roleData = $meta.RoleData
                        try {
                            $result = Invoke-SingleRoleActivation -RoleData $roleData -Justification $justification -EffectiveDuration $meta.EffectiveDuration -TicketInfo $ticketInfo
                            if ($result.Success) {
                                $successCount++
                                $responseId = & $GetActivationResponseId $result
                                if ($responseId) {
                                    Write-Verbose "$($roleData.Type) role activated (sequential fallback) - Response ID: $responseId"
                                }
                                else {
                                    Write-Verbose "$($roleData.Type) role activated (sequential fallback)"
                                }
                            }
                            else {
                                $friendlyError = Get-FriendlyErrorMessage -Exception $result.Error.Exception -ErrorDetails $result.ErrorDetails
                                $activationErrors += "$($roleData.DisplayName): $friendlyError"
                                Write-Warning "Sequential fallback failed for $($roleData.DisplayName): $friendlyError"
                            }
                        }
                        catch {
                            $activationErrors += "$($roleData.DisplayName): $($_.Exception.Message)"
                        }
                    }
                }
            }
        }

        # ── Azure Resource roles: submit in parallel via ARM REST PUT ─────────
        if ($azureNoCtxItems.Count -gt 0) {
            if ($operationSplash -and -not $operationSplash.IsDisposed) {
                $operationSplash.UpdateStatus("Submitting $($azureNoCtxItems.Count) Azure Resource role activation(s) in parallel...", 75)
            }

            # Acquire ARM token once for all parallel requests
            $armTokenForParallel = $null
            try {
                $azCtxParallel = Get-AzContext -ErrorAction SilentlyContinue
                if ($azCtxParallel) {
                    $tokenObj = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop
                    $armTokenForParallel = if ($tokenObj.Token -is [System.Security.SecureString]) {
                        [System.Net.NetworkCredential]::new('', $tokenObj.Token).Password
                    } else { $tokenObj.Token }
                }
            }
            catch { Write-Verbose "Could not acquire ARM token for parallel activation: $($_.Exception.Message)" }

            # Pre-compute activation parameters for all Azure roles
            $azureJobList = [System.Collections.ArrayList]::new()
            foreach ($item in $azureNoCtxItems) {
                $roleData          = $item.Tag
                $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                $scopeOverride     = & $GetAzureScopeOverride $roleData
                $azureParamArgs    = @{
                    RoleData          = $roleData
                    Justification     = $justification
                    EffectiveDuration = $effectiveDuration
                    TicketInfo        = $ticketInfo
                }
                if ($scopeOverride) {
                    $azureParamArgs.AzureTargetScope = $scopeOverride.TargetScope
                    $azureParamArgs.LinkedRoleEligibilityScheduleId = $scopeOverride.LinkedRoleEligibilityScheduleId
                }
                $azureParams = Get-RoleActivationParameters @azureParamArgs
                if ($azureParams.IsReducedScope) {
                    Write-Verbose "Using reduced Azure scope for $($roleData.DisplayName): $($azureParams.OriginalScope) -> $($azureParams.Scope)"
                }
                $null = $azureJobList.Add([PSCustomObject]@{
                    RoleData          = $roleData
                    EffectiveDuration = $effectiveDuration
                    Params            = $azureParams
                })
            }

            $azureRoleNames = @($azureJobList | ForEach-Object { $_.RoleData.DisplayName } | Where-Object { $_ })
            $azureRoleList = $azureRoleNames -join ', '
            Write-Verbose "Submitting Azure Resource activation batch with $($azureJobList.Count) role(s): $azureRoleList"
            if ($operationSplash -and -not $operationSplash.IsDisposed) {
                $statusRoleList = & $GetStatusSafeRoleList $azureRoleNames
                $operationSplash.UpdateStatus("Submitting Azure batch: $statusRoleList", 78)
            }

            # Submit in parallel using ARM REST PUT (avoids Az module context issues in runspaces)
            $azureResults = $azureJobList | ForEach-Object -Parallel {
                $job    = $_
                $params = $job.Params
                $tok    = $using:armTokenForParallel
                try {
                    $requestName = [System.Guid]::NewGuid().ToString()
                    $roleDefId   = if ($params.RoleDefinitionId.StartsWith('/')) {
                        $params.RoleDefinitionId
                    } elseif ($params.OriginalScope -and $params.OriginalScope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                        "/subscriptions/$($matches[1])/providers/Microsoft.Authorization/roleDefinitions/$($params.RoleDefinitionId)"
                    } elseif ($params.Scope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                        "/subscriptions/$($matches[1])/providers/Microsoft.Authorization/roleDefinitions/$($params.RoleDefinitionId)"
                    } else {
                        "/providers/Microsoft.Authorization/roleDefinitions/$($params.RoleDefinitionId)"
                    }
                    $bodyObj = @{
                        properties = @{
                            roleDefinitionId = $roleDefId
                            principalId      = $params.PrincipalId
                            requestType      = $params.RequestType
                            justification    = $params.Justification
                            scheduleInfo     = @{
                                startDateTime = $params.ScheduleInfo.StartDateTime
                                expiration    = @{
                                    type     = $params.ScheduleInfo.Expiration.Type
                                    duration = $params.ScheduleInfo.Expiration.Duration
                                }
                            }
                        }
                    }
                    if ($params.TicketInfo -and $params.TicketInfo.ticketNumber) {
                        $bodyObj.properties.ticketInfo = @{
                            ticketNumber = $params.TicketInfo.ticketNumber
                            ticketSystem = $params.TicketInfo.ticketSystem
                        }
                    }
                    if ($params.LinkedRoleEligibilityScheduleId) {
                        $bodyObj.properties.linkedRoleEligibilityScheduleId = $params.LinkedRoleEligibilityScheduleId
                    }
                    $hdrs    = @{ 'Authorization' = "Bearer $tok"; 'Content-Type' = 'application/json' }
                    $uri     = "https://management.azure.com$($params.Scope)/providers/Microsoft.Authorization/roleAssignmentScheduleRequests/$($requestName)?api-version=2020-10-01"
                    $bodyJson = $bodyObj | ConvertTo-Json -Depth 10 -Compress
                    $resp    = Invoke-RestMethod -Uri $uri -Headers $hdrs -Method Put -Body $bodyJson -ErrorAction Stop
                    [PSCustomObject]@{ Success = $true; RoleData = $job.RoleData; Response = $resp; EffectiveDuration = $job.EffectiveDuration; ActivationScope = $params.Scope; OriginalScope = $params.OriginalScope; IsReducedScope = $params.IsReducedScope }
                }
                catch {
                    [PSCustomObject]@{ Success = $false; RoleData = $job.RoleData; ErrorMessage = $_.Exception.Message; EffectiveDuration = $job.EffectiveDuration; ActivationScope = $params.Scope; OriginalScope = $params.OriginalScope; IsReducedScope = $params.IsReducedScope }
                }
            } -ThrottleLimit 5

            foreach ($azResult in @($azureResults)) {
                $currentRole++
                $roleData          = $azResult.RoleData
                $effectiveDuration = $azResult.EffectiveDuration

                if ($azResult.Success) {
                    $successCount++
                    Write-Verbose "Azure Resource role activated in parallel: $($roleData.DisplayName)"
                    if ($azResult.IsReducedScope) {
                        Write-Verbose "Azure Resource role used reduced scope: $($azResult.OriginalScope) -> $($azResult.ActivationScope)"
                    }
                    try {
                        if (-not (Get-Variable -Name 'AzureActiveOverrides' -Scope Script -ErrorAction SilentlyContinue)) { $script:AzureActiveOverrides = @{} }
                        $endUtc      = (Get-Date).ToUniversalTime().AddHours($effectiveDuration.Hours).AddMinutes($effectiveDuration.Minutes)
                        $roleDefKey  = $roleData.RoleDefinitionId
                        if ($roleDefKey -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefKey = $matches[1] }
                        $fullScope   = if ($azResult.ActivationScope) { $azResult.ActivationScope } elseif ($roleData.FullScope) { $roleData.FullScope } else { $roleData.DirectoryScopeId }
                        $overrideKey = "$fullScope|$roleDefKey"
                        $script:AzureActiveOverrides[$overrideKey] = [PSCustomObject]@{ EndDateTime = $endUtc }
                        Write-Verbose "Recorded Azure active override for $overrideKey"
                        if (-not (Get-Variable -Name 'DirtyAzureSubscriptions' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyAzureSubscriptions = @() }
                        if (-not (Get-Variable -Name 'DirtyManagementGroups'   -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyManagementGroups   = @() }
                        if ($fullScope -match '^/subscriptions/([a-fA-F0-9\-]{36})') {
                            $script:DirtyAzureSubscriptions += $matches[1]
                            $script:DirtyAzureSubscriptions  = @($script:DirtyAzureSubscriptions | Select-Object -Unique)
                        }
                        if ($fullScope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                            $mgName = $matches[1]
                            $script:DirtyManagementGroups += $mgName
                            $script:DirtyManagementGroups  = @($script:DirtyManagementGroups | Select-Object -Unique)
                        }
                    }
                    catch { Write-Verbose "Failed to record Azure active override: $($_.Exception.Message)" }
                }
                else {
                    $friendlyError = Get-FriendlyErrorMessage -Exception ([Exception]::new($azResult.ErrorMessage)) -ErrorDetails $null
                    $activationErrors += "$($roleData.DisplayName): $friendlyError"
                    Write-Warning "Azure Resource parallel activation failed for $($roleData.DisplayName): $($azResult.ErrorMessage)"
                }
            }
        }
        
        if ($operationSplash -and -not $operationSplash.IsDisposed) {
            $operationSplash.UpdateStatus("Completing activation process...", 95)
        }
        
        # Clean up authentication context state
        if ($script:JustCompletedAuthContext) {
            $script:JustCompletedAuthContext = $false
            $script:AuthContextCompletionTime = $null
        }
        
        # Display activation results
        Show-ActivationResults -SuccessCount $successCount -TotalCount $CheckedItems.Count -Errors $activationErrors
        
        # Refresh role lists to reflect changes
        if ($operationSplash -and -not $operationSplash.IsDisposed) {
            $operationSplash.UpdateStatus("Refreshing role lists...", 98)
        }
        
        # Immediate refresh only once; Graph/Azure reflect changes near-instantly for Azure RBAC
        if ($successCount -gt 0) {
            # Clear role cache so ActiveOnly refresh pulls fresh data while preserving Azure cache for delta
            Write-Verbose "Clearing role cache to force fresh data retrieval after activation (single refresh, no pagination wait)"
            $script:CachedEligibleRoles = $null
            $script:CachedActiveRoles = $null
            $script:LastRoleFetchTime = $null
        }

        Write-Verbose "Refreshing role data (single attempt)"
        try {
            # If any activated role requires approval, the role stays in eligible state with
            # "Pending Approval = Yes".  The only way to show that is to refresh the eligible
            # panel (it calls Get-PIMPendingRequests inside).  Otherwise a quick active-only
            # refresh is sufficient.
            $anyNeedsApproval = ($CheckedItems | Where-Object {
                $_.Tag.PSObject.Properties['PolicyInfo'] -and
                $_.Tag.PolicyInfo -and
                $_.Tag.PolicyInfo.PSObject.Properties['RequiresApproval'] -and
                $_.Tag.PolicyInfo.RequiresApproval
            }).Count -gt 0

            if ($anyNeedsApproval) {
                Write-Verbose "At least one role requires approval - refreshing both active and eligible panels"
                # Force re-fetch so pending requests are picked up
                $script:CachedEligibleRoles = $null
                Update-PIMRolesList -Form $Form -RefreshActive -RefreshEligible
            } else {
                Update-PIMRolesList -Form $Form -RefreshActive
            }
            Write-Verbose "Role lists refreshed successfully"
        }
        catch {
            Write-Warning "Failed to refresh role lists: $_"
        }
        
    }
    finally {
        # Ensure splash is closed
        if ($operationSplash -and -not $operationSplash.IsDisposed) {
            $operationSplash.Close()
        }
    }
    
    Write-Verbose "Activation process completed - Success: $successCount, Errors: $($activationErrors.Count)"
}