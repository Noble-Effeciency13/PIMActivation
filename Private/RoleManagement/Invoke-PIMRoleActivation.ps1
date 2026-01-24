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
        
        # Show activation dialog for required or optional information
        if ($policyRequirements.RequiresJustification -or $policyRequirements.RequiresTicket -or $CheckedItems.Count -gt 0) {
            Write-Verbose "Showing activation dialog for justification/ticket requirements"
            $result = Show-PIMActivationDialog -RequiresJustification:$policyRequirements.RequiresJustification `
                -RequiresTicket:$policyRequirements.RequiresTicket `
                -OptionalJustification:$(-not $policyRequirements.RequiresJustification)
            
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
        
        # Group roles by authentication context to minimize authentication prompts
        $rolesByContext = @{}
        $noContextRoles = @()
        
        foreach ($item in $CheckedItems) {
            $roleData = $item.Tag
            
            if ($roleData.PolicyInfo -and $roleData.PolicyInfo.RequiresAuthenticationContext -and $roleData.PolicyInfo.AuthenticationContextId) {
                $contextId = $roleData.PolicyInfo.AuthenticationContextId
                
                if (-not $rolesByContext.ContainsKey($contextId)) {
                    $rolesByContext[$contextId] = @()
                }
                $rolesByContext[$contextId] += $item
            }
            else {
                $noContextRoles += $item
            }
        }
        
        Write-Verbose "Roles grouped by authentication context: $($rolesByContext.Keys.Count) contexts, $($noContextRoles.Count) without context"

        # NOW show the splash form after all user input has been collected
        $operationSplash = Show-OperationSplash -Title "Role Activation" -InitialMessage "Processing role activations..." -ShowProgressBar $true
        $activationErrors = @()
        $successCount = 0
        $totalRoles = $CheckedItems.Count
        $currentRole = 0
        
        # Process roles that require authentication context first, grouped by context
        foreach ($contextId in $rolesByContext.Keys) {
            Write-Verbose "Processing roles for authentication context: $contextId"
            
            # Try to get authentication context token once per context (reuse for multiple roles)
            $authContextToken = Get-AuthenticationContextToken -ContextId $contextId
            
            if (-not $authContextToken) {
                Write-Warning "Failed to obtain authentication context token for context: $contextId. Falling back to individual token acquisition per role."
                
                # Fallback: Process each role individually using the original method
                foreach ($item in $rolesByContext[$contextId]) {
                    $currentRole++
                    $roleData = $item.Tag
                    $progressPercent = [int](($currentRole / $totalRoles) * 100)
                    
                    if ($operationSplash -and -not $operationSplash.IsDisposed) {
                        $scopeInfo = if ($roleData.Scope -and $roleData.Scope -ne "Directory") { " [$($roleData.Scope)]" } else { "" }
                        $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)$scopeInfo ... ($currentRole of $totalRoles)", $progressPercent)
                    }
                    
                    Write-Verbose "Processing: $($roleData.DisplayName) [$($roleData.Type)] with individual auth context token acquisition"
                    
                    # Calculate actual duration based on policy
                    $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                    Write-Verbose "Activation duration: $($effectiveDuration.Hours) hours, $($effectiveDuration.Minutes) minutes"
                    
                    # Use consolidated activation function with fallback method
                    try {
                        $result = Invoke-SingleRoleActivation -RoleData $roleData -Justification $justification -EffectiveDuration $effectiveDuration -TicketInfo $ticketInfo -AuthenticationContextId $contextId -UseFallbackMethod
                        
                        # Handle the result
                        if ($result.Success) {
                            if ($result.IsAzureResource) {
                                # Azure Resource roles handle their own success counting
                                $successCount++
                            }
                            else {
                                Write-Verbose "Role activated via fallback method - Response ID: $($result.Response.id)"
                                $successCount++
                            }
                        }
                        else {
                            if ($result.IsAzureResource) {
                                # Azure Resource errors are already handled in the function
                                $activationErrors += "$($roleData.DisplayName): $($result.ErrorMessage)"
                            }
                            else {
                                $friendlyError = Get-FriendlyErrorMessage -Exception $result.Error.Exception -ErrorDetails $result.ErrorDetails
                                $activationErrors += "$($roleData.DisplayName): $friendlyError"
                                Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
                            }
                        }
                    }
                    catch {
                        $activationErrors += "$($roleData.DisplayName): $($_.Exception.Message)"
                        Write-Warning "Failed to activate $($roleData.DisplayName): $($_.Exception.Message)"
                    }
                }
                continue
            }
            
            Write-Verbose "Successfully obtained authentication context token for context: $contextId"
            
            # Process each role requiring this authentication context using the cached token
            foreach ($item in $rolesByContext[$contextId]) {
                $currentRole++
                $roleData = $item.Tag
                $progressPercent = [int](($currentRole / $totalRoles) * 100)
                
                if ($operationSplash -and -not $operationSplash.IsDisposed) {
                    $scopeInfo = if ($roleData.Scope -and $roleData.Scope -ne "Directory") { " [$($roleData.Scope)]" } else { "" }
                    $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)$scopeInfo... ($currentRole of $totalRoles)", $progressPercent)
                }
                
                Write-Verbose "Processing: $($roleData.DisplayName) [$($roleData.Type)] with cached auth context token"
                
                # Calculate actual duration based on policy
                $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                Write-Verbose "Activation duration: $($effectiveDuration.Hours) hours, $($effectiveDuration.Minutes) minutes"
                
                # Use consolidated activation function with cached authentication context token
                $result = Invoke-SingleRoleActivation -RoleData $roleData -Justification $justification -EffectiveDuration $effectiveDuration -TicketInfo $ticketInfo -AuthContextToken $authContextToken
                
                # Handle the result
                if ($result.Success) {
                    if ($result.IsAzureResource) {
                        # Azure Resource activation already incremented success count
                        Write-Verbose "Azure Resource role activated with authentication context successfully"
                    }
                    else {
                        Write-Verbose "$($roleData.Type) role activated with authentication context - Response ID: $($result.Response.id)"
                        $successCount++
                    }
                }
                else {
                    if ($result.IsAzureResource) {
                        # Azure Resource errors are already added to activationErrors
                        $activationErrors += "$($roleData.DisplayName): $($result.ErrorMessage)"
                    }
                    else {
                        # Log detailed error information
                        Write-Verbose "Activation failed. Error details:"
                        Write-Verbose "Exception: $($result.Error.Exception.Message)"
                        Write-Verbose "Error Details: $($result.ErrorDetails)"
                        
                        $friendlyError = Get-FriendlyErrorMessage -Exception $result.Error.Exception -ErrorDetails $result.ErrorDetails
                        $activationErrors += "$($roleData.DisplayName): $friendlyError"
                        Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
                    }
                }
            }
        }
        
        # Process roles without authentication context
        foreach ($item in $noContextRoles) {
            $currentRole++
            $roleData = $item.Tag
            $progressPercent = [int](($currentRole / $totalRoles) * 100)
            
            if ($operationSplash -and -not $operationSplash.IsDisposed) {
                $scopeInfo = if ($roleData.Scope -and $roleData.Scope -ne "Directory") { " [$($roleData.Scope)]" } else { "" }
                $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)$scopeInfo... ($currentRole of $totalRoles)", $progressPercent)
            }
            
            Write-Verbose "Processing: $($roleData.DisplayName) [$($roleData.Type)]"
            
            # Calculate actual duration based on policy
            $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
            Write-Verbose "Activation duration: $($effectiveDuration.Hours) hours, $($effectiveDuration.Minutes) minutes"
            
            # Use consolidated activation function
            $result = Invoke-SingleRoleActivation -RoleData $roleData -Justification $justification -EffectiveDuration $effectiveDuration -TicketInfo $ticketInfo
            
            # Handle the result
            if ($result.Success) {
                if ($result.IsAzureResource) {
                    Write-Verbose "Azure Resource role activated successfully"
                    $successCount++
                    # Record expected expiration for display using requested effective duration
                    try {
                        if (-not (Get-Variable -Name 'AzureActiveOverrides' -Scope Script -ErrorAction SilentlyContinue)) { $script:AzureActiveOverrides = @{} }
                        $endUtc = (Get-Date).ToUniversalTime().AddHours($effectiveDuration.Hours).AddMinutes($effectiveDuration.Minutes)
                        # Normalize RoleDefinitionId to GUID-only to ensure key matches across sources
                        $roleDefKey = $roleData.RoleDefinitionId
                        if ($roleDefKey -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") {
                            $roleDefKey = $matches[1]
                        }
                        # Ensure FullScope exists for keying
                        $fullScope = $roleData.FullScope
                        if (-not $fullScope) { $fullScope = $roleData.DirectoryScopeId }
                        $overrideKey = "$( $fullScope )|$( $roleDefKey )"
                        $script:AzureActiveOverrides[$overrideKey] = [PSCustomObject]@{ EndDateTime = $endUtc }
                        Write-Verbose "Recorded Azure active override for $overrideKey with expiration $endUtc"
                        # Mark affected scope as dirty for delta refresh
                        if (-not (Get-Variable -Name 'DirtyAzureSubscriptions' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyAzureSubscriptions = @() }
                        if (-not (Get-Variable -Name 'DirtyManagementGroups' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyManagementGroups = @() }
                        if ($roleData.PSObject.Properties['SubscriptionId'] -and $roleData.SubscriptionId) {
                            $script:DirtyAzureSubscriptions += $roleData.SubscriptionId
                            $script:DirtyAzureSubscriptions = @($script:DirtyAzureSubscriptions | Select-Object -Unique)
                            Write-Verbose "Marked subscription $($roleData.SubscriptionId) as dirty for delta refresh"
                        }
                        # If scope is a management group, mark MG as dirty
                        if ($roleData.PSObject.Properties['FullScope'] -and $roleData.FullScope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                            $mgName = $matches[1]
                            $script:DirtyManagementGroups += $mgName
                            $script:DirtyManagementGroups = @($script:DirtyManagementGroups | Select-Object -Unique)
                            Write-Verbose "Marked management group ${mgName} as dirty for delta refresh"
                        }
                    }
                    catch { Write-Verbose "Failed to record Azure active override: $($_.Exception.Message)" }
                }
                else {
                    Write-Verbose "$($roleData.Type) role activated via Microsoft Graph SDK - Response ID: $($result.Response.id)"
                    $successCount++
                }
            }
            else {
                if ($result.IsAzureResource) {
                    # Azure Resource errors
                    $activationErrors += "$($roleData.DisplayName): $($result.ErrorMessage)"
                    Write-Warning "Failed to activate Azure Resource role $($roleData.DisplayName): $($result.ErrorMessage)"
                }
                else {
                    # Log detailed error information
                    Write-Verbose "Microsoft Graph SDK call failed. Error details:"
                    Write-Verbose "Exception: $($result.Error.Exception.Message)"
                    Write-Verbose "Error Details: $($result.ErrorDetails)"
                    
                    $friendlyError = Get-FriendlyErrorMessage -Exception $result.Error.Exception -ErrorDetails $result.ErrorDetails
                    $activationErrors += "$($roleData.DisplayName): $friendlyError"
                    Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
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
            # Per refresh semantics: only refresh ACTIVE roles, do not re-fetch eligible
            Update-PIMRolesList -Form $Form -RefreshActive
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