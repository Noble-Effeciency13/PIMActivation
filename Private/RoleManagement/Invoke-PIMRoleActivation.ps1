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
            RequiresTicket = $false
            RequiresMfa = $false
            RequiresAuthContext = $false
            AuthContextIds = @()
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
        $operationSplash = Show-OperationSplash -Title "Role Activation" -InitialMessage "Processing role activations..." -ShowProgressBar $true        # Process individual role activations
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
                        $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)... ($currentRole of $totalRoles)", $progressPercent)
                    }
                    
                    Write-Verbose "Processing: $($roleData.DisplayName) [$($roleData.Type)] with individual auth context token acquisition"
                    
                    # Calculate actual duration based on policy
                    $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                    Write-Verbose "Activation duration: $($effectiveDuration.Hours) hours, $($effectiveDuration.Minutes) minutes"
                    
                    # Build activation parameters and use original method
                    try {
                        switch ($roleData.Type) {
                            'Entra' {
                                # Build activation parameters
                                $activationParams = @{
                                    action = "selfActivate"
                                    justification = $justification
                                    principalId = $script:CurrentUser.Id
                                    roleDefinitionId = $roleData.RoleDefinitionId
                                    directoryScopeId = if ($roleData.DirectoryScopeId) { $roleData.DirectoryScopeId } else { "/" }
                                    scheduleInfo = @{
                                        startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                                        expiration = @{
                                            duration = "PT$($effectiveDuration.Hours)H$($effectiveDuration.Minutes)M"
                                            type = "afterDuration"
                                        }
                                    }
                                }
                                
                                if ($ticketInfo -and $ticketInfo.ticketNumber) {
                                    $activationParams.ticketInfo = $ticketInfo
                                }
                                
                                Write-Verbose "Falling back to original authentication context method for Entra role"
                                $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Entra' -AuthenticationContextId $contextId
                            }
                            'Group' {
                                # Build activation parameters  
                                $activationParams = @{
                                    accessId = "member"
                                    principalId = $script:CurrentUser.Id
                                    groupId = $roleData.GroupId
                                    action = "selfActivate"
                                    justification = $justification
                                    scheduleInfo = @{
                                        startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                                        expiration = @{
                                            duration = "PT$($effectiveDuration.Hours)H$($effectiveDuration.Minutes)M"
                                            type = "afterDuration"
                                        }
                                    }
                                }
                                
                                if ($ticketInfo -and $ticketInfo.ticketNumber) {
                                    $activationParams.ticketInfo = $ticketInfo
                                }
                                
                                Write-Verbose "Falling back to original authentication context method for Group role"
                                $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Group' -AuthenticationContextId $contextId
                            }
                        }
                        
                        if ($mgResult.Success) {
                            Write-Verbose "Role activated via fallback method - Response ID: $($mgResult.Response.id)"
                            $successCount++
                        }
                        else {
                            $friendlyError = Get-FriendlyErrorMessage -Exception $mgResult.Error.Exception -ErrorDetails $mgResult.ErrorDetails
                            $activationErrors += "$($roleData.DisplayName): $friendlyError"
                            Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
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
                    $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)... ($currentRole of $totalRoles)", $progressPercent)
                }
                
                Write-Verbose "Processing: $($roleData.DisplayName) [$($roleData.Type)] with cached auth context token"
                
                # Calculate actual duration based on policy
                $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
                Write-Verbose "Activation duration: $($effectiveDuration.Hours) hours, $($effectiveDuration.Minutes) minutes"
                
                # Build activation parameters based on role type
                switch ($roleData.Type) {
                    'Entra' {
                        # Check eligibility
                        $eligibilityCheck = Test-PIMRoleEligibility -UserId $script:CurrentUser.Id -RoleDefinitionId $roleData.RoleDefinitionId
                        if (-not $eligibilityCheck.IsEligible) {
                            throw "User is not eligible for this role assignment"
                        }
                        Write-Verbose "Eligibility check completed. IsEligible: $($eligibilityCheck.IsEligible)"
                        
                        # Build activation parameters
                        $activationParams = @{
                            action = "selfActivate"
                            justification = $justification
                            principalId = $script:CurrentUser.Id
                            roleDefinitionId = $roleData.RoleDefinitionId
                            directoryScopeId = if ($roleData.DirectoryScopeId) { $roleData.DirectoryScopeId } else { "/" }
                            scheduleInfo = @{
                                startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                                expiration = @{
                                    duration = "PT$($effectiveDuration.Hours)H$($effectiveDuration.Minutes)M"
                                    type = "afterDuration"
                                }
                            }
                        }
                        
                        # Only add ticketInfo if it exists and has content
                        if ($ticketInfo -and $ticketInfo.ticketNumber) {
                            $activationParams.ticketInfo = $ticketInfo
                        }
                        
                        Write-Verbose "Using role-specific scope: $($activationParams.directoryScopeId)"
                        
                        # Use direct activation with cached authentication context token
                        Write-Verbose "Using cached authentication context token for immediate activation"
                        
                        $mgResult = Invoke-PIMActivationWithAuthContextToken -ActivationParams $activationParams -RoleType 'Entra' -AuthContextToken $authContextToken
                        
                        if ($mgResult.Success) {
                            Write-Verbose "Entra role activated with authentication context - Response ID: $($mgResult.Response.id)"
                            $successCount++
                        }
                        else {
                            # Log detailed error information
                            Write-Verbose "Activation failed. Error details:"
                            Write-Verbose "Exception: $($mgResult.Error.Exception.Message)"
                            Write-Verbose "Error Details: $($mgResult.ErrorDetails)"
                            
                            # Extract friendly error message
                            $friendlyError = Get-FriendlyErrorMessage -Exception $mgResult.Error.Exception -ErrorDetails $mgResult.ErrorDetails
                            $activationErrors += "$($roleData.DisplayName): $friendlyError"
                            
                            Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
                        }
                    }
                    
                    'Group' {
                        $activationParams = @{
                            accessId = "member"
                            principalId = $script:CurrentUser.Id
                            groupId = $roleData.GroupId
                            action = "selfActivate"
                            justification = $justification
                            scheduleInfo = @{
                                startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                                expiration = @{
                                    duration = "PT$($effectiveDuration.Hours)H$($effectiveDuration.Minutes)M"
                                    type = "afterDuration"
                                }
                            }
                        }
                        
                        # Only add ticketInfo if it exists and has content
                        if ($ticketInfo -and $ticketInfo.ticketNumber) {
                            $activationParams.ticketInfo = $ticketInfo
                        }
                        
                        # Use direct activation with cached authentication context token
                        Write-Verbose "Using cached authentication context token for immediate activation"
                        
                        $mgResult = Invoke-PIMActivationWithAuthContextToken -ActivationParams $activationParams -RoleType 'Group' -AuthContextToken $authContextToken
                        
                        if ($mgResult.Success) {
                            Write-Verbose "Group membership activated with authentication context - Response ID: $($mgResult.Response.id)"
                            $successCount++
                        }
                        else {
                            # Log detailed error information
                            Write-Verbose "Activation failed for group. Error details:"
                            Write-Verbose "Exception: $($mgResult.Error.Exception.Message)"
                            Write-Verbose "Error Details: $($mgResult.ErrorDetails)"
                            
                            $friendlyError = Get-FriendlyErrorMessage -Exception $mgResult.Error.Exception -ErrorDetails $mgResult.ErrorDetails
                            $activationErrors += "$($roleData.DisplayName): $friendlyError"
                            
                            Write-Warning "Failed to activate group $($roleData.DisplayName): $friendlyError"
                        }
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
                $operationSplash.UpdateStatus("Activating $($roleData.DisplayName)... ($currentRole of $totalRoles)", $progressPercent)
            }
            
            Write-Verbose "Processing: $($roleData.DisplayName) [$($roleData.Type)]"
            
            # Calculate actual duration based on policy
            $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $roleData.PolicyInfo.MaxDuration
            Write-Verbose "Activation duration: $($effectiveDuration.Hours) hours, $($effectiveDuration.Minutes) minutes"
            
            # Build activation parameters based on role type
            switch ($roleData.Type) {
                'Entra' {
                    # Check eligibility
                    $eligibilityCheck = Test-PIMRoleEligibility -UserId $script:CurrentUser.Id -RoleDefinitionId $roleData.RoleDefinitionId
                    if (-not $eligibilityCheck.IsEligible) {
                        throw "User is not eligible for this role assignment"
                    }
                    Write-Verbose "Eligibility check completed. IsEligible: $($eligibilityCheck.IsEligible)"
                    
                    # Build activation parameters
                    $activationParams = @{
                        action = "selfActivate"
                        justification = $justification
                        principalId = $script:CurrentUser.Id
                        roleDefinitionId = $roleData.RoleDefinitionId
                        directoryScopeId = if ($roleData.DirectoryScopeId) { $roleData.DirectoryScopeId } else { "/" }
                        scheduleInfo = @{
                            startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                            expiration = @{
                                duration = "PT$($effectiveDuration.Hours)H$($effectiveDuration.Minutes)M"
                                type = "afterDuration"
                            }
                        }
                    }
                    
                    # Only add ticketInfo if it exists and has content
                    if ($ticketInfo -and $ticketInfo.ticketNumber) {
                        $activationParams.ticketInfo = $ticketInfo
                    }
                    
                    Write-Verbose "Using role-specific scope: $($activationParams.directoryScopeId)"
                    
                    # Use Microsoft Graph SDK with standard token (no auth context required)
                    Write-Verbose "Using Microsoft Graph SDK for Entra role without authentication context requirement"
                    
                    $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Entra'
                    
                    if ($mgResult.Success) {
                        Write-Verbose "Entra role activated via Microsoft Graph SDK - Response ID: $($mgResult.Response.id)"
                        $successCount++
                    }
                    else {
                        # Log detailed error information
                        Write-Verbose "Microsoft Graph SDK call failed. Error details:"
                        Write-Verbose "Exception: $($mgResult.Error.Exception.Message)"
                        Write-Verbose "Error Details: $($mgResult.ErrorDetails)"
                        
                        $friendlyError = Get-FriendlyErrorMessage -Exception $mgResult.Error.Exception -ErrorDetails $mgResult.ErrorDetails
                        $activationErrors += "$($roleData.DisplayName): $friendlyError"
                        
                        Write-Warning "Failed to activate $($roleData.DisplayName): $friendlyError"
                    }
                }
                
                'Group' {
                    $activationParams = @{
                        accessId = "member"
                        principalId = $script:CurrentUser.Id
                        groupId = $roleData.GroupId
                        action = "selfActivate"
                        justification = $justification
                        scheduleInfo = @{
                            startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                            expiration = @{
                                duration = "PT$($effectiveDuration.Hours)H$($effectiveDuration.Minutes)M"
                                type = "afterDuration"
                            }
                        }
                    }
                    
                    # Only add ticketInfo if it exists and has content
                    if ($ticketInfo -and $ticketInfo.ticketNumber) {
                        $activationParams.ticketInfo = $ticketInfo
                    }
                    
                    # Use Microsoft Graph SDK with standard token (no auth context required)
                    Write-Verbose "Using Microsoft Graph SDK for Group role without authentication context requirement"
                    
                    $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Group'
                    
                    if ($mgResult.Success) {
                        Write-Verbose "Group membership activated via Microsoft Graph SDK - Response ID: $($mgResult.Response.id)"
                        $successCount++
                    }
                    else {
                        # Log detailed error information
                        Write-Verbose "Microsoft Graph SDK call failed for group. Error details:"
                        Write-Verbose "Exception: $($mgResult.Error.Exception.Message)"
                        Write-Verbose "Error Details: $($mgResult.ErrorDetails)"
                        
                        $friendlyError = Get-FriendlyErrorMessage -Exception $mgResult.Error.Exception -ErrorDetails $mgResult.ErrorDetails
                        $activationErrors += "$($roleData.DisplayName): $friendlyError"
                        
                        Write-Warning "Failed to activate group $($roleData.DisplayName): $friendlyError"
                    }
                }
                
                default {
                    $activationErrors += "$($roleData.DisplayName): Unsupported role type '$($roleData.Type)'"
                    Write-Warning "Unsupported role type: $($roleData.Type)"
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
        
        # Add delay after successful activations to allow Microsoft Graph to process changes
        if ($successCount -gt 0) {
            Write-Verbose "Waiting for Microsoft Graph to process role changes..."
            Start-Sleep -Seconds 3  # Increased initial delay
            
            # Clear role cache to ensure fresh data is fetched after activation
            Write-Verbose "Clearing role cache to force fresh data retrieval after activation"
            $script:CachedEligibleRoles = @()
            $script:CachedActiveRoles = @()
            $script:LastRoleFetchTime = $null
        }
        
        Write-Verbose "Refreshing role data"
        $refreshAttempts = 0
        $maxRefreshAttempts = 10
        $refreshSuccessful = $false
        
        while ($refreshAttempts -lt $maxRefreshAttempts -and -not $refreshSuccessful) {
            $refreshAttempts++
            try {
                if ($refreshAttempts -gt 1) {
                    Write-Verbose "Refresh attempt $refreshAttempts of $maxRefreshAttempts after waiting for Graph propagation..."
                    Start-Sleep -Seconds 3  # Reduced retry delay but still adequate
                    
                    # Clear cache again before retry to ensure fresh data
                    Write-Verbose "Clearing role cache before retry attempt"
                    $script:CachedEligibleRoles = @()
                    $script:CachedActiveRoles = @()
                    $script:LastRoleFetchTime = $null
                }
                
                Update-PIMRolesList -Form $Form -RefreshActive -RefreshEligible
                $refreshSuccessful = $true
                Write-Verbose "Role lists refreshed successfully"
            }
            catch {
                Write-Warning "Failed to refresh role lists (attempt $refreshAttempts): $_"
                if ($refreshAttempts -eq $maxRefreshAttempts) {
                    Write-Warning "All refresh attempts failed. You may need to manually refresh the role lists."
                }
            }
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

# Helper function to calculate effective duration
function Get-EffectiveDuration {
    param(
        [int]$RequestedMinutes,
        [int]$MaxDurationHours
    )
    
    $maxMinutes = $MaxDurationHours * 60
    
    if ($RequestedMinutes -gt $maxMinutes) {
        Write-Verbose "Requested duration ($RequestedMinutes minutes) exceeds maximum ($maxMinutes minutes)"
        $hours = [Math]::Floor($maxMinutes / 60)
        $minutes = $maxMinutes % 60
    }
    else {
        $hours = [Math]::Floor($RequestedMinutes / 60)
        $minutes = $RequestedMinutes % 60
    }
    
    return @{
        Hours = $hours
        Minutes = $minutes
        TotalMinutes = ($hours * 60) + $minutes
    }
}

# Helper function to extract user-friendly error messages
function Get-FriendlyErrorMessage {
    param(
        [System.Exception]$Exception,
        [object]$ErrorDetails
    )
    
    $errorMessage = $Exception.Message
    
    # Try to parse structured error details
    if ($ErrorDetails) {
        try {
            $errorObj = $ErrorDetails | ConvertFrom-Json
            if ($errorObj.error.message) {
                $errorMessage = $errorObj.error.message
                
                # Extract specific error codes for common scenarios
                switch ($errorObj.error.code) {
                    'RoleAssignmentRequestAcrsValidationFailed' {
                        return "Authentication context validation failed. The token does not contain the required authentication context claim. Please ensure you've completed the authentication context challenge."
                    }
                    'RoleAssignmentExists' {
                        return "This role is already active or a request is already pending."
                    }
                    'RoleEligibilityScheduleRequestNotFound' {
                        return "You are not eligible for this role. Please check your PIM eligibility."
                    }
                    'RoleDefinitionDoesNotExist' {
                        return "The requested role no longer exists. Please refresh the role list."
                    }
                    'AuthorizationFailed' {
                        return "You don't have permission to activate this role."
                    }
                    'InvalidAuthenticationToken' {
                        return "Your authentication has expired. Please reconnect."
                    }
                    'RequestConflict' {
                        return "Another activation request is already in progress for this role."
                    }
                    default {
                        if ($errorObj.error.innerError -and $errorObj.error.innerError.message) {
                            return "$errorMessage - $($errorObj.error.innerError.message)"
                        }
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not parse error details: $($_.Exception.Message)"
        }
    }
    
    return $errorMessage
}

# Helper function to display activation results
function Show-ActivationResults {
    param(
        [int]$SuccessCount,
        [int]$TotalCount,
        [array]$Errors
    )
    
    if ($Errors.Count -gt 0) {
        $message = "Successfully activated $SuccessCount of $TotalCount role(s).`n`nErrors:`n$($Errors -join "`n")"
        Show-TopMostMessageBox -Message $message -Title "Activation Results" -Icon Warning
    }
    else {
        Show-TopMostMessageBox -Message "Successfully activated all $SuccessCount role(s)!" -Title "Success" -Icon Information
    }
}

# Helper function to get and cache authentication context tokens
function Get-AuthenticationContextToken {
    <#
    .SYNOPSIS
        Gets or retrieves a cached authentication context token for the specified context ID.
    
    .DESCRIPTION
        Manages authentication context tokens by caching them per context ID to avoid
        repeated authentication prompts. Validates token expiry and refreshes as needed.
        Uses Windows Web Account Manager (WAM) for authentication.
    
    .PARAMETER ContextId
        The authentication context ID (e.g., "c3") required by the role policy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ContextId
    )
    
    try {
        Write-Verbose "=== Starting authentication context token acquisition for context: $ContextId ==="
        
        # Initialize the AuthContextTokens hashtable if it doesn't exist
        if (-not $script:AuthContextTokens) {
            $script:AuthContextTokens = @{}
            Write-Verbose "Initialized AuthContextTokens cache"
        }
        
        # Check if we already have a valid token for this context
        if ($script:AuthContextTokens.ContainsKey($ContextId)) {
            $cachedToken = $script:AuthContextTokens[$ContextId]
            
            # Validate token is still fresh (less than 30 minutes old)
            if ($cachedToken.ExpiryTime -and (Get-Date) -lt $cachedToken.ExpiryTime) {
                Write-Verbose "Using cached authentication context token for context: $ContextId (expires: $($cachedToken.ExpiryTime))"
                return $cachedToken.AccessToken
            }
            else {
                Write-Verbose "Cached token for context $ContextId has expired, obtaining fresh token"
                $script:AuthContextTokens.Remove($ContextId)
            }
        }
        
        # Get current Graph context for tenant ID
        $currentContext = Get-MgContext
        $tenantId = if ($currentContext) { $currentContext.TenantId } else { $null }
        
        if (-not $tenantId) {
            throw "No active Microsoft Graph connection. Cannot determine tenant ID. Please ensure you're connected to Microsoft Graph."
        }
        
        Write-Verbose "Current tenant ID: $tenantId"
        Write-Verbose "PowerShell version: $($PSVersionTable.PSVersion)"
        Write-Verbose "Obtaining fresh authentication context token for context: $ContextId"
        
        # Ensure we're running PowerShell Core
        if ($PSEdition -ne "Core") {
            throw "WAM authentication requires PowerShell Core (PowerShell 7+)"
        }
        
        # Build the claims challenge format
        $claimsJson = '{"access_token":{"acrs":{"essential":true,"value":"' + $ContextId + '"}}}'
        Write-Verbose "Claims challenge: $claimsJson"
        
        Write-Verbose "=== Starting WAM authentication setup ==="
        
        # Check for Az.Accounts module (required for WAM dependencies)
        $AzAccountsModule = Get-Module -Name Az.Accounts -ListAvailable
        if ($null -eq $AzAccountsModule) {
            Write-Verbose "Installing Az.Accounts module for WAM dependencies"
            Install-Module -Name Az.Accounts -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
        }
        
        # Import Az.Accounts module
        Import-Module Az.Accounts -ErrorAction Stop
        Write-Verbose "Az.Accounts module loaded"
        
        # Find the location of the Azure.Common assembly
        $LoadedAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Select-Object -ExpandProperty Location
        $AzureCommon = $LoadedAssemblies | Where-Object { $_ -match "\\Modules\\Az.Accounts\\" -and $_ -match "Microsoft.Azure.Common" }
        
        if (-not $AzureCommon) {
            throw "Could not find Microsoft.Azure.Common assembly from Az.Accounts module"
        }
        
        $AzureCommonLocation = $AzureCommon.TrimEnd("Microsoft.Azure.Common.dll")
        Write-Verbose "Azure Common Location: $AzureCommonLocation"
        
        # Locate the required assemblies
        Write-Verbose "Locating required assemblies for WAM"
        $requiredAssemblies = @(
            'Microsoft.IdentityModel.Abstractions.dll',
            'Microsoft.Identity.Client.dll',
            'Microsoft.Identity.Client.Broker.dll',
            'Microsoft.Identity.Client.NativeInterop.dll',
            'Microsoft.Identity.Client.Extensions.Msal.dll',
            'System.Security.Cryptography.ProtectedData.dll'
        )
        
        $assemblies = @{}
        
        foreach ($assemblyFile in $requiredAssemblies) {
            $assemblyName = $assemblyFile.Replace('.dll', '')
            $found = Get-ChildItem -Path $AzureCommonLocation -Filter $assemblyFile -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            if (-not $found) {
                throw "Could not find required assembly: $assemblyFile"
            }
            $assemblies[$assemblyName] = $found
            Write-Verbose "Found $assemblyName at: $found"
        }
        
        # Get System.Diagnostics.TraceSource from .NET Core installation or module dependencies
        $sdts = $null
        
        # First, try to get the module path
        $moduleBase = $null
        $pimModule = Get-Module -Name 'PIMActivation' -ErrorAction SilentlyContinue
        if ($pimModule) {
            $moduleBase = $pimModule.ModuleBase
            Write-Verbose "PIMActivation module base: $moduleBase"
        }
        else {
            # Try to find the module in the PSModulePath
            $modulePaths = $env:PSModulePath -split ';'
            foreach ($path in $modulePaths) {
                $testPath = Join-Path $path 'PIMActivation'
                if (Test-Path $testPath) {
                    $moduleBase = $testPath
                    Write-Verbose "Found PIMActivation module at: $moduleBase"
                    break
                }
            }
        }
        
        # Check if System.Diagnostics.TraceSource is already loaded
        $loadedAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies()
        $tracesourceAssembly = $loadedAssemblies | Where-Object { $_.GetName().Name -eq "System.Diagnostics.TraceSource" }
        
        if ($tracesourceAssembly -and $tracesourceAssembly.Location) {
            $sdts = $tracesourceAssembly.Location
            Write-Verbose "Found System.Diagnostics.TraceSource from loaded assemblies: $sdts"
        }
        else {
            # Try multiple locations for System.Diagnostics.TraceSource
            $searchPaths = @()
            
            # Add module's lib directory if available
            if ($moduleBase) {
                $searchPaths += Join-Path $moduleBase 'lib'
                $searchPaths += Join-Path $moduleBase 'Dependencies'
                $searchPaths += $moduleBase
            }
            
            # Add Az.Accounts location
            $searchPaths += $AzureCommonLocation
            
            # Try .NET Core reference assemblies
            $RuntimeFrameworkMajorVersion = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription.Split()[-1].Split(".")[0]
            $possibleDotNetPaths = @(
                "C:\Program Files\dotnet\packs\Microsoft.NETCore.App.Ref",
                "C:\Program Files (x86)\dotnet\packs\Microsoft.NETCore.App.Ref",
                "$env:ProgramFiles\dotnet\packs\Microsoft.NETCore.App.Ref",
                "${env:ProgramFiles(x86)}\dotnet\packs\Microsoft.NETCore.App.Ref",
                "$env:DOTNET_ROOT\packs\Microsoft.NETCore.App.Ref"
            )
            
            foreach ($dotnetPath in $possibleDotNetPaths) {
                if (Test-Path $dotnetPath) {
                    $dotNetDirectory = Get-ChildItem -Path $dotnetPath -Filter "$RuntimeFrameworkMajorVersion.*" -Directory -ErrorAction SilentlyContinue | 
                        Sort-Object -Property Name -Descending | Select-Object -First 1
                    if ($dotNetDirectory) {
                        $searchPaths += $dotNetDirectory.FullName
                        Write-Verbose "Added .NET reference path: $($dotNetDirectory.FullName)"
                    }
                }
            }
            
            # Add runtime directory as last resort
            $runtimeDir = [System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
            $searchPaths += $runtimeDir
            
            # Search for the assembly
            foreach ($searchPath in $searchPaths) {
                if (Test-Path $searchPath) {
                    $found = Get-ChildItem -Path $searchPath -Filter "System.Diagnostics.TraceSource.dll" -Recurse -File -ErrorAction SilentlyContinue | 
                        Select-Object -First 1
                    if ($found) {
                        $sdts = $found.FullName
                        Write-Verbose "Found System.Diagnostics.TraceSource at: $sdts"
                        break
                    }
                }
            }
            
            # If still not found, check if it's available as a type (might be in GAC)
            if (-not $sdts) {
                try {
                    $traceSourceType = [System.Diagnostics.TraceSource]
                    if ($traceSourceType) {
                        Write-Verbose "System.Diagnostics.TraceSource is available as a type (likely in GAC)"
                        # Continue without loading it explicitly
                    }
                }
                catch {
                    Write-Warning "System.Diagnostics.TraceSource type is not available"
                }
            }
        }
        
        if ($sdts) {
            Write-Verbose "System.Diagnostics.TraceSource located at: $sdts"
        }
        else {
            Write-Warning "Could not locate System.Diagnostics.TraceSource.dll - WAM authentication might still work"
        }
        
        # Load the assemblies
        Write-Verbose "Loading WAM assemblies..."
        $loadedCount = 0
        $failedAssemblies = @()
        
        foreach ($assemblyPath in $assemblies.Values) {
            try {
                # Check if already loaded
                $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($assemblyPath)
                $alreadyLoaded = $loadedAssemblies | Where-Object { $_.GetName().Name -eq $assemblyName }
                
                if ($alreadyLoaded) {
                    Write-Verbose "Assembly already loaded: $assemblyName"
                    $loadedCount++
                }
                else {
                    [void][System.Reflection.Assembly]::LoadFrom($assemblyPath)
                    Write-Verbose "Loaded assembly: $(Split-Path -Leaf $assemblyPath)"
                    $loadedCount++
                }
            }
            catch {
                $failedAssemblies += Split-Path -Leaf $assemblyPath
                Write-Warning "Failed to load assembly $(Split-Path -Leaf $assemblyPath): $_"
            }
        }
        
        if ($sdts) {
            try {
                # Check if already loaded
                $alreadyLoaded = $loadedAssemblies | Where-Object { $_.GetName().Name -eq "System.Diagnostics.TraceSource" }
                
                if ($alreadyLoaded) {
                    Write-Verbose "System.Diagnostics.TraceSource already loaded"
                }
                else {
                    [void][System.Reflection.Assembly]::LoadFrom($sdts)
                    Write-Verbose "Loaded System.Diagnostics.TraceSource"
                }
            }
            catch {
                Write-Warning "Failed to load System.Diagnostics.TraceSource: $_"
            }
        }
        
        # Check if we have the minimum required assemblies
        $criticalAssemblies = @('Microsoft.Identity.Client', 'Microsoft.Identity.Client.Broker')
        $missingCritical = $criticalAssemblies | Where-Object { $failedAssemblies -contains "$_.dll" }
        
        if ($missingCritical) {
            throw "Critical assemblies missing for WAM authentication: $($missingCritical -join ', '). Please ensure Az.Accounts module is properly installed."
        }
        
        Write-Verbose "WAM assembly loading completed - Loaded: $loadedCount, Failed: $($failedAssemblies.Count)"
        
        # If System.Diagnostics.TraceSource couldn't be loaded as file, it might still work if the type exists
        if (-not $sdts -or $failedAssemblies -contains "System.Diagnostics.TraceSource.dll") {
            $sdts = "System.Diagnostics.TraceSource" # Use as type name reference
        }
        
        # C# code for WAM authentication with claims
        $code = @"
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.IdentityModel.Abstractions;
using Microsoft.Identity.Client.NativeInterop;
using Microsoft.Identity.Client.Extensions.Msal;

public class PIMAuthContextHelper
{
    // Get window handle of the console window
    [DllImport("user32.dll", ExactSpelling = true)]
    public static extern IntPtr GetAncestor(IntPtr hwnd, GetAncestorFlags flags);
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    
    public enum GetAncestorFlags
    {   
        GetParent = 1,
        GetRoot = 2,
        GetRootOwner = 3
    }
    
    public static IntPtr GetConsoleOrTerminalWindow()
    {
        IntPtr consoleHandle = GetConsoleWindow();
        if (consoleHandle == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }
        IntPtr handle = GetAncestor(consoleHandle, GetAncestorFlags.GetRootOwner);
        return (handle != IntPtr.Zero) ? handle : consoleHandle;
    }    
    
    // Method for retrieving the access token with authentication context claims
    public static string GetAccessTokenWithAuthContext(string clientId, string tenantId, string redirectUri, string[] scopes, string claimsJson)
    {
        try
        {
            // Run the async method synchronously to avoid deadlocks
            var task = Task.Run(async () => await GetAccessTokenWithAuthContextAsync(clientId, tenantId, redirectUri, scopes, claimsJson));
            
            // Wait with timeout
            if (task.Wait(TimeSpan.FromSeconds(120)))
            {
                return task.Result;
            }
            else
            {
                throw new TimeoutException("Authentication timed out after 120 seconds");
            }
        }
        catch (AggregateException ae)
        {
            // Unwrap aggregate exceptions
            var innerException = ae.InnerException;
            if (innerException != null)
            {
                throw innerException;
            }
            throw;
        }
    }
    
    private static async Task<string> GetAccessTokenWithAuthContextAsync(string clientId, string tenantId, string redirectUri, string[] scopes, string claimsJson)
    {
        // Setup broker options
        var brokerOptions = new BrokerOptions(BrokerOptions.OperatingSystems.Windows)
        {
            Title = "PIM Role Activation - Authentication Context Required"
        };
        
        var authority = $"https://login.microsoftonline.com/{tenantId}";
        
        var appBuilder = PublicClientApplicationBuilder.Create(clientId)
            .WithAuthority(authority)
            .WithBroker(brokerOptions)
            .WithRedirectUri(redirectUri);
        
        // Try to set parent window if available
        var windowHandle = GetConsoleOrTerminalWindow();
        if (windowHandle != IntPtr.Zero)
        {
            appBuilder = appBuilder.WithParentActivityOrWindow(() => windowHandle);
        }
        
        IPublicClientApplication publicClientApp = appBuilder.Build();
        
        // Create cancellation token
        using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(120)))
        {
            try
            {
                // Always do interactive authentication for authentication context
                var result = await publicClientApp
                    .AcquireTokenInteractive(scopes)
                    .WithClaims(claimsJson)
                    .WithPrompt(Prompt.SelectAccount)
                    .WithUseEmbeddedWebView(false) // Force system browser/WAM
                    .ExecuteAsync(cts.Token)
                    .ConfigureAwait(false);
                
                return result.AccessToken;
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException("Authentication was cancelled or timed out");
            }
        }
    }
}
"@
        
        # List of assemblies we need to reference - filter out null/empty values
        $referencedAssemblies = @(
            $assemblies['Microsoft.IdentityModel.Abstractions'],
            $assemblies['Microsoft.Identity.Client'],
            $assemblies['Microsoft.Identity.Client.Broker'],
            $assemblies['Microsoft.Identity.Client.NativeInterop'],
            $assemblies['Microsoft.Identity.Client.Extensions.Msal'],
            $assemblies['System.Security.Cryptography.ProtectedData']
        ) | Where-Object { $_ }
        
        # Add System.Diagnostics.TraceSource if available
        if ($sdts -and $sdts -ne "System.Diagnostics.TraceSource") {
            $referencedAssemblies += $sdts
        }
        
        # Add standard assemblies
        $referencedAssemblies += @("netstandard", "System.Linq", "System.Threading.Tasks")
        
        # Get the access token with WAM
        $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph PowerShell
        $redirectUri = "http://localhost"
        $scopes = @("https://graph.microsoft.com/.default")
        
        Write-Verbose "=== Attempting WAM authentication with claims ==="
        Write-Verbose "Client ID: $clientId"
        Write-Verbose "Redirect URI: $redirectUri"
        Write-Verbose "Scopes: $($scopes -join ', ')"
        Write-Verbose "Referenced assemblies count: $($referencedAssemblies.Count)"
        
        $tokenStartTime = Get-Date
        
        try {
            # Check if type already exists
            $existingType = [System.Type]::GetType("PIMAuthContextHelper")
            if ($existingType) {
                Write-Verbose "PIMAuthContextHelper type already exists, using it"
            }
            else {
                Write-Verbose "Adding PIMAuthContextHelper type"
                Add-Type -ReferencedAssemblies $referencedAssemblies -TypeDefinition $code -Language CSharp -ErrorAction Stop
                Write-Verbose "PIMAuthContextHelper type added successfully"
            }
            
            Write-Verbose "Calling PIMAuthContextHelper.GetAccessTokenWithAuthContext"
            $accessToken = [PIMAuthContextHelper]::GetAccessTokenWithAuthContext($clientId, $tenantId, $redirectUri, $scopes, $claimsJson)
            
            if (-not $accessToken) {
                throw "No access token returned from WAM authentication"
            }
        }
        catch [System.Reflection.ReflectionTypeLoadException] {
            $loaderExceptions = $_.Exception.LoaderExceptions
            foreach ($loaderException in $loaderExceptions) {
                Write-Verbose "Loader exception: $($loaderException.Message)"
            }
            throw "Failed to load types: $($_.Exception.Message)"
        }
        catch {
            Write-Verbose "WAM authentication error: $($_.Exception.Message)"
            Write-Verbose "Exception type: $($_.Exception.GetType().FullName)"
            
            # If it's a specific error about window handle or broker, provide more context
            if ($_.Exception.Message -like "*broker*" -or $_.Exception.Message -like "*window*") {
                Write-Warning "WAM broker authentication failed. This might be due to:"
                Write-Warning "- Running in a non-interactive session"
                Write-Warning "- WAM not being available on this system"
                Write-Warning "- Missing Windows updates"
            }
            
            throw $_
        }
        
        $tokenDuration = (Get-Date) - $tokenStartTime
        Write-Verbose "Successfully obtained authentication context token via WAM in $($tokenDuration.TotalSeconds) seconds"
        Write-Verbose "Token length: $($accessToken.Length)"
        
        # Cache the token for reuse (assume 45 minutes expiry for safety)
        $expiryTime = (Get-Date).AddMinutes(45)
        
        # Validate that the token contains the expected authentication context claim
        $isValidToken = Test-AuthenticationContextToken -AccessToken $accessToken -ExpectedContextId $ContextId
        if (-not $isValidToken) {
            Write-Warning "Authentication context token validation failed - token does not contain expected context claim: $ContextId"
            Write-Verbose "Token might still be valid - continuing anyway"
        }
        
        $script:AuthContextTokens[$ContextId] = @{
            AccessToken = $accessToken
            ExpiryTime = $expiryTime
            ContextId = $ContextId
        }
        
        Write-Verbose "Cached authentication context token for context: $ContextId (expires: $expiryTime)"
        Write-Verbose "=== Authentication context token acquisition completed successfully ==="
        return $accessToken
    }
    catch {
        $errorMessage = "Failed to obtain authentication context token for context $ContextId`: $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Verbose "Exception details:"
        Write-Verbose "  Type: $($_.Exception.GetType().FullName)"
        Write-Verbose "  Message: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Verbose "  Inner Exception: $($_.Exception.InnerException.Message)"
        }
        Write-Verbose "  Stack Trace: $($_.Exception.StackTrace)"
        
        # Return null to let the calling function handle the failure
        return $null
    }
}

# Helper function to perform immediate activation with authentication context token
function Invoke-PIMActivationWithAuthContextToken {
    <#
    .SYNOPSIS
        Performs immediate PIM role activation using a pre-obtained authentication context token.
    
    .DESCRIPTION
        Makes direct REST API calls to activate PIM roles immediately after obtaining the
        correct authentication context token, eliminating timing issues.
    
    .PARAMETER ActivationParams
        Hashtable containing the activation request parameters.
    
    .PARAMETER RoleType
        Type of role being activated ('Entra' or 'Group').
        
    .PARAMETER AuthContextToken
        The authentication context token to use for the activation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ActivationParams,
        
        [Parameter(Mandatory)]
        [ValidateSet('Entra', 'Group')]
        [string]$RoleType,
        
        [Parameter(Mandatory)]
        [string]$AuthContextToken
    )
    
    try {
        Write-Verbose "Performing immediate activation with authentication context token"
        
        # Prepare REST API headers with the authentication context token
        $headers = @{
            'Authorization' = "Bearer $AuthContextToken"
            'Content-Type' = 'application/json'
        }
        
        # Convert activation parameters to the format expected by Microsoft Graph REST API
        $restParams = @{
            Action           = "selfActivate"
            PrincipalId      = $ActivationParams.principalId
            Justification    = $ActivationParams.justification
            ScheduleInfo     = @{
                StartDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
                Expiration    = @{
                    Type     = "AfterDuration"
                    Duration = $ActivationParams.scheduleInfo.expiration.duration
                }
            }
        }
        
        # Add role-specific parameters and determine the correct REST API endpoint
        $apiUri = ""
        switch ($RoleType) {
            'Entra' {
                $restParams.RoleDefinitionId = $ActivationParams.roleDefinitionId
                $restParams.DirectoryScopeId = if ([string]::IsNullOrEmpty($ActivationParams.directoryScopeId)) { "/" } else { $ActivationParams.directoryScopeId }
                $apiUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests"
                
                # Add ticket info only if present and required
                if ($ActivationParams.ContainsKey('ticketInfo') -and $ActivationParams.ticketInfo) {
                    $restParams.TicketInfo = $ActivationParams.ticketInfo
                }
            }
            'Group' {
                $restParams.GroupId = $ActivationParams.groupId
                $restParams.AccessId = $ActivationParams.accessId
                $apiUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentScheduleRequests"
                
                # Add ticket info only if present and required
                if ($ActivationParams.ContainsKey('ticketInfo') -and $ActivationParams.ticketInfo) {
                    $restParams.TicketInfo = $ActivationParams.ticketInfo
                }
            }
        }
        
        Write-Verbose "Submitting immediate PIM activation request"
        Write-Verbose "API URI: $apiUri"
        Write-Verbose "Parameters: $($restParams | ConvertTo-Json -Depth 5 -Compress)"
        
        # Submit activation request using direct REST API call - IMMEDIATE execution
        $requestBody = $restParams | ConvertTo-Json -Depth 5
        $activationStartTime = Get-Date
        $response = Invoke-RestMethod -Uri $apiUri -Headers $headers -Method Post -Body $requestBody -ErrorAction Stop
        $activationDuration = (Get-Date) - $activationStartTime
        
        Write-Verbose "PIM activation successful with authentication context - Response ID: $($response.Id) (completed in $($activationDuration.TotalSeconds) seconds)"
        return @{ Success = $true; Response = $response }
    }
    catch {
        # Enhanced error handling for authentication context activations
        Write-Verbose "Authentication context activation failed: $($_.Exception.Message)"
        $errorDetails = $null
        
        # Capture error details from REST API responses
        if ($_.Exception -is [System.Net.WebException]) {
            $webException = $_.Exception
            if ($webException.Response) {
                try {
                    $responseStream = $webException.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($responseStream)
                    $responseBody = $reader.ReadToEnd()
                    $reader.Close()
                    
                    # Try to parse the error response
                    try {
                        $errorResponse = $responseBody | ConvertFrom-Json
                        if ($errorResponse -and $errorResponse.error) {
                            $errorDetails = $errorResponse.error.message
                            
                            if ($errorResponse.error.code -eq "RoleAssignmentRequestAcrsValidationFailed") {
                                Write-Verbose "Authentication context validation failed - token may be invalid or expired"
                            }
                        } else {
                            $errorDetails = $responseBody
                        }
                    }
                    catch {
                        $errorDetails = $responseBody
                    }
                } catch {
                    $errorDetails = $webException.Message
                }
            }
        }
        
        # Fallback to standard PowerShell error details
        if (-not $errorDetails -and $_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errorDetails = $_.ErrorDetails.Message
        }
        
        return @{ Success = $false; Error = $_; ErrorDetails = $errorDetails }
    }
}

# Helper function to perform PIM role activation using Microsoft Graph SDK
function Invoke-PIMActivationWithMgGraph {
    <#
    .SYNOPSIS
        Performs PIM role activation using the Microsoft Graph PowerShell SDK.
    
    .DESCRIPTION
        Makes Microsoft Graph SDK calls to activate PIM roles for standard scenarios
        that don't require authentication context tokens.
    
    .PARAMETER ActivationParams
        Hashtable containing the activation request parameters.
    
    .PARAMETER RoleType
        Type of role being activated ('Entra' or 'Group').
        
    .PARAMETER AuthenticationContextId
        Optional. The authentication context ID if this activation requires authentication context.
        When provided, this function will use cached authentication context tokens.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ActivationParams,
        
        [Parameter(Mandatory)]
        [ValidateSet('Entra', 'Group')]
        [string]$RoleType,
        
        [Parameter()]
        [string]$AuthenticationContextId
    )
    
    try {
        Write-Verbose "Performing PIM activation using Microsoft Graph SDK"
        Write-Verbose "Role Type: $RoleType"
        
        # If authentication context is required, use the specialized function
        if ($AuthenticationContextId) {
            Write-Verbose "Authentication context required: $AuthenticationContextId"
            
            # Get the cached authentication context token
            $authContextToken = Get-AuthenticationContextToken -ContextId $AuthenticationContextId
            if (-not $authContextToken) {
                throw "Failed to obtain authentication context token for context: $AuthenticationContextId"
            }
            
            # Use the authentication context token function
            return Invoke-PIMActivationWithAuthContextToken -ActivationParams $ActivationParams -RoleType $RoleType -AuthContextToken $authContextToken
        }
        
        Write-Verbose "Using standard Microsoft Graph SDK for activation"
        
        $activationStartTime = Get-Date
        $response = $null
        
        # Submit activation request using Microsoft Graph SDK
        switch ($RoleType) {
            'Entra' {
                Write-Verbose "Activating Entra ID role via Microsoft Graph SDK"
                Write-Verbose "Role Definition ID: $($ActivationParams.roleDefinitionId)"
                Write-Verbose "Principal ID: $($ActivationParams.principalId)"
                Write-Verbose "Directory Scope: $($ActivationParams.directoryScopeId)"
                
                # Build the request body for Entra roles
                $requestBody = @{
                    action = $ActivationParams.action
                    principalId = $ActivationParams.principalId
                    roleDefinitionId = $ActivationParams.roleDefinitionId
                    directoryScopeId = $ActivationParams.directoryScopeId
                    justification = $ActivationParams.justification
                    scheduleInfo = $ActivationParams.scheduleInfo
                }
                
                # Add ticket info if present
                if ($ActivationParams.ContainsKey('ticketInfo') -and $ActivationParams.ticketInfo) {
                    $requestBody.ticketInfo = $ActivationParams.ticketInfo
                }
                
                # Use Microsoft Graph SDK to submit the request
                $response = New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $requestBody -ErrorAction Stop
            }
            
            'Group' {
                Write-Verbose "Activating Group role via Microsoft Graph SDK"
                Write-Verbose "Group ID: $($ActivationParams.groupId)"
                Write-Verbose "Principal ID: $($ActivationParams.principalId)"
                Write-Verbose "Access ID: $($ActivationParams.accessId)"
                
                # Build the request body for Group roles
                $requestBody = @{
                    action = $ActivationParams.action
                    principalId = $ActivationParams.principalId
                    groupId = $ActivationParams.groupId
                    accessId = $ActivationParams.accessId
                    justification = $ActivationParams.justification
                    scheduleInfo = $ActivationParams.scheduleInfo
                }
                
                # Add ticket info if present
                if ($ActivationParams.ContainsKey('ticketInfo') -and $ActivationParams.ticketInfo) {
                    $requestBody.ticketInfo = $ActivationParams.ticketInfo
                }
                
                # Use Microsoft Graph SDK to submit the request
                $response = New-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleRequest -BodyParameter $requestBody -ErrorAction Stop
            }
        }
        
        $activationDuration = (Get-Date) - $activationStartTime
        Write-Verbose "PIM activation successful via Microsoft Graph SDK - Response ID: $($response.Id) (completed in $($activationDuration.TotalSeconds) seconds)"
        
        return @{ Success = $true; Response = $response }
    }
    catch {
        Write-Verbose "Microsoft Graph SDK activation failed: $($_.Exception.Message)"
        $errorDetails = $null
        
        # Extract error details from Microsoft Graph exceptions
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errorDetails = $_.ErrorDetails.Message
        }
        elseif ($_.Exception.Message) {
            $errorDetails = $_.Exception.Message
        }
        
        return @{ Success = $false; Error = $_; ErrorDetails = $errorDetails }
    }
}