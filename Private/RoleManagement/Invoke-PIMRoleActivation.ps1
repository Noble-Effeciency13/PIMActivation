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
            Start-Sleep -Seconds 3
        }
        
        Write-Verbose "Refreshing role data"
        try {
            Update-PIMRolesList -Form $Form -RefreshActive -RefreshEligible
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

function Invoke-PIMActivationWithMgGraph {
    <#
    .SYNOPSIS
        Performs PIM role activation using direct REST API calls with authentication context support.
    
    .DESCRIPTION
        Obtains the correct access token with authentication context claims based on the
        role's policy requirements, then makes direct REST API calls to activate the role.
        Handles PowerShell version compatibility automatically.
    
    .PARAMETER ActivationParams
        Hashtable containing the activation request parameters.
    
    .PARAMETER RoleType
        Type of role being activated ('Entra' or 'Group').
        
    .PARAMETER AuthenticationContextId
        The authentication context ID required by the role policy (e.g., "c3").
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
        # Get current Graph context to determine current authentication state
        $currentContext = Get-MgContext
        $tenantId = if ($currentContext) { $currentContext.TenantId } else { $null }
        
        if (-not $tenantId) {
            throw "No active Microsoft Graph connection. Cannot determine tenant ID."
        }
        
        # Determine the correct token to use based on authentication context requirements
        $accessToken = $null
        
        if ($AuthenticationContextId) {
            Write-Verbose "Role requires authentication context: $AuthenticationContextId"
            Write-Verbose "Obtaining fresh token with authentication context claims"
            
            # Build the claims challenge format - this is what makes it work!
            $claimsJson = '{"access_token":{"acrs":{"essential":true,"value":"' + $AuthenticationContextId + '"}}}'
            
            # Clear any existing tokens to ensure fresh authentication
            Clear-MsalTokenCache -FromDisk
            
            # Use MSAL.PS in current PowerShell session (works in both PS7 and PS5.1)
            Write-Verbose "Using MSAL.PS in current PowerShell session for authentication context token"
            
            # Ensure MSAL.PS module is available
            if (-not (Get-Module -Name MSAL.PS -ListAvailable)) {
                throw "MSAL.PS module is not installed. Please run Install-RequiredModules first."
            }
            
            # Import MSAL.PS module
            try {
                Import-Module MSAL.PS -Force -ErrorAction Stop
                Write-Verbose "MSAL.PS module imported successfully"
            }
            catch {
                throw "Failed to import MSAL.PS module: $($_.Exception.Message)"
            }
            
            # Get token with the specific authentication context claims
            $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
            $scopes = @("https://graph.microsoft.com/.default")
            
            try {
                Write-Verbose "Obtaining authentication context token with claims: $claimsJson"
                $tokenStartTime = Get-Date
                $tokenResult = Get-MsalToken -ClientId $clientId `
                                             -TenantId $tenantId `
                                             -Scopes $scopes `
                                             -Interactive `
                                             -Prompt SelectAccount `
                                             -ExtraQueryParameters @{ "claims" = $claimsJson } `
                                             -ErrorAction Stop
                
                if (-not $tokenResult -or -not $tokenResult.AccessToken) {
                    throw "Failed to obtain authentication context token for context: $AuthenticationContextId"
                }
                
                $accessToken = $tokenResult.AccessToken
                $tokenDuration = (Get-Date) - $tokenStartTime
                Write-Verbose "Successfully obtained authentication context token in $($tokenDuration.TotalSeconds) seconds"
            }
            catch {
                throw "Failed to obtain authentication context token: $($_.Exception.Message)"
            }
        }
        else {
            Write-Verbose "No authentication context required, using current Microsoft Graph SDK token"
            # For roles without authentication context, simply use the existing Microsoft Graph SDK connection
            $accessToken = "USE_MGCONTEXT"
        }
        
        # Choose the appropriate method based on whether we're using existing MgContext or custom token
        if ($accessToken -eq "USE_MGCONTEXT") {
            Write-Verbose "Using Microsoft Graph SDK with existing authentication context"
            
            # Use the Microsoft Graph SDK directly - it will use the existing token
            try {
                if ($RoleType -eq 'Entra') {
                    $response = New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $ActivationParams
                    Write-Verbose "Role activation successful via Microsoft Graph SDK - Response ID: $($response.Id)"
                    
                    return @{
                        Success = $true
                        Response = $response
                        Error = $null
                        ErrorDetails = $null
                        UsedSdk = $true
                    }
                }
                elseif ($RoleType -eq 'Group') {
                    $response = New-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleRequest -BodyParameter $ActivationParams
                    Write-Verbose "Group role activation successful via Microsoft Graph SDK - Response ID: $($response.Id)"
                    
                    return @{
                        Success = $true
                        Response = $response
                        Error = $null
                        ErrorDetails = $null
                        UsedSdk = $true
                    }
                }
            }
            catch {
                Write-Warning "Microsoft Graph SDK call failed: $($_.Exception.Message)"
                return @{
                    Success = $false
                    Response = $null
                    Error = $_
                    ErrorDetails = $_.Exception.Message
                    UsedSdk = $true
                }
            }
        }
        else {
            Write-Verbose "Using direct REST API calls with custom authentication token"
            
            # Prepare REST API headers with the custom authentication token
            $headers = @{
                'Authorization' = "Bearer $accessToken"
                'Content-Type' = 'application/json'
            }
        }
        
        # Convert activation parameters to the format expected by Microsoft Graph REST API
        # Use proper UTC formatting to match authentication context token timing
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
        
        Write-Verbose "Submitting PIM activation request via direct REST API"
        Write-Verbose "API URI: $apiUri"
        Write-Verbose "Parameters: $($restParams | ConvertTo-Json -Depth 5 -Compress)"
        
        # Submit activation request using direct REST API call with immediate execution
        $requestBody = $restParams | ConvertTo-Json -Depth 5
        $activationStartTime = Get-Date
        $response = Invoke-RestMethod -Uri $apiUri -Headers $headers -Method Post -Body $requestBody -ErrorAction Stop
        $activationDuration = (Get-Date) - $activationStartTime
        
        Write-Verbose "PIM activation successful via direct REST API - Response ID: $($response.Id) (completed in $($activationDuration.TotalSeconds) seconds)"
        return @{ Success = $true; Response = $response }
    }
    catch {
        # Enhanced error handling for direct REST API calls
        Write-Verbose "Direct REST API activation failed: $($_.Exception.Message)"
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
                                Write-Verbose "Authentication context validation failed - check policy configuration"
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

# Helper function to get and cache authentication context tokens
function Get-AuthenticationContextToken {
    <#
    .SYNOPSIS
        Gets or retrieves a cached authentication context token for the specified context ID.
    
    .DESCRIPTION
        Manages authentication context tokens by caching them per context ID to avoid
        repeated authentication prompts. Validates token expiry and refreshes as needed.
    
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
        
        # Build the claims challenge format
        $claimsJson = '{"access_token":{"acrs":{"essential":true,"value":"' + $ContextId + '"}}}'
        Write-Verbose "Claims challenge: $claimsJson"
        
        # Clear any existing tokens to ensure fresh authentication
        try {
            Clear-MsalTokenCache -FromDisk -ErrorAction SilentlyContinue
            Write-Verbose "Cleared MSAL token cache"
        }
        catch {
            Write-Verbose "Could not clear MSAL token cache: $($_.Exception.Message)"
        }
        
        $accessToken = $null
        
        # Always use PowerShell 5.1 for MSAL interactive authentication (required for proper UI)
        Write-Verbose "Using PowerShell 5.1 process for MSAL interactive authentication"
        
        # Since MSAL.PS has compatibility issues in PowerShell 5.1, let's use a direct approach
        # We'll use the current PowerShell 7 session but with a method that works consistently
        Write-Verbose "Using fallback method: Direct MSAL authentication in current PowerShell session"
        
        # Try to get authentication context token using current session
        try {
            # Ensure MSAL.PS module is available in current session
            if (-not (Get-Module -Name MSAL.PS)) {
                Import-Module MSAL.PS -Force -ErrorAction Stop
                Write-Verbose "MSAL.PS module imported in current session"
            }
            
            # Get token with the specific authentication context claims
            $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
            $scopes = @("https://graph.microsoft.com/.default")
            
            Write-Verbose "Attempting to acquire token with authentication context claim for context: $ContextId"
            Write-Verbose "Claims JSON: $claimsJson"
            
            $tokenStartTime = Get-Date
            $tokenResult = Get-MsalToken -ClientId $clientId `
                                         -TenantId $tenantId `
                                         -Scopes $scopes `
                                         -Interactive `
                                         -Prompt SelectAccount `
                                         -ExtraQueryParameters @{ "claims" = $claimsJson } `
                                         -ErrorAction Stop
            
            if (-not $tokenResult) {
                throw "Get-MsalToken returned null result"
            }
            
            if (-not $tokenResult.AccessToken) {
                throw "Get-MsalToken returned result without AccessToken property"
            }
            
            $accessToken = $tokenResult.AccessToken
            $tokenDuration = (Get-Date) - $tokenStartTime
            Write-Verbose "Successfully obtained authentication context token directly in current session in $($tokenDuration.TotalSeconds) seconds (length: $($accessToken.Length))"
            
        } catch {
            Write-Verbose "Direct authentication in current session failed: $($_.Exception.Message)"
            Write-Verbose "This is expected if running in a non-interactive context or if MSAL UI components are not available"
            throw "Failed to obtain authentication context token for context $ContextId`: $($_.Exception.Message)"
        }
        
        # Cache the token for reuse (assume 45 minutes expiry for safety)
        $expiryTime = (Get-Date).AddMinutes(45)
        
        # Validate that the token contains the expected authentication context claim
        $isValidToken = Test-AuthenticationContextToken -AccessToken $accessToken -ExpectedContextId $ContextId
        if (-not $isValidToken) {
            Write-Warning "Authentication context token validation failed - token does not contain expected context claim: $ContextId"
            Write-Verbose "Token length: $($accessToken.Length)"
            # Don't cache invalid tokens, but let the calling function decide what to do
            Write-Warning "Token validation failed, but returning token anyway for troubleshooting"
            # return $null
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