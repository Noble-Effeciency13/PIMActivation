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