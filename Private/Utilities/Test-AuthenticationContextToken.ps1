function Test-AuthenticationContextToken {
    <#
    .SYNOPSIS
        Validates that an authentication context token contains the required claims.
    
    .DESCRIPTION
        Decodes and validates JWT tokens to ensure they contain the required authentication
        context (acrs) claim with the expected value. This helps troubleshoot authentication
        context issues by validating tokens before attempting role activations.
    
    .PARAMETER AccessToken
        The JWT access token to validate.
    
    .PARAMETER ExpectedContextId
        The expected authentication context ID (e.g., "c3") that should be present in the token.
    
    .EXAMPLE
        $isValid = Test-AuthenticationContextToken -AccessToken $token -ExpectedContextId "c3"
        Validates that the token contains the "c3" authentication context claim.
    
    .OUTPUTS
        System.Boolean
        Returns $true if the token contains the expected authentication context claim, $false otherwise.
    
    .NOTES
        This function performs basic JWT decoding without signature verification.
        It's intended for validation and troubleshooting purposes only.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,
        
        [Parameter(Mandatory)]
        [string]$ExpectedContextId
    )
    
    try {
        # Basic JWT token validation (format check)
        if (-not $AccessToken -or $AccessToken.Split('.').Count -ne 3) {
            Write-Verbose "Invalid JWT token format"
            return $false
        }
        
        # Extract and decode the payload (second part of JWT)
        $tokenParts = $AccessToken.Split('.')
        $payload = $tokenParts[1]
        
        # Add padding if necessary for Base64 decoding
        $padding = 4 - ($payload.Length % 4)
        if ($padding -ne 4) {
            $payload += '=' * $padding
        }
        
        # Decode from Base64URL
        $payload = $payload.Replace('-', '+').Replace('_', '/')
        $decodedBytes = [System.Convert]::FromBase64String($payload)
        $payloadJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
        
        # Parse the JSON payload
        $tokenData = $payloadJson | ConvertFrom-Json
        
        # Check for authentication context claim (acrs)
        if ($tokenData.acrs) {
            $authContextValue = $tokenData.acrs
            Write-Verbose "Found authentication context claim: $authContextValue"
            
            if ($authContextValue -eq $ExpectedContextId) {
                Write-Verbose "Authentication context token validation successful - contains expected context: $ExpectedContextId"
                return $true
            }
            else {
                Write-Verbose "Authentication context mismatch - expected: $ExpectedContextId, found: $authContextValue"
                return $false
            }
        }
        else {
            Write-Verbose "No authentication context claim (acrs) found in token"
            return $false
        }
    }
    catch {
        Write-Verbose "Failed to validate authentication context token: $($_.Exception.Message)"
        return $false
    }
}
