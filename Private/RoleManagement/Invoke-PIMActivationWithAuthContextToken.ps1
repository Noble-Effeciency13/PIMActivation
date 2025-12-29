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
            'Content-Type'  = 'application/json'
        }
        
        # Convert activation parameters to the format expected by Microsoft Graph REST API
        $restParams = @{
            Action        = "selfActivate"
            PrincipalId   = $ActivationParams.principalId
            Justification = $ActivationParams.justification
            ScheduleInfo  = @{
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
        return @{ Success = $true; Response = $response; IsAzureResource = $false }
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
                        }
                        else {
                            $errorDetails = $responseBody
                        }
                    }
                    catch {
                        $errorDetails = $responseBody
                    }
                }
                catch {
                    $errorDetails = $webException.Message
                }
            }
        }
        
        # Fallback to standard PowerShell error details
        if (-not $errorDetails -and $_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errorDetails = $_.ErrorDetails.Message
        }
        
        return @{ Success = $false; Error = $_; ErrorDetails = $errorDetails; IsAzureResource = $false }
    }
}