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
                    action           = $ActivationParams.action
                    principalId      = $ActivationParams.principalId
                    roleDefinitionId = $ActivationParams.roleDefinitionId
                    directoryScopeId = $ActivationParams.directoryScopeId
                    justification    = $ActivationParams.justification
                    scheduleInfo     = $ActivationParams.scheduleInfo
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
                    action        = $ActivationParams.action
                    principalId   = $ActivationParams.principalId
                    groupId       = $ActivationParams.groupId
                    accessId      = $ActivationParams.accessId
                    justification = $ActivationParams.justification
                    scheduleInfo  = $ActivationParams.scheduleInfo
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
        
        return @{ Success = $true; Response = $response; IsAzureResource = $false }
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
        
        return @{ Success = $false; Error = $_; ErrorDetails = $errorDetails; IsAzureResource = $false }
    }
}