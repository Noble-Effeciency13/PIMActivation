function Invoke-SingleRoleActivation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$RoleData,
        
        [Parameter(Mandatory)]
        [string]$Justification,
        
        [Parameter(Mandatory)]
        [hashtable]$EffectiveDuration,
        
        [hashtable]$TicketInfo,
        
        [string]$AuthContextToken,
        
        [string]$AuthenticationContextId,
        
        [switch]$UseFallbackMethod
    )
    
    Write-Verbose "Activating role: $($RoleData.DisplayName) [Type: $($RoleData.Type)]"
    
    try {
        switch ($RoleData.Type) {
            'Entra' {
                # Check eligibility for Entra roles
                $eligibilityCheck = Test-PIMRoleEligibility -UserId $script:CurrentUser.Id -RoleDefinitionId $RoleData.RoleDefinitionId
                if (-not $eligibilityCheck.IsEligible) {
                    throw "User is not eligible for this role assignment"
                }
                Write-Verbose "Eligibility check completed. IsEligible: $($eligibilityCheck.IsEligible)"
                
                # Get activation parameters
                $activationParams = Get-RoleActivationParameters -RoleData $RoleData -Justification $Justification -EffectiveDuration $EffectiveDuration -TicketInfo $TicketInfo
                
                # Choose activation method
                if ($AuthContextToken) {
                    Write-Verbose "Using cached authentication context token for immediate activation"
                    $mgResult = Invoke-PIMActivationWithAuthContextToken -ActivationParams $activationParams -RoleType 'Entra' -AuthContextToken $AuthContextToken
                }
                elseif ($AuthenticationContextId -and $UseFallbackMethod) {
                    Write-Verbose "Falling back to original authentication context method for Entra role"
                    $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Entra' -AuthenticationContextId $AuthenticationContextId
                }
                else {
                    Write-Verbose "Using Microsoft Graph SDK for Entra role without authentication context requirement"
                    $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Entra'
                }
                
                return $mgResult
            }
            
            'Group' {
                # Get activation parameters
                $activationParams = Get-RoleActivationParameters -RoleData $RoleData -Justification $Justification -EffectiveDuration $EffectiveDuration -TicketInfo $TicketInfo
                
                # Choose activation method
                if ($AuthContextToken) {
                    Write-Verbose "Using cached authentication context token for immediate activation"
                    $mgResult = Invoke-PIMActivationWithAuthContextToken -ActivationParams $activationParams -RoleType 'Group' -AuthContextToken $AuthContextToken
                }
                elseif ($AuthenticationContextId -and $UseFallbackMethod) {
                    Write-Verbose "Falling back to original authentication context method for Group role"
                    $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Group' -AuthenticationContextId $AuthenticationContextId
                }
                else {
                    Write-Verbose "Using Microsoft Graph SDK for Group role without authentication context requirement"
                    $mgResult = Invoke-PIMActivationWithMgGraph -ActivationParams $activationParams -RoleType 'Group'
                }
                
                return $mgResult
            }
            
            'AzureResource' {
                # Get Azure-specific activation parameters
                $azureParams = Get-RoleActivationParameters -RoleData $RoleData -Justification $Justification -EffectiveDuration $EffectiveDuration -TicketInfo $TicketInfo
                
                # Azure Resource roles use direct function call
                $response = Invoke-AzureResourceRoleActivation @azureParams
                
                Write-Verbose "Azure Resource role activated successfully"
                return @{ Success = $true; Response = $response; IsAzureResource = $true }
            }
            
            default {
                throw "Unsupported role type: $($RoleData.Type)"
            }
        }
    }
    catch {
        $errorMessage = Get-FriendlyErrorMessage -Exception $_.Exception -ErrorDetails $_.ErrorDetails
        Write-Warning "Failed to activate $($RoleData.DisplayName): $errorMessage"
        return @{ Success = $false; Error = $_; ErrorMessage = $errorMessage; IsAzureResource = ($RoleData.Type -eq 'AzureResource') }
    }
}