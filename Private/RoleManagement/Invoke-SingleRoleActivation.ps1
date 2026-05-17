function Invoke-SingleRoleActivation {
    <#
    .SYNOPSIS
        Submits one PIM role activation request.

    .DESCRIPTION
        Activates a single Entra, PIM-enabled group, or Azure Resource role. The function is used
        by sequential fallback paths and authentication-context flows, and it accepts an optional
        scheduled start time so fallback requests preserve the same date/time selected in the
        activation dialog.

    .PARAMETER RoleData
        Role metadata for the selected role.

    .PARAMETER Justification
        Justification text to submit with the activation request.

    .PARAMETER EffectiveDuration
        Duration after policy maximum enforcement.

    .PARAMETER TicketInfo
        Optional ticket metadata when required by policy.

    .PARAMETER AuthContextToken
        Optional pre-obtained authentication-context token for direct REST activation.

    .PARAMETER AuthenticationContextId
        Optional authentication context ID used by fallback activation flows.

    .PARAMETER UseFallbackMethod
        Uses the existing authentication-context fallback method for Graph role activation.

    .PARAMETER ScheduleStartTime
        Optional local date/time when the activation should start.

    .PARAMETER AzureTargetScope
        Optional reduced Azure Resource scope for Azure activation.

    .PARAMETER LinkedRoleEligibilityScheduleId
        Optional Azure eligibility schedule ID required for reduced-scope activation.
    #>
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
        
        [switch]$UseFallbackMethod,

        [datetime]$ScheduleStartTime,

        [string]$AzureTargetScope,

        [string]$LinkedRoleEligibilityScheduleId
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
                $activationParamArgs = @{
                    RoleData          = $RoleData
                    Justification     = $Justification
                    EffectiveDuration = $EffectiveDuration
                    TicketInfo        = $TicketInfo
                }
                if ($PSBoundParameters.ContainsKey('ScheduleStartTime')) { $activationParamArgs.ScheduleStartTime = $ScheduleStartTime }
                $activationParams = Get-RoleActivationParameters @activationParamArgs
                
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
                $activationParamArgs = @{
                    RoleData          = $RoleData
                    Justification     = $Justification
                    EffectiveDuration = $EffectiveDuration
                    TicketInfo        = $TicketInfo
                }
                if ($PSBoundParameters.ContainsKey('ScheduleStartTime')) { $activationParamArgs.ScheduleStartTime = $ScheduleStartTime }
                $activationParams = Get-RoleActivationParameters @activationParamArgs
                
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
                $azureParamArgs = @{
                    RoleData          = $RoleData
                    Justification     = $Justification
                    EffectiveDuration = $EffectiveDuration
                    TicketInfo        = $TicketInfo
                }
                if ($PSBoundParameters.ContainsKey('ScheduleStartTime')) { $azureParamArgs.ScheduleStartTime = $ScheduleStartTime }
                if ($AzureTargetScope) { $azureParamArgs.AzureTargetScope = $AzureTargetScope }
                if ($LinkedRoleEligibilityScheduleId) { $azureParamArgs.LinkedRoleEligibilityScheduleId = $LinkedRoleEligibilityScheduleId }

                $azureParams = Get-RoleActivationParameters @azureParamArgs
                
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