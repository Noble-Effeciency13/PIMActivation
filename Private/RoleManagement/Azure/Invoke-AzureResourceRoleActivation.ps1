function Invoke-AzureResourceRoleActivation {
    <#
    .SYNOPSIS
        Activates Azure Resource PIM roles using Azure REST API.
    
    .PARAMETER Scope
        The Azure resource scope (subscription, resource group, or resource).
    
    .PARAMETER RoleDefinitionId
        The role definition ID to activate.
    
    .PARAMETER PrincipalId
        The principal ID (user object ID) requesting activation.
    
    .PARAMETER RequestType
        The request type (typically 'SelfActivate').
    
    .PARAMETER Justification
        Justification text for the activation.
    
    .PARAMETER ScheduleInfo
        Schedule information including start time and expiration.
    
    .PARAMETER TicketInfo
        Optional ticket information if required by policy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Scope,
        
        [Parameter(Mandatory)]
        [string]$RoleDefinitionId,
        
        [Parameter(Mandatory)]
        [string]$PrincipalId,
        
        [Parameter(Mandatory)]
        [string]$RequestType,
        
        [Parameter(Mandatory)]
        [string]$Justification,
        
        [Parameter(Mandatory)]
        [hashtable]$ScheduleInfo,
        
        [hashtable]$TicketInfo
    )
    
    try {
        Write-Verbose "Activating Azure Resource role: $RoleDefinitionId in scope: $Scope"
        
        # Ensure we have an active Azure context
        $azContext = Get-AzContext
        if (-not $azContext) {
            throw "No Azure context available. Please connect to Azure first using Connect-AzAccount."
        }
        
        Write-Verbose "Using Azure context: $($azContext.Account.Id) in tenant $($azContext.Tenant.Id)"
        
        Write-Verbose "Using Az.Resources module for Azure Resource role activation"
        
        # Generate a unique name for the request
        $requestName = [System.Guid]::NewGuid().ToString()
        
        # Ensure the role definition ID is in the correct format (full ARM resource ID)
        $fullRoleDefinitionId = if ($RoleDefinitionId.StartsWith('/')) {
            $RoleDefinitionId  # Already a full resource ID
        } else {
            "$Scope/providers/Microsoft.Authorization/roleDefinitions/$RoleDefinitionId"
        }
        
        # According to Microsoft docs, the correct parameters for New-AzRoleAssignmentScheduleRequest are:
        $activationParams = @{
            Name = $requestName
            Scope = $Scope
            RoleDefinitionId = $fullRoleDefinitionId
            PrincipalId = $PrincipalId
            RequestType = $RequestType
            Justification = $Justification
        }
        
        # Add expiration duration using the correct parameter name
        if ($ScheduleInfo.Expiration -and $ScheduleInfo.Expiration.Duration) {
            $activationParams.ExpirationDuration = $ScheduleInfo.Expiration.Duration
        }
        
        # Add expiration type
        if ($ScheduleInfo.Expiration -and $ScheduleInfo.Expiration.Type) {
            $activationParams.ExpirationType = $ScheduleInfo.Expiration.Type
        }
        
        # Add ticket information if provided using correct parameter names
        if ($TicketInfo -and $TicketInfo.TicketNumber) {
            $activationParams.TicketNumber = $TicketInfo.TicketNumber
            if ($TicketInfo.TicketSystem) {
                $activationParams.TicketSystem = $TicketInfo.TicketSystem
            }
        }
        
        Write-Verbose "Submitting Azure Resource role activation using New-AzRoleAssignmentScheduleRequest"
        Write-Verbose "Parameters: Name=$requestName, Scope=$Scope, RoleDefinitionId=$RoleDefinitionId, PrincipalId=$PrincipalId"
        Write-Verbose "RequestType=$RequestType, ExpirationDuration=$($ScheduleInfo.Expiration.Duration)"
        
        # Use the Az.Resources cmdlet to activate the role
        $response = New-AzRoleAssignmentScheduleRequest @activationParams -ErrorAction Stop
        
        Write-Verbose "Azure Resource role activation request submitted successfully"
        Write-Verbose "Response ID: $($response.Id), Status: $($response.Status), Type: $($response.Type)"
        
        return $response
    }
    catch {
        Write-Verbose "Azure Resource role activation failed: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            try {
                $responseStream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($responseStream)
                $responseBody = $reader.ReadToEnd()
                $reader.Close()
                Write-Verbose "Error response body: $responseBody"
            }
            catch {
                Write-Verbose "Could not read error response body"
            }
        }
        throw "Failed to activate Azure Resource role: $($_.Exception.Message)"
    }
}