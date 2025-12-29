function Initialize-AzureResourceSupport {
    <#
    .SYNOPSIS
        Validates that Azure Resource support is ready for use.
    
    .DESCRIPTION
        This function validates that Azure PowerShell modules required for 
        Azure Resource PIM operations are available and loaded. It uses the
        existing Import-PIMModule function for consistency.
    
    .OUTPUTS
        Boolean - Returns $true if Azure support is available, $false otherwise.
    
    .EXAMPLE
        $azureSupported = Initialize-AzureResourceSupport
        if ($azureSupported) {
            $azureRoles = Get-AzureResourceRoles
        }
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Validating Azure Resource support..."
    
    # Required Azure modules and their key commands
    $requiredAzureModules = @{
        'Az.Accounts' = @('Connect-AzAccount', 'Get-AzContext', 'Disconnect-AzAccount')
        'Az.Resources' = @('Get-AzSubscription', 'Get-AzRoleAssignment')
    }
    
    $failedModules = [System.Collections.ArrayList]::new()
    
    foreach ($moduleSpec in $requiredAzureModules.GetEnumerator()) {
        $moduleName = $moduleSpec.Key
        $requiredCommands = $moduleSpec.Value
        
        # Use Import-PIMModule to load the module
        $loadResult = Import-PIMModule -ModuleName $moduleName
        if (-not $loadResult) {
            Write-Warning "Failed to load required Azure module: $moduleName"
            $null = $failedModules.Add($moduleName)
            continue
        }
        
        # Validate required commands are available
        foreach ($command in $requiredCommands) {
            if (-not (Get-Command -Name $command -ErrorAction SilentlyContinue)) {
                Write-Warning "Required command '$command' is not available from module $moduleName"
                $null = $failedModules.Add("$moduleName (missing $command)")
            }
        }
    }
    
    if ($failedModules.Count -gt 0) {
        Write-Warning "Azure Resource support is not available due to missing modules/commands: $($failedModules -join ', ')"
        return $false
    }
    
    Write-Verbose "Azure Resource support is ready"
    return $true
}