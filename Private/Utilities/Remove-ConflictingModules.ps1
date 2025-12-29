function Remove-ConflictingModules {
    <#
    .SYNOPSIS
        Removes conflicting module versions from the current session
    
    .DESCRIPTION
        Removes all loaded Microsoft Graph and optionally Az modules to prevent assembly conflicts
        when loading the pinned versions.
    
    .PARAMETER IncludeAzureModules
        Also remove Azure PowerShell modules from the session
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeAzureModules
    )
    
    Write-Verbose "Removing potentially conflicting modules from session..."
    
    # Remove all Microsoft Graph modules
    $graphModules = Get-Module -Name Microsoft.Graph*
    if ($graphModules) {
        Write-Verbose "Removing $($graphModules.Count) Microsoft Graph modules"
        $graphModules | Remove-Module -Force
    }
    
    # Remove Az modules if requested
    if ($IncludeAzureModules) {
        $azModules = Get-Module -Name Az.*
        if ($azModules) {
            Write-Verbose "Removing $($azModules.Count) Azure PowerShell modules"
            $azModules | Remove-Module -Force
        }
    }
    
    # Clear module loading state
    $script:ModuleLoadingState = @{}
}