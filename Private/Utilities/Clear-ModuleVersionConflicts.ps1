function Clear-ModuleVersionConflicts {
    <#
    .SYNOPSIS
        Removes conflicting Microsoft Graph module versions to prepare for clean loading.
    
    .DESCRIPTION
        Safely removes loaded Microsoft Graph modules that may cause version conflicts
        with PIM Activation requirements. This is a utility function to help resolve
        assembly loading conflicts.
    
    .EXAMPLE
        Clear-ModuleVersionConflicts
        Removes all loaded Microsoft Graph modules.
        
    .EXAMPLE  
        Clear-ModuleVersionConflicts -ModuleNames @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
        Removes only specified modules.
    
    .NOTES
        This function is automatically available when PIMActivation module is imported.
        Use this if you encounter version conflict errors.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string[]]$ModuleNames = @('Microsoft.Graph.*', 'Az.*')
    )
    
    $removedModules = [System.Collections.ArrayList]::new()
    
    try {
        foreach ($pattern in $ModuleNames) {
            Write-Verbose "Checking for modules matching pattern: $pattern"
            $matchingModules = Get-Module -Name $pattern -ErrorAction SilentlyContinue
            
            foreach ($module in $matchingModules) {
                if ($PSCmdlet.ShouldProcess($module.Name, "Remove Module")) {
                    Write-Verbose "Removing module: $($module.Name) v$($module.Version)"
                    try {
                        Remove-Module -Name $module.Name -Force -ErrorAction Stop
                        $null = $removedModules.Add([PSCustomObject]@{
                            Name = $module.Name
                            Version = $module.Version
                            Status = 'Removed'
                        })
                    }
                    catch {
                        Write-Warning "Failed to remove module $($module.Name): $($_.Exception.Message)"
                        $null = $removedModules.Add([PSCustomObject]@{
                            Name = $module.Name
                            Version = $module.Version
                            Status = 'Failed'
                            Error = $_.Exception.Message
                        })
                    }
                }
            }
        }
        
        if ($removedModules.Count -gt 0) {
            Write-Host "✓ Removed $($removedModules.Count) module(s) to resolve version conflicts" -ForegroundColor Green
            $removedModules | Format-Table Name, Version, Status -AutoSize
            Write-Host "You can now safely import and use PIMActivation module" -ForegroundColor Cyan
        }
        else {
            Write-Host "✓ No conflicting modules found" -ForegroundColor Green
        }
        
        return $removedModules
    }
    catch {
        Write-Error "Error clearing module conflicts: $($_.Exception.Message)"
        throw
    }
}
