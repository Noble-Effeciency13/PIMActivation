function Test-ModuleVersionConflicts {
    <#
    .SYNOPSIS
        Tests for module version conflicts before importing required modules.
    
    .DESCRIPTION
        Checks if any required modules are already loaded with incompatible versions
        and provides guidance on resolution. This prevents assembly loading conflicts
        and ensures clean module loading. Can automatically resolve conflicts when requested.
    
    .PARAMETER RequiredModuleVersions
        Hashtable of module names and their required minimum versions.
    
    .PARAMETER AutoResolve
        Automatically attempt to resolve conflicts by removing incompatible modules.
    
    .PARAMETER Force
        Force removal of conflicting modules without confirmation (use with AutoResolve).
    
    .EXAMPLE
        Test-ModuleVersionConflicts -RequiredModuleVersions $script:RequiredModuleVersions
        
    .EXAMPLE
        Test-ModuleVersionConflicts -RequiredModuleVersions $script:RequiredModuleVersions -AutoResolve
        
    .OUTPUTS
        PSCustomObject with conflict information and recommendations
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$RequiredModuleVersions,
        
        [Parameter()]
        [switch]$AutoResolve,
        
        [Parameter()]
        [switch]$Force
    )
    
    $result = [PSCustomObject]@{
        HasConflicts = $false
        Conflicts = [System.Collections.ArrayList]::new()
        Recommendations = [System.Collections.ArrayList]::new()
        SafeToProceed = $true
        AutoResolutionAttempted = $false
        AutoResolutionSuccess = $false
        ActionsPerformed = [System.Collections.ArrayList]::new()
    }
    
    Write-Verbose "Checking for module version conflicts..."
    
    foreach ($moduleName in $RequiredModuleVersions.Keys) {
        $requiredVersion = [version]$RequiredModuleVersions[$moduleName]
        
        # Check if module is currently loaded
        $loadedModules = Get-Module -Name $moduleName -ErrorAction SilentlyContinue
        
        if ($loadedModules) {
            foreach ($loadedModule in $loadedModules) {
                $loadedVersion = $loadedModule.Version
                
                if ($loadedVersion -lt $requiredVersion) {
                    Write-Verbose "Version conflict detected: $moduleName v$loadedVersion (loaded) < v$requiredVersion (required)"
                    
                    $conflict = [PSCustomObject]@{
                        ModuleName = $moduleName
                        LoadedVersion = $loadedVersion
                        RequiredVersion = $requiredVersion
                        ConflictType = 'IncompatibleVersion'
                        Severity = 'High'
                    }
                    
                    $null = $result.Conflicts.Add($conflict)
                    $result.HasConflicts = $true
                    $result.SafeToProceed = $false
                }
                elseif ($loadedVersion -gt $requiredVersion) {
                    Write-Verbose "Newer version detected: $moduleName v$loadedVersion (loaded) > v$requiredVersion (required)"
                    
                    $conflict = [PSCustomObject]@{
                        ModuleName = $moduleName
                        LoadedVersion = $loadedVersion
                        RequiredVersion = $requiredVersion
                        ConflictType = 'NewerVersion'
                        Severity = 'Low'
                    }
                    
                    $null = $result.Conflicts.Add($conflict)
                    # Newer versions are usually compatible, so we can proceed
                }
                else {
                    Write-Verbose "✓ $moduleName v$loadedVersion matches requirements"
                }
            }
        }
    }
    
    # Generate recommendations based on conflicts
    if ($result.HasConflicts) {
        $highSeverityConflicts = $result.Conflicts | Where-Object { $_.Severity -eq 'High' }
        
        if ($highSeverityConflicts) {
            if ($AutoResolve) {
                Write-Verbose "Auto-resolving module version conflicts..."
                $result.AutoResolutionAttempted = $true
                $resolutionSuccess = $true
                
                foreach ($conflict in $highSeverityConflicts) {
                    try {
                        Write-Verbose "Removing conflicting module: $($conflict.ModuleName) v$($conflict.LoadedVersion)"
                        Remove-Module $conflict.ModuleName -Force -ErrorAction Stop
                        $null = $result.ActionsPerformed.Add("Removed $($conflict.ModuleName) v$($conflict.LoadedVersion)")
                        Write-Verbose "Successfully removed $($conflict.ModuleName) v$($conflict.LoadedVersion)"
                    }
                    catch {
                        Write-Verbose "Failed to remove $($conflict.ModuleName): $($_.Exception.Message)"
                        $null = $result.ActionsPerformed.Add("Failed to remove $($conflict.ModuleName): $($_.Exception.Message)")
                        $resolutionSuccess = $false
                    }
                }
                
                if ($resolutionSuccess) {
                    $result.AutoResolutionSuccess = $true
                    $result.SafeToProceed = $true
                    $result.HasConflicts = $false
                    Write-Verbose "All conflicts resolved automatically"
                }
                else {
                    $null = $result.Recommendations.Add("Some conflicts could not be resolved automatically")
                    $null = $result.Recommendations.Add("Restart PowerShell session to ensure clean module loading")
                }
            }
            else {
                $null = $result.Recommendations.Add("Remove incompatible module versions using: Remove-Module <ModuleName> -Force")
                $null = $result.Recommendations.Add("Restart PowerShell session to ensure clean module loading")
                $null = $result.Recommendations.Add("Re-import PIMActivation module after restart")
                $null = $result.Recommendations.Add("Or run with -AutoResolve parameter for automatic conflict resolution")
                
                # Generate specific commands
                foreach ($conflict in $highSeverityConflicts) {
                    $null = $result.Recommendations.Add("Remove-Module $($conflict.ModuleName) -Force")
                }
            }
        }
        
        $newVersionConflicts = $result.Conflicts | Where-Object { $_.ConflictType -eq 'NewerVersion' }
        if ($newVersionConflicts) {
            $null = $result.Recommendations.Add("Newer module versions detected - this is usually safe but may cause compatibility issues")
            $null = $result.Recommendations.Add("Monitor for any unexpected behavior during PIM operations")
        }
    }
    
    return $result
}
