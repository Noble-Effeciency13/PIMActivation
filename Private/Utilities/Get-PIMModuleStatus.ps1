function Get-PIMModuleStatus {
    <#
    .SYNOPSIS
        Gets the current status of PIM module loading
    
    .DESCRIPTION
        Returns information about which modules are loaded and their versions
    
    .OUTPUTS
        PSCustomObject with module status information
    #>
    [CmdletBinding()]
    param()
    
    $status = [PSCustomObject]@{
        RequiredVersions = $script:RequiredModuleVersions
        LoadedModules    = [System.Collections.ArrayList]::new()
        LoadingState     = $script:ModuleLoadingState
        Compatible       = $false
    }
    
    foreach ($moduleName in $script:RequiredModuleVersions.Keys) {
        $loadedModule = Get-Module -Name $moduleName
        if ($loadedModule) {
            $null = $status.LoadedModules.Add([PSCustomObject]@{
                    Name             = $moduleName
                    LoadedVersion    = $loadedModule.Version
                    RequiredVersion  = $script:RequiredModuleVersions[$moduleName]
                    IsCorrectVersion = ($loadedModule.Version -eq $script:RequiredModuleVersions[$moduleName])
                })
        }
    }
    
    # Test compatibility if modules are loaded
    if ($status.LoadedModules.Count -gt 0) {
        $status.Compatible = Test-PIMModuleCompatibility
    }
    
    return $status
}