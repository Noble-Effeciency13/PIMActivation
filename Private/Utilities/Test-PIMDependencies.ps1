function Test-PIMDependencies {
    <#
    .SYNOPSIS
        Comprehensive dependency check for PIM Activation module.
    
    .DESCRIPTION
        Performs a complete analysis of module dependencies, version compatibility,
        and system requirements for optimal PIM Activation functionality.
    
    .EXAMPLE
        Test-PIMDependencies
        Runs a full dependency check and reports status.
        
    .EXAMPLE
        Test-PIMDependencies -Detailed
        Provides detailed information about each dependency.
    
    .OUTPUTS
        PSCustomObject with dependency status and recommendations
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$Detailed
    )
    
    $result = [PSCustomObject]@{
        OverallStatus = 'Unknown'
        PowerShellVersion = [PSCustomObject]@{
            Current = $PSVersionTable.PSVersion
            Required = [version]'7.0'
            Compatible = $false
        }
        ModuleStatus = [System.Collections.ArrayList]::new()
        Conflicts = [System.Collections.ArrayList]::new()
        Recommendations = [System.Collections.ArrayList]::new()
        ReadyForActivation = $false
    }
    
    # Check PowerShell version
    $result.PowerShellVersion.Compatible = $PSVersionTable.PSVersion -ge [version]'7.0'
    
    if (-not $result.PowerShellVersion.Compatible) {
        $null = $result.Recommendations.Add("Upgrade to PowerShell 7.0 or later from https://aka.ms/powershell")
    }
    
    # Check required modules using our version requirements with suppressed verbose output
    $originalVerbosePreference = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'
    
    if (Get-Variable -Name 'script:RequiredModuleVersions' -Scope Script -ErrorAction SilentlyContinue) {
        $requiredVersions = $script:RequiredModuleVersions
    }
    else {
        # Fallback if called outside module context
        $requiredVersions = @{
            'Microsoft.Graph.Authentication' = '2.29.0'
            'Microsoft.Graph.Users' = '2.29.0'
            'Microsoft.Graph.Identity.DirectoryManagement' = '2.29.0'
            'Microsoft.Graph.Identity.Governance' = '2.29.0'
            'Microsoft.Graph.Groups' = '2.29.0'
            'Microsoft.Graph.Identity.SignIns' = '2.29.0'
            'Az.Accounts' = '5.1.0'
        }
    }
    
    foreach ($moduleName in $requiredVersions.Keys) {
        $requiredVersion = [version]$requiredVersions[$moduleName]
        $loadedModule = Get-Module -Name $moduleName -ErrorAction SilentlyContinue
        $availableModules = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue
        
        $moduleStatus = [PSCustomObject]@{
            Name = $moduleName
            RequiredVersion = $requiredVersion
            LoadedVersion = if ($loadedModule) { $loadedModule.Version } else { $null }
            AvailableVersions = if ($availableModules) { $availableModules.Version | Sort-Object -Descending } else { @() }
            Status = 'Unknown'
            Loaded = [bool]$loadedModule
            Available = [bool]$availableModules
            VersionConflict = $false
        }
        
        # Determine status
        if ($loadedModule) {
            if ($loadedModule.Version -ge $requiredVersion) {
                $moduleStatus.Status = 'Loaded-Compatible'
            }
            else {
                $moduleStatus.Status = 'Loaded-Incompatible'
                $moduleStatus.VersionConflict = $true
                $null = $result.Conflicts.Add("$moduleName v$($loadedModule.Version) is loaded but v$requiredVersion+ is required")
            }
        }
        elseif ($availableModules) {
            $suitableVersions = $availableModules | Where-Object { $_.Version -ge $requiredVersion }
            if ($suitableVersions) {
                $moduleStatus.Status = 'Available-Compatible'
            }
            else {
                $moduleStatus.Status = 'Available-Incompatible'
                $null = $result.Recommendations.Add("Update $moduleName to version $requiredVersion or later")
            }
        }
        else {
            $moduleStatus.Status = 'Not-Available'
            $null = $result.Recommendations.Add("Install $moduleName version $requiredVersion or later")
        }
        
        $null = $result.ModuleStatus.Add($moduleStatus)
    }
    
    # Restore original verbose preference
    $VerbosePreference = $originalVerbosePreference
    
    # Determine overall status
    $incompatibleLoaded = $result.ModuleStatus | Where-Object { $_.Status -eq 'Loaded-Incompatible' }
    $notAvailable = $result.ModuleStatus | Where-Object { $_.Status -eq 'Not-Available' }
    $availableIncompatible = $result.ModuleStatus | Where-Object { $_.Status -eq 'Available-Incompatible' }
    
    if ($incompatibleLoaded) {
        $result.OverallStatus = 'Version-Conflicts'
        $null = $result.Recommendations.Add("Run 'Clear-ModuleVersionConflicts' to resolve version conflicts")
        $null = $result.Recommendations.Add("Restart PowerShell session after clearing conflicts")
    }
    elseif ($notAvailable -or $availableIncompatible) {
        $result.OverallStatus = 'Missing-Dependencies'
        $null = $result.Recommendations.Add("Run 'Install-RequiredModules' to install missing dependencies")
    }
    else {
        $result.OverallStatus = 'Ready'
        $result.ReadyForActivation = $true
    }
    
    # Display results only if detailed or if there are issues
    if ($Detailed -or (-not $result.ReadyForActivation)) {
        Write-Host "`n=== PIM Activation Dependency Check ===" -ForegroundColor Cyan
        Write-Host "PowerShell Version: $($result.PowerShellVersion.Current) " -NoNewline
        if ($result.PowerShellVersion.Compatible) {
            Write-Host "✓" -ForegroundColor Green
        }
        else {
            Write-Host "❌ (Requires 7.0+)" -ForegroundColor Red
        }
        
        Write-Host "`nModule Dependencies:" -ForegroundColor Cyan
        foreach ($module in $result.ModuleStatus) {
            $icon = switch ($module.Status) {
                'Loaded-Compatible' { '✓' }
                'Available-Compatible' { '⚡' }
                'Loaded-Incompatible' { '❌' }
                'Available-Incompatible' { '⚠️' }
                'Not-Available' { '❌' }
                default { '❓' }
            }
            
            $color = switch ($module.Status) {
                'Loaded-Compatible' { 'Green' }
                'Available-Compatible' { 'Yellow' }
                'Loaded-Incompatible' { 'Red' }
                'Available-Incompatible' { 'Red' }
                'Not-Available' { 'Red' }
                default { 'Gray' }
            }
            
            Write-Host "  $icon " -ForegroundColor $color -NoNewline
            Write-Host "$($module.Name) " -NoNewline
            
            if ($module.LoadedVersion) {
                Write-Host "v$($module.LoadedVersion) (loaded)" -ForegroundColor $color
            }
            elseif ($module.AvailableVersions) {
                Write-Host "v$($module.AvailableVersions[0]) (available)" -ForegroundColor $color
            }
            else {
                Write-Host "(not installed)" -ForegroundColor $color
            }
            
            if ($Detailed -and $module.AvailableVersions.Count -gt 1) {
                Write-Host "    Available versions: $($module.AvailableVersions -join ', ')" -ForegroundColor Gray
            }
        }
        
        Write-Host "`nOverall Status: " -NoNewline
        switch ($result.OverallStatus) {
            'Ready' { Write-Host "Ready for PIM Activation ✓" -ForegroundColor Green }
            'Version-Conflicts' { Write-Host "Version Conflicts Detected ❌" -ForegroundColor Red }
            'Missing-Dependencies' { Write-Host "Missing Dependencies ⚠️" -ForegroundColor Yellow }
            default { Write-Host $result.OverallStatus -ForegroundColor Gray }
        }
        
        if ($result.Conflicts.Count -gt 0) {
            Write-Host "`nConflicts:" -ForegroundColor Red
            foreach ($conflict in $result.Conflicts) {
                Write-Host "  ❌ $conflict" -ForegroundColor Red
            }
        }
        
        if ($result.Recommendations.Count -gt 0) {
            Write-Host "`nRecommendations:" -ForegroundColor Cyan
            foreach ($recommendation in $result.Recommendations) {
                Write-Host "  • $recommendation" -ForegroundColor White
            }
        }
        
        Write-Host ""
    }
    else {
        # Silent mode - only log to verbose
        Write-Verbose "Dependency check completed. Status: $($result.OverallStatus), Ready: $($result.ReadyForActivation)"
    }
    return $result
}
