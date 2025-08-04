function Install-RequiredModules {
    <#
    .SYNOPSIS
        Installs required PowerShell modules for PIM activation.
    
    .DESCRIPTION
        Validates and installs necessary Microsoft Graph modules and optionally Azure PowerShell modules.
        Automatically handles NuGet provider setup, repository trust configuration, and module versioning.
        Falls back to CurrentUser scope if not running as administrator.
    
    .PARAMETER RequiredModules
        Array of hashtables containing module specifications with Name and MinVersion properties.
        If not provided, defaults to core Microsoft Graph modules required for PIM operations.
    
    .PARAMETER IncludeAzureModules
        Switch to include Azure PowerShell modules (Az.Accounts, Az.Resources) for Azure resource support.
    
    .EXAMPLE
        Install-RequiredModules
        Installs default Microsoft Graph modules for PIM operations.
    
    .EXAMPLE
        Install-RequiredModules -IncludeAzureModules
        Installs Microsoft Graph modules plus Azure PowerShell modules.
    
    .EXAMPLE
        $modules = @(@{Name='Microsoft.Graph.Users'; MinVersion='2.0.0'})
        Install-RequiredModules -RequiredModules $modules
        Installs only the specified modules.
    
    .OUTPUTS
        PSCustomObject
        Returns object with Success (boolean) and Error (string) properties indicating operation status.
    
    .NOTES
        Requires PowerShell 7 or later.
        Administrative privileges recommended for AllUsers scope installation.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [hashtable[]]$RequiredModules,
        
        [Parameter()]
        [switch]$IncludeAzureModules,
        
        [Parameter()]
        [switch]$Force,
        
        [Parameter()]
        [switch]$AutoResolveConflicts
    )
    
    $result = [PSCustomObject]@{
        Success = $true
        Error = $null
    }
    
    try {
        # Initialize module list with defaults if not provided
        if (-not $RequiredModules) {
            Write-Verbose "Using default Microsoft Graph module set"
            $moduleList = [System.Collections.ArrayList]::new()
            $null = $moduleList.AddRange(@(
                @{Name = 'Microsoft.Graph.Authentication'; MinVersion = '2.29.0'},
                @{Name = 'Microsoft.Graph.Users'; MinVersion = '2.29.0'},
                @{Name = 'Microsoft.Graph.Identity.DirectoryManagement'; MinVersion = '2.29.0'},
                @{Name = 'Microsoft.Graph.Identity.Governance'; MinVersion = '2.29.0'},
                @{Name = 'Microsoft.Graph.Groups'; MinVersion = '2.29.0'},
                @{Name = 'Microsoft.Graph.Identity.SignIns'; MinVersion = '2.29.0'},
                @{Name = 'Az.Accounts'; MinVersion = '5.1.0'}
            ))
            
            if ($IncludeAzureModules) {
                Write-Verbose "Including Azure PowerShell modules"
                $null = $moduleList.Add(@{Name = 'Az.Resources'; MinVersion = '6.0.0'})
            }
            
            $RequiredModules = $moduleList.ToArray()
        }
        
        # Check for version conflicts before proceeding
        Write-Host "🔍 Checking for module version conflicts..." -ForegroundColor Yellow
        $moduleVersionMap = @{}
        foreach ($module in $RequiredModules) {
            $moduleVersionMap[$module.Name] = $module.MinVersion
        }
        
        if ($AutoResolveConflicts) {
            Write-Verbose "Checking for module version conflicts..."
            $conflictResult = Test-ModuleVersionConflicts -RequiredModuleVersions $moduleVersionMap -AutoResolve:$Force -Force:$Force
            
            if ($conflictResult.HasConflicts -and -not $conflictResult.AutoResolutionSuccess) {
                Write-Warning "Some version conflicts could not be resolved automatically."
                if ($conflictResult.Recommendations.Count -gt 0) {
                    Write-Host "`nRecommendations:" -ForegroundColor Yellow
                    foreach ($recommendation in $conflictResult.Recommendations) {
                        Write-Host "  • $recommendation" -ForegroundColor White
                    }
                }
                
                if (-not $Force) {
                    $userChoice = Read-Host "`nContinue with installation anyway? (y/N)"
                    if ($userChoice -ne 'y') {
                        $result.Success = $false
                        $result.Error = "Installation cancelled due to version conflicts"
                        return $result
                    }
                }
            }
            elseif ($conflictResult.AutoResolutionSuccess) {
                Write-Verbose "Version conflicts resolved automatically"
            }
            else {
                Write-Verbose "No version conflicts detected"
            }
        }
        else {
            Write-Verbose "Skipping automatic conflict resolution"
        }
        
        $conflictCheck = Test-ModuleVersionConflicts -RequiredModuleVersions $moduleVersionMap
        
        if ($conflictCheck.HasConflicts) {
            Write-Warning "Module version conflicts detected!"
            
            foreach ($conflict in $conflictCheck.Conflicts) {
                if ($conflict.Severity -eq 'High') {
                    Write-Warning "❌ $($conflict.ModuleName): Loaded v$($conflict.LoadedVersion) < Required v$($conflict.RequiredVersion)"
                }
                else {
                    Write-Warning "⚠️  $($conflict.ModuleName): Loaded v$($conflict.LoadedVersion) > Required v$($conflict.RequiredVersion) (newer version detected)"
                }
            }
            
            if (-not $conflictCheck.SafeToProceed) {
                $result.Success = $false
                $result.Error = "Incompatible module versions are currently loaded. Please restart PowerShell session and try again."
                
                Write-Warning "Resolution steps:"
                foreach ($recommendation in $conflictCheck.Recommendations) {
                    Write-Warning "  • $recommendation"
                }
                
                return $result
            }
            else {
                Write-Warning "Proceeding with newer module versions - monitor for compatibility issues"
            }
        }
        else {
            Write-Verbose "✓ No module version conflicts detected"
        }
        
        # Determine installation scope based on privileges
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
        $installScope = if ($isAdmin) { 'AllUsers' } else { 'CurrentUser' }
        Write-Verbose "Installation scope: $installScope"
        
        # Process each required module
        foreach ($module in $RequiredModules) {
            Write-Verbose "Processing module: $($module.Name) (min version: $($module.MinVersion))"
            
            # Check if module is already loaded with sufficient version
            $loadedModule = Get-Module -Name $module.Name -ErrorAction SilentlyContinue
            if ($loadedModule -and $loadedModule.Version -ge $module.MinVersion) {
                Write-Verbose "✓ $($module.Name) v$($loadedModule.Version) already loaded"
                continue
            }
            
            # Check for suitable installed version
            $availableModules = Get-Module -ListAvailable -Name $module.Name -ErrorAction SilentlyContinue
            $suitableModule = $availableModules | 
                Where-Object { $_.Version -ge $module.MinVersion } | 
                Sort-Object Version -Descending | 
                Select-Object -First 1
            
            if ($suitableModule) {
                Write-Verbose "Found suitable version: $($module.Name) v$($suitableModule.Version)"
                try {
                    Import-Module -Name $module.Name -MinimumVersion $module.MinVersion -ErrorAction Stop
                    Write-Verbose "✓ $($module.Name) imported successfully"
                    continue
                }
                catch {
                    Write-Verbose "Import failed, proceeding with installation: $($_.Exception.Message)"
                }
            }
            
            # Install module if not available or insufficient version
            Write-Verbose "Installing $($module.Name)..."
            
            try {
                # Ensure NuGet provider is available
                $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
                if (-not $nugetProvider -or $nugetProvider.Version -lt '2.8.5.201') {
                    Write-Verbose "Installing NuGet provider..."
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope $installScope -ErrorAction Stop
                }
                
                # Configure PSGallery as trusted repository
                $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
                if ($psGallery.InstallationPolicy -ne 'Trusted') {
                    Write-Verbose "Configuring PSGallery as trusted repository"
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
                }
                
                # Install the module
                $installParams = @{
                    Name = $module.Name
                    MinimumVersion = $module.MinVersion
                    Scope = $installScope
                    Force = $true
                    AllowClobber = $true
                    Repository = 'PSGallery'
                    ErrorAction = 'Stop'
                }
                Install-Module @installParams
                
                # Import the newly installed module
                $importParams = @{
                    Name = $module.Name
                    MinimumVersion = $module.MinVersion
                    ErrorAction = 'Stop'
                }
                Import-Module @importParams
                Write-Verbose "✓ $($module.Name) installed and imported successfully"
            }
            catch {
                # Fallback: retry with CurrentUser scope only
                try {
                    Write-Verbose "Retrying installation with CurrentUser scope..."
                    $fallbackParams = @{
                        Name = $module.Name
                        MinimumVersion = $module.MinVersion
                        Scope = 'CurrentUser'
                        Force = $true
                        AllowClobber = $true
                        ErrorAction = 'Stop'
                    }
                    Install-Module @fallbackParams
                    
                    Import-Module @importParams
                    Write-Verbose "✓ $($module.Name) installed successfully (fallback)"
                }
                catch {
                    throw "Failed to install $($module.Name): $($_.Exception.Message)"
                }
            }
        }
        
        # Final validation of all required modules
        Write-Verbose "Validating module installation..."
        foreach ($module in $RequiredModules) {
            $loadedModule = Get-Module -Name $module.Name -ErrorAction SilentlyContinue
            if (-not $loadedModule) {
                throw "$($module.Name) failed to load after installation"
            }
            if ($loadedModule.Version -lt $module.MinVersion) {
                throw "$($module.Name) v$($loadedModule.Version) loaded but v$($module.MinVersion) required"
            }
        }
        
        Write-Verbose "All required modules validated successfully"
    }
    catch {
        $result.Success = $false
        $result.Error = $_.Exception.Message
        Write-Verbose "Installation failed: $($_.Exception.Message)"
    }
    
    return $result
}