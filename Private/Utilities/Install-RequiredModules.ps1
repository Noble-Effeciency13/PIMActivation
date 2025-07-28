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
        [switch]$IncludeAzureModules
    )
    
    $result = [PSCustomObject]@{
        Success = $true
        Error = $null
    }
    
    try {
        # Initialize module list with defaults if not provided
        if (-not $RequiredModules) {
            Write-Verbose "Using default Microsoft Graph module set"
            $RequiredModules = @(
                @{Name = 'Microsoft.Graph.Authentication'; MinVersion = '2.0.0'},
                @{Name = 'Microsoft.Graph.Users'; MinVersion = '2.0.0'},
                @{Name = 'Microsoft.Graph.Identity.DirectoryManagement'; MinVersion = '2.0.0'},
                @{Name = 'Microsoft.Graph.Identity.Governance'; MinVersion = '2.0.0'},
                @{Name = 'MSAL.PS'; MinVersion = '4.36.1'}
            )
            
            if ($IncludeAzureModules) {
                Write-Verbose "Including Azure PowerShell modules"
                $RequiredModules += @(
                    @{Name = 'Az.Accounts'; MinVersion = '2.0.0'},
                    @{Name = 'Az.Resources'; MinVersion = '6.0.0'}
                )
            }
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
                Install-Module -Name $module.Name `
                             -MinimumVersion $module.MinVersion `
                             -Scope $installScope `
                             -Force `
                             -AllowClobber `
                             -Repository PSGallery `
                             -ErrorAction Stop
                
                # Import the newly installed module
                Import-Module -Name $module.Name -MinimumVersion $module.MinVersion -ErrorAction Stop
                Write-Verbose "✓ $($module.Name) installed and imported successfully"
            }
            catch {
                # Fallback: retry with CurrentUser scope only
                try {
                    Write-Verbose "Retrying installation with CurrentUser scope..."
                    Install-Module -Name $module.Name `
                                 -MinimumVersion $module.MinVersion `
                                 -Scope CurrentUser `
                                 -Force `
                                 -AllowClobber `
                                 -ErrorAction Stop
                    
                    Import-Module -Name $module.Name -MinimumVersion $module.MinVersion -ErrorAction Stop
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