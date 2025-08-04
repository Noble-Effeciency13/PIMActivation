function Initialize-PIMModules {
    <#
    .SYNOPSIS
        Initializes and loads required modules with version pinning and just-in-time loading
    
    .DESCRIPTION
        This function handles the initialization of required modules for PIM operations.
        It ensures only the exact required versions are loaded and removes other versions
        from the session to prevent assembly conflicts.
        
        Uses just-in-time loading - modules are only imported when actually needed.
    
    .PARAMETER Force
        Forces reinitialization even if modules are already loaded
        
    .OUTPUTS
        PSCustomObject with Success and Error properties
    #>
    [CmdletBinding()]
    param(
        [switch]$Force
    )
    
    # Define pinned module versions (working combination)
    $script:RequiredModuleVersions = @{
        'Microsoft.Graph.Authentication' = '2.29.1'
        'Microsoft.Graph.Users' = '2.29.1'
        'Microsoft.Graph.Identity.DirectoryManagement' = '2.29.1'
        'Microsoft.Graph.Identity.Governance' = '2.29.1'
        'Az.Accounts' = '5.1.0'
    }
    
    $result = [PSCustomObject]@{
        Success = $true
        Error = $null
        LoadedModules = @()
    }
    
    try {
        Write-Verbose "Initializing PIM modules with version pinning..."
        
        # Remove any currently loaded conflicting modules
        if ($Force) {
            Write-Verbose "Force flag specified - removing all loaded Graph and Az modules"
            Remove-ConflictingModules
        }
        
        # Validate required module availability
        foreach ($moduleSpec in $script:RequiredModuleVersions.GetEnumerator()) {
            $moduleName = $moduleSpec.Key
            $requiredVersion = $moduleSpec.Value
            
            Write-Verbose "Checking availability of $moduleName version $requiredVersion"
            
            $availableModule = Get-Module -Name $moduleName -ListAvailable | 
                Where-Object { $_.Version -eq $requiredVersion } |
                Select-Object -First 1
            
            if (-not $availableModule) {
                $errorMsg = "Required module $moduleName version $requiredVersion is not installed. Please run: Install-Module -Name $moduleName -RequiredVersion $requiredVersion -Force"
                Write-Error $errorMsg
                $result.Success = $false
                $result.Error = $errorMsg
                return $result
            }
        }
        
        # Initialize module loading state tracking
        if (-not $script:ModuleLoadingState) {
            $script:ModuleLoadingState = @{}
        }
        
        Write-Verbose "All required modules are available. Modules will be loaded just-in-time."
        $result.Success = $true
        
    } catch {
        $result.Success = $false
        $result.Error = $_.Exception.Message
        Write-Error "Failed to initialize PIM modules: $($_.Exception.Message)"
    }
    
    return $result
}

function Import-PIMModule {
    <#
    .SYNOPSIS
        Imports a specific PIM module with version checking and conflict removal
    
    .DESCRIPTION
        Just-in-time module loading function that ensures only the correct version
        of a module is loaded and removes any conflicting versions from the session.
    
    .PARAMETER ModuleName
        Name of the module to import
        
    .PARAMETER Force
        Force reimport even if already loaded
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement', 'Microsoft.Graph.Identity.Governance', 'Az.Accounts')]
        [string]$ModuleName,
        
        [switch]$Force
    )
    
    # Check if already loaded correctly
    if (-not $Force -and $script:ModuleLoadingState[$ModuleName] -eq 'Loaded') {
        Write-Verbose "$ModuleName is already loaded correctly"
        return $true
    }
    
    try {
        $requiredVersion = $script:RequiredModuleVersions[$ModuleName]
        Write-Verbose "Loading $ModuleName version $requiredVersion"
        
        # Remove any currently loaded versions of this module
        $loadedModule = Get-Module -Name $ModuleName
        if ($loadedModule) {
            if ($loadedModule.Version -ne $requiredVersion) {
                Write-Verbose "Removing currently loaded version $($loadedModule.Version) of $ModuleName"
                Remove-Module -Name $ModuleName -Force
            } else {
                Write-Verbose "$ModuleName version $requiredVersion is already loaded"
                $script:ModuleLoadingState[$ModuleName] = 'Loaded'
                return $true
            }
        }
        
        # Import the specific version
        $moduleToImport = Get-Module -Name $ModuleName -ListAvailable | 
            Where-Object { $_.Version -eq $requiredVersion } |
            Select-Object -First 1
            
        if (-not $moduleToImport) {
            throw "Required version $requiredVersion of $ModuleName is not available"
        }
        
        Import-Module -ModuleInfo $moduleToImport -Force -Global
        $script:ModuleLoadingState[$ModuleName] = 'Loaded'
        
        Write-Verbose "Successfully loaded $ModuleName version $requiredVersion"
        return $true
        
    } catch {
        Write-Error "Failed to import $ModuleName`: $($_.Exception.Message)"
        $script:ModuleLoadingState[$ModuleName] = 'Failed'
        return $false
    }
}

function Remove-ConflictingModules {
    <#
    .SYNOPSIS
        Removes conflicting module versions from the current session
    
    .DESCRIPTION
        Removes all loaded Microsoft Graph and Az modules to prevent assembly conflicts
        when loading the pinned versions.
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Removing potentially conflicting modules from session..."
    
    # Remove all Microsoft Graph modules
    $graphModules = Get-Module -Name Microsoft.Graph*
    if ($graphModules) {
        Write-Verbose "Removing $($graphModules.Count) Microsoft Graph modules"
        $graphModules | Remove-Module -Force
    }
    
    # Remove Az.Accounts
    $azModule = Get-Module -Name Az.Accounts
    if ($azModule) {
        Write-Verbose "Removing Az.Accounts module"
        Remove-Module -Name Az.Accounts -Force
    }
    
    # Clear module loading state
    $script:ModuleLoadingState = @{}
}

function Test-PIMModuleCompatibility {
    <#
    .SYNOPSIS
        Tests if the current module combination is compatible
    
    .DESCRIPTION
        Performs a quick compatibility test to verify the loaded modules work together
    
    .OUTPUTS
        Boolean indicating compatibility
    #>
    [CmdletBinding()]
    param()
    
    try {
        # Ensure required modules are loaded
        $authLoaded = Import-PIMModule -ModuleName 'Microsoft.Graph.Authentication'
        if (-not $authLoaded) {
            return $false
        }
        
        # Test the problematic method signature
        try {
            # This will fail if there's a signature mismatch
            Connect-MgGraph -Identity -ErrorAction Stop 2>$null
        } catch {
            if ($_.Exception.Message -like "*Method not found*AuthenticateAsync*") {
                Write-Warning "Module compatibility issue detected: AuthenticateAsync method signature mismatch"
                return $false
            } elseif ($_.Exception.Message -like "*No account*" -or $_.Exception.Message -like "*identity*") {
                # Expected error - method signatures are compatible
                return $true
            } else {
                Write-Verbose "Unexpected error during compatibility test: $($_.Exception.Message)"
                return $true  # Assume compatible if it's not the signature issue
            }
        }
        
        return $true
        
    } catch {
        Write-Warning "Compatibility test failed: $($_.Exception.Message)"
        return $false
    }
}

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
        LoadedModules = @()
        LoadingState = $script:ModuleLoadingState
        Compatible = $false
    }
    
    foreach ($moduleName in $script:RequiredModuleVersions.Keys) {
        $loadedModule = Get-Module -Name $moduleName
        if ($loadedModule) {
            $status.LoadedModules += [PSCustomObject]@{
                Name = $moduleName
                LoadedVersion = $loadedModule.Version
                RequiredVersion = $script:RequiredModuleVersions[$moduleName]
                IsCorrectVersion = ($loadedModule.Version -eq $script:RequiredModuleVersions[$moduleName])
            }
        }
    }
    
    # Test compatibility if modules are loaded
    if ($status.LoadedModules.Count -gt 0) {
        $status.Compatible = Test-PIMModuleCompatibility
    }
    
    return $status
}
