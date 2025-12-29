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
        
    .PARAMETER IncludeAzureModules
        Includes Azure PowerShell modules in the initialization process
        
    .OUTPUTS
        PSCustomObject with Success and Error properties
    #>
    [CmdletBinding()]
    param(
        [switch]$Force,
        [switch]$IncludeAzureModules
    )
    
    # Define pinned module versions (working combination)
    $script:RequiredModuleVersions = @{
        'Microsoft.Graph.Authentication'               = '2.29.1'
        'Microsoft.Graph.Users'                        = '2.29.1'
        'Microsoft.Graph.Identity.DirectoryManagement' = '2.29.1'
        'Microsoft.Graph.Identity.Governance'          = '2.29.1'
        'Microsoft.Graph.Groups'                       = '2.29.1'
        'Microsoft.Graph.Identity.SignIns'             = '2.29.1'
    }
    
    # Add Azure modules if requested
    if ($IncludeAzureModules) {
        $script:RequiredModuleVersions['Az.Accounts'] = '5.1.0'
        $script:RequiredModuleVersions['Az.Resources'] = '8.1.0'
    }
    
    # All modules now use minimum version checking for better compatibility
    $script:MinimumVersionModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users', 
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Identity.Governance',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Identity.SignIns'
    )
    
    # Add Azure modules to minimum version check if requested
    if ($IncludeAzureModules) {
        $script:MinimumVersionModules += @('Az.Accounts', 'Az.Resources')
    }
    
    $result = [PSCustomObject]@{
        Success       = $true
        Error         = $null
        LoadedModules = @()
    }
    
    try {
        Write-Verbose "Initializing PIM modules with version pinning..."
        if ($IncludeAzureModules) {
            Write-Verbose "Including Azure PowerShell modules in initialization"
        }
        
        # Remove any currently loaded conflicting modules
        if ($Force) {
            Write-Verbose "Force flag specified - removing all loaded Graph and Az modules"
            Remove-ConflictingModules -IncludeAzureModules:$IncludeAzureModules
        }
        
        # Validate required module availability
        # Temporarily suppress verbose output during Get-Module operations
        $currentVerbose = $VerbosePreference
        $VerbosePreference = 'SilentlyContinue'
        
        try {
            foreach ($moduleSpec in $script:RequiredModuleVersions.GetEnumerator()) {
                $moduleName = $moduleSpec.Key
                $requiredVersion = $moduleSpec.Value
                
                # Restore verbose for our own output
                $VerbosePreference = $currentVerbose
                Write-Verbose "Checking availability of $moduleName minimum version $requiredVersion"
                $VerbosePreference = 'SilentlyContinue'
                
                # For all modules, check if we have the required version or higher
                $availableModule = Get-Module -Name $moduleName -ListAvailable | 
                Where-Object { $_.Version -ge [Version]$requiredVersion } |
                Sort-Object Version -Descending |
                Select-Object -First 1
                
                if (-not $availableModule) {
                    $VerbosePreference = $currentVerbose
                    $errorMsg = "Required module $moduleName minimum version $requiredVersion is not installed. Please run: Install-Module -Name $moduleName -MinimumVersion $requiredVersion -Force"
                    Write-Error $errorMsg
                    $result.Success = $false
                    $result.Error = $errorMsg
                    return $result
                }
            }
        }
        finally {
            # Always restore the original verbose preference
            $VerbosePreference = $currentVerbose
        }
        
        # Initialize module loading state tracking
        if (-not $script:ModuleLoadingState) {
            $script:ModuleLoadingState = @{}
        }
        
        Write-Verbose "All required modules are available. Modules will be loaded just-in-time."
        $result.Success = $true
        
    }
    catch {
        $result.Success = $false
        $result.Error = $_.Exception.Message
        Write-Error "Failed to initialize PIM modules: $($_.Exception.Message)"
    }
    
    return $result
}