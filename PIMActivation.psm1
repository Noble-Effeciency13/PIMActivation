#Requires -Version 7.0
# Note: Required modules are declared in the manifest and handled by internal dependency management
# This allows for both PowerShell Gallery automatic installation and development scenarios

# Set strict mode for better error handling
Set-StrictMode -Version Latest

#region Module Setup

# Module-level variables
$script:ModuleRoot = $PSScriptRoot
$script:ModuleName = Split-Path -Path $script:ModuleRoot -Leaf

# Token storage variables
$script:CurrentAccessToken = $null
$script:TokenExpiry = $null

# User context variables
$script:CurrentUser = $null
$script:GraphContext = $null

# Configuration variables
$script:IncludeEntraRoles = $true
$script:IncludeGroups = $true
$script:IncludeAzureResources = $false

# Startup parameters (for restarts)
$script:StartupParameters = @{}

# Restart flag
$script:RestartRequested = $false

# Policy cache
if (-not (Test-Path Variable:script:PolicyCache)) {
    $script:PolicyCache = @{}
}

# Authentication context cache
if (-not (Test-Path Variable:script:AuthenticationContextCache)) {
    $script:AuthenticationContextCache = @{}
}

# Entra policies loaded flag
if (-not (Test-Path Variable:script:EntraPoliciesLoaded)) {
    $script:EntraPoliciesLoaded = $false
}

# Role data cache to avoid repeated API calls during refresh operations
if (-not (Test-Path Variable:script:CachedEligibleRoles)) {
    $script:CachedEligibleRoles = @()
}

if (-not (Test-Path Variable:script:CachedActiveRoles)) {
    $script:CachedActiveRoles = @()
}

if (-not (Test-Path Variable:script:LastRoleFetchTime)) {
    $script:LastRoleFetchTime = $null
}

if (-not (Test-Path Variable:script:RoleCacheValidityMinutes)) {
    $script:RoleCacheValidityMinutes = 5  # Cache roles for 5 minutes
}

# Authentication context variables - now supporting multiple contexts
$script:CurrentAuthContextToken = $null  # Deprecated - kept for backwards compatibility
$script:AuthContextTokens = @{}  # New: Hashtable of contextId -> token
$script:JustCompletedAuthContext = $null
$script:AuthContextCompletionTime = $null

# Module loading state for just-in-time loading
$script:ModuleLoadingState = @{}
$script:RequiredModuleVersions = @{
    'Microsoft.Graph.Authentication' = '2.29.0'
    'Microsoft.Graph.Users' = '2.29.0'
    'Microsoft.Graph.Identity.DirectoryManagement' = '2.29.0'
    'Microsoft.Graph.Identity.Governance' = '2.29.0'
    'Microsoft.Graph.Groups' = '2.29.0'
    'Microsoft.Graph.Identity.SignIns' = '2.29.0'
    'Az.Accounts' = '5.1.0'
}

#endregion Module Setup

#region Import Functions

# Import all functions from subdirectories
$functionFolders = [System.Collections.ArrayList]::new()
$null = $functionFolders.AddRange(@(
    'Authentication',
    'RoleManagement', 
    'UI',
    'Utilities'
))

# Note: Profiles folder contains placeholder functions for planned features
$null = $functionFolders.Add('Profiles')

# Import private functions from organized folders
# Temporarily suppress verbose output during function imports to reduce noise
$originalVerbosePreference = $VerbosePreference
$VerbosePreference = 'SilentlyContinue'

foreach ($folder in $functionFolders) {
    $folderPath = Join-Path -Path "$script:ModuleRoot\Private" -ChildPath $folder
    if (Test-Path -Path $folderPath) {
        $functions = Get-ChildItem -Path $folderPath -Filter '*.ps1' -File -ErrorAction SilentlyContinue
        
        foreach ($function in $functions) {
            try {
                . $function.FullName
            }
            catch {
                Write-Error -Message "Failed to import function $($function.FullName): $_"
            }
        }
    }
}

# Import remaining private functions from root Private folder
$privateRoot = Get-ChildItem -Path "$script:ModuleRoot\Private" -Filter '*.ps1' -File -ErrorAction SilentlyContinue
foreach ($import in $privateRoot) {
    try {
        . $import.FullName
    }
    catch {
        Write-Error -Message "Failed to import function $($import.FullName): $_"
    }
}

# Import public functions
$Public = @(Get-ChildItem -Path "$script:ModuleRoot\Public" -Filter '*.ps1' -File -ErrorAction SilentlyContinue)
foreach ($import in $Public) {
    try {
        . $import.FullName
    }
    catch {
        Write-Error -Message "Failed to import function $($import.FullName): $_"
    }
}

# Restore original verbose preference
$VerbosePreference = $originalVerbosePreference

#endregion Import Functions

#region Export Module Members

# Export public functions
if ($Public -and $Public.Count -gt 0) {
    Export-ModuleMember -Function $Public.BaseName -Alias *
}

#endregion Export Module Members

#region Module Initialization

# Smart dependency resolution - handles both development and production scenarios
# This allows the module to work regardless of how it's imported
$script:DependenciesValidated = $false

function Install-MissingPIMModules {
    <#
    .SYNOPSIS
        Automatically installs missing required modules during import
    #>
    [CmdletBinding()]
    param()
    
    $missingModules = [System.Collections.ArrayList]::new()
    
    # Check each required module
    foreach ($moduleSpec in $script:RequiredModuleVersions.GetEnumerator()) {
        $moduleName = $moduleSpec.Key
        $requiredVersion = [version]$moduleSpec.Value
        
        # Check if suitable version is available
        $availableModule = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | 
            Where-Object { $_.Version -ge $requiredVersion } | 
            Sort-Object Version -Descending | 
            Select-Object -First 1
            
        if (-not $availableModule) {
            $null = $missingModules.Add(@{
                Name = $moduleName
                RequiredVersion = $requiredVersion
            })
        }
    }
    
    # Install missing modules if any
    if ($missingModules.Count -gt 0) {
        Write-Verbose "Installing missing dependencies..."
        
        # Ensure NuGet provider and PSGallery trust (silent setup)
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nugetProvider -or $nugetProvider.Version -lt '2.8.5.201') {
            Write-Verbose "Installing NuGet provider..."
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
        }
        
        $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($psGallery.InstallationPolicy -ne 'Trusted') {
            Write-Verbose "Configuring PSGallery as trusted..."
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }
        
        # Install each missing module
        foreach ($module in $missingModules) {
            try {
                Write-Verbose "Installing $($module.Name) v$($module.RequiredVersion)..."
                $originalInformationPreference = $InformationPreference
                $originalProgressPreference = $ProgressPreference
                $InformationPreference = 'SilentlyContinue'
                $ProgressPreference = 'SilentlyContinue'
                
                Install-Module -Name $module.Name -MinimumVersion $module.RequiredVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop | Out-Null
                
                $InformationPreference = $originalInformationPreference
                $ProgressPreference = $originalProgressPreference
                Write-Verbose "Successfully installed $($module.Name)"
            }
            catch {
                $InformationPreference = $originalInformationPreference
                $ProgressPreference = $originalProgressPreference
                Write-Warning "Failed to install $($module.Name): $($_.Exception.Message)"
            }
        }
        
        Write-Verbose "Dependencies installation completed."
    }
}

function Test-PIMModuleDependencies {
    <#
    .SYNOPSIS
        Internal function to validate and import required module dependencies
    #>
    [CmdletBinding()]
    param()
    
    if ($script:DependenciesValidated) {
        return $true
    }
    
    # First, try to install any missing modules
    try {
        Install-MissingPIMModules
    }
    catch {
        Write-Verbose "Module installation failed: $($_.Exception.Message)"
    }
    
    # Now validate and import modules
    $failedModules = [System.Collections.ArrayList]::new()
    
    foreach ($moduleSpec in $script:RequiredModuleVersions.GetEnumerator()) {
        $moduleName = $moduleSpec.Key
        $requiredVersion = [version]$moduleSpec.Value
        
        $loadedModule = Get-Module -Name $moduleName -ErrorAction SilentlyContinue
        if (-not $loadedModule) {
            $availableModule = Get-Module -ListAvailable -Name $moduleName | 
                Where-Object { $_.Version -ge $requiredVersion } | 
                Sort-Object Version -Descending | 
                Select-Object -First 1
                
            if ($availableModule) {
                try {
                    Import-Module -Name $moduleName -MinimumVersion $requiredVersion -ErrorAction Stop -Force
                    Write-Verbose "Imported $moduleName v$($availableModule.Version)"
                }
                catch {
                    $null = $failedModules.Add("$moduleName (import failed: $($_.Exception.Message))")
                }
            }
            else {
                $null = $failedModules.Add("$moduleName v$requiredVersion+ (not available)")
            }
        }
        elseif ($loadedModule.Version -lt $requiredVersion) {
            $null = $failedModules.Add("$moduleName (loaded: v$($loadedModule.Version), required: v$requiredVersion+)")
        }
    }
    
    if ($failedModules.Count -gt 0) {
        $errorMessage = @"
Required module dependencies could not be resolved:
$($failedModules | ForEach-Object { "  - $_" } | Out-String)
Try running: Start-PIMActivation -Force
"@
        Write-Warning $errorMessage
        return $false
    }
    
    $script:DependenciesValidated = $true
    return $true
}

# Attempt to resolve dependencies during module import (with error handling)
try {
    $null = Test-PIMModuleDependencies
    Write-Verbose "PIMActivation module dependencies resolved successfully"
}
catch {
    # Don't fail the import, just warn
    Write-Warning "Dependency resolution during import encountered issues: $($_.Exception.Message)"
    Write-Host "You can resolve this by running: Start-PIMActivation" -ForegroundColor Yellow
}

#endregion Module Initialization

#region Cleanup

# Clean up variables
Remove-Variable -Name Private, Public, functionFolders, folder, folderPath, functions, function, privateRoot, import -ErrorAction SilentlyContinue

#endregion Cleanup