#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.29.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Users'; ModuleVersion = '2.29.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Identity.DirectoryManagement'; ModuleVersion = '2.29.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Identity.Governance'; ModuleVersion = '2.29.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Groups'; ModuleVersion = '2.29.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Identity.SignIns'; ModuleVersion = '2.29.0' }
#Requires -Modules @{ ModuleName = 'Az.Accounts'; ModuleVersion = '5.1.0' }

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

# Dependencies are loaded on-demand when Start-PIMActivation is called
# This ensures clean module loading and avoids import-time dependency issues
# All dependency management is handled automatically by Start-PIMActivation

#endregion Module Initialization

#region Cleanup

# Clean up variables
Remove-Variable -Name Private, Public, functionFolders, folder, folderPath, functions, function, privateRoot, import -ErrorAction SilentlyContinue

#endregion Cleanup