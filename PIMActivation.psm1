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
$script:AuthContextTokens = @{} 
$script:JustCompletedAuthContext = $null
$script:AuthContextCompletionTime = $null

# Module loading state for just-in-time loading
$script:ModuleLoadingState = @{}
$script:RequiredModuleVersions = @{
    'Microsoft.Graph.Authentication'               = '2.29.0'
    'Microsoft.Graph.Users'                        = '2.29.0'
    'Microsoft.Graph.Identity.DirectoryManagement' = '2.29.0'
    'Microsoft.Graph.Identity.Governance'          = '2.29.0'
    'Microsoft.Graph.Groups'                       = '2.29.0'
    'Microsoft.Graph.Identity.SignIns'             = '2.29.0'
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

# Import all private functions from any subfolder
$privateRoot = Join-Path $script:ModuleRoot 'Private'
if (Test-Path -Path $privateRoot) {
    $privateFiles = Get-ChildItem -Path $privateRoot -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue | Where-Object { $_ -is [System.IO.FileInfo] }
    foreach ($file in $privateFiles) {
        try {
            . $file.FullName
        }
        catch {
            Write-Error -Message "Failed to import function $($file.FullName): $_"
        }
    }
}

# Import all public functions from any subfolder (if you ever nest them)
$publicRoot = Join-Path $script:ModuleRoot 'Public'
$Public = @()
if (Test-Path -Path $publicRoot) {
    $Public = Get-ChildItem -Path $publicRoot -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue | Where-Object { $_ -is [System.IO.FileInfo] }
    foreach ($import in $Public) {
        try {
            . $import.FullName
        }
        catch {
            Write-Error -Message "Failed to import function $($import.FullName): $_"
        }
    }
}

# Restore original verbose preference
$VerbosePreference = $originalVerbosePreference

#endregion Import Functions

#region Export Module Members

# Export public functions by filename
if ($Public) {
    $publicFiles = @($Public)
    $functionNames = $publicFiles.BaseName | Sort-Object -Unique
    if ($functionNames) {
        Export-ModuleMember -Function $functionNames -Alias *
    }
}

#endregion Export Module Members

#region Module Initialization

# Smart dependency resolution - handles both development and production scenarios
# This allows the module to work regardless of how it's imported
$script:DependenciesValidated = $false

Write-Verbose "PIMActivation module loaded. Use Start-PIMActivation to begin."
Write-Verbose "Dependencies will be validated and installed automatically when needed."

#endregion Module Initialization

#region Cleanup

# Clean up variables
Remove-Variable -Name Private, Public, functionFolders, folder, folderPath, functions, function, privateRoot, import -ErrorAction SilentlyContinue

#endregion Cleanup