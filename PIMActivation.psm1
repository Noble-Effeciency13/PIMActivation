#Requires -Version 7.0

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

# Authentication context variables - now supporting multiple contexts
$script:CurrentAuthContextToken = $null  # Deprecated - kept for backwards compatibility
$script:AuthContextTokens = @{}  # New: Hashtable of contextId -> token
$script:JustCompletedAuthContext = $null
$script:AuthContextCompletionTime = $null

# Module loading state for just-in-time loading
$script:ModuleLoadingState = @{}
$script:RequiredModuleVersions = @{
    'Microsoft.Graph.Authentication' = '2.29.1'
    'Microsoft.Graph.Users' = '2.29.1'
    'Microsoft.Graph.Identity.DirectoryManagement' = '2.29.1'
    'Microsoft.Graph.Identity.Governance' = '2.29.1'
    'Az.Accounts' = '5.1.0'
}

#endregion Module Setup

#region Import Functions

# Import all functions from subdirectories
$functionFolders = @(
    'Authentication',
    'RoleManagement', 
    'UI',
    'Utilities'
)

# Note: Profiles folder contains placeholder functions for planned features
$functionFolders += 'Profiles'

# Import private functions from organized folders
foreach ($folder in $functionFolders) {
    $folderPath = Join-Path -Path "$script:ModuleRoot\Private" -ChildPath $folder
    if (Test-Path -Path $folderPath) {
        Write-Verbose "Importing functions from $folder"
        $functions = Get-ChildItem -Path $folderPath -Filter '*.ps1' -File -ErrorAction SilentlyContinue
        
        foreach ($function in $functions) {
            try {
                Write-Verbose "Importing $($function.Name)"
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
        Write-Verbose "Importing $($import.Name)"
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
        Write-Verbose "Importing public function $($import.Name)"
        . $import.FullName
    }
    catch {
        Write-Error -Message "Failed to import function $($import.FullName): $_"
    }
}

#endregion Import Functions

#region Export Module Members

# Export public functions
Export-ModuleMember -Function $Public.BaseName -Alias *

#endregion Export Module Members

#region Cleanup

# Clean up variables
Remove-Variable -Name Private, Public, functionFolders, folder, folderPath, functions, function, privateRoot, import -ErrorAction SilentlyContinue

#endregion Cleanup