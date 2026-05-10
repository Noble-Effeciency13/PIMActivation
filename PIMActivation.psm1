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

# UI re-entrancy guard for the Azure Resources checkbox toggle
$script:_suppressAzureToggle = $false

# Startup parameters (for restarts)
$script:StartupParameters = @{}

# Control runtime update-notification behavior. Can be set by consumers to suppress the PSGallery check.
$script:SuppressUpdateNotification = $false

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

# Eagerly import all required Graph modules at module-load time so that the first call
# to Start-PIMActivation does not have to pay the module-import cost.
# Az modules are NOT loaded here because they are optional / conditional at runtime.
$script:DependenciesValidated = $false

try {
    # Initialize-PIMModules sets $script:RequiredModuleVersions and validates availability
    $script:_preloadResult = Initialize-PIMModules
    if ($script:_preloadResult.Success) {
        # Import Graph modules in dependency order
        $script:_graphModuleOrder = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Identity.Governance',
            'Microsoft.Graph.Identity.SignIns',
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Users'
        )
        $allLoaded = $true
        Write-Host "PIMActivation: loading dependencies..." -ForegroundColor Cyan
        foreach ($mod in $script:_graphModuleOrder) {
            $loadedOk = Import-PIMModule -ModuleName $mod
            if ($loadedOk) {
                $loadedModule = Get-Module -Name $mod -ErrorAction SilentlyContinue
                $versionStr   = if ($loadedModule) { " v$($loadedModule.Version)" } else { '' }
                Write-Host "  [√] $mod$versionStr" -ForegroundColor Green
            }
            else {
                $allLoaded = $false
                Write-Host "  [!] $mod – failed to load, will retry on Start-PIMActivation" -ForegroundColor Yellow
            }
        }
        $script:DependenciesValidated = $allLoaded
        if ($allLoaded) {
            Write-Host "PIMActivation: all dependencies loaded. Run Start-PIMActivation to begin." -ForegroundColor Cyan
        }
    }
    else {
        # Modules not yet installed; Start-PIMActivation will install them on first call
        Write-Warning "PIMActivation: required modules are not installed ($($script:_preloadResult.Error)). Run Start-PIMActivation to install them automatically."
    }
    Remove-Variable -Name _preloadResult, _graphModuleOrder -Scope Script -ErrorAction SilentlyContinue
}
catch {
    # Non-fatal – Start-PIMActivation will retry dependency resolution
    Write-Verbose "PIMActivation: module pre-load error: $($_.Exception.Message). Dependencies will be resolved on first run."
}

# Check PowerShell Gallery for newer module version and notify on import (best-effort)
try {
    if (-not $script:SuppressUpdateNotification) {
        if (Get-Command -Name Find-Module -Module PowerShellGet -ErrorAction SilentlyContinue) {
            try {
                $remote = Find-Module -Name $script:ModuleName -Repository PSGallery -ErrorAction SilentlyContinue
                if ($remote -and $remote.Version) {
                    $localVersion = $null
                    try {
                        $manifest = Join-Path $script:ModuleRoot "$($script:ModuleName).psd1"
                        if (Test-Path $manifest) {
                            $md = Import-PowerShellDataFile -Path $manifest
                            $localVersion = [Version]($md.ModuleVersion)
                        }
                    } catch {}

                    if (-not $localVersion) {
                        # Fallback to module manifest discovery
                        $localVersion = [Version]('0.0.0')
                    }

                    $remoteVersion = [Version]($remote.Version.ToString())
                    if ($remoteVersion -gt $localVersion) {
                        $psgalleryUrl = "https://www.powershellgallery.com/packages/$($script:ModuleName)/$($remoteVersion)"
                        $msg = @()
                        $msg += "A newer version of the module '$($script:ModuleName)' is available on the PowerShell Gallery."
                        $msg += "Installed version: $localVersion"
                        $msg += "Latest version:    $remoteVersion"
                        $msg += "To update this module, run:     Update-Module -Name $($script:ModuleName) -Force"
                        $msg += "If Update-Module is unavailable, you can install the latest version with: Install-Module -Name $($script:ModuleName) -Force"
                        $msg += "More information: $psgalleryUrl"
                        Write-Warning ($msg -join "`n")
                    }
                }
            } catch {
                Write-Verbose "PSGallery version check failed: $($_.Exception.Message)"
            }
        }
    }
} catch {
    Write-Verbose "Update notification check encountered an error: $($_.Exception.Message)"
}
#endregion Module Initialization

#region Cleanup

# Clean up variables
Remove-Variable -Name Private, Public, functionFolders, folder, folderPath, functions, function, privateRoot, import -ErrorAction SilentlyContinue


#endregion Cleanup
