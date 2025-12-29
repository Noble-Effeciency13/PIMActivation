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
        [ValidateSet(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Identity.Governance',
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.SignIns',
            'Az.Accounts',
            'Az.Resources'
        )]
        [string]$ModuleName,
        [switch]$Force
    )

    if (-not $Force -and $script:ModuleLoadingState[$ModuleName] -eq 'Loaded') {
        Write-Verbose "$ModuleName is already loaded correctly"
        return $true
    }

    try {
        $requiredVersionString = $script:RequiredModuleVersions[$ModuleName]
        $minVersion = $null
        if ($requiredVersionString) {
            # RequiredModuleVersions holds minimums for Azure when IncludeAzureModules was used
            $minVersion = [Version]$requiredVersionString
        }

        Write-Verbose ("Preparing to load {0}{1}" -f $ModuleName, $(if ($minVersion) { " (minimum $minVersion)" } else { "" }))

        $currentVerbose = $VerbosePreference
        $VerbosePreference = 'SilentlyContinue'
        try {
            # Remove currently loaded module if below minimum
            $loaded = Get-Module -Name $ModuleName
            if ($loaded) {
                if ($minVersion -and $loaded.Version -lt $minVersion) {
                    $VerbosePreference = $currentVerbose
                    Write-Verbose "Removing $ModuleName v$($loaded.Version) (below minimum $minVersion)"
                    $VerbosePreference = 'SilentlyContinue'
                    Remove-Module -Name $ModuleName -Force
                }
                else {
                    $VerbosePreference = $currentVerbose
                    Write-Verbose "$ModuleName v$($loaded.Version) already loaded$(if ($minVersion) { " (meets minimum $minVersion)" })"
                    $script:ModuleLoadingState[$ModuleName] = 'Loaded'
                    return $true
                }
            }

            # Select best available version
            $available = Get-Module -Name $ModuleName -ListAvailable |
                Where-Object { -not $minVersion -or $_.Version -ge $minVersion } |
                Sort-Object Version -Descending |
                Select-Object -First 1

            if (-not $available) {
                $VerbosePreference = $currentVerbose
                if ($minVersion) {
                    Write-Error "No available $ModuleName meets minimum version $minVersion"
                }
                else {
                    Write-Error "$ModuleName is not installed or not discoverable (Get-Module -ListAvailable returned nothing)"
                }
                $script:ModuleLoadingState[$ModuleName] = 'Failed'
                return $false
            }

            # Import the selected version
            $VerbosePreference = $currentVerbose
            Write-Verbose "Importing $ModuleName v$($available.Version)"
            $VerbosePreference = 'SilentlyContinue'
            Import-Module -ModuleInfo $available -Force -Global
        }
        finally {
            $VerbosePreference = $currentVerbose
        }

        $script:ModuleLoadingState[$ModuleName] = 'Loaded'
        $actual = Get-Module -Name $ModuleName
        Write-Verbose "Successfully loaded $ModuleName v$($actual.Version)"
        return $true
    }
    catch {
        Write-Error "Failed to import ${ModuleName}: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Error "Inner: $($_.Exception.InnerException.Message)"
        }
        $script:ModuleLoadingState[$ModuleName] = 'Failed'
        return $false
    }
}