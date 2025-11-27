@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'PIMActivation.psm1'
    
    # Version number of this module.
    ModuleVersion = '1.2.6'
    
    # Supported PSEditions - Requires PowerShell Core (7+)
    CompatiblePSEditions = @('Core')
    
    # ID used to uniquely identify this module
    GUID = 'a3f4b8e2-9c7d-4e5f-b6a9-8d7c6b5a4f3e'
    
    # Author of this module
    Author = 'Sebastian Flæng Markdanner'
    
    # Company or vendor of this module
    CompanyName = 'Cloudy With a Change Of Security'
    
    # Copyright statement for this module
    Copyright = '(c) 2025 Sebastian Flæng Markdanner. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'PowerShell module for managing Microsoft Entra ID Privileged Identity Management (PIM) role activations through a modern GUI interface. Supports authentication context, bulk operations, and policy compliance. Developed with AI assistance. Requires PowerShell 7+.'
    
    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '7.0'
    
    # Script to run after the module is imported
    ScriptsToProcess = @()
    
    # Required modules - conditionally enforced based on availability
    # Auto-installation logic in PSM1 handles missing modules
    RequiredModules = @()
    
    # Functions to export from this module
    FunctionsToExport = @(
        'Start-PIMActivation'
    )
    
    # Cmdlets to export from this module
    CmdletsToExport = @()
    
    # Variables to export from this module
    VariablesToExport = @()
    
    # Aliases to export from this module
    AliasesToExport = @()
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData = @{
        PSData = @{
            # Tags applied to this module for online gallery discoverability
            Tags = @('PIM', 'PrivilegedIdentityManagement', 'EntraID', 'AzureAD', 'Identity', 'Governance', 'RBAC', 'GUI', 'Authentication', 'ConditionalAccess', 'Security', 'Microsoft', 'Graph')
            
            # A URL to the license for this module.
            LicenseUri = 'https://github.com/Noble-Effeciency13/PIMActivation/blob/main/LICENSE'
            
            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/Noble-Effeciency13/PIMActivation'
            
            # A URL to an icon representing this module.
            IconUri = 'https://raw.githubusercontent.com/Noble-Effeciency13/PIMActivation/main/Resources/icon.png'
            
            # ReleaseNotes
            ReleaseNotes = @'
## PIMActivation v1.2.6

### ✅ Added
- Support for custom app registration for Microsoft Graph delegated auth. New parameters `ClientId` and `TenantId` are available on `Start-PIMActivation` and `Connect-PIMServices`. When both are provided, the module authenticates using the specified app registration; otherwise, it falls back to the default interactive flow.

### 🔧 Fixes
- Resolved Microsoft Graph query limitations when collecting role policies for large sets (e.g., >20 eligible roles of the same type). Implemented chunked batching and a REST-based path with pagination so policies are fetched reliably at scale.
- Added robust fallback to per-item fetching when the service rejects complex filters or returns zero results.
- Corrected control flow and ensured `-ErrorAction Stop` on policy assignment calls so fallbacks always trigger when needed.
- Addressed a transient InvalidResource/InvalidFilter regression introduced during the fix and removed it.

### ⚡ Improvements
- Performance: Replaced array concatenations with `ArrayList`/`AddRange` in hot paths (role collection and batch aggregations).
- Stability: Flattened `ArrayList` before mapping policies; treat `InvalidResource` like `InvalidFilter` for resilient behavior.
- Caching: Memoized scope and AU display name lookups in `Get-ScopeDisplayName` to reduce repeated Graph calls.

### 📚 More
- Changelog: https://github.com/Noble-Effeciency13/PIMActivation/blob/main/CHANGELOG.md
- Releases:  https://github.com/Noble-Effeciency13/PIMActivation/releases

PowerShell module for Microsoft Entra ID PIM role activations with a modern GUI. Requires PowerShell 7+.
'@
            # Flag to indicate whether the module requires explicit user acceptance
            RequireLicenseAcceptance = $false
        }
    }
}