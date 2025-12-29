@{
    # Script module or binary module file associated with this manifest.
    RootModule           = 'PIMActivation.psm1'
    
    # Version number of this module.
    ModuleVersion        = '2.0.0'
    
    # Supported PSEditions - Requires PowerShell Core (7+)
    CompatiblePSEditions = @('Core')
    
    # ID used to uniquely identify this module
    GUID                 = 'a3f4b8e2-9c7d-4e5f-b6a9-8d7c6b5a4f3e'
    
    # Author of this module
    Author               = 'Sebastian Flæng Markdanner'
    
    # Company or vendor of this module
    CompanyName          = 'Cloudy With a Change Of Security'
    
    # Copyright statement for this module
    Copyright            = '(c) 2025 Sebastian Flæng Markdanner. All rights reserved.'

    # Description of the functionality provided by this module
    Description          = 'PowerShell module for managing Microsoft Entra ID Privileged Identity Management (PIM) role activations through a modern GUI interface. Supports Entra ID roles, PIM-enabled groups, and Azure Resource roles. Features authentication context, bulk operations, and policy compliance. Developed with AI assistance. Requires PowerShell 7+.'
    
    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion    = '7.0'
    
    # Script to run after the module is imported
    ScriptsToProcess     = @()
    
    # Required modules - conditionally enforced based on availability
    # Auto-installation logic in PSM1 handles missing modules
    RequiredModules      = @()
    
    # Functions to export from this module
    FunctionsToExport    = @(
        'Start-PIMActivation'
    )
    
    # Cmdlets to export from this module
    CmdletsToExport      = @()
    
    # Variables to export from this module
    VariablesToExport    = @()
    
    # Aliases to export from this module
    AliasesToExport      = @()
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData          = @{
        PSData = @{
            # Tags applied to this module for online gallery discoverability
            Tags                     = @('PIM', 'PrivilegedIdentityManagement', 'EntraID', 'AzureAD', 'Azure', 'AzureResources', 'Identity', 'Governance', 'RBAC', 'GUI', 'Authentication', 'ConditionalAccess', 'Security', 'Microsoft', 'Graph')
            
            # A URL to the license for this module.
            LicenseUri               = 'https://github.com/Noble-Effeciency13/PIMActivation/blob/main/LICENSE'
            
            # A URL to the main website for this project.
            ProjectUri               = 'https://github.com/Noble-Effeciency13/PIMActivation'
            
            # A URL to an icon representing this module.
            IconUri                  = 'https://raw.githubusercontent.com/Noble-Effeciency13/PIMActivation/main/Resources/icon.png'
            
            # ReleaseNotes
            ReleaseNotes             = @'
## PIMActivation v2.0.0 - Azure Resources & Parallel Processing Engine

### 🚀 Major New Features
- **Azure Resource Roles Support**: Full integration with Azure Resource PIM for subscription, resource group, and individual resource role management
- **Parallel Processing Engine**: High-performance concurrent execution for all operations with real-time progress tracking
- **Enhanced Role Display**: Azure roles display with [Azure] prefix and portal-aligned resource/scope columns
- **Cross-Subscription Support**: Automatic enumeration and management across all accessible Azure subscriptions
- **Modular Architecture**: Split functions into individual files for better maintainability

### ⚡ Performance Features
- **Parallel Processing by Default**: Concurrent execution for Azure, Entra, and Group operations
- **Real-Time Progress Tracking**: Enhanced verbose output with emoji indicators (🚀, ✅, ❌) and timing metrics
- **Smart Throttling**: Default ThrottleLimit of 10 concurrent operations, configurable up to 50
- **Thread-Safe Operations**: ConcurrentBag and ConcurrentDictionary for safe parallel result aggregation

### ✅ Added
- Complete Azure Resource role activation and deactivation support
- Select All button for bulk role selection in GUI
- `Get-AzureResourceRoles` function with parallel subscription processing
- `Initialize-AzureResourceSupport` for Azure module management
- `DisableParallelProcessing` parameter for sequential processing when needed
- Enhanced scope parsing for Azure ARM resource hierarchies
- Support for both PIM-eligible and active Azure Resource role assignments

### 🔧 Enhanced Performance
- All v1.2.x optimizations preserved and extended:
  - ArrayList-based collections for optimal memory usage
  - Batch API operations reducing Graph calls by 85%
  - Memoized scope display name lookups
  - Intelligent role deduplication and caching
  - NEW: Parallel processing across all role types and policy operations

### 📋 Requirements
- PowerShell 7.0+ (required for parallel processing engine)
- Az.Accounts 5.1.0+ and Az.Resources 6.0.0+ (auto-installed for Azure resources)
- Microsoft Graph PowerShell modules (existing requirements preserved)

### 📚 More
- Changelog: https://github.com/Noble-Effeciency13/PIMActivation/blob/main/CHANGELOG.md
- Blog Post: https://www.chanceofsecurity.com/post/microsoft-entra-pim-bulk-role-activation-tool
- Releases:  https://github.com/Noble-Effeciency13/PIMActivation/releases

PowerShell module for comprehensive PIM role management across Entra ID, Groups, and Azure Resources with parallel processing engine and modern GUI.
'@
            # Flag to indicate whether the module requires explicit user acceptance
            RequireLicenseAcceptance = $false
        }
    }
}