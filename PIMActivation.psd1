@{
    # Script module or binary module file associated with this manifest.
    RootModule           = 'PIMActivation.psm1'
    
    # Version number of this module.
    ModuleVersion        = '2.1.0'
    
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
## PIMActivation v2.1.0 - Patch & Enhancements

### ✅ Enhancements
- Management group display names: management-group scopes are now shown with their friendly display name (or `/` for tenant root) instead of raw MG IDs.
- Inherited eligible role suppression: subscription-scoped inherited eligible roles are suppressed when the same role is available at the management-group level to avoid duplicate activation entries.
- Temporary activation detection: initial tenant-root and management-group active assignments are enriched with PIM activation schedule Start/End windows so temporarily activated roles show expiry rather than appearing permanently active.
- Role definition normalization: role definition identifiers are normalized (GUID) during deduplication to eliminate duplicates caused by full-path vs GUID variants.
- Import-time PSGallery notification: on import the module performs a best-effort check against the PowerShell Gallery and warns when a newer release is available. The notification follows Microsoft module style and provides Update-Module / Install-Module examples. This check can be suppressed via `$script:SuppressUpdateNotification`.

### 🛠️ Fixes (Community Contribution)
- Activation/Deactivation Scope and Safety: Added explicit `Scope` support when activating and deactivating Azure PIM roles and improved error handling to prevent attempting to deactivate a role that was activated less than the required 5-minute window. (Thanks to Lukas Gosling (@l-gosling) for this contribution.)

### ⚡ Notes
- These changes are additive and preserve existing public APIs. They improve display fidelity and de-duplication for Azure resource roles and make temporary activations visible as such in the UI.

PowerShell module for comprehensive PIM role management across Entra ID, Groups, and Azure Resources with parallel processing engine and modern GUI.
'@
            # Flag to indicate whether the module requires explicit user acceptance
            RequireLicenseAcceptance = $false
        }
    }
}