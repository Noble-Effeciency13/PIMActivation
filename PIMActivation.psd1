@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'PIMActivation.psm1'
    
    # Version number of this module.
    ModuleVersion = '1.2.1'
    
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
    
    # Modules are dynamically installed and imported by the module's initialization logic
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
## PIMActivation v1.2.1

### � What's New
- Enhanced automatic dependency resolution
- Added -Force parameter for fully automated setup
- Cleaner console output with reduced verbose noise
- Improved error handling and user guidance

### � Full Release Notes
For complete release notes, changelog, and detailed information:
- **GitHub Releases**: https://github.com/Noble-Effeciency13/PIMActivation/releases
- **Changelog**: https://github.com/Noble-Effeciency13/PIMActivation/blob/main/CHANGELOG.md
- **Documentation**: https://github.com/Noble-Effeciency13/PIMActivation/blob/main/README.md

### � Getting Started
```powershell
Install-Module PIMActivation -Scope CurrentUser
Start-PIMActivation
```

PowerShell module for Microsoft Entra ID Privileged Identity Management (PIM) role activations through a modern GUI interface.
'@
            # Flag to indicate whether the module requires explicit user acceptance
            RequireLicenseAcceptance = $false
            
            # External module dependencies that are not captured by RequiredModules
            ExternalModuleDependencies = @()
        }
    }
}