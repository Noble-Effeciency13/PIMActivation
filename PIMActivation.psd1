@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'PIMActivation.psm1'
    
    # Version number of this module.
    ModuleVersion = '1.0.0'
    
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
    Description = 'PowerShell module for managing Microsoft Entra ID Privileged Identity Management (PIM) role activations through a modern GUI interface. Supports authentication context, bulk operations, and policy compliance. Requires PowerShell 7+.'
    
    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '7.0'
    
    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0' },
        @{ ModuleName = 'Microsoft.Graph.Users'; ModuleVersion = '2.0.0' },
        @{ ModuleName = 'Microsoft.Graph.Identity.DirectoryManagement'; ModuleVersion = '2.0.0' },
        @{ ModuleName = 'Microsoft.Graph.Identity.Governance'; ModuleVersion = '2.0.0' },
        @{ ModuleName = 'MSAL.PS'; ModuleVersion = '4.37.0' }
    )
    
    # Functions to export from this module
    FunctionsToExport = @('Start-PIMActivation')
    
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
## Release Notes v1.0.0

### 🎉 Initial Release
- **Modern GUI Interface**: Clean Windows Forms application for PIM role management
- **Multi-Role Support**: Activate Microsoft Entra ID roles and PIM-enabled security groups
- **Authentication Context**: Seamless handling of Conditional Access authentication context policies
- **Bulk Operations**: Select and activate multiple roles simultaneously with policy validation
- **PowerShell Compatibility**: Requires PowerShell 7+ for optimal performance and modern language features
- **Policy Compliance**: Automatic detection of MFA, justification, and ticket requirements
- **Real-time Updates**: Live monitoring of active assignments and pending requests

### 🔧 Technical Features
- Hybrid PowerShell architecture for optimal MSAL.PS compatibility
- Direct REST API calls for authentication context preservation  
- Automatic module dependency management
- Comprehensive error handling and user feedback

### 📋 Requirements
- Windows Operating System
- PowerShell 7+ (Download from https://aka.ms/powershell)
- Microsoft Graph PowerShell modules (auto-installed)
- Appropriate Entra ID permissions for PIM role management

For detailed usage instructions, see the README.md file.
'@
            # Flag to indicate whether the module requires explicit user acceptance
            RequireLicenseAcceptance = $false
            
            # External module dependencies that are not captured by RequiredModules
            ExternalModuleDependencies = @()
        }
    }
}