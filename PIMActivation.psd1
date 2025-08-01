@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'PIMActivation.psm1'
    
    # Version number of this module.
    ModuleVersion = '1.2.0'
    
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
    
    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        @{ ModuleName = 'Microsoft.Graph.Authentication'; RequiredVersion = '2.29.1' },
        @{ ModuleName = 'Microsoft.Graph.Users'; RequiredVersion = '2.29.1' },
        @{ ModuleName = 'Microsoft.Graph.Identity.DirectoryManagement'; RequiredVersion = '2.29.1' },
        @{ ModuleName = 'Microsoft.Graph.Identity.Governance'; RequiredVersion = '2.29.1' },
        @{ ModuleName = 'Az.Accounts'; RequiredVersion = '5.1.0' }
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
## Release Notes v1.2.0

### 🚀 Major Performance Enhancements
- **Batch API Operations**: Complete rewrite of role fetching logic using batch operations (85% reduction in API calls)
- **Intelligent Duplicate Role Handling**: Advanced algorithm for managing multiple instances of same role with proper group attribution
- **Enhanced Group-Role Attribution**: Sophisticated cross-referencing system showing which groups provide which roles
- **Comprehensive Error Handling**: Bulletproof property access protection preventing common PowerShell errors

### 🎯 UI/UX Improvements
- **Smooth Progress Flow**: Coordinated progress tracking across all loading phases (no more backwards jumps)
- **Group Visibility**: ProvidedRoles functionality shows exactly which roles each group membership provides
- **Proper Expiration Attribution**: Duplicate roles now show individual expiration times based on their providing groups
- **Enhanced Resource Display**: Shows "Entra ID (via Group: GroupName)" for group-derived roles

### 🔧 Technical Improvements
- **Advanced Array Handling**: @() wrapper implementation preventing .Count property errors
- **Safe Property Access**: PSObject.Properties pattern for bulletproof property checking
- **Intelligent Caching**: Enhanced cache invalidation system with proper timing
- **Defensive Coding**: Comprehensive try-catch blocks around all critical operations

### 🔍 Debugging & Logging
- **Enhanced Verbose Logging**: Detailed progress tracking with differentiated handling for groups vs Entra roles
- **Sophisticated Matching Logic**: Priority-based group assignment with temporal vs permanent preferences
- **Cross-Reference Validation**: Extensive debugging for group-role relationship verification

## Release Notes v1.1.1

### Added
- **Just-in-Time Module Loading**: New `Initialize-PIMModules` system that loads modules only when needed
- **Version Pinning**: Exact module version enforcement to prevent compatibility issues
- **Assembly Conflict Prevention**: Automatic removal of conflicting module versions from session
- Module loading state tracking and compatibility validation

### Changed
- **Updated Module Versions**: Now uses Microsoft.Graph 2.29.1 + Az.Accounts 5.1.0 (tested working combination)
- Replaced legacy `Install-RequiredModules` with new `Initialize-PIMModules` function
- Improved module initialization in `Start-PIMActivation` function
- Updated CI/CD workflow to use latest compatible module versions

### Removed
- **Scripts Folder**: Removed compatibility testing tools (no longer needed with version pinning)
- Legacy module installation and validation code
- Outdated module version requirements

### Fixed
- Resolved `AuthenticateAsync` method signature compatibility issues
- Improved module loading reliability and error handling
- Enhanced troubleshooting guidance for version conflicts

## Release Notes v1.1.0

### ⚡ Major Improvements
- **WAM Authentication**: Implemented Windows Web Account Manager (WAM) for reliable authentication
- **Removed MSAL.PS Dependency**: Now uses direct MSAL.NET calls for better reliability and performance
- **Enhanced Authentication Context**: Improved handling of conditional access policies

### 🔧 Technical Changes
- Direct integration with Az.Accounts MSAL assemblies
- Eliminated PowerShell 5.1 fallback - now fully PowerShell 7+ native
- Improved error handling and timeout management
- Better assembly loading and management

## Release Notes v1.0.1

### 🔧 Bug Fixes
- Fixed authentication context token acquisition for conditional access policies
- Enhanced error handling for authentication scenarios
- Improved MSAL.PS integration for more reliable interactive authentication prompts
- Fixed timing issues with authentication context token validation

### 🆕 New Features
- Added token caching to minimize re-authentication prompts
- Enhanced authentication context flow with better error messages
- Improved handling of authentication timeouts and cancellation

### 🔧 Technical Changes
- Better integration with MSAL.PS for authentication context scenarios
- Enhanced token validation and refresh logic
- Improved error handling for authentication context failures

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
- Direct REST API calls for authentication context preservation  
- Automatic module dependency management
- Comprehensive error handling and user feedback

### 📋 Requirements
- Windows Operating System
- PowerShell 7+ (Download from https://aka.ms/powershell)
- Microsoft Graph PowerShell modules (auto-installed)
- Az.Accounts module for WAM authentication support
- Appropriate Entra ID permissions for PIM role management

### 📝 Development Note
This module was developed with the assistance of AI tools (GitHub Copilot and Claude), combining AI-accelerated development with human expertise in Microsoft identity and security workflows.

For detailed usage instructions, see the README.md file.
'@
            # Flag to indicate whether the module requires explicit user acceptance
            RequireLicenseAcceptance = $false
            
            # External module dependencies that are not captured by RequiredModules
            ExternalModuleDependencies = @()
        }
    }
}