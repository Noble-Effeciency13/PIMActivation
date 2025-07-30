# Changelog

All notable changes to the PIMActivation PowerShell module will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Planned Features
- **Azure Resource Roles**: Support for Azure subscription and resource-level PIM roles
- **Profile Management**: Save and quickly activate frequently used role combinations and accounts
- **Scheduling**: Plan role activations for future times
- **Enhanced Reporting**: Built-in activation history and analytics

---

## [1.1.1] - 2025-07-30

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

### Documentation
- Updated troubleshooting wiki with latest version compatibility information
- Revised module installation and conflict resolution guidance
- Added prevention tips for maintaining module compatibility

---

## [1.1.0] - 2025-07-30

### Added
- **WAM Authentication**: Implemented Windows Web Account Manager (WAM) for more reliable authentication
- Direct integration with Az.Accounts MSAL assemblies for better performance
- Enhanced timeout management and error handling for authentication flows

### Credits
- WAM implementation inspired by Trevor Jones' excellent blog post: [Getting an access token for Microsoft Entra in PowerShell using the Web Account Manager (WAM) broker in Windows](https://smsagent.blog/2024/11/28/getting-an-access-token-for-microsoft-entra-in-powershell-using-the-web-account-manager-wam-broker-in-windows/)

### Removed
- **MSAL.PS Dependency**: Removed dependency on MSAL.PS module
- PowerShell 5.1 fallback processes - now fully PowerShell 7+ native

### Changed
- Improved authentication context handling with direct MSAL.NET calls
- Better assembly loading and management from Az.Accounts module
- Enhanced error messages and troubleshooting information

### Fixed
- Resolved authentication hanging issues with proper async/sync handling
- Improved reliability of authentication prompts on Windows 10/11

---

## [1.0.1] - 2025-07-29

### Fixed
- Enhanced error handling for authentication context scenarios

### Changed
- Authentication context tokens are now cached per context ID to minimize authentication prompts
- Improved token validation and expiry management

---

## [1.0.0] - 2025-07-28

### 🎉 Initial Release

#### Added
- **Modern GUI Interface**: Clean Windows Forms application for PIM role management
- **Multi-Role Support**: Activate Microsoft Entra ID roles and PIM-enabled security groups
- **Authentication Context Support**: Seamless handling of Conditional Access authentication context policies
- **Bulk Operations**: Select and activate multiple roles simultaneously with policy validation
- **PowerShell Compatibility**: Requires PowerShell 7+ for optimal performance and modern language features
- **Policy Compliance**: Automatic detection of MFA, justification, and ticket requirements
- **Real-time Updates**: Live monitoring of active assignments and pending requests
- **Account Management**: Easy account switching without application restart

#### Technical Features
- Direct REST API calls for authentication context preservation
- Automatic module dependency management
- Comprehensive error handling and user feedback
- Full PowerShell 7+ compatibility with modern language features

#### Requirements
- Windows Operating System (Windows 10/11 or Windows Server 2016+)
- PowerShell 7+
- Microsoft Graph PowerShell modules (auto-installed if missing)
- Appropriate Entra ID permissions for PIM role management

#### Known Limitations
- Azure Resource roles not yet supported (planned for v2.0.0)
- Profile management not yet available (planned for v2.0.0)
- Windows-only GUI support

---

## Version History

- **v1.0.1** (2025-07-29): Bug fixes for authentication context and MSAL.PS reliability
- **v1.0.0** (2025-07-28): Initial release with core PIM activation functionality

---

[Unreleased]: https://github.com/Noble-Effeciency13/PIMActivation/compare/v1.0.1...HEAD
[1.0.1]: https://github.com/Noble-Effeciency13/PIMActivation/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/Noble-Effeciency13/PIMActivation/releases/tag/v1.0.0