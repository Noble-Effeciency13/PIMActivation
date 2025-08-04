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

## [1.2.4] - 2025-08-04

### Fixed
- **Module Compatibility**: Changed from exact version requirements to minimum version checking for all dependencies
- **Missing Dependencies**: Added support for Microsoft.Graph.Groups and Microsoft.Graph.Identity.SignIns modules
- **Version Flexibility**: Module now accepts specified version or higher for better compatibility with existing installations

### Technical Improvements
- **Minimum Version Logic**: All modules now use minimum version checking (`-ge`) instead of exact matching (`-eq`)
- **Enhanced Error Messages**: Clear installation instructions for minimum version requirements
- **Complete Module Coverage**: Ensured all required Graph modules are properly validated and loaded
- **Future Compatibility**: Better support for newer module versions while maintaining stability

### Changed
- **Az.Accounts**: Now requires minimum version 5.1.0 (previously exact 5.1.0)
- **Microsoft.Graph Modules**: All Graph modules now use minimum version 2.29.1 (previously exact 2.29.1)
- **Module Loading**: Improved logic to select best available version that meets minimum requirements

---

## [1.2.3] - 2025-08-04

### Fixed
- **Dependency Management**: Fixed automatic module installation for missing dependencies during import
- **Silent Import**: Module import is now completely silent unless verbose mode is enabled
- **Development Workflow**: Resolved issues with local module import blocking when dependencies were missing

### Technical Improvements
- **ArrayList Performance**: Optimized dependency collection using ArrayLists instead of regular arrays
- **Quiet Installation**: Added silent installation with progress suppression for cleaner user experience
- **Verbose Support**: Detailed output available via `-Verbose` parameter for troubleshooting
- **Intelligent Installation**: Automatically configures NuGet provider and PSGallery trust as needed
- **Progress Feedback**: Clear visual feedback during automatic module installation process
- **Error Resilience**: Graceful handling of installation failures with helpful error messages
- **Universal Compatibility**: Same codebase works for local development and PowerShell Gallery distribution

---

## [1.2.1] - 2025-08-04

### Added
- **Automatic Dependency Resolution**: Enhanced `Start-PIMActivation` with automatic conflict detection and module installation
- **Force Parameter**: Added `-Force` parameter to `Start-PIMActivation` for fully automated dependency resolution
- **Clean Console Output**: Suppressed verbose output noise while preserving debugging capabilities when requested
- **Module Requirements**: Added explicit `#Requires` statements for all required Microsoft Graph modules and versions

### Fixed
- **Verbose Output Noise**: Suppressed "Populating RepositorySourceLocation" and other unwanted Get-Module verbose messages
- **Function Import Verbosity**: Eliminated verbose function import messages during module loading  
- **Console Clutter**: Removed duplicate dependency checking logic and unnecessary output during normal operation
- **Performance**: Replaced `+=` operators with `ArrayList.Add()` for better performance in loops and array operations
- **Code Readability**: Replaced backtick line continuations with parameter splatting for improved maintainability

### Technical Improvements
- **Resolve-PIMDependencies**: New internal function for comprehensive dependency resolution with retry logic
- **Enhanced Error Messages**: Improved user guidance for dependency resolution issues
- **Code Quality**: Performance optimizations including ArrayList usage and cleaner parameter handling
- **Maintainability**: Improved code structure with splatting instead of line continuations

---

## [1.2.0] - 2025-07-31

### Added
- **Batch Role Fetching**: New `Get-PIMRolesBatch` function that retrieves all roles and policies in optimized batch operations (85% reduction in API calls)
- **Batch Policy Processing**: New `Get-PIMPoliciesBatch` function that fetches multiple role policies simultaneously
- **Advanced Duplicate Role Handling**: Sophisticated MemberType-based classification system (Direct vs Group vs Inherited) for managing multiple instances of same role
- **Intelligent Group Attribution Algorithm**: Smart matching logic with unused group preference preventing overcorrection issues in duplicate role scenarios
- **MemberType Detection System**: Precise identification of role assignment sources using Microsoft Graph MemberType field ('Direct', 'Group', 'Inherited')
- **Group-Role Attribution System**: ProvidedRoles functionality showing which roles each group membership provides with cross-referencing validation
- **Enhanced Cross-Referencing**: Sophisticated logic for matching group-derived Entra roles with their providing groups using expiration correlation
- **Priority-Based Role Distribution**: Advanced algorithm distributing multiple inherited roles across respective providing groups
- **Advanced Array Handling**: @() wrapper implementation preventing .Count property errors in PowerShell
- **Safe Property Access**: PSObject.Properties pattern for bulletproof property checking
- **Enhanced Verbose Logging**: Detailed progress tracking with differentiated handling for groups vs Entra roles and attribution decisions

### Changed
- **Performance Enhancement**: Role loading now uses batch API operations instead of individual calls, significantly reducing load times
- **Improved UI Responsiveness**: Batch processing provides better progress feedback and faster role list population
- **Optimized Policy Retrieval**: Policies are now fetched once and cached for all applicable roles
- **Smooth Progress Flow**: Coordinated progress tracking across all loading phases (86%-98% range)
- **Enhanced Resource Display**: Shows "Entra ID (via Group: GroupName)" for group-derived roles
- **Proper Expiration Attribution**: Duplicate roles now show individual expiration times based on their providing groups

### Fixed
- **Cache Refresh Issues**: Resolved active roles not updating immediately after activation/deactivation
- **Progress Bar Jumps**: Fixed backwards progress jumps from 60% to 10% during initialization
- **Property Access Errors**: Comprehensive error handling for .Count, RoleDefinitionId, and EndDateTime properties
- **Group Visibility**: Enhanced cross-referencing logic for proper group-role relationship display
- **Duplicate Role Overcorrection**: Fixed issue where direct assignments were incorrectly attributed to groups when duplicates existed
- **AI Administrator Attribution Bug**: Resolved overcorrection causing both direct and group-derived AI Administrator roles to show group attribution
- **Helpdesk Administrator Distribution**: Fixed attribution issue where multiple Helpdesk Administrator roles from different groups were all attributed to same group
- **MemberType Classification**: Corrected detection logic to properly distinguish MemberType 'Group' vs 'Direct' vs 'Inherited' assignments
- **Single Direct Role Logic**: Enhanced logic to only apply group attribution when no inherited roles of same type exist
- **Defensive Coding**: Comprehensive try-catch blocks around all critical operations preventing crashes

### Technical
- **MemberType-Based Classification**: Sophisticated detection algorithm using Microsoft Graph MemberType field for accurate role source identification
- **Unused Group Preference Algorithm**: Smart attribution logic preferring unused groups for inherited role distribution across multiple providing groups
- **Overcorrection Prevention**: Enhanced single direct role logic preventing incorrect group attribution when inherited duplicates exist
- **Priority-Based Group Assignment**: Sophisticated matching logic with temporal vs permanent preferences and expiration correlation
- **Intelligent Caching**: Enhanced cache invalidation system with proper timing coordination
- **Comprehensive Error Handling**: Bulletproof property access protection preventing common PowerShell errors
- **Cross-Reference Validation**: Extensive debugging for group-role relationship verification and duplicate role scenarios

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