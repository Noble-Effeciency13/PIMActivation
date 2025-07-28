# Changelog

All notable changes to the PIMActivation PowerShell module will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
- Hybrid PowerShell architecture for optimal MSAL.PS compatibility
- Direct REST API calls for authentication context preservation
- Automatic module dependency management
- Comprehensive error handling and user feedback
- Support for PowerShell 7+ with hybrid PowerShell 5.1 processes for optimal MSAL.PS compatibility

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

## [Unreleased]

### Planned Features
- **Azure Resource Roles**: Support for Azure subscription and resource-level PIM roles
- **Profile Management**: Save and quickly activate frequently used role combinations
- **Scheduling**: Plan role activations for future times
- **Enhanced Reporting**: Built-in activation history and analytics
- **Automation Integration**: PowerShell DSC and Azure Automation support

---

## Version History

- **v1.0.0** (2025-07-28): Initial release with core PIM activation functionality
