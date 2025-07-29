# PIMActivation PowerShell Module

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/PIMActivation.svg)](https://www.powershellgallery.com/packages/PIMActivation)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/PIMActivation.svg)](https://www.powershellgallery.com/packages/PIMActivation)
[![Publish to PowerShell Gallery](https://github.com/Noble-Effeciency13/PIMActivation/actions/workflows/publish-to-psgallery.yml/badge.svg)](https://github.com/Noble-Effeciency13/PIMActivation/actions/workflows/publish-to-psgallery.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A comprehensive PowerShell module for managing Microsoft Entra ID Privileged Identity Management (PIM) role activations through an intuitive graphical interface. Streamline your privileged access workflows with support for authentication context, bulk activations, and policy compliance.

> 📖 **Read the full blog post**: [Microsoft Entra PIM Bulk Role Activation Tool](https://www.chanceofsecurity.com/post/microsoft-entra-pim-bulk-role-activation-tool) on [Chance of Security](https://www.chanceofsecurity.com/)

![PIM Activation Interface](https://img.shields.io/badge/GUI-Windows%20Forms-blue?style=flat-square)
![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?style=flat-square)

## ✨ Key Features

- 🎨 **Modern GUI Interface** - Clean, responsive Windows Forms application with real-time updates
- 🔐 **Multi-Role Support** - Activate Microsoft Entra ID roles and PIM-enabled security groups
- ⚡ **Bulk Operations** - Select and activate multiple roles simultaneously with policy validation
- 🛡️ **Authentication Context Support** - Seamless handling of Conditional Access authentication context requirements
- ⏱️ **Flexible Duration** - Configurable activation periods from 1 hour to policy maximum (typically 8-24 hours)
- 📋 **Policy Compliance** - Automatic detection and handling of MFA, justification, and ticket requirements
- 🔄 **Real-time Monitoring** - Live view of active assignments and pending requests
- 👤 **Account Management** - Easy account switching without application restart
- 🔧 **PowerShell Compatibility** - Requires PowerShell 7+ for optimal performance and modern language features

## 🚀 Quick Start

### Installation

#### From PowerShell Gallery (Recommended)
```powershell
# Install for current user
Install-Module -Name PIMActivation -Scope CurrentUser

# Install system-wide (requires admin)
Install-Module -Name PIMActivation -Scope AllUsers
```

#### From GitHub Source
```powershell
# Clone and import
git clone https://github.com/Noble-Effeciency13/PIMActivation.git
cd PIMActivation
Import-Module .\PIMActivation.psd1
```

### First Run
```powershell
# Launch the PIM activation interface
Start-PIMActivation
```

On first launch, you'll be prompted to authenticate with Microsoft Graph using your organizational account.

## 📋 Prerequisites

### System Requirements
- **Windows Operating System** (Windows 10/11 or Windows Server 2016+)
- **PowerShell 7+** (Download from [https://aka.ms/powershell](https://aka.ms/powershell))
- **.NET Framework 4.7.2+** (for Windows Forms support)

### Required PowerShell Modules
The following modules will be automatically installed if missing:
- `Microsoft.Graph.Authentication` (2.0+)
- `Microsoft.Graph.Users`
- `Microsoft.Graph.Identity.DirectoryManagement`
- `Microsoft.Graph.Identity.Governance`
- `MSAL.PS` (4.37+)

### Microsoft Entra ID Permissions
Your account needs the following **delegated** permissions:

#### For Entra ID Role Management
- `RoleEligibilitySchedule.ReadWrite.Directory`
- `RoleAssignmentSchedule.ReadWrite.Directory`
- `RoleManagementPolicy.Read.Directory`
- `Directory.Read.All`

#### For PIM Group Management
- `PrivilegedAccess.ReadWrite.AzureADGroup`
- `RoleManagementPolicy.Read.AzureADGroup`

#### Base Permissions
- `User.Read`
- `Policy.Read.ConditionalAccess` (for authentication context support)

> **Note:** These permissions are typically granted through PIM role assignments like "Privileged Role Administrator" or "Global Administrator".

## 💡 Usage Examples

### Basic Operations
```powershell
# Launch with default settings (Entra roles and groups)
Start-PIMActivation

# Show only Entra ID directory roles
Start-PIMActivation -IncludeEntraRoles

# Show only PIM-enabled security groups
Start-PIMActivation -IncludeGroups

# Force account selection dialog
Start-PIMActivation -ForceNewAccount
```

### Advanced Scenarios
```powershell
# For organizations with authentication context policies
# The module automatically handles conditional access requirements

# For bulk activations
# 1. Launch Start-PIMActivation
# 2. Use Ctrl+Click to select multiple roles
# 3. Click "Activate Selected Roles"
# 4. Set duration and justification
# 5. Complete any required authentication challenges
```

## 🔧 Configuration

### Authentication Context Support
The module automatically detects and handles authentication context requirements from Conditional Access policies. When a role requires additional authentication (like privileged access workstations), the module will:

1. Detect the authentication context requirement
2. Acquire the appropriate token with correct claims
3. Handle the activation seamlessly

### Module Settings
```powershell
# View current Graph connection
Get-MgContext

# Clear cached tokens (useful for troubleshooting)
Disconnect-MgGraph
```

## � Supported Role Types

| Role Type | Support Status | Notes |
|-----------|---------------|-------|
| **Entra ID Directory Roles** | ✅ Full Support | Global Admin, User Admin, etc. |
| **PIM-Enabled Security Groups** | ✅ Full Support | Groups with PIM governance enabled |
| **Azure Resource Roles** | 🚧 Planned | Subscription and resource-level roles |

## 🛠️ Troubleshooting

### Common Issues

**Authentication Failures**
```powershell
# Clear authentication cache
Disconnect-MgGraph
Clear-MsalTokenCache

# Restart with fresh authentication
Start-PIMActivation -ForceNewAccount
```

**PowerShell Version Issues**
- The module requires PowerShell 7+ for modern language features and optimal performance
- For authentication context scenarios, it may launch Windows PowerShell 5.1 processes internally for optimal MSAL.PS compatibility

**Permission Errors**
- Ensure your account has the required PIM role assignments
- Check that the necessary Graph API permissions are consented for your organization

### Verbose Logging
```powershell
# Enable detailed logging for troubleshooting
$VerbosePreference = 'Continue'
Start-PIMActivation -Verbose
```

## 🔒 Security Considerations

- **Credential Management**: Uses Microsoft Graph delegated permissions, no credentials are stored
- **Token Handling**: Leverages MSAL.PS for secure token management with automatic refresh
- **Authentication Context**: Properly handles conditional access policies and authentication challenges
- **Audit Trail**: All role activations are logged in Entra ID audit logs

## 🗺️ Roadmap

### Version 2.0.0 (Planned)
- **Azure Resource Roles**: Support for Azure subscription and resource-level PIM roles
- **Profile Management**: Save and quickly activate frequently used role combinations
- **Scheduling**: Plan role activations for future times
- **Reporting**: Built-in activation history and analytics

### Version 2.1.0 (Future)
- **REST API Support**: Direct API calls for environments without PowerShell modules
- **Cross-Platform**: Linux and macOS support with PowerShell 7+
- **Automation**: PowerShell DSC and Azure Automation integration

## 🤝 Contributing

I welcome contributions! Please see my [Contributing Guidelines](CONTRIBUTING.md) for details.

### Development Setup

```powershell
# Clone the repository
git clone https://github.com/Noble-Effeciency13/PIMActivation.git
cd PIMActivation

# Import module for development
Import-Module .\PIMActivation.psd1 -Force

# Run tests (when available)
Invoke-Pester
```

### Areas for Contribution
- 🧪 **Testing**: Unit tests and integration tests
- 📚 **Documentation**: Examples, tutorials, and API documentation
- 🔧 **Features**: Azure resource roles, profile management
- 🐛 **Bug Fixes**: Issue resolution and performance improvements

## 🤖 Development Transparency

This module was developed using modern AI-assisted programming practices, combining AI tools (GitHub Copilot and Claude) with human expertise in Microsoft identity and security workflows. All code has been thoroughly reviewed, tested, and validated in production environments.

The authentication context implementation particularly benefited from AI assistance in solving complex token management and timing challenges. The result is production-ready code that leverages the efficiency of AI-assisted development while maintaining high standards of quality and security.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- **GitHub Issues**: [Report bugs or request features](https://github.com/Noble-Effeciency13/PimActivation/issues)
- **Documentation**: [Wiki and guides](https://github.com/Noble-Effeciency13/PimActivation/wiki)
- **Discussions**: [Community discussions](https://github.com/Noble-Effeciency13/PimActivation/discussions)
- **Blog Post**: [Detailed solution walkthrough](https://www.chanceofsecurity.com/post/microsoft-entra-pim-bulk-role-activation-tool)
- **Author's Blog**: [Chance of Security](https://www.chanceofsecurity.com/)

## 🙏 Acknowledgments

- Microsoft Graph PowerShell SDK team for excellent API support
- MSAL.PS contributors for authentication library
- PowerShell community for best practices and feedback

---

**Made with ❤️ for the PowerShell and Microsoft Entra ID community**
