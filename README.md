# PIMActivation PowerShell Module

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/PIMActivation.svg)](https://www.powershellgallery.com/packages/PIMActivation)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/PIMActivation.svg)](https://www.powershellgallery.com/packages/PIMActivation)
[![Publish to PowerShell Gallery](https://github.com/Noble-Effeciency13/PIMActivation/actions/workflows/PSGalleryPublish.yml/badge.svg)](https://github.com/Noble-Effeciency13/PIMActivation/actions/workflows/PSGalleryPublish.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A comprehensive PowerShell module for managing Privileged Identity Management (PIM) role activations across Microsoft Entra ID, PIM-enabled groups, and Azure Resources through an intuitive graphical interface. Streamline your privileged access workflows with support for authentication context, bulk activations, and policy compliance across your entire Azure and Microsoft 365 environment.

> 📖 **Read the full blog post**: [PIMActivation: The Ultimate Tool for Microsoft Entra PIM Bulk Role Activation](https://www.chanceofsecurity.com/post/microsoft-entra-pim-bulk-role-activation-tool) on [Chance of Security](https://www.chanceofsecurity.com/)

![PIM Activation Interface](https://img.shields.io/badge/GUI-Windows%20Forms-blue?style=flat-square)
![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?style=flat-square)

## ✨ Key Features

- 🎨 **Modern GUI Interface** - Clean, responsive Windows Forms application with real-time updates
- 🔐 **Multi-Role Support** - Activate Microsoft Entra ID roles, PIM-enabled security groups, and Azure Resource roles
- ⚡ **Parallel Processing Engine** - Lightning-fast parallel execution with real-time progress tracking and emoji indicators
- 🚀 **High-Performance Batch API** - 85% reduction in API calls through intelligent batching, caching, and concurrent operations
- 🎯 **Advanced Duplicate Role Handling** - Sophisticated MemberType-based classification system for managing roles with multiple assignment paths
- 🛡️ **Authentication Context Support** - Seamless handling of Conditional Access authentication context requirements
- ⏱️ **Flexible Duration** - Configurable activation periods from 30 minutes to 24 hours, depending on policy maximum
- 📋 **Policy Compliance** - Automatic detection and handling of MFA, justification, and ticket requirements
- 🔄 **Up-to-Date Snapshot** - Shows current active and pending assignments based on the latest refresh or user action
- 👤 **Account Management** - Easy account switching without application restart
- 🔧 **PowerShell Compatibility** - Requires PowerShell 7+ for optimal parallel processing performance and modern language features

## 📸 Screenshots

### Main Interface
![PIM Activation Main Interface](https://github.com/user-attachments/assets/27557d5b-9060-45b4-bd61-dbccb96b6493)

*The main PIM activation interface showing eligible roles, active assignments, and activation options with policy requirements. Features intelligent group-role attribution, advanced duplicate role handling with MemberType classification, and smooth progress tracking with batch API performance enhancements.*

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

### Use a specific app registration (optional)
If your organization requires using a dedicated app registration for delegated auth, provide ClientId and TenantId:

```powershell
Start-PIMActivation -ClientId "<appId>" -TenantId "<tenantId>"
```

When both are provided, authentication uses the supplied app; otherwise, the default interactive flow is used.

## 📋 Prerequisites

### System Requirements
- **Windows Operating System** (Windows 10/11 or Windows Server 2016+)
- **PowerShell 7+** (Download from [https://aka.ms/powershell](https://aka.ms/powershell))
- **.NET Framework 4.7.2+** (for Windows Forms support)

### Required PowerShell Modules
The following modules will be automatically installed when you first run `Start-PIMActivation`:

#### Microsoft Graph (for Entra ID and Groups)
- `Microsoft.Graph.Authentication` (2.29.0+)
- `Microsoft.Graph.Users` (2.29.0+)
- `Microsoft.Graph.Identity.DirectoryManagement` (2.29.0+)
- `Microsoft.Graph.Identity.Governance` (2.29.0+)
- `Microsoft.Graph.Groups` (2.29.0+)
- `Microsoft.Graph.Identity.SignIns` (2.29.0+)

#### Azure PowerShell (for Azure Resources)
- `Az.Accounts` (5.1.0+) - provides authentication and context management
- `Az.Resources` (6.0.0+) - required for Azure Resource PIM role management

**Note:** Dependencies are automatically resolved when you run `Start-PIMActivation`. If you encounter issues, try running the command with the `-Force` parameter for fully automated resolution.

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

#### For Azure Resource Management
- **Azure RBAC Reader** or higher at subscription level
- **Privileged Role Administrator** for PIM-eligible resource role management
- **Access to Azure subscriptions** where resource roles are assigned

#### Base Permissions
- `User.Read`
- `Policy.Read.ConditionalAccess` (for authentication context support)

## 💡 Usage Examples

### Basic Operations
```powershell
# Launch with default settings (parallel processing enabled, Entra roles and groups)
Start-PIMActivation

# Include Azure Resource roles with parallel processing (fast!)
Start-PIMActivation -IncludeAzureResources

# Include all role types with optimized parallel execution
Start-PIMActivation -IncludeEntraRoles -IncludeGroups -IncludeAzureResources

# Use a specific app registration for delegated auth
Start-PIMActivation -ClientId "<appId>" -TenantId "<tenantId>"

# Show only Entra ID directory roles
Start-PIMActivation -IncludeEntraRoles

# Show only PIM-enabled security groups
Start-PIMActivation -IncludeGroups

# Show only Azure Resource roles
Start-PIMActivation -IncludeAzureResources
```

### Performance and Parallel Processing
```powershell
# Default: Parallel processing with ThrottleLimit 10 (fastest)
Start-PIMActivation -IncludeAzureResources

# Increase parallel operations for very large environments
Start-PIMActivation -IncludeAzureResources -ThrottleLimit 15

# Disable parallel processing for troubleshooting or compatibility
Start-PIMActivation -IncludeAzureResources -DisableParallelProcessing

# Custom throttle with parallel processing disabled
Start-PIMActivation -DisableParallelProcessing -ThrottleLimit 1

# Enable verbose output to see parallel processing performance
$VerbosePreference = 'Continue'
Start-PIMActivation -IncludeAzureResources -Verbose
```

### Advanced Scenarios
```powershell
# For organizations with authentication context policies
# The module automatically handles conditional access requirements

# For bulk activations
# 1. Launch Start-PIMActivation
# 2. Select multiple roles
# 3. Set duration
# 4. Click "Activate Roles"
# 5. Fill out justification, and ticket info if required
# 6. Complete any required authentication challenges
```

## 🚀 Parallel Processing Engine

### Performance Features
The module includes a powerful parallel processing engine that dramatically improves performance:

- **Default Parallel Execution**: All operations run in parallel by default (PowerShell 7+ required)
- **Real-Time Progress Tracking**: Visual progress with emoji indicators (🚀, ✅, ❌) and timing metrics
- **Intelligent Throttling**: Default ThrottleLimit of 10 concurrent operations, adjustable up to 50
- **Thread-Safe Operations**: Concurrent collections ensure safe parallel execution
- **Enhanced Verbose Output**: Detailed logging shows parallel operation progress and performance gains

### Parallel Processing Control
```powershell
# Default: Parallel processing enabled (fastest)
Start-PIMActivation

# Increase concurrency for large environments
Start-PIMActivation -ThrottleLimit 20

# Disable parallel processing if needed
Start-PIMActivation -DisableParallelProcessing

# See parallel processing performance
$VerbosePreference = 'Continue'
Start-PIMActivation -Verbose
```

### Performance Impact
- **Azure Subscriptions**: Processes multiple subscriptions concurrently
- **Policy Retrieval**: Fetches Entra and Group policies in parallel
- **Real-Time Feedback**: Shows progress like "Processing 5 subscriptions in parallel"
- **Timing Metrics**: Displays completion times, e.g., "Completed in 3.2s"

## 🔧 Configuration

### Authentication Context Support
The module automatically detects and handles authentication context requirements from Conditional Access policies. When a role requires additional authentication, the module will:

1. Detect the authentication context requirement for each selected roles
2. Group roles by context ID
3. Prompt re-authentication pr. context ID, utilizing WAM
4. Handle the activation seamlessly

### Module Settings
```powershell
# View current Graph connection
Get-MgContext

# Clear cached tokens (useful for troubleshooting)
Disconnect-MgGraph
```

## 📊 Supported Role Types

| Role Type | Support Status | Notes |
|-----------|---------------|-------|
| **Entra ID Directory Roles** | ✅ Full Support | Global Admin, User Admin, etc. |
| **PIM-Enabled Security Groups** | ✅ Full Support | Groups with PIM governance enabled |
| **Azure Resource Roles** | ✅ Full Support | Subscription, resource group, and individual resource roles |

### Azure Resource Role Features
- **Multi-Subscription Support**: Automatically enumerates roles across all accessible Azure subscriptions
- **Scope Hierarchy**: Supports tenant root, management group, subscription, resource group, and individual resource scopes
- **Inheritance Detection**: Distinguishes between direct assignments and inherited roles from higher scopes
- **Silent SSO**: Seamlessly authenticates to Azure PowerShell using your existing Graph authentication context
- **Resource Type Parsing**: Intelligently displays resource names and types (Storage Account, Virtual Machine, etc.)
- **PIM Integration**: Full support for PIM-eligible Azure Resource role activation and deactivation

## 🛠️ Troubleshooting

### Common Issues

**Authentication Failures**
```powershell
# Clear authentication cache
Disconnect-MgGraph

# Restart with fresh authentication
Start-PIMActivation
```

**PowerShell Version Issues**
- The module requires PowerShell 7+ for modern language features and WAM authentication support
- WAM (Windows Web Account Manager) provides more reliable authentication on Windows 10/11

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
- **Token Handling**: Leverages WAM (Windows Web Account Manager) for secure token management with automatic refresh
- **Authentication Context**: Properly handles conditional access policies and authentication challenges
- **Audit Trail**: All role activations are logged in Entra ID audit logs

## 🗺️ Roadmap

### Version 2.1.0 (Planned)
- **Profile Management**: Save and quickly activate frequently used role and account combinations
- **Scheduling**: Plan role activations for future times
- **Enhanced Reporting**: Built-in activation history and analytics
- **Persistent Settings**: Save parallel processing and throttle preferences

### Wishlist features
- **Cross-Platform**: Linux and macOS Support
- **Backwards compatibility**: Support for Powershell 5.1
- **Mobile app**: Mobile app for PIM Activations on the go
- **Reporting**: Built-in activation history and analytics
- **Automation integration**: Integration with different automation systems - still a liquid idea

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

- **Trevor Jones** for his excellent blog post on [WAM authentication in PowerShell](https://smsagent.blog/2024/11/28/getting-an-access-token-for-microsoft-entra-in-powershell-using-the-web-account-manager-wam-broker-in-windows/) which was instrumental in implementing reliable authentication
- PowerShell community for best practices and feedback

---

**Made with ❤️ for the PowerShell and Microsoft Entra ID community**
