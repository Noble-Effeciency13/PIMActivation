# Contributing to PIMActivation

Thank you for your interest in contributing to PIMActivation! This document provides guidelines and information for contributors.

## 🤝 How to Contribute

### Reporting Issues
- Use the [GitHub Issues](https://github.com/Noble-Effeciency13/PIMActivation/issues) page
- Search existing issues before creating a new one
- Provide detailed information including:
  - PowerShell version
  - Operating system
  - Steps to reproduce
  - Expected vs actual behavior
  - Error messages (if any)

### Suggesting Features
- Open a [Feature Request](https://github.com/Noble-Effeciency13/PIMActivation/issues/new?template=feature_request.md)
- Describe the use case and benefits
- Consider implementation complexity

### Pull Requests
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test thoroughly
5. Commit with clear messages (`git commit -m 'Add amazing feature'`)
6. Push to your branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## 🛠️ Development Setup

### Prerequisites
- Windows 10/11 or Windows Server 2016+
- PowerShell 7+ (Download from [https://aka.ms/powershell](https://aka.ms/powershell))
- Git for version control
- Visual Studio Code (recommended)

### Local Development
```powershell
# Clone the repository
git clone https://github.com/Noble-Effeciency13/PIMActivation.git
cd PIMActivation

# Import the module for testing
Import-Module .\PIMActivation.psd1 -Force

# Test the module
Start-PIMActivation
```

### Testing
Currently, the project uses manual testing. Automated tests are welcome contributions!

```powershell
# Manual testing checklist
# 1. Module imports without errors
# 2. GUI launches successfully
# 3. Authentication works
# 4. Role activation functions properly
# 5. Error handling works as expected
```

## 📝 Coding Standards

### PowerShell Best Practices
- Follow [PowerShell Best Practices](https://docs.microsoft.com/en-us/powershell/scripting/developer/cmdlet/strongly-encouraged-development-guidelines)
- Use approved verbs for function names
- Include comprehensive help documentation
- Use `Write-Verbose` for debugging information
- Handle errors gracefully with try/catch blocks

### Code Style
- Use 4 spaces for indentation
- Place opening braces on the same line
- Use meaningful variable and function names
- Include comments for complex logic
- Follow PowerShell naming conventions

### Documentation
- Update README.md for user-facing changes
- Update CHANGELOG.md for all changes
- Include inline comments for complex code
- Update help documentation for new functions

## 🎯 Priority Areas

### High Priority
- **Unit Tests**: PowerShell Pester tests
- **Azure Resource Roles**: Implementation of Azure resource support
- **Error Handling**: Improved user feedback and error recovery
- **Performance**: Optimization of API calls and UI responsiveness

### Medium Priority
- **Profile Management**: Save/load role activation profiles
- **Logging**: Structured logging with configurable levels
- **Internationalization**: Support for multiple languages
- **Accessibility**: Improved GUI accessibility

### Low Priority
- **Themes**: Dark mode and custom themes
- **Keyboard Shortcuts**: Additional hotkeys for power users
- **Export/Import**: Configuration backup and restore

## 🔍 Code Review Process

### For Contributors
- Ensure code follows style guidelines
- Test on PowerShell 7+ (Core edition)
- Update documentation as needed
- Keep PRs focused and atomic

### For Maintainers
- Review for functionality and style
- Test on PowerShell 7+ environments
- Verify documentation updates
- Check for breaking changes

## 📋 Release Process

### Version Numbering
We follow [Semantic Versioning](https://semver.org/):
- **Major** (X.0.0): Breaking changes
- **Minor** (0.X.0): New features, backward compatible
- **Patch** (0.0.X): Bug fixes, backward compatible

### Release Checklist
1. Update version in `PIMActivation.psd1`
2. Update `CHANGELOG.md`
3. Update `README.md` if needed
4. Test thoroughly
5. Create GitHub release
6. Publish to PowerShell Gallery

## 🏷️ Labels and Milestones

### Issue Labels
- `bug`: Something isn't working
- `enhancement`: New feature or request
- `documentation`: Improvements or additions to docs
- `good first issue`: Good for newcomers
- `help wanted`: Extra attention is needed
- `priority-high`: Critical issues
- `priority-low`: Nice to have

### Milestones
- `v1.1.0`: Bug fixes and minor improvements
- `v2.0.0`: Azure resource roles and profiles
- `v2.1.0`: Cross-platform support

## 🤔 Questions?

- Check the [Wiki](https://github.com/Noble-Effeciency13/PIMActivation/wiki)
- Join [Discussions](https://github.com/Noble-Effeciency13/PIMActivation/discussions)
- Ask in issues with the `question` label
- Read the [detailed blog post](https://www.chanceofsecurity.com/post/microsoft-entra-pim-bulk-role-activation-tool) about this solution
- Visit [Chance of Security](https://www.chanceofsecurity.com/) for more security insights

## 📄 License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for helping make PIMActivation better! 🎉
