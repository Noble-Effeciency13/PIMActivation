---
name: Bug report
about: Create a report to help us improve PIMActivation
title: "[BUG] "
labels: bug
assignees: ''

---

## Describe the bug
A clear and concise description of what the bug is.

## To Reproduce
Steps to reproduce the behavior:
1. Launch module with command: '...'
2. Select roles: '...'
3. Click button: '...'
4. See error

## Expected behavior
A clear and concise description of what you expected to happen.

## Error Details
If applicable, paste the full error message:
```powershell
# Paste error message here
```

## Screenshots
If applicable, add screenshots of the GUI or error messages.

## Environment (please complete the following information):
- **OS**: [e.g. Windows 11, Windows Server 2022]
- **PowerShell Version**: [Run `$PSVersionTable.PSVersion`]
- **Module Version**: [Run `Get-Module PIMActivation | Select-Object Version`]
- **Microsoft Graph Module Version**: [Run `Get-Module Microsoft.Graph.Authentication | Select-Object Version`]

## Role Configuration:
- **Role Types Affected**: [Entra ID roles, PIM Groups, Azure Resources]
- **Authentication Context Required**: [Yes/No - if yes, which context?]
- **MFA Required**: [Yes/No]
- **Approval Required**: [Yes/No]

## Additional context
Add any other context about the problem here, such as:
- When did this start happening?
- Does it happen consistently?
- Any recent changes to your environment?

## Verbose Output (if possible)
```powershell
# Run with: $VerbosePreference = 'Continue'; Start-PIMActivation -Verbose
# Paste relevant verbose output here
```
