@{
    # Script module or binary module file associated with this manifest.
    RootModule           = 'PIMActivation.psm1'
    
    # Version number of this module.
    ModuleVersion        = '2.2.0'
    
    # Supported PSEditions - Requires PowerShell Core (7+)
    CompatiblePSEditions = @('Core')
    
    # ID used to uniquely identify this module
    GUID                 = 'a3f4b8e2-9c7d-4e5f-b6a9-8d7c6b5a4f3e'
    
    # Author of this module
    Author               = 'Sebastian Flæng Markdanner'
    
    # Company or vendor of this module
    CompanyName          = 'Cloudy With a Change Of Security'
    
    # Copyright statement for this module
    Copyright            = '(c) 2025 Sebastian Flæng Markdanner. All rights reserved.'

    # Description of the functionality provided by this module
    Description          = 'PowerShell module for managing Microsoft Entra ID Privileged Identity Management (PIM) role activations through a modern GUI interface. Supports Entra ID roles, PIM-enabled groups, Azure Resource roles, scheduled activations, activation profiles, Azure reduced scope, authentication-context batching, bulk operations, persistent policy metadata caching, and policy compliance. Developed with AI assistance. Requires PowerShell 7+.'
    
    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion    = '7.0'
    
    # Script to run after the module is imported
    ScriptsToProcess     = @()
    
    # Required modules are validated and imported by PIMActivation.psm1 to support
    # faster startup and local development scenarios.
    RequiredModules      = @()
    
    # Functions to export from this module
    FunctionsToExport    = @(
        'Start-PIMActivation'
    )
    
    # Cmdlets to export from this module
    CmdletsToExport      = @()
    
    # Variables to export from this module
    VariablesToExport    = @()
    
    # Aliases to export from this module
    AliasesToExport      = @()
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData          = @{
        PSData = @{
            # Tags applied to this module for online gallery discoverability
            Tags                     = @('PIM', 'PrivilegedIdentityManagement', 'EntraID', 'AzureAD', 'Azure', 'AzureResources', 'Identity', 'Governance', 'RBAC', 'GUI', 'Authentication', 'ConditionalAccess', 'Security', 'Microsoft', 'Graph', 'ScheduledActivations', 'ActivationProfiles', 'PolicyCache')
            
            # A URL to the license for this module.
            LicenseUri               = 'https://github.com/Noble-Effeciency13/PIMActivation/blob/main/LICENSE'
            
            # A URL to the main website for this project.
            ProjectUri               = 'https://github.com/Noble-Effeciency13/PIMActivation'
            
            # A URL to an icon representing this module.
            IconUri                  = 'https://raw.githubusercontent.com/Noble-Effeciency13/PIMActivation/main/Resources/icon.png'
            
            # ReleaseNotes
            ReleaseNotes             = @'
## PIMActivation v2.2.0 - Scheduling, Activation Profiles, Reduced Scope, and Policy Cache

## What's included

• Activation Profiles: save frequently used role selections as named local profiles under %LOCALAPPDATA%\PIMActivation\ActivationProfiles. A new Activation Profiles button in the header gives quick access to saved profiles. Any current role selection can be saved from the activation dialog, and profiles support one-click launch with pre-filled roles and duration. Profiles can be updated or deleted within the activation flow.
• Scheduled Activations: choose a future local date/time for regular and profile-based activation requests within the selected role eligibility window. The chosen time is honoured across Entra ID, Azure Resource, and authentication-context flows.
• Azure Reduced Scope: Azure Resource role activations can optionally target a narrower effective scope using a guided picker through subscription, resource group, and resource. The last-used scope path is remembered for repeat activations.
• Persistent Policy Cache: PIM policy metadata is cached under %LOCALAPPDATA%\PIMActivation\PolicyCache in tenant-scoped folders, reducing API calls at startup. Stale entries are revalidated in the background to keep policy data current.
• Azure Scope Display: Azure Resource scopes in the eligible and active role lists now show as Sub: <subscription>, RG: <resource group>, or Resource: <name> (and MG: <name> or Tenant Root for higher-level scopes) so the effective scope is readable without inspecting the raw ARM path.
• Administrative Unit Scope Column: Entra ID and Group role assignments scoped to an Administrative Unit now show Administrative Unit in the Scope column with the AU name in the Resource column, making the scope kind visible at a glance.
• Authentication Context Claim ID Display: authentication-context requirements show as their claim ID, for example Required (C2), removing the need for Conditional Access display-name lookups at startup and avoiding access-denied noise in restricted environments.
• Activation Progress Visibility: the activation splash shows a grouped batch overview shared across Entra ID (Graph) and Azure Resource (ARM) channels so a multi-role activation reads as a single logical operation rather than separate mini-batches.
• Adaptive Operation Splash: the operation splash auto-resizes on each status update to fit the current message. Long batch overviews and multi-line messages are no longer clipped.
• Approval-Required Activation Refresh: the eligible roles list refreshes after submission of an approval-required activation so the Pending Approval column reflects the newly submitted request without a manual refresh.
• Full Refresh Behavior: Full Refresh clears both in-memory role/policy data and the on-disk persistent policy cache before rebuilding, ensuring fresh policy requirements from source.
• Faster Startup and Dependency Loading: Microsoft Graph and Azure (Az.Accounts, Az.Resources) modules are now validated and loaded at module import time rather than during Start-PIMActivation. The GUI opens noticeably faster, Azure Resource role support is ready immediately without the previous mid-launch module-install pause, and import-time errors surface clearly before the activation workflow starts.
• Azure Resource Role and Policy Collection: Azure Resource role enumeration and ARM policy parsing now more accurately reflect active, eligible, direct, inherited, and provisioned assignment states, so the eligible and active lists better match what the Azure portal shows.

## Fixes

• Azure Resource Eligible Role Discovery: resolved a catch-22 where listing Azure Resource PIM eligibility could fail because eligibility enumeration appeared to require an existing Azure role. Eligible role discovery now uses the ARM asTarget() filter so users can enumerate their own Azure Resource PIM eligibility from a clean state without any pre-existing Azure role assignment.

## Notes

• Local files under %LOCALAPPDATA%\PIMActivation\ store metadata only. Tokens, refresh tokens, authentication-context tokens, activation request bodies, scheduled start times, justifications, ticket values, and secrets are not persisted.
• Multi-channel activations still use the appropriate channel per role (Graph, Azure Resource Manager, authentication-context step-up); the unified batch overview reflects the logical operation without merging the underlying HTTP calls.
'@
            # Flag to indicate whether the module requires explicit user acceptance
            RequireLicenseAcceptance = $false
        }
    }
}