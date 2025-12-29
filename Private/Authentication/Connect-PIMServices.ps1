function Connect-PIMServices {
    <#
    .SYNOPSIS
        Establishes authenticated connections to Microsoft Graph and Azure services for PIM operations.

    .DESCRIPTION
        Creates authenticated connections to Microsoft services based on the specified role types.
        Uses just-in-time module loading with version pinning to ensure compatibility.
        Handles Microsoft Graph authentication for Entra ID roles and groups, and Azure Resource Manager
        authentication for Azure resource roles. Also ensures Azure context is reset and re-scoped when
        switching accounts so Azure roles are correctly discovered for the active identity.

    .PARAMETER IncludeEntraRoles
        Connect for Entra ID role management.

    .PARAMETER IncludeGroups
        Connect for privileged group management.

    .PARAMETER IncludeAzureResources
        Connect for Azure resource role management (requires Graph + Az).

    .PARAMETER ForceNewAccount
        Forces account picker and clears Graph/Azure contexts (useful when switching accounts).

    .PARAMETER ClientId
        Optional app registration ClientId for Graph auth.

    .PARAMETER TenantId
        Optional tenant for the provided app registration.

    .OUTPUTS
        PSCustomObject
        Properties:
        - Success       : [bool]
        - Error         : [string]
        - GraphContext  : [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphPowerShellContext]
        - CurrentUser   : [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]
        - AzureContext  : [Microsoft.Azure.Commands.Common.Authentication.Abstractions.IAzureContextContainer]

    .LINK
        https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/
    #>
    [CmdletBinding()]
    param(
        [Parameter(HelpMessage = "Connect for Entra ID role management")]
        [switch]$IncludeEntraRoles,

        [Parameter(HelpMessage = "Connect for privileged group management")]
        [switch]$IncludeGroups,

        [Parameter(HelpMessage = "Connect for Azure resource role management")]
        [switch]$IncludeAzureResources,

        [Parameter(HelpMessage = "Force account picker to appear")]
        [switch]$ForceNewAccount,

        [Parameter(HelpMessage = "Client ID of the app registration to use for Graph auth")]
        [string]$ClientId,

        [Parameter(HelpMessage = "Tenant ID to use with the specified app registration")]
        [string]$TenantId,

        [Parameter(HelpMessage = "Optional loading splash control to update UI status/progress")]
        [PSCustomObject]$SplashForm
    )

    # Result object
    $result = [PSCustomObject]@{
        Success      = $false
        Error        = $null
        GraphContext = $null
        CurrentUser  = $null
        AzureContext = $null
    }

    # Local helper to safely update splash status
    function _UpdateStatus([string]$status, [int]$progress = -1) {
        try {
            if ($PSBoundParameters.ContainsKey('SplashForm') -and $SplashForm -and $SplashForm.SyncHash -and -not $SplashForm.IsDisposed) {
                Update-LoadingStatus -SplashForm $SplashForm -Status $status -Progress $progress
            }
        } catch { }
    }

    try {
        # Initialize/pin modules needed for Graph/Az
        $moduleInit = Initialize-PIMModules
        if (-not $moduleInit.Success) {
            $result.Error = "Failed to initialize PIM modules: $($moduleInit.Error)"
            return $result
        }

        # --- Microsoft Graph ---
        if ($IncludeEntraRoles -or $IncludeGroups -or $IncludeAzureResources) {
            Write-Verbose "Initializing Microsoft Graph connection..."

            # JIT load Graph modules
            _UpdateStatus "Loading modules..." 40
            if (-not (Import-PIMModule -ModuleName 'Microsoft.Graph.Authentication')) {
                $result.Error = "Failed to load Microsoft.Graph.Authentication"
                return $result
            }
            Initialize-WebAssembly

            if ($IncludeEntraRoles) {
                if (-not (Import-PIMModule -ModuleName 'Microsoft.Graph.Identity.DirectoryManagement')) {
                    $result.Error = "Failed to load Microsoft.Graph.Identity.DirectoryManagement"
                    return $result
                }
                if (-not (Import-PIMModule -ModuleName 'Microsoft.Graph.Identity.Governance')) {
                    $result.Error = "Failed to load Microsoft.Graph.Identity.Governance"
                    return $result
                }
            }
            if ($IncludeEntraRoles -or $IncludeGroups) {
                if (-not (Import-PIMModule -ModuleName 'Microsoft.Graph.Users')) {
                    $result.Error = "Failed to load Microsoft.Graph.Users"
                    return $result
                }
            }

            # JIT load Az before any auth if Azure is requested
            if ($IncludeAzureResources) {
                Write-Verbose "Importing Az.Accounts and Az.Resources modules before authentication"
                _UpdateStatus "Loading Azure modules..." 50
                if (-not (Import-PIMModule -ModuleName 'Az.Accounts')) {
                    $result.Error = "Failed to load Az.Accounts"
                    return $result
                }
                if (-not (Import-PIMModule -ModuleName 'Az.Resources')) {
                    $result.Error = "Failed to load Az.Resources"
                    return $result
                }
            }

            # Clear previous contexts when switching accounts
            if ($ForceNewAccount) {
                Write-Verbose "Clearing existing Graph and Azure contexts (ForceNewAccount)"
                for ($i = 0; $i -lt 2; $i++) {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                    Start-Sleep -Milliseconds 150
                }
                try {
                    # Remove any persisted Az contexts to avoid tenant/subscription bleed-through
                    Clear-AzContext -Force -ErrorAction SilentlyContinue | Out-Null
                    Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
                } catch { }
                # Clear cached role data to force fresh fetches
                $script:CachedEligibleRoles = @()
                $script:CachedActiveRoles   = @()
                $script:LastRoleFetchTime   = $null
            }

            # Graph scopes
            $graphScopes = @(
                'User.Read'
                'Directory.Read.All'
                'RoleEligibilitySchedule.ReadWrite.Directory'
                'RoleAssignmentSchedule.ReadWrite.Directory'
                'PrivilegedAccess.ReadWrite.AzureADGroup'
                'RoleManagementPolicy.Read.Directory'
                'RoleManagementPolicy.Read.AzureADGroup'
                'Policy.Read.ConditionalAccess'
            )
            if ($IncludeAzureResources) {
                # Broader scope enables ARM PIM operations via Graph SSO
                $graphScopes += 'RoleManagement.ReadWrite.Directory'
            }

            try {
                Write-Verbose "Authenticating to Microsoft Graph..."
                _UpdateStatus "Authenticating user..." 60
                if ($PSBoundParameters.ContainsKey('ClientId') -and $PSBoundParameters.ContainsKey('TenantId') -and $ClientId -and $TenantId) {
                    Write-Verbose "Using provided app registration (ClientId=$ClientId, TenantId=$TenantId)"
                    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Scopes $graphScopes -NoWelcome -ErrorAction Stop
                } else {
                    Write-Verbose "Using default interactive authentication"
                    Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
                }

                $context = Get-MgContext
                if (-not $context) {
                    $result.Error = "Microsoft Graph connection failed - no authentication context available"
                    return $result
                }

                $result.GraphContext     = $context
                $script:CurrentTenantId  = $context.TenantId
                $script:CurrentGraphUser = $context.Account

                Write-Verbose "Microsoft Graph connection established successfully"

                # Current user
                if ($context.Account) {
                    Write-Verbose "Retrieving current user profile..."
                    _UpdateStatus "Loading user profile..." 70
                    $currentUser = Get-MgUser -UserId $context.Account -ErrorAction Stop
                    $result.CurrentUser    = $currentUser
                    $script:CurrentUser    = $currentUser
                    Write-Verbose "Authenticated as: $($currentUser.UserPrincipalName)"
                    try { Save-LastUsedAccount -UserPrincipalName $currentUser.UserPrincipalName } catch { }
                }
            }
            catch {
                $result.Error = "Microsoft Graph authentication failed: $($_.Exception.Message)"
                Write-Verbose "Graph connection error: $($_.Exception.GetType().Name) - $($_.Exception.Message)"
                return $result
            }
        }

        # --- Azure Resource Manager ---
        if ($IncludeAzureResources) {
            Write-Verbose "Initializing Azure Resource Manager connection..."
            _UpdateStatus "Connecting to Azure Resource Manager..." 75
            try {
                if (-not $result.GraphContext -or -not $result.GraphContext.Account) {
                    throw "Azure Resource Manager connection requires valid Graph context"
                }

                $connectedAccount = $result.GraphContext.Account
                $connectedTenant  = $result.GraphContext.TenantId
                if (-not $connectedTenant) {
                    throw "No tenant available from Graph context for Azure authentication"
                }

                Write-Verbose "Using Microsoft Graph $connectedAccount context for Azure authentication"
                Write-Verbose "Scoping Azure connection to tenant: $connectedTenant"
                $azureContext = Connect-AzAccount -AccountId $connectedAccount -Tenant $connectedTenant -ErrorAction Stop
                if (-not $azureContext) { throw "Connect-AzAccount returned no context" }

                $result.AzureContext = $azureContext
                $script:AzureContext = $azureContext.Context
                $script:AzureConnectedAccount = $azureContext.Context.Account.Id
                $script:CurrentTenantId       = $azureContext.Context.Tenant.Id

                Write-Verbose "Azure Resource Manager connection established successfully"
                _UpdateStatus "Azure connection established" 80
                Write-Verbose "Connected to Azure with account: $($azureContext.Context.Account.Id)"
                Write-Verbose "Connected tenant: $($azureContext.Context.Tenant.Id)"

                # Enumerate subscriptions in current tenant and select a default context
                $subscriptions = Get-AzSubscription -ErrorAction SilentlyContinue | Where-Object {
                    $_.TenantId -eq $azureContext.Context.Tenant.Id -and $_.State -eq 'Enabled'
                }
                if (-not $subscriptions) {
                    # Fallback to HomeTenantId property used in some environments
                    $subscriptions = Get-AzSubscription -ErrorAction SilentlyContinue | Where-Object {
                        $_.HomeTenantId -eq $azureContext.Context.Tenant.Id -and $_.State -eq 'Enabled'
                    }
                }

                if ($subscriptions) {
                    $script:AzureSubscriptions = @($subscriptions)
                    $subscriptionCount = $script:AzureSubscriptions.Count
                    Write-Verbose "Found $subscriptionCount subscription(s) in tenant $($azureContext.Context.Tenant.Id)"

                    $defaultSub = $script:AzureSubscriptions[0]
                    Write-Verbose "Selecting subscription $($defaultSub.Name) ($($defaultSub.Id))"
                    Select-AzSubscription -SubscriptionId $defaultSub.Id -Tenant $azureContext.Context.Tenant.Id -ErrorAction SilentlyContinue | Out-Null

                    # Publish the selected subscription for downstream role queries
                    $script:AzureDefaultSubscriptionId = $defaultSub.Id
                    _UpdateStatus "Selected subscription: $($defaultSub.Name)" 85
                } else {
                    Write-Verbose "No Azure subscriptions accessible with current account in this tenant"
                    $script:AzureSubscriptions = @()
                    $script:AzureDefaultSubscriptionId = $null
                }
            }
            catch {
                $result.Error = "Azure Resource Manager authentication failed: $($_.Exception.Message)"
                Write-Verbose "Azure connection error: $($_.Exception.GetType().Name) - $($_.Exception.Message)"
                return $result
            }
        }

        $result.Success = $true
        Write-Verbose "All requested service connections established successfully"
        _UpdateStatus "All services connected" 90
    }
    catch {
        $result.Error = "Service connection failed: $($_.Exception.Message)"
        Write-Verbose "Unexpected error: $($_.Exception.GetType().Name) - $($_.Exception.Message)"
    }

    return $result
}