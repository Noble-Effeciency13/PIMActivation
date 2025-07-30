function Initialize-WebAssembly {
    <#
    .SYNOPSIS
        Initializes System.Web assembly for URL encoding operations.
    
    .DESCRIPTION
        Loads the System.Web assembly required for HttpUtility.UrlEncode operations
        used in authentication context handling. This is a non-critical operation
        with fallback methods available if loading fails.
    
    .EXAMPLE
        Initialize-WebAssembly
        
        Loads the System.Web assembly for URL encoding functionality.
    
    .NOTES
        This function is called internally during PIM service connections.
        Failure to load this assembly is non-critical as fallback methods exist.
    #>
    [CmdletBinding()]
    param()
    
    try {
        Add-Type -AssemblyName System.Web -ErrorAction Stop
        Write-Verbose "Successfully loaded System.Web assembly"
    }
    catch {
        Write-Verbose "System.Web assembly load failed: $($_.Exception.Message). Using fallback methods."
    }
}

function Connect-PIMServices {
    <#
    .SYNOPSIS
        Establishes authenticated connections to Microsoft Graph and Azure services for PIM operations.
    
    .DESCRIPTION
        Creates authenticated connections to Microsoft services based on the specified role types.
        Handles Microsoft Graph authentication for Entra ID roles and groups, and Azure Resource Manager
        authentication for Azure resource roles. Manages module version conflicts and provides
        comprehensive connection validation.
    
    .PARAMETER IncludeEntraRoles
        Specifies whether to establish connection for Entra ID role management.
        Requires Microsoft Graph connection with appropriate role management scopes.
    
    .PARAMETER IncludeGroups
        Specifies whether to establish connection for privileged group management.
        Requires Microsoft Graph connection with group management scopes.
    
    .PARAMETER IncludeAzureResources
        Specifies whether to establish connection for Azure resource role management.
        Requires both Microsoft Graph and Azure Resource Manager connections.
    
    .PARAMETER ForceNewAccount
        Forces the account picker to appear even if already authenticated.
        Useful for switching between different user accounts or when authentication issues occur.
    
    .OUTPUTS
        PSCustomObject
        Returns a connection result object with the following properties:
        - Success: Boolean indicating overall connection success
        - Error: String containing error message if connection failed
        - GraphContext: Microsoft Graph context object if Graph connection established
        - CurrentUser: Current user object from Microsoft Graph
        - AzureContext: Azure context object if Azure connection established
    
    .EXAMPLE
        Connect-PIMServices -IncludeEntraRoles
        
        Connects to Microsoft Graph for Entra ID role management operations.
    
    .EXAMPLE
        Connect-PIMServices -IncludeEntraRoles -IncludeGroups -IncludeAzureResources
        
        Establishes connections for all PIM operation types: Entra roles, groups, and Azure resources.
    
    .EXAMPLE
        $connection = Connect-PIMServices -IncludeEntraRoles -ForceNewAccount
        if ($connection.Success) {
            Write-Host "Connected as: $($connection.CurrentUser.UserPrincipalName)"
        }
        
        Forces account selection and displays the connected user information.
    
    .NOTES
        This function requires the following PowerShell modules to be installed:
        - Microsoft.Graph.Authentication
        - Microsoft.Graph.Users
        - Microsoft.Graph.Identity.DirectoryManagement
        - Microsoft.Graph.Identity.Governance
        - Az.Accounts (if IncludeAzureResources is specified)
        
        The function automatically handles module version conflicts by removing and reimporting
        the latest available versions.
    
    .LINK
        https://docs.microsoft.com/en-us/powershell/microsoftgraph/
    
    .LINK
        https://docs.microsoft.com/en-us/azure/active-directory/privileged-identity-management/
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
        [switch]$ForceNewAccount
    )
    
    # Initialize connection result object
    $result = [PSCustomObject]@{
        Success = $false
        Error = $null
        GraphContext = $null
        CurrentUser = $null
        AzureContext = $null
    }
    
    try {
        # Establish Microsoft Graph connection if required
        if ($IncludeEntraRoles -or $IncludeGroups) {
            Write-Verbose "Initializing Microsoft Graph connection..."
            
            # Define required Graph modules
            $requiredModules = @(
                'Microsoft.Graph.Authentication',
                'Microsoft.Graph.Users',
                'Microsoft.Graph.Identity.DirectoryManagement',
                'Microsoft.Graph.Identity.Governance',
                'Az.Accounts'
            )
            
            # Resolve module version conflicts and ensure latest versions are loaded
            foreach ($moduleName in $requiredModules) {
                $loadedModules = @(Get-Module -Name $moduleName -ErrorAction SilentlyContinue)
                
                if ($loadedModules.Count -gt 1) {
                    Write-Verbose "Resolving version conflict for $moduleName..."
                    $loadedModules | Remove-Module -Force -ErrorAction SilentlyContinue
                    
                    $latestModule = Get-Module -ListAvailable -Name $moduleName | 
                                   Sort-Object Version -Descending | 
                                   Select-Object -First 1
                    
                    if ($latestModule) {
                        Import-Module -Name $moduleName -RequiredVersion $latestModule.Version -Force -ErrorAction Stop
                        Write-Verbose "Loaded $moduleName v$($latestModule.Version)"
                    }
                }
                elseif ($loadedModules.Count -eq 0) {
                    $availableModule = Get-Module -ListAvailable -Name $moduleName | 
                                      Sort-Object Version -Descending | 
                                      Select-Object -First 1
                    
                    if ($availableModule) {
                        Import-Module -Name $moduleName -RequiredVersion $availableModule.Version -Force -ErrorAction Stop
                        Write-Verbose "Loaded $moduleName v$($availableModule.Version)"
                    }
                    else {
                        $result.Error = "Required module '$moduleName' not found. Install with: Install-Module -Name $moduleName"
                        return $result
                    }
                }
            }
            
            # Initialize System.Web assembly for URL encoding
            Initialize-WebAssembly
            
            # Clear existing Graph context if forced account selection requested
            if ($ForceNewAccount) {
                Write-Verbose "Clearing existing authentication context..."
                
                # Preserve authentication context token if it exists and is recent
                $preserveAuthToken = $false
                if ($script:CurrentAuthContextToken -and $script:AuthContextCompletionTime) {
                    $timeSinceAuth = (Get-Date) - $script:AuthContextCompletionTime
                    if ($timeSinceAuth.TotalMinutes -lt 30) {  # Preserve token if less than 30 minutes old
                        $preserveAuthToken = $true
                        Write-Verbose "Preserving recent authentication context token (age: $([math]::Round($timeSinceAuth.TotalMinutes, 2)) minutes)"
                    }
                }
                
                for ($i = 0; $i -lt 3; $i++) {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                    Start-Sleep -Milliseconds 200
                }
                
                # Clear token cache but preserve auth context token if recent
                if (-not $preserveAuthToken) {
                    Write-Verbose "Clearing authentication context tokens (too old or doesn't exist)"
                    $script:CurrentAuthContextToken = $null
                    $script:CurrentAuthContextRefreshToken = $null
                    $script:AuthContextTokens = @{}
                    $script:JustCompletedAuthContext = $false
                    $script:AuthContextCompletionTime = $null
                }
            }
            
            # Define Microsoft Graph permission scopes
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
                        
            try {
                # Establish Graph connection
                Write-Verbose "Authenticating to Microsoft Graph..."
                Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
                
                # Validate connection establishment
                $context = Get-MgContext
                if (-not $context) {
                    $result.Error = "Microsoft Graph connection failed - no authentication context available"
                    return $result
                }
                
                $result.GraphContext = $context
                Write-Verbose "Microsoft Graph connection established successfully"
                
                # Retrieve current user information
                if ($context.Account) {
                    Write-Verbose "Retrieving current user profile..."
                    $currentUser = Get-MgUser -UserId $context.Account -ErrorAction Stop
                    $result.CurrentUser = $currentUser
                    Write-Verbose "Authenticated as: $($currentUser.UserPrincipalName)"
                    
                    # Persist last used account for future sessions
                    try {
                        Save-LastUsedAccount -UserPrincipalName $currentUser.UserPrincipalName
                    }
                    catch {
                        Write-Verbose "Unable to save account preference: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                $result.Error = "Microsoft Graph authentication failed: $($_.Exception.Message)"
                Write-Verbose "Graph connection error: $($_.Exception.GetType().Name) - $($_.Exception.Message)"
                return $result
            }
        }
        
        # Establish Azure Resource Manager connection if required
        if ($IncludeAzureResources) {
            Write-Warning "Azure Resource role management is not yet implemented. Skipping Azure connection for version 2.0.0."
            Write-Verbose "Azure resource support placeholder - connection skipped"
            
            # Set result properties to indicate feature not available
            $result.AzureContext = [PSCustomObject]@{
                Status = "Not Implemented"
                Message = "Azure Resource support planned for version 2.0.0"
            }
        }
        
        $result.Success = $true
        Write-Verbose "All requested service connections established successfully"
    }
    catch {
        $result.Error = "Service connection failed: $($_.Exception.Message)"
        Write-Verbose "Unexpected error: $($_.Exception.GetType().Name) - $($_.Exception.Message)"
    }
    
    return $result
}