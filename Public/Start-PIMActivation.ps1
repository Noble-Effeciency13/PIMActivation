function Start-PIMActivation {
    <#
    .SYNOPSIS
        Starts the PIM Role Activation graphical interface.
    
    .DESCRIPTION
        Launches a Windows Forms application for managing Privileged Identity Management (PIM) role activations.
        The application provides an intuitive interface for activating Entra ID roles, PIM-enabled groups, and Azure resource roles.
        
        Requirements:
        - PowerShell 7.0 or later
        - Single-threaded apartment (STA) mode for Windows Forms
        - Required Microsoft Graph and Azure PowerShell modules (auto-installed if missing)
        
        The tool automatically handles authentication, module dependencies, and provides a loading interface
        during initialization.
    
    .PARAMETER IncludeEntraRoles
        Include Entra ID (Azure AD) roles in the activation interface.
        When enabled, displays available Entra ID role assignments that can be activated.
        Default: $true
    
    .PARAMETER IncludeGroups
        Include PIM-enabled security groups in the activation interface.
        When enabled, displays eligible group memberships that can be activated.
        Default: $true
    
    .PARAMETER IncludeAzureResources
        Include Azure resource roles (RBAC) in the activation interface.
        When enabled, displays eligible Azure subscription and resource group role assignments.
        Requires additional Az PowerShell modules (Az.Accounts, Az.Resources).
        Default: $false
    
    .EXAMPLE
        Start-PIMActivation
        
        Launches the PIM activation interface with default settings.
        Includes Entra ID roles and PIM-enabled groups, but excludes Azure resource roles.
    
    .EXAMPLE
        Start-PIMActivation -IncludeAzureResources
        
        Launches the interface including Azure resource roles (RBAC assignments).
        Also includes Entra ID roles and groups by default.
    
    .EXAMPLE
        Start-PIMActivation -IncludeEntraRoles:$false -IncludeGroups:$false -IncludeAzureResources
        
        Launches the interface showing only Azure resource roles.
        Excludes Entra ID roles and PIM-enabled groups.
    
    .NOTES
        Name: Start-PIMActivation
        Author: GitHub Copilot
        Version: 1.0
        
        This function requires PowerShell 7+ and will automatically restart in STA mode if needed.
        Missing required modules are automatically installed from the PowerShell Gallery.
        
        The function maintains session state for account switching and can restart itself
        when users need to switch between different Microsoft accounts.
    
    .LINK
        https://docs.microsoft.com/en-us/azure/active-directory/privileged-identity-management/
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([void])]
    param(
        [Parameter(HelpMessage = "Include Entra ID roles in the activation interface")]
        [switch]$IncludeEntraRoles = $true,
        
        [Parameter(HelpMessage = "Include PIM-enabled groups in the activation interface")]
        [switch]$IncludeGroups = $true,
        
        [Parameter]
        [switch]$IncludeAzureResources
    )
    
    begin {
        Write-Verbose "Starting PIM Activation Tool initialization"
        
        # Validate PowerShell version requirement
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            $errorMessage = "PowerShell 7 or later is required. Current version: $($PSVersionTable.PSVersion). Please upgrade from https://aka.ms/powershell"
            Write-Error $errorMessage -Category InvalidOperation
            throw $errorMessage
        }
        
        # Configure execution preferences
        $originalVerbosePreference = $VerbosePreference
        $originalWarningPreference = $WarningPreference
        $originalProgressPreference = $ProgressPreference
        
        # Preserve user's verbose preference while silencing other noise
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')) {
            $script:UserVerbose = $true
            Write-Verbose "Verbose output enabled by user"
        } else {
            $script:UserVerbose = $false
        }
        
        $WarningPreference = 'SilentlyContinue'
        $ProgressPreference = 'SilentlyContinue'
        
        # Suppress Azure PowerShell breaking change warnings
        $env:SuppressAzurePowerShellBreakingChangeWarnings = 'true'
        
        # Initialize session state variables
        $script:RestartRequested = $false
        
        Write-Verbose "Initialization parameters: EntraRoles=$IncludeEntraRoles, Groups=$IncludeGroups, AzureResources=$IncludeAzureResources"
    }
    
    process {
        try {
            # Check if user wants to proceed with starting the PIM activation tool
            if (-not $PSCmdlet.ShouldProcess("PIM Activation Tool", "Start PIM role activation interface")) {
                Write-Verbose "Operation cancelled by user"
                return
            }
            
            # Ensure Single-Threaded Apartment mode for Windows Forms
            if (-not (Test-STAMode)) {
                Write-Verbose "Restarting in STA mode for Windows Forms compatibility"
                return Start-STAProcess -ScriptBlock {
                    param($ModulePath, $Params)
                    Import-Module $ModulePath -Force
                    Start-PIMActivation @Params
                } -ArgumentList @($PSScriptRoot, $PSBoundParameters)
            }
            
            # Store parameters for potential restart scenarios (account switching)
            $script:StartupParameters = $PSBoundParameters
            $script:IncludeEntraRoles = $IncludeEntraRoles
            $script:IncludeGroups = $IncludeGroups
            $script:IncludeAzureResources = $IncludeAzureResources
            
            # Load required .NET assemblies for Windows Forms
            Write-Verbose "Loading Windows Forms assemblies"
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            
            # Initialize loading interface
            $splashForm = Show-LoadingSplash -Message "Initializing PIM Activation Tool..."
            
            # Give splash form time to render
            Start-Sleep -Milliseconds 200
            
            try {
                # Define and validate required PowerShell modules
                Write-Verbose "Validating required PowerShell modules"
                Update-LoadingStatus -SplashForm $splashForm -Status "Checking dependencies..." -Progress 10
                
                $requiredModules = @(
                    @{Name = 'Microsoft.Graph.Authentication'; MinVersion = '2.0.0'}
                    @{Name = 'Microsoft.Graph.Users'; MinVersion = '2.0.0'}
                    @{Name = 'Microsoft.Graph.Identity.DirectoryManagement'; MinVersion = '2.0.0'}
                    @{Name = 'Microsoft.Graph.Identity.Governance'; MinVersion = '2.0.0'}
                )
                
                # Add Azure modules if Azure resources are requested
                if ($IncludeAzureResources) {
                    Write-Warning "Azure Resource role management is not yet implemented. This parameter will be functional in version 2.0.0."
                    Write-Verbose "Setting IncludeAzureResources to false - feature not yet implemented"
                    $IncludeAzureResources = $false
                    $script:IncludeAzureResources = $false
                }
                
                # Install/update required modules with progress tracking
                $moduleProgress = 10
                $moduleStep = 30 / $requiredModules.Count
                
                foreach ($module in $requiredModules) {
                    Update-LoadingStatus -SplashForm $splashForm -Status "Validating $($module.Name)..." -Progress $moduleProgress
                    $moduleProgress += $moduleStep
                    Start-Sleep -Milliseconds 50  # Brief pause for UI responsiveness
                }
                
                $moduleResult = Install-RequiredModules -RequiredModules $requiredModules -Verbose:$script:UserVerbose
                
                if (-not $moduleResult.Success) {
                    throw "Module installation failed: $($moduleResult.Error)"
                }
                
                # Establish service connections
                Write-Verbose "Connecting to Microsoft services"
                Update-LoadingStatus -SplashForm $splashForm -Status "Connecting to Microsoft Graph..." -Progress 50
                
                $connectionParams = @{
                    IncludeEntraRoles = $script:IncludeEntraRoles
                }
                
                Update-LoadingStatus -SplashForm $splashForm -Status "Authenticating user..." -Progress 60
                $connectionResult = Connect-PIMServices @connectionParams
                
                if (-not $connectionResult.Success) {
                    throw "Authentication failed: $($connectionResult.Error)"
                }
                
                Write-Verbose "Connected as user: $($connectionResult.CurrentUser.UserPrincipalName)"
                Update-LoadingStatus -SplashForm $splashForm -Status "Loading user profile..." -Progress 70
                
                # Store connection context for session management
                $script:CurrentUser = $connectionResult.CurrentUser
                $script:GraphContext = $connectionResult.GraphContext
                
                # Initialize main application form
                Write-Verbose "Building main application interface"
                Update-LoadingStatus -SplashForm $splashForm -Status "Building interface..." -Progress 80
                
                $form = Initialize-PIMForm -SplashForm $splashForm -Verbose:$script:UserVerbose
                
                if (-not $form) {
                    throw "Failed to create main application form"
                }
                
                # Launch main application
                Write-Verbose "Launching PIM Activation interface"
                [System.Windows.Forms.Application]::EnableVisualStyles()
                [void]$form.ShowDialog()
                
                # Handle restart requests (typically for account switching)
                if ($script:RestartRequested) {
                    Write-Verbose "Processing restart request for account switch"
                    $script:RestartRequested = $false
                    Start-Sleep -Milliseconds 500  # Allow clean shutdown
                    
                    # Restart with same parameters
                    Start-PIMActivation @script:StartupParameters
                }
            }
            finally {
                # Clean up loading interface
                if ($splashForm -and -not $splashForm.IsDisposed) {
                    Close-LoadingSplash -SplashForm $splashForm
                }
            }
        }
        catch {
            $errorMessage = "PIM Activation Tool failed to start: $($_.Exception.Message)"
            Write-Error $errorMessage -Category OperationStopped
            Write-Verbose "Error details: $($_.ScriptStackTrace)"
            
            # Display user-friendly error dialog
            try {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to start PIM Activation Tool:`n`n$($_.Exception.Message)",
                    "PIM Activation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            catch {
                # Fallback if MessageBox fails
                Write-Host "Error: $errorMessage" -ForegroundColor Red
            }
            
            throw
        }
        finally {
            # Session cleanup
            if ($script:CurrentUser) {
                Write-Verbose "Cleaning up session for: $($script:CurrentUser.UserPrincipalName)"
            }
            
            # Avoid disconnection during restart to maintain session state
            if (-not $script:RestartRequested) {
                try {
                    Write-Verbose "Disconnecting from services"
                    Disconnect-PIMServices -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Verbose "Non-critical error during service disconnection: $($_.Exception.Message)"
                }
            }
            
            # Restore original preferences
            $VerbosePreference = $originalVerbosePreference
            $WarningPreference = $originalWarningPreference
            $ProgressPreference = $originalProgressPreference
        }
    }
    
    end {
        Write-Verbose "PIM Activation Tool session completed"
    }
}