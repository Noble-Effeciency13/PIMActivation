function Get-AuthenticationContextToken {
    <#
    .SYNOPSIS
        Gets or retrieves a cached authentication context token for the specified context ID.
    
    .DESCRIPTION
        Manages authentication context tokens by caching them per context ID to avoid
        repeated authentication prompts. Validates token expiry and refreshes as needed.
        Uses Windows Web Account Manager (WAM) for authentication.
    
    .PARAMETER ContextId
        The authentication context ID (e.g., "c3") required by the role policy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ContextId
    )
    
    try {
        Write-Verbose "=== Starting authentication context token acquisition for context: $ContextId ==="
        
        # Initialize the AuthContextTokens hashtable if it doesn't exist
        if (-not $script:AuthContextTokens) {
            $script:AuthContextTokens = @{}
            Write-Verbose "Initialized AuthContextTokens cache"
        }
        
        # Check if we already have a valid token for this context
        if ($script:AuthContextTokens.ContainsKey($ContextId)) {
            $cachedToken = $script:AuthContextTokens[$ContextId]
            
            # Validate token is still fresh (less than 30 minutes old)
            if ($cachedToken.ExpiryTime -and (Get-Date) -lt $cachedToken.ExpiryTime) {
                Write-Verbose "Using cached authentication context token for context: $ContextId (expires: $($cachedToken.ExpiryTime))"
                return $cachedToken.AccessToken
            }
            else {
                Write-Verbose "Cached token for context $ContextId has expired, obtaining fresh token"
                $script:AuthContextTokens.Remove($ContextId)
            }
        }
        
        # Get current Graph context for tenant ID
        $currentContext = Get-MgContext
        $tenantId = if ($currentContext) { $currentContext.TenantId } else { $null }
        
        if (-not $tenantId) {
            throw "No active Microsoft Graph connection. Cannot determine tenant ID. Please ensure you're connected to Microsoft Graph."
        }
        
        Write-Verbose "Current tenant ID: $tenantId"
        Write-Verbose "PowerShell version: $($PSVersionTable.PSVersion)"
        Write-Verbose "Obtaining fresh authentication context token for context: $ContextId"
        
        # Ensure we're running PowerShell Core
        if ($PSEdition -ne "Core") {
            throw "WAM authentication requires PowerShell Core (PowerShell 7+)"
        }
        
        # Build the claims challenge format
        $claimsJson = '{"access_token":{"acrs":{"essential":true,"value":"' + $ContextId + '"}}}'
        Write-Verbose "Claims challenge: $claimsJson"
        
        Write-Verbose "=== Starting WAM authentication setup ==="
        
        # Check if Az.Accounts module is already loaded
        $LoadedAzAccountsModule = Get-Module -Name Az.Accounts
        if ($null -eq $LoadedAzAccountsModule) {
            # Check for Az.Accounts module (required for WAM dependencies)
            $AzAccountsModule = Get-Module -Name Az.Accounts -ListAvailable
            if ($null -eq $AzAccountsModule) {
                Write-Verbose "Installing Az.Accounts module for WAM dependencies"
                Install-Module -Name Az.Accounts -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
            }
            
            # Import Az.Accounts module
            Write-Verbose "Loading Az.Accounts module for WAM dependencies"
            Import-Module Az.Accounts -ErrorAction Stop -Verbose:$false
            Write-Verbose "Az.Accounts module loaded"
        }
        else {
            Write-Verbose "Az.Accounts module already loaded (version $($LoadedAzAccountsModule.Version))"
        }
        
        # Find the location of the Azure.Common assembly
        $LoadedAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Select-Object -ExpandProperty Location
        $AzureCommon = $LoadedAssemblies | Where-Object { $_ -match "\\Modules\\Az.Accounts\\" -and $_ -match "Microsoft.Azure.Common" }
        
        if (-not $AzureCommon) {
            throw "Could not find Microsoft.Azure.Common assembly from Az.Accounts module"
        }
        
        $AzureCommonLocation = $AzureCommon.TrimEnd("Microsoft.Azure.Common.dll")
        Write-Verbose "Azure Common Location: $AzureCommonLocation"
        
        # Locate the required assemblies
        Write-Verbose "Locating required assemblies for WAM"
        $requiredAssemblies = @(
            'Microsoft.IdentityModel.Abstractions.dll',
            'Microsoft.Identity.Client.dll',
            'Microsoft.Identity.Client.Broker.dll',
            'Microsoft.Identity.Client.NativeInterop.dll',
            'Microsoft.Identity.Client.Extensions.Msal.dll',
            'System.Security.Cryptography.ProtectedData.dll'
        )
        
        $assemblies = @{}
        
        foreach ($assemblyFile in $requiredAssemblies) {
            $assemblyName = $assemblyFile.Replace('.dll', '')
            $found = Get-ChildItem -Path $AzureCommonLocation -Filter $assemblyFile -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            if (-not $found) {
                throw "Could not find required assembly: $assemblyFile"
            }
            $assemblies[$assemblyName] = $found
            Write-Verbose "Found $assemblyName at: $found"
        }
        
        # Get System.Diagnostics.TraceSource from .NET Core installation or module dependencies
        $sdts = $null
        
        # First, try to get the module path
        $moduleBase = $null
        $pimModule = Get-Module -Name 'PIMActivation' -ErrorAction SilentlyContinue
        if ($pimModule) {
            $moduleBase = $pimModule.ModuleBase
            Write-Verbose "PIMActivation module base: $moduleBase"
        }
        else {
            # Try to find the module in the PSModulePath
            $modulePaths = $env:PSModulePath -split ';'
            foreach ($path in $modulePaths) {
                $testPath = Join-Path $path 'PIMActivation'
                if (Test-Path $testPath) {
                    $moduleBase = $testPath
                    Write-Verbose "Found PIMActivation module at: $moduleBase"
                    break
                }
            }
        }
        
        # Check if System.Diagnostics.TraceSource is already loaded
        $loadedAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies()
        $tracesourceAssembly = $loadedAssemblies | Where-Object { $_.GetName().Name -eq "System.Diagnostics.TraceSource" }
        
        if ($tracesourceAssembly -and $tracesourceAssembly.Location) {
            $sdts = $tracesourceAssembly.Location
            Write-Verbose "Found System.Diagnostics.TraceSource from loaded assemblies: $sdts"
        }
        else {
            # Try multiple locations for System.Diagnostics.TraceSource
            $searchPaths = @()
            
            # Add module's lib directory if available
            if ($moduleBase) {
                $searchPaths += Join-Path $moduleBase 'lib'
                $searchPaths += Join-Path $moduleBase 'Dependencies'
                $searchPaths += $moduleBase
            }
            
            # Add Az.Accounts location
            $searchPaths += $AzureCommonLocation
            
            # Try .NET Core reference assemblies
            $RuntimeFrameworkMajorVersion = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription.Split()[-1].Split(".")[0]
            $possibleDotNetPaths = @(
                "C:\Program Files\dotnet\packs\Microsoft.NETCore.App.Ref",
                "C:\Program Files (x86)\dotnet\packs\Microsoft.NETCore.App.Ref",
                "$env:ProgramFiles\dotnet\packs\Microsoft.NETCore.App.Ref",
                "${env:ProgramFiles(x86)}\dotnet\packs\Microsoft.NETCore.App.Ref",
                "$env:DOTNET_ROOT\packs\Microsoft.NETCore.App.Ref"
            )
            
            foreach ($dotnetPath in $possibleDotNetPaths) {
                if (Test-Path $dotnetPath) {
                    $dotNetDirectory = Get-ChildItem -Path $dotnetPath -Filter "$RuntimeFrameworkMajorVersion.*" -Directory -ErrorAction SilentlyContinue | 
                    Sort-Object -Property Name -Descending | Select-Object -First 1
                    if ($dotNetDirectory) {
                        $searchPaths += $dotNetDirectory.FullName
                        Write-Verbose "Added .NET reference path: $($dotNetDirectory.FullName)"
                    }
                }
            }
            
            # Add runtime directory as last resort
            $runtimeDir = [System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
            $searchPaths += $runtimeDir
            
            # Search for the assembly
            foreach ($searchPath in $searchPaths) {
                if (Test-Path $searchPath) {
                    $found = Get-ChildItem -Path $searchPath -Filter "System.Diagnostics.TraceSource.dll" -Recurse -File -ErrorAction SilentlyContinue | 
                    Select-Object -First 1
                    if ($found) {
                        $sdts = $found.FullName
                        Write-Verbose "Found System.Diagnostics.TraceSource at: $sdts"
                        break
                    }
                }
            }
            
            # If still not found, check if it's available as a type (might be in GAC)
            if (-not $sdts) {
                try {
                    $traceSourceType = [System.Diagnostics.TraceSource]
                    if ($traceSourceType) {
                        Write-Verbose "System.Diagnostics.TraceSource is available as a type (likely in GAC)"
                        # Continue without loading it explicitly
                    }
                }
                catch {
                    Write-Warning "System.Diagnostics.TraceSource type is not available"
                }
            }
        }
        
        if ($sdts) {
            Write-Verbose "System.Diagnostics.TraceSource located at: $sdts"
        }
        else {
            Write-Warning "Could not locate System.Diagnostics.TraceSource.dll - WAM authentication might still work"
        }
        
        # Load the assemblies
        Write-Verbose "Loading WAM assemblies..."
        $loadedCount = 0
        $failedAssemblies = @()
        
        foreach ($assemblyPath in $assemblies.Values) {
            try {
                # Check if already loaded
                $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($assemblyPath)
                $alreadyLoaded = $loadedAssemblies | Where-Object { $_.GetName().Name -eq $assemblyName }
                
                if ($alreadyLoaded) {
                    Write-Verbose "Assembly already loaded: $assemblyName"
                    $loadedCount++
                }
                else {
                    [void][System.Reflection.Assembly]::LoadFrom($assemblyPath)
                    Write-Verbose "Loaded assembly: $(Split-Path -Leaf $assemblyPath)"
                    $loadedCount++
                }
            }
            catch {
                $failedAssemblies += Split-Path -Leaf $assemblyPath
                Write-Warning "Failed to load assembly $(Split-Path -Leaf $assemblyPath): $_"
            }
        }
        
        if ($sdts) {
            try {
                # Check if already loaded
                $alreadyLoaded = $loadedAssemblies | Where-Object { $_.GetName().Name -eq "System.Diagnostics.TraceSource" }
                
                if ($alreadyLoaded) {
                    Write-Verbose "System.Diagnostics.TraceSource already loaded"
                }
                else {
                    [void][System.Reflection.Assembly]::LoadFrom($sdts)
                    Write-Verbose "Loaded System.Diagnostics.TraceSource"
                }
            }
            catch {
                Write-Warning "Failed to load System.Diagnostics.TraceSource: $_"
            }
        }
        
        # Check if we have the minimum required assemblies
        $criticalAssemblies = @('Microsoft.Identity.Client', 'Microsoft.Identity.Client.Broker')
        $missingCritical = $criticalAssemblies | Where-Object { $failedAssemblies -contains "$_.dll" }
        
        if ($missingCritical) {
            throw "Critical assemblies missing for WAM authentication: $($missingCritical -join ', '). Please ensure Az.Accounts module is properly installed."
        }
        
        Write-Verbose "WAM assembly loading completed - Loaded: $loadedCount, Failed: $($failedAssemblies.Count)"
        
        # If System.Diagnostics.TraceSource couldn't be loaded as file, it might still work if the type exists
        if (-not $sdts -or $failedAssemblies -contains "System.Diagnostics.TraceSource.dll") {
            $sdts = "System.Diagnostics.TraceSource" # Use as type name reference
        }
        
        # C# code for WAM authentication with claims
        $code = @"
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.IdentityModel.Abstractions;
using Microsoft.Identity.Client.NativeInterop;
using Microsoft.Identity.Client.Extensions.Msal;

public class PIMAuthContextHelper
{
    // Get window handle of the console window
    [DllImport("user32.dll", ExactSpelling = true)]
    public static extern IntPtr GetAncestor(IntPtr hwnd, GetAncestorFlags flags);
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    
    public enum GetAncestorFlags
    {   
        GetParent = 1,
        GetRoot = 2,
        GetRootOwner = 3
    }
    
    public static IntPtr GetConsoleOrTerminalWindow()
    {
        IntPtr consoleHandle = GetConsoleWindow();
        if (consoleHandle == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }
        IntPtr handle = GetAncestor(consoleHandle, GetAncestorFlags.GetRootOwner);
        return (handle != IntPtr.Zero) ? handle : consoleHandle;
    }    
    
    // Method for retrieving the access token with authentication context claims
    public static string GetAccessTokenWithAuthContext(string clientId, string tenantId, string redirectUri, string[] scopes, string claimsJson)
    {
        try
        {
            // Run the async method synchronously to avoid deadlocks
            var task = Task.Run(async () => await GetAccessTokenWithAuthContextAsync(clientId, tenantId, redirectUri, scopes, claimsJson));
            
            // Wait with timeout
            if (task.Wait(TimeSpan.FromSeconds(120)))
            {
                return task.Result;
            }
            else
            {
                throw new TimeoutException("Authentication timed out after 120 seconds");
            }
        }
        catch (AggregateException ae)
        {
            // Unwrap aggregate exceptions
            var innerException = ae.InnerException;
            if (innerException != null)
            {
                throw innerException;
            }
            throw;
        }
    }
    
    private static async Task<string> GetAccessTokenWithAuthContextAsync(string clientId, string tenantId, string redirectUri, string[] scopes, string claimsJson)
    {
        // Setup broker options
        var brokerOptions = new BrokerOptions(BrokerOptions.OperatingSystems.Windows)
        {
            Title = "PIM Role Activation - Authentication Context Required"
        };
        
        var authority = $"https://login.microsoftonline.com/{tenantId}";
        
        var appBuilder = PublicClientApplicationBuilder.Create(clientId)
            .WithAuthority(authority)
            .WithBroker(brokerOptions)
            .WithRedirectUri(redirectUri);
        
        // Try to set parent window if available
        var windowHandle = GetConsoleOrTerminalWindow();
        if (windowHandle != IntPtr.Zero)
        {
            appBuilder = appBuilder.WithParentActivityOrWindow(() => windowHandle);
        }
        
        IPublicClientApplication publicClientApp = appBuilder.Build();
        
        // Create cancellation token
        using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(120)))
        {
            try
            {
                // Always do interactive authentication for authentication context
                var result = await publicClientApp
                    .AcquireTokenInteractive(scopes)
                    .WithClaims(claimsJson)
                    .WithPrompt(Prompt.SelectAccount)
                    .WithUseEmbeddedWebView(false) // Force system browser/WAM
                    .ExecuteAsync(cts.Token)
                    .ConfigureAwait(false);
                
                return result.AccessToken;
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException("Authentication was cancelled or timed out");
            }
        }
    }
}
"@
        
        # List of assemblies we need to reference - filter out null/empty values
        $referencedAssemblies = @(
            $assemblies['Microsoft.IdentityModel.Abstractions'],
            $assemblies['Microsoft.Identity.Client'],
            $assemblies['Microsoft.Identity.Client.Broker'],
            $assemblies['Microsoft.Identity.Client.NativeInterop'],
            $assemblies['Microsoft.Identity.Client.Extensions.Msal'],
            $assemblies['System.Security.Cryptography.ProtectedData']
        ) | Where-Object { $_ }
        
        # Add System.Diagnostics.TraceSource if available
        if ($sdts -and $sdts -ne "System.Diagnostics.TraceSource") {
            $referencedAssemblies += $sdts
        }
        
        # Add standard assemblies
        $referencedAssemblies += @("netstandard", "System.Linq", "System.Threading.Tasks")
        
        # Resolve tenantId safely
        $hasClientId = ($script:StartupParameters -is [hashtable]) -and $script:StartupParameters.ContainsKey('ClientId') -and -not [string]::IsNullOrWhiteSpace($script:StartupParameters['ClientId'])
        $hasTenantId = ($script:StartupParameters -is [hashtable]) -and $script:StartupParameters.ContainsKey('TenantId') -and -not [string]::IsNullOrWhiteSpace($script:StartupParameters['TenantId'])

        # Get the access token with WAM
        $clientId = if ($hasClientId) {
            $script:StartupParameters['ClientId']   # custom app client id
        }
        else {
            "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # PowerShell public client
        }

        # Avoid accessing an unset script:CurrentTenantId
        $currentTenantAvailable = (Get-Variable -Name 'CurrentTenantId' -Scope Script -ErrorAction SilentlyContinue) -and -not [string]::IsNullOrWhiteSpace($script:CurrentTenantId)

        $tenantId = if ($hasTenantId) {
            $script:StartupParameters['TenantId']
        }
        elseif ($currentTenantAvailable) {
            $script:CurrentTenantId
        }
        else {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx -and $ctx.PSObject.Properties['TenantId'] -and $ctx.TenantId) { $ctx.TenantId }
            elseif ($ctx -and $ctx.PSObject.Properties['Tenant'] -and $ctx.Tenant) { $ctx.Tenant }
            else { throw "Could not determine tenant ID" }
        }

        if ([string]::IsNullOrWhiteSpace($clientId)) { throw "Client ID could not be determined and is null or empty" }
        if ([string]::IsNullOrWhiteSpace($tenantId)) { throw "Tenant ID could not be determined and is null or empty" }

        $redirectUri = "http://localhost"
        $scopes = @("https://graph.microsoft.com/.default")

        Write-Verbose "=== Attempting WAM authentication with claims ==="
        Write-Verbose "Client ID: $clientId"
        Write-Verbose "Redirect URI: $redirectUri"
        Write-Verbose "Scopes: $($scopes -join ', ')"
        Write-Verbose "Referenced assemblies count: $($referencedAssemblies.Count)"
        
        $tokenStartTime = Get-Date
        
        try {
            # Check if type already exists
            $existingType = [System.Type]::GetType("PIMAuthContextHelper")
            if ($existingType) {
                Write-Verbose "PIMAuthContextHelper type already exists, using it"
            }
            else {
                Write-Verbose "Adding PIMAuthContextHelper type"
                Add-Type -ReferencedAssemblies $referencedAssemblies -TypeDefinition $code -Language CSharp -ErrorAction Stop
                Write-Verbose "PIMAuthContextHelper type added successfully"
            }
            
            Write-Verbose "Calling PIMAuthContextHelper.GetAccessTokenWithAuthContext"
            $accessToken = [PIMAuthContextHelper]::GetAccessTokenWithAuthContext($clientId, $tenantId, $redirectUri, $scopes, $claimsJson)
            
            if (-not $accessToken) {
                throw "No access token returned from WAM authentication"
            }
        }
        catch [System.Reflection.ReflectionTypeLoadException] {
            $loaderExceptions = $_.Exception.LoaderExceptions
            foreach ($loaderException in $loaderExceptions) {
                Write-Verbose "Loader exception: $($loaderException.Message)"
            }
            throw "Failed to load types: $($_.Exception.Message)"
        }
        catch {
            Write-Verbose "WAM authentication error: $($_.Exception.Message)"
            Write-Verbose "Exception type: $($_.Exception.GetType().FullName)"
            
            # If it's a specific error about window handle or broker, provide more context
            if ($_.Exception.Message -like "*broker*" -or $_.Exception.Message -like "*window*") {
                Write-Warning "WAM broker authentication failed. This might be due to:"
                Write-Warning "- Running in a non-interactive session"
                Write-Warning "- WAM not being available on this system"
                Write-Warning "- Missing Windows updates"
            }
            
            throw $_
        }
        
        $tokenDuration = (Get-Date) - $tokenStartTime
        Write-Verbose "Successfully obtained authentication context token via WAM in $($tokenDuration.TotalSeconds) seconds"
        Write-Verbose "Token length: $($accessToken.Length)"
        
        # Cache the token for reuse (assume 45 minutes expiry for safety)
        $expiryTime = (Get-Date).AddMinutes(45)
        
        # Validate that the token contains the expected authentication context claim
        $isValidToken = Test-AuthenticationContextToken -AccessToken $accessToken -ExpectedContextId $ContextId
        if (-not $isValidToken) {
            Write-Warning "Authentication context token validation failed - token does not contain expected context claim: $ContextId"
            Write-Verbose "Token might still be valid - continuing anyway"
        }
        
        $script:AuthContextTokens[$ContextId] = @{
            AccessToken = $accessToken
            ExpiryTime  = $expiryTime
            ContextId   = $ContextId
        }
        
        Write-Verbose "Cached authentication context token for context: $ContextId (expires: $expiryTime)"
        Write-Verbose "=== Authentication context token acquisition completed successfully ==="
        return $accessToken
    }
    catch {
        $errorMessage = "Failed to obtain authentication context token for context $ContextId`: $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Verbose "Exception details:"
        Write-Verbose "  Type: $($_.Exception.GetType().FullName)"
        Write-Verbose "  Message: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Verbose "  Inner Exception: $($_.Exception.InnerException.Message)"
        }
        Write-Verbose "  Stack Trace: $($_.Exception.StackTrace)"
        
        # Return null to let the calling function handle the failure
        return $null
    }
}