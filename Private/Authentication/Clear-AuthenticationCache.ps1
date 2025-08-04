function Clear-AuthenticationCache {
    <#
    .SYNOPSIS
        Clears all cached authentication tokens to force fresh authentication.
    
    .DESCRIPTION
        Removes cached tokens from MSAL cache and provides guidance for browser cookie cleanup
        to ensure a completely fresh authentication flow. This function helps resolve authentication
        issues by clearing stored credentials and tokens.
    
    .EXAMPLE
        Clear-AuthenticationCache
        Clears all authentication caches and returns $true if successful.
    
    .OUTPUTS
        System.Boolean
        Returns $true if cache clearing was successful, $false otherwise.
    
    .NOTES
        - MSAL token cache is automatically cleared from the local PowerShell cache folder
        - Browser cookies may need to be cleared manually as they cannot be safely removed while browsers are running
        - Consider closing all browser instances before running this function for complete cleanup
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    try {
        Write-Verbose "Starting authentication cache cleanup..."
        
        # Clear module-level authentication context token cache
        if ($script:AuthContextTokens) {
            $tokenCount = $script:AuthContextTokens.Count
            $script:AuthContextTokens.Clear()
            Write-Verbose "Cleared $tokenCount cached authentication context tokens"
        }
        
        # Clear legacy authentication context variables
        $script:CurrentAuthContextToken = $null
        $script:JustCompletedAuthContext = $null
        $script:AuthContextCompletionTime = $null
        
        # Clear MSAL token cache
        try {
            $cacheFolder = Join-Path $env:LOCALAPPDATA "Microsoft\PowerShell\TokenCache"
            if (Test-Path $cacheFolder) {
                Remove-Item -Path "$cacheFolder\*" -Force -Recurse -ErrorAction SilentlyContinue
                Write-Verbose "Successfully cleared MSAL token cache from: $cacheFolder"
            } else {
                Write-Verbose "MSAL token cache folder not found - no cache to clear"
            }
        }
        catch {
            Write-Warning "Failed to clear MSAL token cache: $($_.Exception.Message)"
        }
        
        # Provide browser cookie guidance
        try {
            $edgeCookies = Join-Path $env:LOCALAPPDATA "Microsoft\Edge\User Data\Default\Cookies"
            $chromeCookies = Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data\Default\Cookies"
            
            $browserPaths = [System.Collections.ArrayList]::new()
            if (Test-Path $edgeCookies) { $null = $browserPaths.Add("Edge") }
            if (Test-Path $chromeCookies) { $null = $browserPaths.Add("Chrome") }
            
            if ($browserPaths.Count -gt 0) {
                Write-Verbose "Browser cookie stores detected for: $($browserPaths -join ', ')"
                Write-Verbose "Recommendation: Close all browser instances and clear Microsoft/Azure cookies manually for complete cleanup"
            }
        }
        catch {
            Write-Verbose "Unable to check browser cookie locations: $($_.Exception.Message)"
        }
        
        Write-Verbose "Authentication cache cleanup completed successfully"
        return $true
    }
    catch {
        Write-Error "Failed to clear authentication cache: $($_.Exception.Message)"
        return $false
    }
}