function Test-PIMModuleCompatibility {
    <#
    .SYNOPSIS
        Tests if the current module combination is compatible
    
    .DESCRIPTION
        Performs a quick compatibility test to verify the loaded modules work together
    
    .OUTPUTS
        Boolean indicating compatibility
    #>
    [CmdletBinding()]
    param()
    
    try {
        # Ensure required modules are loaded
        $authLoaded = Import-PIMModule -ModuleName 'Microsoft.Graph.Authentication'
        if (-not $authLoaded) {
            return $false
        }
        
        # Test the problematic method signature
        try {
            # This will fail if there's a signature mismatch
            Connect-MgGraph -Identity -ErrorAction Stop 2>$null
        }
        catch {
            if ($_.Exception.Message -like "*Method not found*AuthenticateAsync*") {
                Write-Warning "Module compatibility issue detected: AuthenticateAsync method signature mismatch"
                return $false
            }
            elseif ($_.Exception.Message -like "*No account*" -or $_.Exception.Message -like "*identity*") {
                # Expected error - method signatures are compatible
                return $true
            }
            else {
                Write-Verbose "Unexpected error during compatibility test: $($_.Exception.Message)"
                return $true  # Assume compatible if it's not the signature issue
            }
        }
        
        return $true
        
    }
    catch {
        Write-Warning "Compatibility test failed: $($_.Exception.Message)"
        return $false
    }
}