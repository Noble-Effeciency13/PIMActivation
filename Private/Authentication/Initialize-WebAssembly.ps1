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