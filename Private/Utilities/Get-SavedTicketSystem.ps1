function Get-SavedTicketSystem {
    <#
    .SYNOPSIS
        Retrieves the last used ticket system preference.
    
    .DESCRIPTION
        Gets the saved ticket system preference from the user's local profile.
    
    .OUTPUTS
        String - The saved ticket system name, or $null if not found.
    #>
    [CmdletBinding()]
    param()
    
    $prefsPath = Join-Path $env:LOCALAPPDATA "PIMActivation\preferences.json"
    
    if (Test-Path $prefsPath) {
        try {
            $prefs = Get-Content $prefsPath -Raw | ConvertFrom-Json
            return $prefs.TicketSystem
        }
        catch {
            Write-Verbose "Failed to load ticket system preference: $_"
        }
    }
    
    return $null
}