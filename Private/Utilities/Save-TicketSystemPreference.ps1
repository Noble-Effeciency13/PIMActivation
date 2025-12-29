function Save-TicketSystemPreference {
    <#
    .SYNOPSIS
        Saves the user's ticket system preference.
    
    .PARAMETER System
        The ticket system name to save.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$System
    )
    
    $prefsDir = Join-Path $env:LOCALAPPDATA "PIMActivation"
    $prefsPath = Join-Path $prefsDir "preferences.json"
    
    # Ensure directory exists
    if (-not (Test-Path $prefsDir)) {
        New-Item -ItemType Directory -Path $prefsDir -Force | Out-Null
    }
    
    try {
        # Load existing preferences or create new
        $prefs = if (Test-Path $prefsPath) {
            Get-Content $prefsPath -Raw | ConvertFrom-Json
        }
        else {
            [PSCustomObject]@{}
        }
        
        # Update ticket system
        $prefs | Add-Member -NotePropertyName TicketSystem -NotePropertyValue $System -Force
        
        # Save
        $prefs | ConvertTo-Json | Set-Content $prefsPath -Force
        Write-Verbose "Saved ticket system preference: $System"
    }
    catch {
        Write-Warning "Failed to save ticket system preference: $_"
    }
}