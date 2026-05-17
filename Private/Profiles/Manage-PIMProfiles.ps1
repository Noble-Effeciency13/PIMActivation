function Manage-PIMProfiles {
    <#
    .SYNOPSIS
        Performs simple activation profile management actions.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Get', 'Delete')]
        [string]$Action = 'Get',

        [string]$ProfileName
    )

    switch ($Action) {
        'Get' {
            return Get-PIMActivationProfiles
        }
        'Delete' {
            if ([string]::IsNullOrWhiteSpace($ProfileName)) {
                throw 'ProfileName is required when deleting an activation profile.'
            }

            $profile = Get-PIMActivationProfiles | Where-Object { $_.Name -eq $ProfileName } | Select-Object -First 1
            if (-not $profile -or -not $profile.PSObject.Properties['FilePath'] -or -not (Test-Path -Path $profile.FilePath)) {
                return $false
            }

            Remove-Item -Path $profile.FilePath -Force
            return $true
        }
    }
}