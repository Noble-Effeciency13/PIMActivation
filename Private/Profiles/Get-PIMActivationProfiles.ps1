function Get-PIMActivationProfiles {
    <#
    .SYNOPSIS
        Retrieves saved PIM activation profiles from the local user profile.
    #>
    [CmdletBinding()]
    param()

    $profilePath = Get-PIMActivationProfileStorePath
    if (-not (Test-Path -Path $profilePath)) {
        return @()
    }

    $profiles = [System.Collections.ArrayList]::new()
    foreach ($file in @(Get-ChildItem -Path $profilePath -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        try {
            $profile = Get-Content -Path $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if (-not $profile.PSObject.Properties['FilePath']) {
                $profile | Add-Member -MemberType NoteProperty -Name FilePath -Value $file.FullName
            }
            $null = $profiles.Add($profile)
        }
        catch {
            Write-Warning "Failed to read activation profile '$($file.Name)': $($_.Exception.Message)"
        }
    }

    return @($profiles | Sort-Object Name)
}