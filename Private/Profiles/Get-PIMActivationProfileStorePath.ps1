function Get-PIMActivationProfileStorePath {
    [CmdletBinding()]
    param()

    $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        $localAppData = $env:LOCALAPPDATA
    }
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        throw 'Unable to resolve the local application data folder for activation profiles.'
    }

    return (Join-Path $localAppData 'PIMActivation\ActivationProfiles')
}
