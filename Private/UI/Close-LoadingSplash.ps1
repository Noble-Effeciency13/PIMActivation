function Close-LoadingSplash {
    <#
    .SYNOPSIS
        Closes the loading splash screen and cleans up resources.
    
    .PARAMETER SplashForm
        The splash screen control object returned by Show-LoadingSplash.
    
    .EXAMPLE
        Close-LoadingSplash -SplashForm $splash
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$SplashForm
    )

    if ($SplashForm -and -not $SplashForm.IsDisposed) {
        $SplashForm.Close()
    }
}