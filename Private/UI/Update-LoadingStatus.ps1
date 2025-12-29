function Update-LoadingStatus {
    <#
    .SYNOPSIS
        Updates the splash screen status and progress.
    
    .PARAMETER SplashForm
        The splash screen control object returned by Show-LoadingSplash.
    
    .PARAMETER Status
        New status message to display.
    
    .PARAMETER Progress
        Progress percentage (0-100). Optional.
    
    .EXAMPLE
        Update-LoadingStatus -SplashForm $splash -Status "Processing..." -Progress 75
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$SplashForm,
        
        [Parameter(Mandatory)]
        [string]$Status,

        [int]$Progress = -1
    )

    if ($SplashForm -and $SplashForm.SyncHash -and -not $SplashForm.IsDisposed) {
        $SplashForm.UpdateStatus($Status, $Progress)
    }
}