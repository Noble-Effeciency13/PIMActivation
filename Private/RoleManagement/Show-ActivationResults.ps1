function Show-ActivationResults {
    param(
        [int]$SuccessCount,
        [int]$TotalCount,
        [array]$Errors
    )
    
    if ($Errors.Count -gt 0) {
        $message = "Successfully activated $SuccessCount of $TotalCount role(s).`n`nErrors:`n$($Errors -join "`n")"
        Show-TopMostMessageBox -Message $message -Title "Activation Results" -Icon Warning
    }
    else {
        Show-TopMostMessageBox -Message "Successfully activated all $SuccessCount role(s)!" -Title "Success" -Icon Information
    }
}