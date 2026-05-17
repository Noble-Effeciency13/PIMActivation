function Show-ActivationResults {
    <#
    .SYNOPSIS
        Displays a summary of activation request results.

    .DESCRIPTION
        Shows a top-most message box summarizing successful and failed activation requests. When a
        scheduled start time is provided, the summary uses scheduled activation wording instead of
        implying the roles are already active.

    .PARAMETER SuccessCount
        Number of activation requests that were submitted successfully.

    .PARAMETER TotalCount
        Total number of requested role activations.

    .PARAMETER Errors
        Error messages for role activation requests that failed.

    .PARAMETER ScheduledStartTime
        Optional local start time for scheduled activation requests.
    #>
    [CmdletBinding()]
    param(
        [int]$SuccessCount,
        [int]$TotalCount,
        [array]$Errors,
        [datetime]$ScheduledStartTime
    )

    $isScheduled = $PSBoundParameters.ContainsKey('ScheduledStartTime') -and $ScheduledStartTime
    $actionText = if ($isScheduled) { 'scheduled' } else { 'activated' }
    $scheduleSuffix = if ($isScheduled) { " for $($ScheduledStartTime.ToString('yyyy-MM-dd HH:mm'))" } else { '' }
    
    if ($Errors.Count -gt 0) {
        $message = "Successfully $actionText $SuccessCount of $TotalCount role(s)$scheduleSuffix.`n`nErrors:`n$($Errors -join "`n")"
        Show-TopMostMessageBox -Message $message -Title "Activation Results" -Icon Warning
    }
    else {
        Show-TopMostMessageBox -Message "Successfully $actionText all $SuccessCount role(s)$scheduleSuffix!" -Title "Success" -Icon Information
    }
}