function Get-EffectiveDuration {
    param(
        [int]$RequestedMinutes,
        [int]$MaxDurationHours
    )
    
    $maxMinutes = $MaxDurationHours * 60
    
    if ($RequestedMinutes -gt $maxMinutes) {
        Write-Verbose "Requested duration ($RequestedMinutes minutes) exceeds maximum ($maxMinutes minutes)"
        $hours = [Math]::Floor($maxMinutes / 60)
        $minutes = $maxMinutes % 60
    }
    else {
        $hours = [Math]::Floor($RequestedMinutes / 60)
        $minutes = $RequestedMinutes % 60
    }
    
    return @{
        Hours        = $hours
        Minutes      = $minutes
        TotalMinutes = ($hours * 60) + $minutes
    }
}