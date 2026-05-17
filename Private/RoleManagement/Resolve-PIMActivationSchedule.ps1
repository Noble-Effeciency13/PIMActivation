function Resolve-PIMActivationSchedule {
    <#
    .SYNOPSIS
        Resolves and validates a PIM activation start time.

    .DESCRIPTION
        Converts a requested local activation start time to UTC and calculates the scheduling
        window for the selected roles when eligibility start and end metadata is available.
        The returned object is used by the activation dialog and request builders so regular
        activations and activation-profile launches use the same scheduling rules.

    .PARAMETER RoleItems
        The selected role ListView items or role objects. When a ListView item is supplied, its
        Tag property is used as the role data source.

    .PARAMETER RequestedDuration
        Hashtable containing the requested activation duration. Supports Hours, Minutes, and
        TotalMinutes keys.

    .PARAMETER ScheduleStartTime
        The local date and time selected for the activation start.

    .PARAMETER Scheduled
        When specified, validates ScheduleStartTime as a future scheduled activation. When omitted,
        the function returns the current schedule window without requiring a future start.

    .EXAMPLE
        Resolve-PIMActivationSchedule -RoleItems $checkedItems -RequestedDuration $duration -ScheduleStartTime $start -Scheduled

        Validates a requested scheduled start time and returns UTC request metadata.

    .OUTPUTS
        PSCustomObject
        Returns IsValid, ErrorMessage, StartLocal, StartUtc, StartUtcString, MinStartLocal,
        MaxStartLocal, and MaxStartRoleName properties.
    #>
    [CmdletBinding()]
    param(
        [object[]]$RoleItems = @(),

        [hashtable]$RequestedDuration,

        [datetime]$ScheduleStartTime,

        [switch]$Scheduled
    )

    $nowLocal = Get-Date
    $nowUtc = $nowLocal.ToUniversalTime()
    $requestedTotalMinutes = 480

    if ($RequestedDuration) {
        if ($RequestedDuration.ContainsKey('TotalMinutes') -and $null -ne $RequestedDuration.TotalMinutes) {
            $requestedTotalMinutes = [int]$RequestedDuration.TotalMinutes
        }
        else {
            $hours = if ($RequestedDuration.ContainsKey('Hours') -and $null -ne $RequestedDuration.Hours) { [int]$RequestedDuration.Hours } else { 8 }
            $minutes = if ($RequestedDuration.ContainsKey('Minutes') -and $null -ne $RequestedDuration.Minutes) { [int]$RequestedDuration.Minutes } else { 0 }
            $requestedTotalMinutes = ($hours * 60) + $minutes
        }
    }

    $toUtc = {
        param($Value)

        if ($null -eq $Value) { return $null }
        if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) { return $null }

        try {
            if ($Value -is [System.DateTimeOffset]) {
                return $Value.UtcDateTime
            }

            if ($Value -is [datetime]) {
                $dateTimeValue = [datetime]$Value
                if ($dateTimeValue.Kind -eq [System.DateTimeKind]::Unspecified) {
                    $dateTimeValue = [datetime]::SpecifyKind($dateTimeValue, [System.DateTimeKind]::Local)
                }
                return $dateTimeValue.ToUniversalTime()
            }

            $offsetValue = [System.DateTimeOffset]::Parse(
                [string]$Value,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::AssumeUniversal
            )
            return $offsetValue.UtcDateTime
        }
        catch {
            return $null
        }
    }

    $getRoleData = {
        param($Item)

        if (-not $Item) { return $null }
        if ($Item.GetType().FullName -eq 'System.Windows.Forms.ListViewItem' -and $Item.PSObject.Properties['Tag']) {
            return $Item.Tag
        }
        if ($Item.PSObject.Properties['Tag'] -and $Item.Tag) {
            return $Item.Tag
        }
        return $Item
    }

    $getRoleName = {
        param($RoleData)

        if (-not $RoleData) { return 'Selected role' }
        if ($RoleData.PSObject.Properties['DisplayName'] -and $RoleData.DisplayName) { return [string]$RoleData.DisplayName }
        if ($RoleData.PSObject.Properties['Name'] -and $RoleData.Name) { return [string]$RoleData.Name }
        if ($RoleData.PSObject.Properties['RoleName'] -and $RoleData.RoleName) { return [string]$RoleData.RoleName }
        return 'Selected role'
    }

    $getDateCandidate = {
        param(
            $RoleData,
            [string]$PropertyName,
            [string]$NestedPropertyName
        )

        if (-not $RoleData) { return $null }
        if ($RoleData.PSObject.Properties[$PropertyName] -and $RoleData.$PropertyName) {
            return $RoleData.$PropertyName
        }

        if ($RoleData.PSObject.Properties['Assignment'] -and $RoleData.Assignment) {
            $assignment = $RoleData.Assignment
            if ($assignment.PSObject.Properties[$PropertyName] -and $assignment.$PropertyName) {
                return $assignment.$PropertyName
            }
            if ($assignment.PSObject.Properties['properties'] -and $assignment.properties) {
                $properties = $assignment.properties
                if ($properties.PSObject.Properties[$NestedPropertyName] -and $properties.$NestedPropertyName) {
                    return $properties.$NestedPropertyName
                }
            }
        }

        return $null
    }

    $minStartUtc = $nowUtc
    $maxStartUtc = $null
    $maxStartRoleName = $null

    foreach ($item in @($RoleItems)) {
        $roleData = & $getRoleData $item
        if (-not $roleData) { continue }

        $roleName = & $getRoleName $roleData
        $roleStartUtc = & $toUtc (& $getDateCandidate $roleData 'StartDateTime' 'startDateTime')
        $roleEndUtc = & $toUtc (& $getDateCandidate $roleData 'EndDateTime' 'endDateTime')

        if ($roleStartUtc -and $roleStartUtc -gt $minStartUtc) {
            $minStartUtc = $roleStartUtc
        }

        if ($roleEndUtc) {
            $maxDurationHours = 8
            if ($roleData.PSObject.Properties['PolicyInfo'] -and $roleData.PolicyInfo -and $roleData.PolicyInfo.PSObject.Properties['MaxDuration'] -and $roleData.PolicyInfo.MaxDuration) {
                try { $maxDurationHours = [int]$roleData.PolicyInfo.MaxDuration } catch { $maxDurationHours = 8 }
            }

            $effectiveDuration = Get-EffectiveDuration -RequestedMinutes $requestedTotalMinutes -MaxDurationHours $maxDurationHours
            $effectiveMinutes = if ($effectiveDuration.ContainsKey('TotalMinutes')) { [int]$effectiveDuration.TotalMinutes } else { ([int]$effectiveDuration.Hours * 60) + [int]$effectiveDuration.Minutes }
            $roleMaxStartUtc = $roleEndUtc.AddMinutes(-1 * $effectiveMinutes)

            if (-not $maxStartUtc -or $roleMaxStartUtc -lt $maxStartUtc) {
                $maxStartUtc = $roleMaxStartUtc
                $maxStartRoleName = $roleName
            }
        }
    }

    $selectedLocal = if ($PSBoundParameters.ContainsKey('ScheduleStartTime')) { [datetime]$ScheduleStartTime } else { $nowLocal }
    if ($selectedLocal.Kind -eq [System.DateTimeKind]::Unspecified) {
        $selectedLocal = [datetime]::SpecifyKind($selectedLocal, [System.DateTimeKind]::Local)
    }
    else {
        $selectedLocal = $selectedLocal.ToLocalTime()
    }

    $selectedUtc = $selectedLocal.ToUniversalTime()
    $isValid = $true
    $errorMessage = ''

    if ($Scheduled) {
        if ($selectedUtc -lt $nowUtc.AddMinutes(-1)) {
            $isValid = $false
            $errorMessage = 'Scheduled activation start time cannot be in the past.'
        }
        elseif ($selectedUtc -lt $minStartUtc.AddMinutes(-1)) {
            $isValid = $false
            $errorMessage = "Scheduled activation start time is before the selected role eligibility window opens. Earliest start: $($minStartUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm'))."
        }
        elseif ($maxStartUtc -and $selectedUtc -gt $maxStartUtc) {
            $isValid = $false
            $roleSuffix = if ($maxStartRoleName) { " for '$maxStartRoleName'" } else { '' }
            $errorMessage = "Scheduled activation start time is outside the selected role eligibility window$roleSuffix. Latest start: $($maxStartUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm'))."
        }
    }

    [PSCustomObject]@{
        IsScheduled      = [bool]$Scheduled
        IsValid          = $isValid
        ErrorMessage     = $errorMessage
        StartLocal       = $selectedLocal
        StartUtc         = $selectedUtc
        StartUtcString   = $selectedUtc.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
        MinStartLocal    = $minStartUtc.ToLocalTime()
        MaxStartLocal    = if ($maxStartUtc) { $maxStartUtc.ToLocalTime() } else { $null }
        MaxStartRoleName = $maxStartRoleName
    }
}