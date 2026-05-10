function Test-AzureReducedScope {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OriginalScope,

        [Parameter(Mandatory)]
        [string]$TargetScope
    )

    $normalizeScope = {
        param([string]$Scope)

        if ([string]::IsNullOrWhiteSpace($Scope)) { return $null }

        $normalized = $Scope.Trim()
        if (-not $normalized.StartsWith('/')) { $normalized = "/$normalized" }

        while ($normalized.Length -gt 1 -and $normalized.EndsWith('/')) {
            $normalized = $normalized.Substring(0, $normalized.Length - 1)
        }

        return $normalized
    }

    $original = & $normalizeScope $OriginalScope
    $target = & $normalizeScope $TargetScope

    if (-not $original) {
        return [PSCustomObject]@{
            IsValid        = $false
            IsReducedScope = $false
            OriginalScope  = $null
            TargetScope    = $target
            ErrorMessage   = 'Original Azure role scope is missing.'
        }
    }

    if (-not $target) {
        return [PSCustomObject]@{
            IsValid        = $false
            IsReducedScope = $false
            OriginalScope  = $original
            TargetScope    = $target
            ErrorMessage   = 'Target scope is required when reduced scope is enabled.'
        }
    }

    if ($target -eq $original) {
        return [PSCustomObject]@{
            IsValid        = $true
            IsReducedScope = $false
            OriginalScope  = $original
            TargetScope    = $target
            ErrorMessage   = $null
        }
    }

    $knownScopePattern = '^(/|/subscriptions/[a-fA-F0-9\-]{36}($|/)|/providers/Microsoft\.Management/managementGroups/[^/]+($|/))'
    if ($target -notmatch $knownScopePattern) {
        return [PSCustomObject]@{
            IsValid        = $false
            IsReducedScope = $true
            OriginalScope  = $original
            TargetScope    = $target
            ErrorMessage   = "Target scope '$target' is not a recognized ARM scope."
        }
    }

    if ($original -eq '/') {
        return [PSCustomObject]@{
            IsValid        = $true
            IsReducedScope = $true
            OriginalScope  = $original
            TargetScope    = $target
            ErrorMessage   = $null
        }
    }

    if ($original -match '^/subscriptions/([a-fA-F0-9\-]{36})(/.*)?$') {
        $originalPrefix = $original + '/'
        $isDescendant = $target.StartsWith($originalPrefix, [System.StringComparison]::OrdinalIgnoreCase)
        if (-not $isDescendant) {
            return [PSCustomObject]@{
                IsValid        = $false
                IsReducedScope = $true
                OriginalScope  = $original
                TargetScope    = $target
                ErrorMessage   = "Target scope '$target' is not below original scope '$original'."
            }
        }
    }
    elseif ($original -match '^/providers/Microsoft\.Management/managementGroups/[^/]+$') {
        if ($target -notmatch '^(/subscriptions/[a-fA-F0-9\-]{36}($|/)|/providers/Microsoft\.Management/managementGroups/[^/]+($|/))') {
            return [PSCustomObject]@{
                IsValid        = $false
                IsReducedScope = $true
                OriginalScope  = $original
                TargetScope    = $target
                ErrorMessage   = "Target scope '$target' is not valid for a management group eligibility."
            }
        }
    }
    else {
        return [PSCustomObject]@{
            IsValid        = $false
            IsReducedScope = $true
            OriginalScope  = $original
            TargetScope    = $target
            ErrorMessage   = "Original scope '$original' is not a supported Azure Resource eligibility scope."
        }
    }

    return [PSCustomObject]@{
        IsValid        = $true
        IsReducedScope = $true
        OriginalScope  = $original
        TargetScope    = $target
        ErrorMessage   = $null
    }
}
