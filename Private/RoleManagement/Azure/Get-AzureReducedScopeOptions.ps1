function Get-AzureReducedScopeOptions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OriginalScope
    )

    $getPropertyValue = {
        param(
            [object]$InputObject,
            [string[]]$PropertyNames
        )

        if (-not $InputObject) { return $null }
        foreach ($propertyName in $PropertyNames) {
            if ($InputObject.PSObject.Properties[$propertyName] -and -not [string]::IsNullOrWhiteSpace([string]$InputObject.$propertyName)) {
                return [string]$InputObject.$propertyName
            }
        }

        return $null
    }

    $getNestedPropertyValue = {
        param(
            [object]$InputObject,
            [string[]]$PropertyNames
        )

        $directValue = & $getPropertyValue $InputObject $PropertyNames
        if ($directValue) { return $directValue }

        if ($InputObject -and $InputObject.PSObject.Properties['properties'] -and $InputObject.properties) {
            return & $getPropertyValue $InputObject.properties $PropertyNames
        }

        return $null
    }

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

    $getSubscriptionIdFromScope = {
        param([string]$Scope)

        if ($Scope -and $Scope -match '^/subscriptions/([^/]+)') { return $matches[1] }
        return $null
    }

    $getResourceGroupFromScope = {
        param([string]$Scope)

        if ($Scope -and $Scope -match '^/subscriptions/[^/]+/resourceGroups/([^/]+)') { return [System.Uri]::UnescapeDataString($matches[1]) }
        return $null
    }

    $getScopeDisplayName = {
        param(
            [string]$Scope,
            [object]$Entry,
            [string]$ResourceType
        )

        $displayName = & $getNestedPropertyValue $Entry @('displayName', 'resourceDisplayName', 'resourceName', 'name')
        if (-not [string]::IsNullOrWhiteSpace($displayName)) { return $displayName }

        if ($Scope -match '^/subscriptions/[^/]+/resourceGroups/([^/]+)/providers/[^/]+/[^/]+/([^/]+)$') {
            return [System.Uri]::UnescapeDataString($matches[2])
        }
        if ($Scope -match '^/subscriptions/[^/]+/resourceGroups/([^/]+)$') {
            return [System.Uri]::UnescapeDataString($matches[1])
        }
        if ($Scope -match '^/subscriptions/([^/]+)$') {
            return $matches[1]
        }
        if ($Scope -match '^/providers/Microsoft\.Management/managementGroups/([^/]+)$') {
            return $matches[1]
        }

        if (-not [string]::IsNullOrWhiteSpace($ResourceType)) { return $ResourceType }
        return $Scope
    }

    $original = & $normalizeScope $OriginalScope
    if ([string]::IsNullOrWhiteSpace($original)) { throw 'OriginalScope is required when listing Azure reduced-scope options.' }

    $tokenObj = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop
    $token = if ($tokenObj.Token -is [System.Security.SecureString]) {
        [System.Net.NetworkCredential]::new('', $tokenObj.Token).Password
    }
    else {
        $tokenObj.Token
    }

    $headers = @{
        Authorization = "Bearer $token"
        'Content-Type' = 'application/json'
    }

    $items = [System.Collections.ArrayList]::new()
    $seenScopes = @{}
    $scopePath = if ($original -eq '/') { '' } else { $original }
    $uri = "https://management.azure.com$scopePath/providers/Microsoft.Authorization/eligibleChildResources?api-version=2020-10-01-preview"

    Write-Verbose "Get-AzureReducedScopeOptions: listing eligible child resources for original scope '$original'."

    while ($uri) {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        foreach ($entry in @($response.value)) {
            $scope = $null
            if ($entry.PSObject.Properties['properties'] -and $entry.properties) {
                $scope = & $getPropertyValue $entry.properties @('scope', 'resourceId', 'id')
            }
            if ([string]::IsNullOrWhiteSpace($scope)) {
                $scope = & $getPropertyValue $entry @('scope', 'resourceId', 'id')
            }
            $scope = & $normalizeScope $scope
            if ([string]::IsNullOrWhiteSpace($scope) -or $scope.Equals($original, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($seenScopes.ContainsKey($scope)) { continue }
            $seenScopes[$scope] = $true

            $resourceType = & $getNestedPropertyValue $entry @('resourceType', 'type')
            $displayName = & $getScopeDisplayName $scope $entry $resourceType
            if (-not [string]::IsNullOrWhiteSpace($resourceType) -and $displayName -notlike "*($resourceType)") {
                $displayName = "$displayName ($resourceType)"
            }

            $null = $items.Add([PSCustomObject]@{
                DisplayName    = $displayName
                Name           = $displayName
                SubscriptionId = & $getSubscriptionIdFromScope $scope
                ResourceGroup  = & $getResourceGroupFromScope $scope
                ResourceId     = $scope
                Scope          = $scope
                Type           = $resourceType
            })
        }

        $uri = if ($response.PSObject.Properties['nextLink'] -and $response.nextLink) {
            if ($response.nextLink -match '^https?://') { $response.nextLink } else { "https://management.azure.com$($response.nextLink)" }
        }
        else {
            $null
        }
    }

    Write-Verbose "Get-AzureReducedScopeOptions: returning $($items.Count) eligible child resource option(s)."

    return @($items | Sort-Object DisplayName)
}
