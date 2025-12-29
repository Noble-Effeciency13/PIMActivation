function Get-AuthenticationContextsBatch {
    <#
    .SYNOPSIS
        Retrieves authentication contexts in batch for better performance.
    
    .PARAMETER ContextIds
        Array of authentication context IDs to fetch.
    
    .PARAMETER ContextCache
        Hashtable to store the fetched contexts in.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ContextIds,
        
        [Parameter(Mandatory)]
        [hashtable]$ContextCache
    )
    
    # Ensure we have arrays to work with for input parameters
    if (-not $ContextIds) {
        $ContextIds = @()
    } elseif ($ContextIds -isnot [array]) {
        $ContextIds = @($ContextIds)
    }
    
    Write-Verbose "Batch fetching $($ContextIds.Count) authentication contexts"
    
    foreach ($contextId in $ContextIds) {
        try {
            # Skip if already cached locally
            if ($ContextCache.ContainsKey($contextId)) {
                continue
            }
            
            # Check script-level cache first
            if ($script:AuthenticationContextCache.ContainsKey($contextId)) {
                Write-Verbose "Using cached authentication context: $contextId"
                $ContextCache[$contextId] = $script:AuthenticationContextCache[$contextId]
                continue
            }
            
            # Fetch from Graph API if not in any cache
            $context = Get-MgIdentityConditionalAccessAuthenticationContextClassReference -AuthenticationContextClassReferenceId $contextId -ErrorAction Stop
            if ($context) {
                $ContextCache[$contextId] = $context
                # Also cache in script-level cache for future use
                $script:AuthenticationContextCache[$contextId] = $context
                Write-Verbose "Cached authentication context: $contextId - $($context.DisplayName)"
            }
        }
        catch {
            # Suppress 403 Forbidden errors (common when the user lacks access); log as verbose only
            $errMsg = $_.Exception.Message
            $statusCode = $null
            try {
                if ($_.Exception.PSObject.Properties["ResponseStatusCode"]) { $statusCode = $_.Exception.ResponseStatusCode }
                elseif ($_.Exception.PSObject.Properties["StatusCode"]) { $statusCode = $_.Exception.StatusCode }
                elseif ($_.PSObject.Properties["CategoryInfo"]) {
                    # Some Graph exceptions surface status in the message
                    if ($errMsg -match "403|Forbidden") { $statusCode = 403 }
                }
            } catch { }

            if ($statusCode -eq 403 -or ($errMsg -match "403|Forbidden")) {
                Write-Verbose "Authentication context $contextId not accessible (403 Forbidden). Suppressing warning."
            }
            else {
                Write-Verbose "Failed to fetch authentication context $contextId : $errMsg"
            }
            continue
        }
    }
    
    Write-Verbose "Completed batch authentication context fetch"
}