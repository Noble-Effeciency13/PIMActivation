function Invoke-PIMRoleDeactivation {
    <#
    .SYNOPSIS
        Deactivates selected active PIM roles.
    
    .DESCRIPTION
        Handles the deactivation of active PIM roles including:
        - Both Entra ID directory roles and PIM-enabled groups
        - Progress tracking with splash screen
        - Comprehensive error handling
    
    .PARAMETER CheckedItems
        Array of checked ListView items representing the active roles to deactivate.
    
    .PARAMETER Form
        Reference to the main form for UI updates.
    
    .EXAMPLE
        Invoke-PIMRoleDeactivation -CheckedItems $selectedRoles -Form $mainForm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$CheckedItems,
        
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form
    )
    
    Write-Verbose "Starting deactivation process for $($CheckedItems.Count) role(s)"
    
    # Initialize splash variable
    $operationSplash = $null
    $getGroupAccessId = {
        param($RoleData)

        $accessCandidates = @()
        if ($RoleData.PSObject.Properties['AccessId'] -and $RoleData.AccessId) {
            $accessCandidates += $RoleData.AccessId
        }
        if ($RoleData.PSObject.Properties['Assignment'] -and $RoleData.Assignment -and $RoleData.Assignment.PSObject.Properties['AccessId'] -and $RoleData.Assignment.AccessId) {
            $accessCandidates += $RoleData.Assignment.AccessId
        }
        if ($RoleData.PSObject.Properties['MemberType'] -and $RoleData.MemberType) {
            $accessCandidates += $RoleData.MemberType
        }
        if ($RoleData.PSObject.Properties['MembershipType'] -and $RoleData.MembershipType) {
            $accessCandidates += $RoleData.MembershipType
        }

        foreach ($candidate in $accessCandidates) {
            $normalizedAccessId = ([string]$candidate).Trim().ToLowerInvariant()
            if ($normalizedAccessId -in @('member', 'owner')) { return $normalizedAccessId }
        }

        return 'member'
    }
    
    try {
        # Confirm deactivation first (before showing splash)
        $roleNames = @($CheckedItems | ForEach-Object { 
            if ($_.Tag.Scope -and $_.Tag.Scope -ne "Directory") {
                "$($_.Tag.DisplayName) [$($_.Tag.Scope)]"
            }
            else {
                $_.Tag.DisplayName
            }
        })
        $message = "Are you sure you want to deactivate the following role(s)?`n`n$($roleNames -join "`n")"
        
        $result = Show-TopMostMessageBox -Message $message -Title "Confirm Deactivation" -Buttons YesNo -Icon Question
        
        if ($result -ne 'Yes') {
            Write-Verbose "Deactivation cancelled by user"
            return
        }
        
        # Show operation splash AFTER user confirms
        $operationSplash = Show-OperationSplash -Title "Role Deactivation" -InitialMessage "Preparing role deactivation..." -ShowProgressBar $true
        
        # ── Build deactivation job list ──────────────────────────────────────────
        $deactivationErrors      = @()
        $successCount            = 0
        $totalRoles              = @($CheckedItems).Count
        $currentRole             = 0
        # Track role data for successfully-submitted deactivations so we can suppress
        # them in the active list during Graph/ARM propagation lag.
        $successfullyDeactivated = [System.Collections.ArrayList]::new()

        # Pre-resolve any missing ScheduleIds (must be done sequentially before batching)
        $operationSplash.UpdateStatus("Resolving active schedule IDs...", 10)
        $resolvedJobs = [System.Collections.ArrayList]::new()

        foreach ($item in $CheckedItems) {
            $roleData = $item.Tag
            $jobEntry = [PSCustomObject]@{
                Item      = $item
                RoleData  = $roleData
                ScheduleId = $null
                Error     = $null
            }

            try {
                switch ($roleData.Type) {
                    'Entra' {
                        $sid = if ($roleData.ScheduleId) { $roleData.ScheduleId } else {
                            $active = @(Get-MgRoleManagementDirectoryRoleAssignmentSchedule `
                                -Filter "principalId eq '$($script:CurrentUser.Id)' and roleDefinitionId eq '$($roleData.RoleDefinitionId)'" `
                                -ErrorAction SilentlyContinue)
                            if ($active -and $active.Count -gt 0) { $active[0].Id }
                            else { throw "Could not find active assignment schedule for: $($roleData.DisplayName)" }
                        }
                        $jobEntry.ScheduleId = $sid
                    }
                    'Group' {
                        if (-not $roleData.GroupId) { throw "Missing GroupId for group deactivation: $($roleData.DisplayName)" }
                        $accessId = & $getGroupAccessId $roleData
                        $sid = if ($roleData.ScheduleId) { $roleData.ScheduleId } else {
                            $active = @(Get-MgIdentityGovernancePrivilegedAccessGroupAssignmentSchedule `
                                -Filter "principalId eq '$($script:CurrentUser.Id)' and groupId eq '$($roleData.GroupId)' and accessId eq '$accessId'" `
                                -ErrorAction SilentlyContinue)
                            if ($active -and $active.Count -gt 0) { $active[0].Id }
                            else { throw "Could not find active group assignment schedule for: $($roleData.DisplayName). It may not currently be active." }
                        }
                        $jobEntry.ScheduleId = $sid
                    }
                    'AzureResource' {
                        if (-not $roleData.RoleDefinitionId -or -not $roleData.FullScope) {
                            throw "Missing Azure Resource role details for deactivation: $($roleData.DisplayName)"
                        }
                        # Azure Resource deactivation does not require a pre-fetched ScheduleId
                    }
                    default { throw "Unsupported role type: $($roleData.Type)" }
                }
            }
            catch {
                $jobEntry.Error = if ($_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
                Write-Warning "Pre-resolution failed for $($roleData.DisplayName): $($jobEntry.Error)"
            }

            $null = $resolvedJobs.Add($jobEntry)
        }

        # Separate into error / Graph / Azure batches
        $graphJobs = @($resolvedJobs | Where-Object { -not $_.Error -and $_.RoleData.Type -in @('Entra', 'Group') })
        $azureJobs = @($resolvedJobs | Where-Object { -not $_.Error -and $_.RoleData.Type -eq 'AzureResource' })

        # Add pre-resolution failures directly to deactivationErrors
        foreach ($failed in @($resolvedJobs | Where-Object { $_.Error })) {
            $currentRole++
            $deactivationErrors += "$($failed.RoleData.DisplayName): $($failed.Error)"
        }

        # ── Graph roles: parallel REST (fall back to $batch if no Graph token) ─
        if ($graphJobs.Count -gt 0) {
            $operationSplash.UpdateStatus("Submitting $($graphJobs.Count) Graph deactivation(s) in parallel...", 30)

            # Try to acquire a Graph access token from the Microsoft.Graph SDK's
            # own session (it carries the PIM scopes granted by Connect-MgGraph;
            # an Az-acquired token would lack those scopes and return
            # PermissionScopeNotGranted). If unavailable, fall through to $batch.
            $graphTokDeact = $null
            try {
                $probe = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id' `
                    -OutputType HttpResponseMessage -ErrorAction Stop
                if ($probe -and $probe.RequestMessage -and $probe.RequestMessage.Headers -and
                    $probe.RequestMessage.Headers.Authorization -and $probe.RequestMessage.Headers.Authorization.Parameter) {
                    $graphTokDeact = $probe.RequestMessage.Headers.Authorization.Parameter
                }
            } catch {
                Write-Verbose "Could not extract Graph SDK token, will use Invoke-MgGraphRequest `$batch: $($_.Exception.Message)"
            }

            $currentUserIdGraph = $script:CurrentUser.Id

            if ($graphTokDeact) {
                # Build per-job request specs for parallel execution
                $graphSpecs = [System.Collections.ArrayList]::new()
                foreach ($job in $graphJobs) {
                    $rd = $job.RoleData
                    if ($rd.Type -eq 'Entra') {
                        $spec = [PSCustomObject]@{
                            Job  = $job
                            Uri  = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests'
                            Body = @{
                                action                   = 'selfDeactivate'
                                principalId              = $currentUserIdGraph
                                roleDefinitionId         = $rd.RoleDefinitionId
                                directoryScopeId         = if ($rd.DirectoryScopeId) { $rd.DirectoryScopeId } else { '/' }
                                roleAssignmentScheduleId = $job.ScheduleId
                                justification            = 'Deactivated via PowerShell'
                            }
                        }
                    } else {
                        $accessId = & $getGroupAccessId $rd
                        $spec = [PSCustomObject]@{
                            Job  = $job
                            Uri  = 'https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentScheduleRequests'
                            Body = @{
                                action               = 'selfDeactivate'
                                principalId          = $currentUserIdGraph
                                groupId              = $rd.GroupId
                                accessId             = $accessId
                                assignmentScheduleId = $job.ScheduleId
                                justification        = 'Deactivated via PowerShell'
                            }
                        }
                    }
                    $null = $graphSpecs.Add($spec)
                }

                $graphResults = $graphSpecs | ForEach-Object -ThrottleLimit 5 -Parallel {
                    $spec    = $_
                    $tok     = $using:graphTokDeact
                    $bodyTxt = $spec.Body | ConvertTo-Json -Depth 8 -Compress
                    $headers = @{
                        Authorization  = "Bearer $tok"
                        'Content-Type' = 'application/json'
                    }
                    try {
                        $resp = Invoke-RestMethod -Method Post -Uri $spec.Uri -Headers $headers -Body $bodyTxt -ErrorAction Stop
                        [PSCustomObject]@{
                            Job      = $spec.Job
                            Success  = $true
                            Response = $resp
                            Error    = $null
                        }
                    } catch {
                        # Try to extract Graph error.message from the response body
                        $errMsg = $_.Exception.Message
                        try {
                            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                                $parsed = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop
                                if ($parsed.error -and $parsed.error.message) { $errMsg = $parsed.error.message }
                            }
                        } catch {}
                        [PSCustomObject]@{
                            Job      = $spec.Job
                            Success  = $false
                            Response = $null
                            Error    = $errMsg
                        }
                    }
                }

                foreach ($r in @($graphResults)) {
                    $currentRole++
                    $rd = $r.Job.RoleData
                    if ($r.Success) {
                        Write-Verbose "$($rd.Type) role deactivated via parallel REST: $($rd.DisplayName)"
                        $successCount++
                        $null = $successfullyDeactivated.Add($rd)
                    } else {
                        $deactivationErrors += "$($rd.DisplayName): $($r.Error)"
                        Write-Warning "Parallel deactivation failed for $($rd.DisplayName): $($r.Error)"
                    }
                }
            }
            else {
                # Fallback: Microsoft Graph $batch API via Invoke-MgGraphRequest
                $batchBodies = [System.Collections.ArrayList]::new()
                $batchMeta   = [System.Collections.ArrayList]::new()

                foreach ($job in $graphJobs) {
                    $batchId = "$($batchBodies.Count + 1)"
                    $rd      = $job.RoleData

                    if ($rd.Type -eq 'Entra') {
                        $reqBody = @{
                            action                    = 'selfDeactivate'
                            principalId               = $currentUserIdGraph
                            roleDefinitionId          = $rd.RoleDefinitionId
                            directoryScopeId          = if ($rd.DirectoryScopeId) { $rd.DirectoryScopeId } else { '/' }
                            roleAssignmentScheduleId  = $job.ScheduleId
                            justification             = 'Deactivated via PowerShell'
                        }
                        $reqUrl = '/roleManagement/directory/roleAssignmentScheduleRequests'
                    } else {
                        $accessId = & $getGroupAccessId $rd
                        $reqBody = @{
                            action               = 'selfDeactivate'
                            principalId          = $currentUserIdGraph
                            groupId              = $rd.GroupId
                            accessId             = $accessId
                            assignmentScheduleId = $job.ScheduleId
                            justification        = 'Deactivated via PowerShell'
                        }
                        $reqUrl = '/identityGovernance/privilegedAccess/group/assignmentScheduleRequests'
                    }

                    $null = $batchBodies.Add(@{
                        id      = $batchId
                        method  = 'POST'
                        url     = $reqUrl
                        headers = @{ 'Content-Type' = 'application/json' }
                        body    = $reqBody
                    })
                    $null = $batchMeta.Add([PSCustomObject]@{ BatchId = $batchId; Job = $job })
                }

                $BATCH_SIZE = 20
                for ($bi = 0; $bi -lt $batchBodies.Count; $bi += $BATCH_SIZE) {
                    $chunkEnd  = [Math]::Min($bi + $BATCH_SIZE - 1, $batchBodies.Count - 1)
                    $chunk     = @($batchBodies[$bi..$chunkEnd])
                    $batchBody = @{ requests = $chunk }
                    try {
                        Write-Verbose "Submitting Graph deactivation batch ($($chunk.Count) request(s))"
                        $batchResp = Invoke-MgGraphRequest -Method POST -Uri '$batch' -Body $batchBody -ErrorAction Stop
                        foreach ($resp in @($batchResp.responses)) {
                            $meta = $batchMeta | Where-Object { $_.BatchId -eq $resp.id } | Select-Object -First 1
                            if (-not $meta) { continue }
                            $currentRole++
                            if ($resp.status -ge 200 -and $resp.status -lt 300) {
                                Write-Verbose "$($meta.Job.RoleData.Type) role deactivated via batch"
                                $successCount++
                                $null = $successfullyDeactivated.Add($meta.Job.RoleData)
                            } else {
                                $errMsg = if ($resp.body -and $resp.body.error -and $resp.body.error.message) { $resp.body.error.message } else { "HTTP $($resp.status)" }
                                $deactivationErrors += "$($meta.Job.RoleData.DisplayName): $errMsg"
                                Write-Warning "Batch deactivation failed for $($meta.Job.RoleData.DisplayName): $errMsg"
                            }
                        }
                    }
                    catch {
                        Write-Warning "Graph batch deactivation failed, falling back to sequential: $($_.Exception.Message)"
                        foreach ($reqItem in $chunk) {
                            $meta = $batchMeta | Where-Object { $_.BatchId -eq $reqItem.id } | Select-Object -First 1
                            if (-not $meta) { continue }
                            $currentRole++
                            $rd = $meta.Job.RoleData
                            try {
                                if ($rd.Type -eq 'Entra') {
                                    $rb = @{
                                        principalId              = $currentUserIdGraph
                                        action                   = 'selfDeactivate'
                                        justification            = 'Deactivated via PowerShell'
                                        roleDefinitionId         = $rd.RoleDefinitionId
                                        directoryScopeId         = if ($rd.DirectoryScopeId) { $rd.DirectoryScopeId } else { '/' }
                                        roleAssignmentScheduleId = $meta.Job.ScheduleId
                                    }
                                    New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $rb -ErrorAction Stop | Out-Null
                                } else {
                                    $accessId = & $getGroupAccessId $rd
                                    $rb = @{
                                        principalId          = $currentUserIdGraph
                                        groupId              = $rd.GroupId
                                        action               = 'selfDeactivate'
                                        justification        = 'Deactivated via PowerShell'
                                        accessId             = $accessId
                                        assignmentScheduleId = $meta.Job.ScheduleId
                                    }
                                    New-MgIdentityGovernancePrivilegedAccessGroupAssignmentScheduleRequest -BodyParameter $rb -ErrorAction Stop | Out-Null
                                }
                                Write-Verbose "$($rd.Type) role deactivated via sequential fallback"
                                $successCount++
                                $null = $successfullyDeactivated.Add($rd)
                            }
                            catch {
                                $errMsg = if ($_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
                                $deactivationErrors += "$($rd.DisplayName): $errMsg"
                            }
                        }
                    }
                }
            }
        }

        # ── Azure Resource roles: parallel ARM REST PUT ───────────────────────
        if ($azureJobs.Count -gt 0) {
            $operationSplash.UpdateStatus("Submitting $($azureJobs.Count) Azure Resource deactivation(s) in parallel...", 60)

            # Acquire ARM token once
            $armTokDeact = $null
            try {
                $azCtxDeact = Get-AzContext -ErrorAction SilentlyContinue
                if ($azCtxDeact) {
                    $tokenObj = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop
                    $armTokDeact = if ($tokenObj.Token -is [System.Security.SecureString]) {
                        [System.Net.NetworkCredential]::new('', $tokenObj.Token).Password
                    } else { $tokenObj.Token }
                }
            }
            catch { Write-Verbose "Could not acquire ARM token for parallel deactivation: $($_.Exception.Message)" }

            # SelfDeactivate always targets the current user, regardless of how the role
            # is held (direct vs. group-inherited). Azure role objects expose ObjectId
            # (which may be a group id), not PrincipalId, so we pass the caller id
            # explicitly via $using:.
            $currentUserIdDeact = $script:CurrentUser.Id

            $azureDeactResults = $azureJobs | ForEach-Object -Parallel {
                $job = $_
                $rd  = $job.RoleData
                $tok = $using:armTokDeact
                $callerId = $using:currentUserIdDeact
                try {
                    if (-not $tok) { throw "ARM access token unavailable; cannot submit SelfDeactivate request." }
                    if (-not $callerId) { throw "Current user id unavailable; cannot submit SelfDeactivate request." }
                    $requestName = [System.Guid]::NewGuid().ToString()
                    $roleDefId   = if ($rd.RoleDefinitionId.StartsWith('/')) {
                        $rd.RoleDefinitionId
                    } else {
                        "$($rd.FullScope)/providers/Microsoft.Authorization/roleDefinitions/$($rd.RoleDefinitionId)"
                    }
                    $bodyObj = @{
                        properties = @{
                            roleDefinitionId = $roleDefId
                            principalId      = $callerId
                            requestType      = 'SelfDeactivate'
                            justification    = 'Deactivated via PowerShell'
                        }
                    }
                    $hdrs    = @{ 'Authorization' = "Bearer $tok"; 'Content-Type' = 'application/json' }
                    $uri     = "https://management.azure.com$($rd.FullScope)/providers/Microsoft.Authorization/roleAssignmentScheduleRequests/$($requestName)?api-version=2020-10-01"
                    $bodyJson = $bodyObj | ConvertTo-Json -Depth 8 -Compress
                    $resp    = Invoke-RestMethod -Uri $uri -Headers $hdrs -Method Put -Body $bodyJson -ErrorAction Stop
                    [PSCustomObject]@{ Success = $true; Job = $job; Response = $resp }
                }
                catch {
                    [PSCustomObject]@{ Success = $false; Job = $job; ErrorMessage = $_.Exception.Message }
                }
            } -ThrottleLimit 5

            # Best-effort scrub of plaintext bearer token from memory.
            $armTokDeact = $null

            foreach ($azResult in @($azureDeactResults)) {
                $currentRole++
                $rd = $azResult.Job.RoleData
                if ($azResult.Success) {
                    Write-Verbose "Azure Resource role deactivated in parallel: $($rd.DisplayName)"
                    $successCount++
                    $null = $successfullyDeactivated.Add($rd)
                } else {
                    $deactivationErrors += "$($rd.DisplayName): $($azResult.ErrorMessage)"
                    Write-Warning "Azure Resource parallel deactivation failed for $($rd.DisplayName): $($azResult.ErrorMessage)"
                }
            }
        }


        $operationSplash.UpdateStatus("Completing deactivation process...", 95)
        
        # Close splash before showing results dialog
        if ($operationSplash -and -not $operationSplash.IsDisposed) {
            $operationSplash.Close()
            $operationSplash = $null
        }
        
        # Display results - always show a message to the user
        $errorCount = @($deactivationErrors).Count
        Write-Verbose "Deactivation complete. Success: $successCount, Errors: $errorCount"
        
        if ($errorCount -gt 0) {
            $message = "Successfully deactivated: $successCount of $totalRoles role(s)`n`nErrors ($errorCount):`n`n$($deactivationErrors -join "`n`n")"
            Show-TopMostMessageBox -Message $message -Title "Deactivation Results" -Icon Warning
        }
        elseif ($successCount -gt 0) {
            Show-TopMostMessageBox -Message "Successfully deactivated all $successCount role(s)!" -Title "Success" -Icon Information
        }
        
        # Clear role cache to ensure fresh data is fetched after deactivation
        if ($successCount -gt 0) {
            Write-Verbose "Clearing role cache to force fresh data retrieval after deactivation"
            $script:CachedEligibleRoles = $null
            $script:CachedActiveRoles = $null
            $script:LastRoleFetchTime = $null

            # Register short-lived suppression entries so just-deactivated roles do not
            # re-appear in the active list during Graph/ARM propagation lag (typically
            # several seconds to a minute). Entries auto-expire so a future refresh
            # will show roles that truly remained active (e.g., failed deactivation
            # discovered out-of-band).
            try {
                if (-not (Get-Variable -Name 'RecentlyDeactivated' -Scope Script -ErrorAction SilentlyContinue) -or -not $script:RecentlyDeactivated) {
                    $script:RecentlyDeactivated = @{}
                }
                $suppressUntil = (Get-Date).AddMinutes(5)
                foreach ($rd in $successfullyDeactivated) {
                    $key = $null
                    try {
                        switch ($rd.Type) {
                            'AzureResource' {
                                $rdef = $rd.RoleDefinitionId
                                if ($rdef -match '/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})') { $rdef = $matches[1] }
                                $fs = if ($rd.PSObject.Properties['FullScope']) { $rd.FullScope } elseif ($rd.PSObject.Properties['Scope']) { $rd.Scope } else { $null }
                                if ($fs -and $rdef) { $key = "Azure|$fs|$rdef" }
                            }
                            'Entra' {
                                $rdef = $rd.RoleDefinitionId
                                $ds   = if ($rd.PSObject.Properties['DirectoryScopeId'] -and $rd.DirectoryScopeId) { $rd.DirectoryScopeId } else { '/' }
                                if ($rdef) { $key = "Entra|$rdef|$ds" }
                            }
                            'Group' {
                                $gid = if ($rd.PSObject.Properties['GroupId']) { $rd.GroupId } else { $null }
                                $aid = & $getGroupAccessId $rd
                                if ($gid -and $aid) { $key = "Group|$gid|$aid" }
                            }
                        }
                    } catch {}
                    if ($key) {
                        $script:RecentlyDeactivated[$key] = $suppressUntil
                        Write-Verbose "Registered active-list suppression for $key (until $suppressUntil)"
                        # Drop any stale activation injection record for this role.
                        if ((Get-Variable -Name 'RecentlyActivated' -Scope Script -ErrorAction SilentlyContinue) -and $script:RecentlyActivated -and $script:RecentlyActivated.ContainsKey($key)) {
                            $null = $script:RecentlyActivated.Remove($key)
                        }
                    }
                }
            } catch { Write-Verbose "Failed to register deactivation suppression entries: $($_.Exception.Message)" }

            # Mark affected Azure subscriptions as dirty for delta refresh and clear any override expirations
            try {
                foreach ($item in $CheckedItems) {
                    $roleData = $item.Tag
                    if ($roleData -and $roleData.Type -eq 'AzureResource') {
                        if (-not (Get-Variable -Name 'DirtyAzureSubscriptions' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyAzureSubscriptions = @() }
                        if ($roleData.PSObject.Properties['SubscriptionId'] -and $roleData.SubscriptionId) {
                            $script:DirtyAzureSubscriptions += $roleData.SubscriptionId
                            $script:DirtyAzureSubscriptions = @($script:DirtyAzureSubscriptions | Select-Object -Unique)
                            Write-Verbose "Marked subscription $($roleData.SubscriptionId) as dirty after deactivation"
                        }
                        # If management group scope, mark MG dirty for delta refresh
                        if ($roleData.PSObject.Properties['FullScope'] -and $roleData.FullScope -match "^/providers/Microsoft\.Management/managementGroups/([^/]+)$") {
                            if (-not (Get-Variable -Name 'DirtyManagementGroups' -Scope Script -ErrorAction SilentlyContinue)) { $script:DirtyManagementGroups = @() }
                            $mgName = $matches[1]
                            $script:DirtyManagementGroups += $mgName
                            $script:DirtyManagementGroups = @($script:DirtyManagementGroups | Select-Object -Unique)
                            Write-Verbose "Marked management group ${mgName} as dirty after deactivation"
                        }

                        # Remove any Azure active override expiration for this role/scope
                        if (Get-Variable -Name 'AzureActiveOverrides' -Scope Script -ErrorAction SilentlyContinue) {
                            $roleDefKey = $roleData.RoleDefinitionId
                            if ($roleDefKey -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $roleDefKey = $matches[1] }
                            $overrideKey = "$($roleData.FullScope)|$($roleDefKey)"
                            if ($script:AzureActiveOverrides.ContainsKey($overrideKey)) {
                                $null = $script:AzureActiveOverrides.Remove($overrideKey)
                                Write-Verbose "Cleared Azure active override for $overrideKey after deactivation"
                            }
                            # Also remove from AzureRolesCache if present
                            if (Get-Variable -Name 'AzureRolesCache' -Scope Script -ErrorAction SilentlyContinue) {
                                $script:AzureRolesCache = @($script:AzureRolesCache | Where-Object {
                                        if ($_.PSObject.Properties['RoleDefinitionId'] -and $_.PSObject.Properties['FullScope']) {
                                            $rd = $_.RoleDefinitionId; if ($rd -match "/providers/Microsoft\.Authorization/roleDefinitions/([a-fA-F0-9\-]{36})") { $rd = $matches[1] }
                                            -not ($rd -eq $roleDefKey -and $_.FullScope -eq $roleData.FullScope)
                                        }
                                        else { $true }
                                    })
                                Write-Verbose "Pruned deactivated Azure role from AzureRolesCache for key $overrideKey"
                            }
                        }
                    }
                }
            }
            catch { Write-Verbose "Post-deactivation delta marking failed: $($_.Exception.Message)" }
        }
        
        try {
            # Per refresh semantics: only refresh ACTIVE roles after deactivation
            Update-PIMRolesList -Form $Form -RefreshActive
        }
        catch {
            Write-Warning "Failed to refresh role lists: $_"
        }
        
    }
    finally {
        # Ensure splash is closed
        if ($operationSplash -and -not $operationSplash.IsDisposed) {
            $operationSplash.Close()
        }
    }
    
    Write-Verbose "Deactivation process completed - Success: $successCount, Errors: $errorCount"
}