function Get-AzureMemberType {
    <#
    .SYNOPSIS
        Determines if an Azure role assignment is Direct or Inherited.
    
    .DESCRIPTION
        Analyzes the assignment scope relative to the current subscription and
        principal type to determine if the role assignment is direct or inherited.
        Matches Azure portal inheritance logic.
    
    .PARAMETER AssignmentScope
        The scope of the Azure role assignment.
    
    .PARAMETER CurrentSubscriptionId
        The ID of the current subscription being processed.
    
    .PARAMETER PrincipalType
        The type of principal (User, Group, ServicePrincipal).
    
    .PARAMETER IsEligible
        Whether this is an eligible assignment (affects inheritance logic).
    
    .PARAMETER ObjectId
        The object ID of the principal for additional validation.
    
    .OUTPUTS
        String indicating "Direct", "Inherited", or "Group"
    #>
    param(
        [Parameter(Mandatory)]
        [string]$AssignmentScope,
        
        [Parameter(Mandatory)]
        [string]$CurrentSubscriptionId,
        
        [Parameter()]
        [string]$PrincipalType = "User",
        
        [Parameter()]
        [bool]$IsEligible = $false,
        
        [Parameter()]
        [string]$ObjectId
    )
    
    Write-Verbose "Analyzing member type for scope: $AssignmentScope (PrincipalType: $PrincipalType, CurrentSub: $CurrentSubscriptionId)"
    
    # Group-based assignments are always "Group" type
    if ($PrincipalType -eq 'Group') {
        Write-Verbose "Assignment is Group-based -> Group"
        return "Group"
    }
    
    # Management group assignments are ALWAYS inherited
    if ($AssignmentScope -match "^/providers/Microsoft\.Management/managementGroups/") {
        Write-Verbose "Assignment at management group scope -> Inherited"
        return "Inherited"
    }
    
    # Tenant root assignments (scope = "/") are ALWAYS inherited  
    if ($AssignmentScope -eq "/") {
        Write-Verbose "Assignment at tenant root scope -> Inherited"
        return "Inherited"
    }
    
    # Cross-subscription assignments are inherited
    if ($AssignmentScope -match "^/subscriptions/([^/]+)") {
        $assignmentSubscriptionId = $matches[1]
        if ($assignmentSubscriptionId -ne $CurrentSubscriptionId) {
            Write-Verbose "Assignment from different subscription ($assignmentSubscriptionId) -> Inherited"
            return "Inherited"
        }
    }
    
    # For assignments within the current subscription
    if ($AssignmentScope -match "^/subscriptions/$CurrentSubscriptionId") {
        # Exact subscription level assignments
        if ($AssignmentScope -eq "/subscriptions/$CurrentSubscriptionId") {
            Write-Verbose "Assignment at current subscription level -> Direct"
            return "Direct"
        }
        
        # Resource group level assignments are direct
        if ($AssignmentScope -match "^/subscriptions/$CurrentSubscriptionId/resourceGroups/[^/]+$") {
            Write-Verbose "Assignment at resource group level -> Direct"
            return "Direct"
        }
        
        # Individual resource assignments are direct
        if ($AssignmentScope -match "^/subscriptions/$CurrentSubscriptionId/resourceGroups/.+") {
            Write-Verbose "Assignment at resource level -> Direct"
            return "Direct"
        }
    }
    
    # Special case: Check if the scope suggests inheritance from a higher level
    # This handles cases where the assignment might be coming from a scope above the current context
    
    # If the scope doesn't contain the current subscription, it's likely inherited
    if ($AssignmentScope -notmatch $CurrentSubscriptionId) {
        # Check if it's a well-known inherited scope pattern
        if ($AssignmentScope -match "^/$" -or 
            $AssignmentScope -match "^/providers/" -or
            $AssignmentScope -match "managementGroups") {
            Write-Verbose "Assignment from higher scope pattern -> Inherited"
            return "Inherited"
        }
    }
    
    # For any assignment that doesn't fit the above patterns but is clearly from a broader scope
    # Check scope hierarchy depth - if it's "shorter" than our subscription scope, it's likely inherited
    $assignmentParts = ($AssignmentScope -split '/').Where({ $_ -ne '' })
    $subscriptionParts = ("/subscriptions/$CurrentSubscriptionId" -split '/').Where({ $_ -ne '' })
    
    if ($assignmentParts.Count -lt $subscriptionParts.Count) {
        Write-Verbose "Assignment scope shorter than subscription scope -> Inherited"
        return "Inherited"
    }
    
    # Default case - if we can't determine inheritance, assume direct
    Write-Verbose "Could not determine inheritance pattern -> Direct (default)"
    return "Direct"
}