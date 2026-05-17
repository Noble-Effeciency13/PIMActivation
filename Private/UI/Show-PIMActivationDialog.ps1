function Show-PIMActivationDialog {
    <#
    .SYNOPSIS
        Displays a dialog for PIM role activation requirements.
    
    .DESCRIPTION
        Shows a Windows Forms dialog to collect scheduling, justification, ticket information,
        and optional Azure reduced scope required for PIM role activation. Returns user input or
        cancellation status.
    
    .PARAMETER RequiresJustification
        Specifies that justification text is mandatory for activation.
    
    .PARAMETER RequiresTicket
        Specifies that a ticket number is mandatory for activation.
    
    .PARAMETER OptionalJustification
        Displays justification field as optional with recommended usage note.

    .PARAMETER ShowAzureReducedScope
        Displays an optional Azure reduced-scope picker for Azure Resource role activations.

    .PARAMETER ProfileRoleItems
        Selected role items used for profile save/update actions and activation schedule validation.

    .PARAMETER ProfileDefaultDuration
        Requested duration used for profile save/update actions and activation schedule validation.
    
    .EXAMPLE
        Show-PIMActivationDialog -RequiresJustification
        Shows dialog with required justification field.
    
    .EXAMPLE
        Show-PIMActivationDialog -RequiresTicket -OptionalJustification
        Shows dialog with required ticket field and optional justification.
    
    .OUTPUTS
        PSCustomObject
        Returns object with Justification, TicketNumber, AzureReducedScope, ScheduleForLater,
        ScheduledStartTime, ScheduledStartTimeUtc, and Cancelled properties.
    
    .NOTES
        Requires System.Windows.Forms assembly for GUI display.
    #>
    [CmdletBinding()]
    param(
        [switch]$RequiresJustification,
        [switch]$RequiresTicket,
        [switch]$OptionalJustification,
        [switch]$ShowAzureReducedScope,
        [object[]]$AzureRoleItems = @(),
        [object[]]$ProfileRoleItems = @(),
        [hashtable]$ProfileDefaultDuration,
        [string]$ActivationProfileName = '',
        [switch]$AllowSaveAsProfile,
        [switch]$AllowProfileManagement
    )
    
    # Initialize result object
    $result = [PSCustomObject]@{
        Justification = ""
        TicketNumber  = ""
        TicketSystem  = "ServiceNow"
        AzureReducedScope = ""
        ScheduleForLater = $false
        ScheduledStartTime = $null
        ScheduledStartTimeUtc = $null
        ProfileSaved  = $false
        ProfileDeleted = $false
        Cancelled     = $true
    }
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text            = "Role Activation Requirements"
        Size            = [System.Drawing.Size]::new(500, 350)
        StartPosition   = 'CenterScreen'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        BackColor       = [System.Drawing.Color]::White
        TopMost         = $true
        ShowInTaskbar   = $true
    }
    
    $y = 10
    $justificationControl = $null
    $txtTicket = $null
    $cmbTicketSystem = $null
    $chkUseReducedScope = $null
    $reducedScopePanel = $null
    $lblReducedOriginalScopeValue = $null
    $lblReducedSelectedScopeValue = $null
    $lblReducedScopeStatus = $null
    $btnReducedScopeReset = $null
    $chkScheduleActivation = $null
    $dtpScheduleDate = $null
    $dtpScheduleTime = $null
    $lblScheduleStatus = $null
    $reducedScopeState = [ordered]@{
        OriginalScope        = ''
        OriginalDisplayName  = ''
        CurrentParentScope   = ''
        SelectedScope        = ''
        SelectedDisplayName  = ''
        SuppressChildChange  = $false
    }

    $saveActivationProfile = {
        param(
            [string]$ProfileName,
            [string]$SuccessMessage
        )

        if ([string]::IsNullOrWhiteSpace($ProfileName)) { return $null }
        if (-not $ProfileRoleItems -or $ProfileRoleItems.Count -eq 0) {
            Show-TopMostMessageBox -Message 'No roles are available to save in this activation profile.' -Title 'Activation Profile' -Icon Warning
            return $null
        }

        $duration = if ($ProfileDefaultDuration) { $ProfileDefaultDuration } else { @{ Hours = 8; Minutes = 0; TotalMinutes = 480 } }
        $savedProfile = Save-PIMActivationProfile -ProfileName $ProfileName -SelectedRoles @($ProfileRoleItems) -DefaultDuration $duration
        $result.ProfileSaved = $true
        Show-TopMostMessageBox -Message ($SuccessMessage -f $savedProfile.Name, $savedProfile.Roles.Count) -Title 'Activation Profile' -Icon Information
        return $savedProfile
    }

    if (-not [string]::IsNullOrWhiteSpace($ActivationProfileName)) {
        $form.Text = "Activate Profile - $ActivationProfileName"

        $lblProfile = New-Object System.Windows.Forms.Label -Property @{
            Text      = "Activation Profile: $ActivationProfileName"
            Location  = [System.Drawing.Point]::new(10, $y)
            Size      = [System.Drawing.Size]::new(460, 22)
            Font      = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
        }
        $form.Controls.Add($lblProfile)
        $y += 28
    }

    $scheduleWindow = Resolve-PIMActivationSchedule -RoleItems $ProfileRoleItems -RequestedDuration $ProfileDefaultDuration
    $defaultScheduleStart = (Get-Date).AddMinutes(30)
    $defaultScheduleStart = $defaultScheduleStart.AddSeconds(-1 * $defaultScheduleStart.Second).AddMilliseconds(-1 * $defaultScheduleStart.Millisecond)
    if ($scheduleWindow.MaxStartLocal -and $defaultScheduleStart -gt $scheduleWindow.MaxStartLocal) {
        $defaultScheduleStart = $scheduleWindow.MaxStartLocal
    }
    if ($defaultScheduleStart -lt (Get-Date)) {
        $defaultScheduleStart = Get-Date
    }

    $scheduleGroup = New-Object System.Windows.Forms.GroupBox -Property @{
        Text      = 'Activation Schedule'
        Location  = [System.Drawing.Point]::new(10, $y)
        Size      = [System.Drawing.Size]::new(460, 108)
        Font      = [System.Drawing.Font]::new('Segoe UI', 9)
        ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
    }

    $chkScheduleActivation = New-Object System.Windows.Forms.CheckBox -Property @{
        Name      = 'chkScheduleActivation'
        Text      = 'Schedule for later'
        Location  = [System.Drawing.Point]::new(12, 22)
        Size      = [System.Drawing.Size]::new(150, 22)
        Font      = [System.Drawing.Font]::new('Segoe UI', 9)
        Cursor    = [System.Windows.Forms.Cursors]::Hand
    }
    $scheduleGroup.Controls.Add($chkScheduleActivation)

    $lblScheduleDate = New-Object System.Windows.Forms.Label -Property @{
        Text     = 'Date'
        Location = [System.Drawing.Point]::new(20, 52)
        Size     = [System.Drawing.Size]::new(38, 22)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    $scheduleGroup.Controls.Add($lblScheduleDate)

    $dtpScheduleDate = New-Object System.Windows.Forms.DateTimePicker -Property @{
        Name     = 'dtpScheduleDate'
        Location = [System.Drawing.Point]::new(60, 49)
        Size     = [System.Drawing.Size]::new(150, 23)
        Format   = [System.Windows.Forms.DateTimePickerFormat]::Short
        Enabled  = $false
        Font     = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    try { $dtpScheduleDate.MinDate = (Get-Date).Date } catch { }
    if ($scheduleWindow.MaxStartLocal -and $scheduleWindow.MaxStartLocal.Date -ge $dtpScheduleDate.MinDate) {
        try { $dtpScheduleDate.MaxDate = $scheduleWindow.MaxStartLocal.Date } catch { }
    }
    try { $dtpScheduleDate.Value = $defaultScheduleStart.Date } catch { }
    $scheduleGroup.Controls.Add($dtpScheduleDate)

    $lblScheduleTime = New-Object System.Windows.Forms.Label -Property @{
        Text     = 'Time'
        Location = [System.Drawing.Point]::new(225, 52)
        Size     = [System.Drawing.Size]::new(40, 22)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    $scheduleGroup.Controls.Add($lblScheduleTime)

    $dtpScheduleTime = New-Object System.Windows.Forms.DateTimePicker -Property @{
        Name         = 'dtpScheduleTime'
        Location     = [System.Drawing.Point]::new(265, 49)
        Size         = [System.Drawing.Size]::new(80, 23)
        Format       = [System.Windows.Forms.DateTimePickerFormat]::Custom
        CustomFormat = 'HH:mm'
        ShowUpDown   = $true
        Enabled      = $false
        Font         = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    try { $dtpScheduleTime.Value = [datetime]::Today.AddHours($defaultScheduleStart.Hour).AddMinutes($defaultScheduleStart.Minute) } catch { }
    $scheduleGroup.Controls.Add($dtpScheduleTime)

    $scheduleStatusText = if ($scheduleWindow.MaxStartLocal) {
        "Latest start: $($scheduleWindow.MaxStartLocal.ToString('yyyy-MM-dd HH:mm'))"
    }
    else {
        'Future start uses your local time.'
    }
    $lblScheduleStatus = New-Object System.Windows.Forms.Label -Property @{
        Text      = $scheduleStatusText
        Location  = [System.Drawing.Point]::new(20, 76)
        Size      = [System.Drawing.Size]::new(420, 24)
        ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
        Font      = [System.Drawing.Font]::new('Segoe UI', 8)
    }
    $scheduleGroup.Controls.Add($lblScheduleStatus)

    $chkScheduleActivation.Add_CheckedChanged({
        $enabled = $chkScheduleActivation.Checked
        $dtpScheduleDate.Enabled = $enabled
        $dtpScheduleTime.Enabled = $enabled
    })

    $form.Controls.Add($scheduleGroup)
    $y += 118

    $getSelectedScheduleStart = {
        $dateValue = $dtpScheduleDate.Value
        $timeValue = $dtpScheduleTime.Value
        $combined = [datetime]::new(
            $dateValue.Year,
            $dateValue.Month,
            $dateValue.Day,
            $timeValue.Hour,
            $timeValue.Minute,
            0
        )
        return [datetime]::SpecifyKind($combined, [System.DateTimeKind]::Local)
    }

    $getSelectedReducedScope = {
        if ($reducedScopeState['SelectedScope']) { return [string]$reducedScopeState['SelectedScope'] }
        return ''
    }

    $getRoleDataFromItem = {
        param([object]$RoleItem)

        if ($RoleItem -is [System.Windows.Forms.ListViewItem]) { return $RoleItem.Tag }
        return $RoleItem
    }

    $getObjectPropertyValue = {
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

    $getRoleScope = {
        param([object]$RoleData)

        return & $getObjectPropertyValue $RoleData @('FullScope', 'Scope', 'DirectoryScopeId')
    }

    $getOriginalScopeDisplayName = {
        param(
            [object]$RoleData,
            [string]$Scope
        )

        $displayName = & $getObjectPropertyValue $RoleData @('ScopeDisplayName', 'ResourceDisplayName', 'SubscriptionName')
        if ([string]::IsNullOrWhiteSpace($displayName) -or $displayName -eq 'Subscription') {
            if ($Scope -match '^/subscriptions/[^/]+/resourceGroups/([^/]+)') {
                $displayName = [System.Uri]::UnescapeDataString($matches[1])
            }
            elseif ($Scope -match '^/subscriptions/([^/]+)$') {
                $displayName = & $getObjectPropertyValue $RoleData @('SubscriptionName')
                if ([string]::IsNullOrWhiteSpace($displayName) -or $displayName -eq 'Subscription') { $displayName = $matches[1] }
            }
            elseif ($Scope -match '^/providers/Microsoft\.Management/managementGroups/([^/]+)') {
                $displayName = $matches[1]
            }
        }

        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $Scope }

        $scopeType = & $getObjectPropertyValue $RoleData @('ScopeType')
        if (-not [string]::IsNullOrWhiteSpace($scopeType) -and $displayName -notlike "*($scopeType)") {
            return "$displayName ($scopeType)"
        }

        return $displayName
    }

    $getOriginalScopeOptionsFromSelectedRoles = {
        $scopeOptionsByScope = [ordered]@{}

        foreach ($roleItem in @($AzureRoleItems)) {
            $roleData = & $getRoleDataFromItem $roleItem
            if (-not $roleData) { continue }

            $scope = & $getRoleScope $roleData
            if ([string]::IsNullOrWhiteSpace($scope)) { continue }
            if (-not $scope.StartsWith('/')) { $scope = "/$scope" }
            while ($scope.Length -gt 1 -and $scope.EndsWith('/')) {
                $scope = $scope.Substring(0, $scope.Length - 1)
            }

            if ($scopeOptionsByScope.Contains($scope)) { continue }
            $displayName = & $getOriginalScopeDisplayName $roleData $scope

            $scopeOptionsByScope[$scope] = [PSCustomObject]@{
                DisplayName    = $displayName
                Name           = $displayName
                SubscriptionId = if ($scope -match '^/subscriptions/([^/]+)') { $matches[1] } else { $null }
                ResourceGroup  = $null
                ResourceId     = $null
                Scope          = $scope
                Type           = 'OriginalScope'
            }
        }

        return @($scopeOptionsByScope.Values | Sort-Object DisplayName)
    }

    $updateReducedScopeLabels = {
        if ($lblReducedOriginalScopeValue) {
            $originalText = if ($reducedScopeState['OriginalDisplayName']) { $reducedScopeState['OriginalDisplayName'] } else { 'Not loaded' }
            $lblReducedOriginalScopeValue.Text = $originalText
        }

        if ($lblReducedSelectedScopeValue) {
            $selectedText = if ($reducedScopeState['SelectedDisplayName']) { $reducedScopeState['SelectedDisplayName'] } else { 'No reduced scope selected' }
            $lblReducedSelectedScopeValue.Text = $selectedText
        }

        if ($btnReducedScopeReset) {
            $btnReducedScopeReset.Enabled = -not [string]::IsNullOrWhiteSpace([string]$reducedScopeState['SelectedScope'])
        }
    }

    $loadEligibleChildScopes = {
        param([string]$ParentScope)

        try {
            $cmbReducedResourceGroup.Items.Clear()
            $cmbReducedResourceGroup.Enabled = $false
            $cmbReducedResource.Items.Clear()
            $cmbReducedResource.Enabled = $false

            if ([string]::IsNullOrWhiteSpace($ParentScope)) { return }

            $reducedScopeState['CurrentParentScope'] = $ParentScope
            $lblReducedScopeStatus.Text = 'Loading eligible child scopes...'
            [System.Windows.Forms.Application]::DoEvents()

            Write-Verbose "Azure reduced scope: loading eligible child resources for scope '$ParentScope'."
            $childScopes = @(Get-AzureReducedScopeOptions -OriginalScope $ParentScope)

            $reducedScopeState['SuppressChildChange'] = $true
            try {
                foreach ($childScope in $childScopes) {
                    [void]$cmbReducedResourceGroup.Items.Add($childScope)
                }
                $cmbReducedResourceGroup.SelectedIndex = -1
            }
            finally {
                $reducedScopeState['SuppressChildChange'] = $false
            }

            $cmbReducedResourceGroup.Enabled = $childScopes.Count -gt 0
            if ($childScopes.Count -gt 0) {
                $lblReducedScopeStatus.Text = if ($reducedScopeState['SelectedScope']) {
                    'Choose a deeper child scope, or click OK to use the selected scope.'
                }
                else {
                    'Choose an eligible child scope.'
                }
            }
            else {
                $lblReducedScopeStatus.Text = if ($reducedScopeState['SelectedScope']) {
                    'No deeper eligible child scopes were returned. Click OK to use the selected scope.'
                }
                else {
                    'No eligible child scopes were returned for the original scope.'
                }
            }
        }
        catch {
            $lblReducedScopeStatus.Text = 'Unable to load eligible child scopes.'
            Show-TopMostMessageBox -Message "Unable to load eligible child scopes: $($_.Exception.Message)" -Title 'Reduced Scope' -Icon Warning
        }
    }

    $loadReducedScopeSubscriptions = {
        try {
            $lblReducedScopeStatus.Text = 'Loading original scopes...'
            $cmbReducedResourceGroup.Enabled = $false
            $cmbReducedResource.Enabled = $false
            $cmbReducedResourceGroup.Items.Clear()
            $cmbReducedResource.Items.Clear()
            [System.Windows.Forms.Application]::DoEvents()

            Write-Verbose 'Azure reduced scope: loading original scope options from selected Azure role(s).'
            $originalScopes = @(& $getOriginalScopeOptionsFromSelectedRoles)
            Write-Verbose "Azure reduced scope: selected Azure role(s) provided $($originalScopes.Count) original scope option(s)."

            if ($originalScopes.Count -eq 0) {
                $reducedScopeState['OriginalScope'] = ''
                $reducedScopeState['OriginalDisplayName'] = ''
                $reducedScopeState['CurrentParentScope'] = ''
                $reducedScopeState['SelectedScope'] = ''
                $reducedScopeState['SelectedDisplayName'] = ''
                & $updateReducedScopeLabels
                $lblReducedScopeStatus.Text = 'No original Azure scopes were available for reduced scope.'
                return
            }

            if ($originalScopes.Count -gt 1) {
                $reducedScopeState['OriginalScope'] = ''
                $reducedScopeState['OriginalDisplayName'] = 'Multiple original scopes selected'
                $reducedScopeState['CurrentParentScope'] = ''
                $reducedScopeState['SelectedScope'] = ''
                $reducedScopeState['SelectedDisplayName'] = ''
                & $updateReducedScopeLabels
                $lblReducedScopeStatus.Text = 'Reduced scope requires selected Azure roles to share one original scope.'
                return
            }

            $originalScope = $originalScopes[0]
            $reducedScopeState['OriginalScope'] = [string]$originalScope.Scope
            $reducedScopeState['OriginalDisplayName'] = [string]$originalScope.DisplayName
            $reducedScopeState['CurrentParentScope'] = [string]$originalScope.Scope
            $reducedScopeState['SelectedScope'] = ''
            $reducedScopeState['SelectedDisplayName'] = ''
            & $updateReducedScopeLabels
            & $loadEligibleChildScopes ([string]$originalScope.Scope)
        }
        catch {
            $lblReducedScopeStatus.Text = 'Unable to load reduced-scope options.'
            $chkUseReducedScope.Checked = $false
            Show-TopMostMessageBox -Message "Unable to load Azure reduced-scope options: $($_.Exception.Message)" -Title 'Reduced Scope' -Icon Warning
        }
    }

    $loadReducedScopeResourceGroups = {
        if ($reducedScopeState['SuppressChildChange']) { return }
        if (-not $cmbReducedResourceGroup.SelectedItem) { return }

        $selectedChildScope = $cmbReducedResourceGroup.SelectedItem
        if (-not $selectedChildScope.PSObject.Properties['Scope'] -or [string]::IsNullOrWhiteSpace([string]$selectedChildScope.Scope)) { return }

        $reducedScopeState['SelectedScope'] = [string]$selectedChildScope.Scope
        $reducedScopeState['SelectedDisplayName'] = [string]$selectedChildScope.DisplayName
        & $updateReducedScopeLabels
        & $loadEligibleChildScopes ([string]$selectedChildScope.Scope)
    }

    $loadReducedScopeResources = {
        if ([string]::IsNullOrWhiteSpace([string]$reducedScopeState['OriginalScope'])) { return }

        $reducedScopeState['SelectedScope'] = ''
        $reducedScopeState['SelectedDisplayName'] = ''
        & $updateReducedScopeLabels
        & $loadEligibleChildScopes ([string]$reducedScopeState['OriginalScope'])
    }
    
    # Add justification field if required or optional
    if ($RequiresJustification -or $OptionalJustification) {
        $labelText = if ($RequiresJustification) { "Justification (required):" } else { "Justification (optional - recommended):" }
        
        $lblJust = New-Object System.Windows.Forms.Label -Property @{
            Text      = $labelText
            Location  = [System.Drawing.Point]::new(10, $y)
            Size      = [System.Drawing.Size]::new(460, 20)
            Font      = [System.Drawing.Font]::new("Segoe UI", 9)
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
        }
        $y += 25
        
        $txtJust = New-Object System.Windows.Forms.TextBox -Property @{
            Name          = "txtJustification"
            Location      = [System.Drawing.Point]::new(10, $y)
            Size          = [System.Drawing.Size]::new(460, 80)
            Multiline     = $true
            AcceptsReturn = $true
            ScrollBars    = 'Vertical'
            Text          = "PowerShell activation"
        }
        $y += 90
        
        $form.Controls.AddRange(@($lblJust, $txtJust))
        # Store the justification textbox in a variable for later use
        $justificationControl = $txtJust
    }
    
    # Add ticket field if required
    if ($RequiresTicket) {
        $lblTicket = New-Object System.Windows.Forms.Label -Property @{
            Text     = "Ticket Number *"
            Location = [System.Drawing.Point]::new(10, $y)
            Size     = [System.Drawing.Size]::new(120, 23)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($lblTicket)
        
        $txtTicket = New-Object System.Windows.Forms.TextBox -Property @{
            Name     = "txtTicket"
            Location = [System.Drawing.Point]::new(130, $y)
            Size     = [System.Drawing.Size]::new(280, 23)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($txtTicket)
        
        $y += 30
        
        # Add ticket system dropdown
        $lblTicketSystem = New-Object System.Windows.Forms.Label -Property @{
            Text     = "Ticket System"
            Location = [System.Drawing.Point]::new(10, $y)
            Size     = [System.Drawing.Size]::new(120, 23)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($lblTicketSystem)
        
        $cmbTicketSystem = New-Object System.Windows.Forms.ComboBox -Property @{
            Name          = "cmbTicketSystem"
            Location      = [System.Drawing.Point]::new(130, $y)
            Size          = [System.Drawing.Size]::new(280, 23)
            Font          = [System.Drawing.Font]::new("Segoe UI", 9)
            DropDownStyle = 'DropDownList'
        }
        
        # Add common ticket systems
        $ticketSystems = @('ServiceNow', 'Jira', 'Azure DevOps', 'ServiceDesk Plus', 'BMC Remedy', 'Cherwell', 'Other')
        $cmbTicketSystem.Items.AddRange($ticketSystems)
        
        # Try to use saved preference or default to ServiceNow
        $savedSystem = Get-SavedTicketSystem
        if ($savedSystem -and $ticketSystems -contains $savedSystem) {
            $cmbTicketSystem.SelectedItem = $savedSystem
        }
        else {
            $cmbTicketSystem.SelectedIndex = 0  # Default to ServiceNow
        }
        
        $form.Controls.Add($cmbTicketSystem)
        
        $y += 35
    }

    if ($ShowAzureReducedScope) {
        $chkUseReducedScope = New-Object System.Windows.Forms.CheckBox -Property @{
            Name      = 'chkUseAzureReducedScope'
            Text      = 'Use Reduced Scope'
            Location  = [System.Drawing.Point]::new(10, $y)
            Size      = [System.Drawing.Size]::new(180, 23)
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
        }
        $form.Controls.Add($chkUseReducedScope)
        $y += 28

        $reducedScopePanel = New-Object System.Windows.Forms.Panel -Property @{
            Name        = 'pnlAzureReducedScope'
            Location    = [System.Drawing.Point]::new(10, $y)
            Size        = [System.Drawing.Size]::new(460, 158)
            BackColor   = [System.Drawing.Color]::White
            BorderStyle = [System.Windows.Forms.BorderStyle]::None
            Visible     = $false
        }

        $lblSubscription = New-Object System.Windows.Forms.Label -Property @{
            Text     = 'Original Scope'
            Location = [System.Drawing.Point]::new(0, 5)
            Size     = [System.Drawing.Size]::new(115, 23)
            Font     = [System.Drawing.Font]::new('Segoe UI', 9)
        }
        $reducedScopePanel.Controls.Add($lblSubscription)

        $lblReducedOriginalScopeValue = New-Object System.Windows.Forms.Label -Property @{
            Name         = 'lblAzureReducedOriginalScope'
            Text         = 'Not loaded'
            Location     = [System.Drawing.Point]::new(120, 5)
            Size         = [System.Drawing.Size]::new(330, 23)
            Font         = [System.Drawing.Font]::new('Segoe UI', 9)
            AutoEllipsis = $true
        }
        $reducedScopePanel.Controls.Add($lblReducedOriginalScopeValue)

        $lblResourceGroup = New-Object System.Windows.Forms.Label -Property @{
            Text     = 'Selected Scope'
            Location = [System.Drawing.Point]::new(0, 35)
            Size     = [System.Drawing.Size]::new(115, 23)
            Font     = [System.Drawing.Font]::new('Segoe UI', 9)
        }
        $reducedScopePanel.Controls.Add($lblResourceGroup)

        $lblReducedSelectedScopeValue = New-Object System.Windows.Forms.Label -Property @{
            Name         = 'lblAzureReducedSelectedScope'
            Text         = 'No reduced scope selected'
            Location     = [System.Drawing.Point]::new(120, 35)
            Size         = [System.Drawing.Size]::new(250, 23)
            Font         = [System.Drawing.Font]::new('Segoe UI', 9)
            AutoEllipsis = $true
        }
        $reducedScopePanel.Controls.Add($lblReducedSelectedScopeValue)

        $btnReducedScopeReset = New-Object System.Windows.Forms.Button -Property @{
            Name      = 'btnAzureReducedScopeReset'
            Text      = 'Reset'
            Location  = [System.Drawing.Point]::new(380, 33)
            Size      = [System.Drawing.Size]::new(70, 25)
            Font      = [System.Drawing.Font]::new('Segoe UI', 8)
            Enabled   = $false
            Cursor    = [System.Windows.Forms.Cursors]::Hand
        }
        $reducedScopePanel.Controls.Add($btnReducedScopeReset)

        $lblEligibleChild = New-Object System.Windows.Forms.Label -Property @{
            Text     = 'Eligible Child'
            Location = [System.Drawing.Point]::new(0, 67)
            Size     = [System.Drawing.Size]::new(115, 23)
            Font     = [System.Drawing.Font]::new('Segoe UI', 9)
        }
        $reducedScopePanel.Controls.Add($lblEligibleChild)

        $cmbReducedResourceGroup = New-Object System.Windows.Forms.ComboBox -Property @{
            Name          = 'cmbAzureReducedChildScope'
            Location      = [System.Drawing.Point]::new(120, 65)
            Size          = [System.Drawing.Size]::new(330, 23)
            Font          = [System.Drawing.Font]::new('Segoe UI', 9)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            DisplayMember = 'DisplayName'
            Enabled       = $false
        }
        $reducedScopePanel.Controls.Add($cmbReducedResourceGroup)

        $lblResource = New-Object System.Windows.Forms.Label -Property @{
            Text     = 'Resource'
            Location = [System.Drawing.Point]::new(0, 95)
            Size     = [System.Drawing.Size]::new(115, 23)
            Font     = [System.Drawing.Font]::new('Segoe UI', 9)
            Visible  = $false
        }
        $reducedScopePanel.Controls.Add($lblResource)

        $cmbReducedResource = New-Object System.Windows.Forms.ComboBox -Property @{
            Name          = 'cmbAzureReducedResource'
            Location      = [System.Drawing.Point]::new(120, 93)
            Size          = [System.Drawing.Size]::new(330, 23)
            Font          = [System.Drawing.Font]::new('Segoe UI', 9)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            DisplayMember = 'DisplayName'
            Enabled       = $false
            Visible       = $false
        }
        $reducedScopePanel.Controls.Add($cmbReducedResource)

        $lblReducedScopeStatus = New-Object System.Windows.Forms.Label -Property @{
            Text      = 'Choose an eligible child scope.'
            Location  = [System.Drawing.Point]::new(120, 95)
            Size      = [System.Drawing.Size]::new(330, 20)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
            Font      = [System.Drawing.Font]::new('Segoe UI', 8)
        }
        $reducedScopePanel.Controls.Add($lblReducedScopeStatus)

        $lblReducedScopeNote = New-Object System.Windows.Forms.Label -Property @{
            Text      = "After choosing a child scope, the next eligible child level loads automatically. Click OK at the level you want to activate."
            Location  = [System.Drawing.Point]::new(120, 115)
            Size      = [System.Drawing.Size]::new(330, 34)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
            Font      = [System.Drawing.Font]::new('Segoe UI', 8)
        }
        $reducedScopePanel.Controls.Add($lblReducedScopeNote)
        $form.Controls.Add($reducedScopePanel)

        $chkUseReducedScope.Add_CheckedChanged({
            $reducedScopePanel.Visible = $chkUseReducedScope.Checked
            if ($chkUseReducedScope.Checked -and [string]::IsNullOrWhiteSpace([string]$reducedScopeState['OriginalScope'])) {
                & $loadReducedScopeSubscriptions
            }
        })

        $cmbReducedResourceGroup.Add_SelectedIndexChanged({ & $loadReducedScopeResourceGroups })
        $btnReducedScopeReset.Add_Click({ & $loadReducedScopeResources })

        $y += 163
    }
    
    # Add optional justification note
    if ($OptionalJustification -and -not $RequiresJustification) {
        $lblNote = New-Object System.Windows.Forms.Label -Property @{
            Text      = "Note: While justification is optional, providing a clear reason helps with audit trails."
            Location  = [System.Drawing.Point]::new(10, $y)
            Size      = [System.Drawing.Size]::new(460, 40)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
            Font      = [System.Drawing.Font]::new("Segoe UI", 8)
        }
        $y += 45
        $form.Controls.Add($lblNote)
    }
    
    $showProfileButtons = $AllowSaveAsProfile -or ($AllowProfileManagement -and -not [string]::IsNullOrWhiteSpace($ActivationProfileName))
    $clientWidth = if ($showProfileButtons) { 604 } else { 484 }
    $form.ClientSize = [System.Drawing.Size]::new($clientWidth, [Math]::Max(311, $y + 75))
    $buttonY = $form.ClientSize.Height - 45
    $cancelButtonX = $form.ClientSize.Width - 95
    $okButtonX = $cancelButtonX - 100

    # Create OK button with styling and validation
    $okButton = New-Object System.Windows.Forms.Button -Property @{
        Text         = "OK"
        DialogResult = [System.Windows.Forms.DialogResult]::None
        Location     = [System.Drawing.Point]::new($okButtonX, $buttonY)
        Size         = [System.Drawing.Size]::new(80, 30)
        BackColor    = [System.Drawing.Color]::FromArgb(0, 103, 184)
        ForeColor    = [System.Drawing.Color]::White
        FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
        Font         = [System.Drawing.Font]::new("Segoe UI", 9)
        Cursor       = [System.Windows.Forms.Cursors]::Hand
        Anchor       = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    $okButton.FlatAppearance.BorderSize = 0
    
    # Add button hover effects
    $okButton.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 158) })
    $okButton.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 184) })
    
    # Validate required fields on OK click
    $okButton.Add_Click({
            $isValid = $true
            $validationMessage = ""
        
            # Validate ticket if required
            if ($RequiresTicket) {
                $ticketText = $form.Controls["txtTicket"].Text.Trim()
                if ([string]::IsNullOrWhiteSpace($ticketText)) {
                    $isValid = $false
                    $validationMessage += "Ticket number is required.{0}" -f [Environment]::NewLine
                }
            }
        
            # Validate justification if required
            if ($RequiresJustification -and $justificationControl -and [string]::IsNullOrWhiteSpace($justificationControl.Text)) {
                Show-TopMostMessageBox -Message "Justification is required for these roles." -Title "Validation Error" -Icon Warning
                return
            }
        
            if ($RequiresTicket -and $txtTicket -and [string]::IsNullOrWhiteSpace($txtTicket.Text)) {
                Show-TopMostMessageBox -Message "Ticket number is required for these roles." -Title "Validation Error" -Icon Warning
                return
            }
        
            if ($isValid) {
                # Set result values
                $result.Justification = if ($justificationControl) { $justificationControl.Text.Trim() } else { "" }
            
                if ($RequiresTicket) {
                    $result.TicketNumber = $form.Controls["txtTicket"].Text.Trim()
                    $result.TicketSystem = $cmbTicketSystem.SelectedItem.ToString()
                
                    # Save ticket system preference
                    Save-TicketSystemPreference -System $result.TicketSystem
                }

                if ($ShowAzureReducedScope -and $chkUseReducedScope -and $chkUseReducedScope.Checked) {
                    $selectedScope = & $getSelectedReducedScope
                    if ([string]::IsNullOrWhiteSpace($selectedScope)) {
                        Show-TopMostMessageBox -Message 'Choose an eligible child scope for reduced scope.' -Title 'Validation Error' -Icon Warning
                        return
                    }
                    $result.AzureReducedScope = $selectedScope
                }
                elseif ($ShowAzureReducedScope) {
                    $result.AzureReducedScope = ''
                }

                if ($chkScheduleActivation -and $chkScheduleActivation.Checked) {
                    $selectedScheduleStart = & $getSelectedScheduleStart
                    $scheduleResolution = Resolve-PIMActivationSchedule -RoleItems $ProfileRoleItems -RequestedDuration $ProfileDefaultDuration -ScheduleStartTime $selectedScheduleStart -Scheduled
                    if (-not $scheduleResolution.IsValid) {
                        Show-TopMostMessageBox -Message $scheduleResolution.ErrorMessage -Title 'Schedule Validation' -Icon Warning
                        return
                    }

                    $result.ScheduleForLater = $true
                    $result.ScheduledStartTime = $scheduleResolution.StartLocal
                    $result.ScheduledStartTimeUtc = $scheduleResolution.StartUtcString
                }
                else {
                    $result.ScheduleForLater = $false
                    $result.ScheduledStartTime = $null
                    $result.ScheduledStartTimeUtc = $null
                }
            
                $result.Cancelled = $false
                $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $form.Close()
            }
            else {
                Show-TopMostMessageBox -Message $validationMessage.TrimEnd() -Title "Validation Error" -Icon Warning
            }
        })
    
    # Create Cancel button with styling
    $cancelButton = New-Object System.Windows.Forms.Button -Property @{
        Text         = "Cancel"
        DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        Location     = [System.Drawing.Point]::new($cancelButtonX, $buttonY)
        Size         = [System.Drawing.Size]::new(80, 30)
        BackColor    = [System.Drawing.Color]::White
        ForeColor    = [System.Drawing.Color]::FromArgb(32, 31, 30)
        FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
        Font         = [System.Drawing.Font]::new("Segoe UI", 9)
        Cursor       = [System.Windows.Forms.Cursors]::Hand
        Anchor       = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
    
    $cancelButton.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245) })
    $cancelButton.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::White })

    $profileButtons = @()
    $profileButtonX = 10

    if ($AllowSaveAsProfile -and [string]::IsNullOrWhiteSpace($ActivationProfileName)) {
        $saveProfileButton = New-Object System.Windows.Forms.Button -Property @{
            Text      = 'Save as Profile'
            Location  = [System.Drawing.Point]::new($profileButtonX, $buttonY)
            Size      = [System.Drawing.Size]::new(115, 30)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
        }
        $saveProfileButton.FlatAppearance.BorderSize = 1
        $saveProfileButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $saveProfileButton.Add_MouseEnter({
            $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $this.ForeColor = [System.Drawing.Color]::White
        })
        $saveProfileButton.Add_MouseLeave({
            $this.BackColor = [System.Drawing.Color]::White
            $this.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        })
        $saveProfileButton.Add_Click({
            try {
                $nameResult = Show-PIMProfileNameDialog -Title 'Save Activation Profile'
                if ($nameResult.Cancelled) { return }
                & $saveActivationProfile $nameResult.ProfileName "Saved activation profile '{0}' with {1} role(s)."
            }
            catch {
                Show-TopMostMessageBox -Message "Failed to save activation profile: $($_.Exception.Message)" -Title 'Activation Profile' -Icon Error
            }
        })
        $profileButtons += $saveProfileButton
        $profileButtonX += 125
    }

    if ($AllowProfileManagement -and -not [string]::IsNullOrWhiteSpace($ActivationProfileName)) {
        $updateProfileButton = New-Object System.Windows.Forms.Button -Property @{
            Text      = 'Update Profile'
            Location  = [System.Drawing.Point]::new($profileButtonX, $buttonY)
            Size      = [System.Drawing.Size]::new(110, 30)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
        }
        $updateProfileButton.FlatAppearance.BorderSize = 1
        $updateProfileButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $updateProfileButton.Add_MouseEnter({
            $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $this.ForeColor = [System.Drawing.Color]::White
        })
        $updateProfileButton.Add_MouseLeave({
            $this.BackColor = [System.Drawing.Color]::White
            $this.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        })
        $updateProfileButton.Add_Click({
            try {
                $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    "Update activation profile '$ActivationProfileName' with the current role set and duration?",
                    'Update Activation Profile',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                & $saveActivationProfile $ActivationProfileName "Updated activation profile '{0}' with {1} role(s)."
            }
            catch {
                Show-TopMostMessageBox -Message "Failed to update activation profile: $($_.Exception.Message)" -Title 'Activation Profile' -Icon Error
            }
        })
        $profileButtons += $updateProfileButton
        $profileButtonX += 120

        $deleteProfileButton = New-Object System.Windows.Forms.Button -Property @{
            Text      = 'Delete'
            Location  = [System.Drawing.Point]::new($profileButtonX, $buttonY)
            Size      = [System.Drawing.Size]::new(80, 30)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(164, 38, 44)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
        }
        $deleteProfileButton.FlatAppearance.BorderSize = 1
        $deleteProfileButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(164, 38, 44)
        $deleteProfileButton.Add_MouseEnter({
            $this.BackColor = [System.Drawing.Color]::FromArgb(164, 38, 44)
            $this.ForeColor = [System.Drawing.Color]::White
        })
        $deleteProfileButton.Add_MouseLeave({
            $this.BackColor = [System.Drawing.Color]::White
            $this.ForeColor = [System.Drawing.Color]::FromArgb(164, 38, 44)
        })
        $deleteProfileButton.Add_Click({
            try {
                $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    "Delete activation profile '$ActivationProfileName'?",
                    'Delete Activation Profile',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) { return }

                $deleted = Manage-PIMProfiles -Action Delete -ProfileName $ActivationProfileName
                if ($deleted) {
                    $result.ProfileDeleted = $true
                    $result.Cancelled = $true
                    Show-TopMostMessageBox -Message "Deleted activation profile '$ActivationProfileName'." -Title 'Activation Profile' -Icon Information
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $form.Close()
                }
                else {
                    Show-TopMostMessageBox -Message "Activation profile '$ActivationProfileName' was not found." -Title 'Activation Profile' -Icon Warning
                }
            }
            catch {
                Show-TopMostMessageBox -Message "Failed to delete activation profile: $($_.Exception.Message)" -Title 'Activation Profile' -Icon Error
            }
        })
        $profileButtons += $deleteProfileButton
    }
    
    # Add controls and set form properties
    $form.Controls.AddRange(@($okButton, $cancelButton) + $profileButtons)
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    
    # Ensure the form is brought to front and activated
    $form.Add_Shown({
            $this.Activate()
            $this.BringToFront()
            $this.TopMost = $true
            $this.Focus()
        })

    # Show dialog and process result
    $dialogResult = $form.ShowDialog()
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        if (($RequiresJustification -or $OptionalJustification) -and $justificationControl) {
            $result.Justification = if ([string]::IsNullOrWhiteSpace($justificationControl.Text)) { "PowerShell activation" } else { $justificationControl.Text }
        }
        
        if ($RequiresTicket -and $txtTicket -and -not [string]::IsNullOrWhiteSpace($txtTicket.Text)) {
            $result.TicketNumber = $txtTicket.Text
        }

        if ($ShowAzureReducedScope -and $chkUseReducedScope -and $chkUseReducedScope.Checked) {
            $result.AzureReducedScope = & $getSelectedReducedScope
        }
    }
    else {
        $result.Cancelled = $true
    }
    
    $form.Dispose()
    return $result
}