function Invoke-PIMActivationProfileSaveFromForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form
    )

    $eligibleMatches = $Form.Controls.Find('lstEligible', $true)
    $eligibleListView = if ($eligibleMatches -and $eligibleMatches.Count -gt 0) { $eligibleMatches[0] } else { $null }
    if (-not $eligibleListView -or $eligibleListView.CheckedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $Form,
            'Please select at least one eligible role before saving an activation profile.',
            'Activation Profile',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return $null
    }

    $nameResult = Show-PIMProfileNameDialog
    if ($nameResult.Cancelled) { return $null }

    $hoursMatches = $Form.Controls.Find('cmbHours', $true)
    $minutesMatches = $Form.Controls.Find('cmbMinutes', $true)
    $hoursControl = if ($hoursMatches -and $hoursMatches.Count -gt 0) { $hoursMatches[0] } else { $null }
    $minutesControl = if ($minutesMatches -and $minutesMatches.Count -gt 0) { $minutesMatches[0] } else { $null }
    $hours = if ($hoursControl -and $hoursControl.SelectedItem -ne $null) { [int]$hoursControl.SelectedItem } else { 8 }
    $minutes = if ($minutesControl -and $minutesControl.SelectedItem -ne $null) { [int]$minutesControl.SelectedItem } else { 0 }

    $duration = @{
        Hours        = $hours
        Minutes      = $minutes
        TotalMinutes = ($hours * 60) + $minutes
    }

    $savedProfile = Save-PIMActivationProfile -ProfileName $nameResult.ProfileName -SelectedRoles @($eligibleListView.CheckedItems) -DefaultDuration $duration
    [System.Windows.Forms.MessageBox]::Show(
        $Form,
        "Saved activation profile '$($savedProfile.Name)' with $($savedProfile.Roles.Count) role(s).",
        'Activation Profile',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    return $savedProfile
}

function Get-PIMActivationProfileMatchedItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form,

        [Parameter(Mandatory)]
        [object]$Profile,

        [switch]$ApplyCheckState
    )

    $eligibleMatches = $Form.Controls.Find('lstEligible', $true)
    $eligibleListView = if ($eligibleMatches -and $eligibleMatches.Count -gt 0) { $eligibleMatches[0] } else { $null }
    if (-not $eligibleListView) { return @() }

    $profileKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($Profile.PSObject.Properties['Roles'] -and $Profile.Roles) {
        foreach ($profileRole in @($Profile.Roles)) {
            if ($profileRole -and $profileRole.PSObject.Properties['Key'] -and $profileRole.Key) {
                [void]$profileKeys.Add([string]$profileRole.Key)
            }
        }
    }

    $matchedItems = [System.Collections.ArrayList]::new()
    foreach ($item in $eligibleListView.Items) {
        if ($item -isnot [System.Windows.Forms.ListViewItem]) { continue }

        if ($ApplyCheckState) { $item.Checked = $false }
        if (-not $item.Tag) { continue }

        $itemKey = Get-PIMActivationProfileRoleKey -RoleData $item.Tag
        if ($profileKeys.Contains($itemKey)) {
            if ($ApplyCheckState) { $item.Checked = $true }
            [void]$matchedItems.Add($item)
        }
    }

    return @($matchedItems)
}

function Set-PIMActivationProfileDurationOnForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form,

        [object]$Profile
    )

    $hoursMatches = $Form.Controls.Find('cmbHours', $true)
    $minutesMatches = $Form.Controls.Find('cmbMinutes', $true)
    $hoursControl = if ($hoursMatches -and $hoursMatches.Count -gt 0) { $hoursMatches[0] } else { $null }
    $minutesControl = if ($minutesMatches -and $minutesMatches.Count -gt 0) { $minutesMatches[0] } else { $null }

    if ($Profile.PSObject.Properties['DefaultDuration'] -and $Profile.DefaultDuration) {
        if ($hoursControl -and $Profile.DefaultDuration.PSObject.Properties['Hours']) {
            $hoursValue = [int]$Profile.DefaultDuration.Hours
            if ($hoursControl.Items.Contains($hoursValue)) { $hoursControl.SelectedItem = $hoursValue }
        }
        if ($minutesControl -and $Profile.DefaultDuration.PSObject.Properties['Minutes']) {
            $minutesValue = [int]$Profile.DefaultDuration.Minutes
            if ($minutesControl.Items.Contains($minutesValue)) { $minutesControl.SelectedItem = $minutesValue }
        }
    }

    $hours = if ($hoursControl -and $hoursControl.SelectedItem -ne $null) { [int]$hoursControl.SelectedItem } else { 8 }
    $minutes = if ($minutesControl -and $minutesControl.SelectedItem -ne $null) { [int]$minutesControl.SelectedItem } else { 0 }

    return @{
        Hours        = $hours
        Minutes      = $minutes
        TotalMinutes = ($hours * 60) + $minutes
    }
}

function Invoke-PIMActivationProfileActivationFromForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Form]$Form,

        [Parameter(Mandatory)]
        [object]$Profile
    )

    $matchedItems = @(Get-PIMActivationProfileMatchedItems -Form $Form -Profile $Profile -ApplyCheckState)
    $profileName = if ($Profile.PSObject.Properties['Name'] -and $Profile.Name) { $Profile.Name } else { 'Unnamed profile' }

    if ($matchedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $Form,
            "No currently eligible roles matched activation profile '$profileName'. Refresh roles or update the profile.",
            'Activation Profile',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $script:RequestedDuration = Set-PIMActivationProfileDurationOnForm -Form $Form -Profile $Profile

    try {
        Invoke-PIMRoleActivation -CheckedItems $matchedItems -Form $Form -ActivationProfile $Profile
    }
    finally {
        $eligibleMatches = $Form.Controls.Find('lstEligible', $true)
        $eligibleListView = if ($eligibleMatches -and $eligibleMatches.Count -gt 0) { $eligibleMatches[0] } else { $null }
        if ($eligibleListView) {
            foreach ($item in $eligibleListView.Items) {
                if ($item -is [System.Windows.Forms.ListViewItem]) { $item.Checked = $false }
            }
        }
    }
}

function Show-PIMActivationProfilesMenu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Control]$SourceControl
    )

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $menu.ShowImageMargin = $false
    $menu.MinimumSize = [System.Drawing.Size]::new(360, 0)

    $saveItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{
        Text = 'Save current selection...'
    }
    $saveItem.Add_Click({
        param($sender, $eventArgs)

        $sourceControl = $sender.Owner.SourceControl
        $form = if ($sourceControl) { $sourceControl.FindForm() } else { $null }
        if (-not $form) { return }

        try {
            Invoke-PIMActivationProfileSaveFromForm -Form $form | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                "Failed to save activation profile: $($_.Exception.Message)",
                'Activation Profile',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })
    [void]$menu.Items.Add($saveItem)
    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    try {
        $profiles = @(Get-PIMActivationProfiles)
    }
    catch {
        $profiles = @()
        $errorItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{
            Text    = 'Failed to load profiles'
            Enabled = $false
        }
        [void]$menu.Items.Add($errorItem)
    }

    if ($profiles.Count -eq 0) {
        $emptyItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{
            Text    = 'No saved profiles'
            Enabled = $false
        }
        [void]$menu.Items.Add($emptyItem)
    }
    else {
        foreach ($profile in $profiles) {
            $profileName = if ($profile.PSObject.Properties['Name'] -and $profile.Name) { $profile.Name } else { 'Unnamed profile' }
            $roleCount = if ($profile.PSObject.Properties['Roles'] -and $profile.Roles) { @($profile.Roles).Count } else { 0 }

            $rowPanel = New-Object System.Windows.Forms.Panel -Property @{
                Size      = [System.Drawing.Size]::new(350, 32)
                BackColor = [System.Drawing.Color]::White
            }

            $rowState = [PSCustomObject]@{
                Menu          = $menu
                SourceControl = $SourceControl
                Profile       = $profile
            }

            $profileLabel = New-Object System.Windows.Forms.Label -Property @{
                Text         = "$profileName ($roleCount role(s))"
                Location     = [System.Drawing.Point]::new(8, 7)
                Size         = [System.Drawing.Size]::new(252, 18)
                Font         = [System.Drawing.Font]::new('Segoe UI', 9)
                ForeColor    = [System.Drawing.Color]::FromArgb(32, 31, 30)
                BackColor    = [System.Drawing.Color]::White
                Cursor       = [System.Windows.Forms.Cursors]::Hand
                AutoEllipsis = $true
                Tag          = $rowState
            }
            $profileLabel.Add_Click({
                $state = $this.Tag
                $form = if ($state.SourceControl) { $state.SourceControl.FindForm() } else { $null }
                if (-not $form -or -not $state.Profile) { return }

                try {
                    $state.Menu.Close()
                    Invoke-PIMActivationProfileActivationFromForm -Form $form -Profile $state.Profile
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        $form,
                        "Failed to open activation profile: $($_.Exception.Message)",
                        'Activation Profile',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            })
            $rowPanel.Controls.Add($profileLabel)

            $editButton = New-Object System.Windows.Forms.Button -Property @{
                Text      = '✎'
                Location  = [System.Drawing.Point]::new(268, 4)
                Size      = [System.Drawing.Size]::new(32, 24)
                BackColor = [System.Drawing.Color]::White
                ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                Font      = [System.Drawing.Font]::new('Segoe UI', 9)
                Cursor    = [System.Windows.Forms.Cursors]::Hand
                Tag       = $rowState
            }
            $editButton.FlatAppearance.BorderSize = 1
            $editButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
            $editButton.Add_MouseEnter({
                $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                $this.ForeColor = [System.Drawing.Color]::White
            })
            $editButton.Add_MouseLeave({
                $this.BackColor = [System.Drawing.Color]::White
                $this.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            })
            $editButton.Add_Click({
                $state = $this.Tag
                $form = if ($state.SourceControl) { $state.SourceControl.FindForm() } else { $null }
                if (-not $form -or -not $state.Profile) { return }

                try {
                    $state.Menu.Close()
                    Invoke-PIMActivationProfileActivationFromForm -Form $form -Profile $state.Profile
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        $form,
                        "Failed to open activation profile: $($_.Exception.Message)",
                        'Activation Profile',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            })
            $rowPanel.Controls.Add($editButton)

            $deleteButton = New-Object System.Windows.Forms.Button -Property @{
                Text      = '✕'
                Location  = [System.Drawing.Point]::new(308, 4)
                Size      = [System.Drawing.Size]::new(32, 24)
                BackColor = [System.Drawing.Color]::White
                ForeColor = [System.Drawing.Color]::FromArgb(164, 38, 44)
                FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                Font      = [System.Drawing.Font]::new('Segoe UI', 9)
                Cursor    = [System.Windows.Forms.Cursors]::Hand
                Tag       = $rowState
            }
            $deleteButton.FlatAppearance.BorderSize = 1
            $deleteButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
            $deleteButton.Add_MouseEnter({
                $this.BackColor = [System.Drawing.Color]::FromArgb(164, 38, 44)
                $this.ForeColor = [System.Drawing.Color]::White
            })
            $deleteButton.Add_MouseLeave({
                $this.BackColor = [System.Drawing.Color]::White
                $this.ForeColor = [System.Drawing.Color]::FromArgb(164, 38, 44)
            })
            $deleteButton.Add_Click({
                $state = $this.Tag
                $form = if ($state.SourceControl) { $state.SourceControl.FindForm() } else { $null }
                if (-not $form -or -not $state.Profile) { return }

                $name = if ($state.Profile.PSObject.Properties['Name'] -and $state.Profile.Name) { [string]$state.Profile.Name } else { 'Unnamed profile' }
                $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    "Delete activation profile '$name'?",
                    'Delete Activation Profile',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) { return }

                try {
                    $deleted = Manage-PIMProfiles -Action Delete -ProfileName $name
                    if ($deleted) {
                        $state.Menu.Close()
                        [System.Windows.Forms.MessageBox]::Show(
                            $form,
                            "Deleted activation profile '$name'.",
                            'Activation Profile',
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Information
                        ) | Out-Null
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show(
                            $form,
                            "Activation profile '$name' was not found.",
                            'Activation Profile',
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        ) | Out-Null
                    }
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        $form,
                        "Failed to delete activation profile: $($_.Exception.Message)",
                        'Activation Profile',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            })
            $rowPanel.Controls.Add($deleteButton)

            $profileHost = New-Object System.Windows.Forms.ToolStripControlHost($rowPanel)
            $profileHost.AutoSize = $false
            $profileHost.Size = $rowPanel.Size
            $profileHost.Margin = [System.Windows.Forms.Padding]::new(0)
            $profileHost.Padding = [System.Windows.Forms.Padding]::new(0)
            [void]$menu.Items.Add($profileHost)
        }
    }

    $menu.Show($SourceControl, [System.Drawing.Point]::new(0, $SourceControl.Height))
}
