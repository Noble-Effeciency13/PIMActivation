function New-PIMActiveRolesPanel {
    <#
    .SYNOPSIS
        Creates a panel displaying currently active PIM roles with a modern UI design.
    
    .DESCRIPTION
        Creates a Windows Forms panel containing a ListView for displaying active PIM roles.
        The panel includes a header with title and role count, and a custom-styled ListView
        with columns for role details. Features owner-drawn headers and hover effects.
    
    .EXAMPLE
        $activeRolesPanel = New-PIMActiveRolesPanel
        $form.Controls.Add($activeRolesPanel)
        
        Creates and adds the active roles panel to a form.
    
    .OUTPUTS
        System.Windows.Forms.Panel
        Returns a panel containing the active roles ListView with header.
    
    .NOTES
        The ListView uses owner-drawn headers for custom styling and includes
        double buffering for smooth rendering performance.
    #>
    [CmdletBinding()]
    param()
    
    # Create main container panel
    $panel = New-Object System.Windows.Forms.Panel -Property @{
        Name        = 'pnlActive'
        BackColor   = [System.Drawing.Color]::White
        BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        Dock        = [System.Windows.Forms.DockStyle]::Fill
        Padding     = New-Object System.Windows.Forms.Padding(0)
    }
    
    # Create header panel with Microsoft blue background
    $headerPanel = New-Object System.Windows.Forms.Panel -Property @{
        Height    = 70
        BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        Dock      = [System.Windows.Forms.DockStyle]::Top
    }
    $panel.Controls.Add($headerPanel)
    
    # Create title label
    $lblTitle = New-Object System.Windows.Forms.Label -Property @{
        Text      = 'Active Roles'
        Location  = [System.Drawing.Point]::new(15, 12)
        Size      = [System.Drawing.Size]::new(200, 25)
        Font      = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        ForeColor = [System.Drawing.Color]::White
        BackColor = [System.Drawing.Color]::Transparent
        Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
    }
    $headerPanel.Controls.Add($lblTitle)
    

    
    # Create role count label (right-aligned)
    $lblCount = New-Object System.Windows.Forms.Label -Property @{
        Name      = 'lblActiveCount'
        Text      = '0 roles active'
        Location  = [System.Drawing.Point]::new(0, 27)
        Size      = [System.Drawing.Size]::new(150, 15)
        Font      = [System.Drawing.Font]::new("Segoe UI", 8)
        ForeColor = [System.Drawing.Color]::White
        BackColor = [System.Drawing.Color]::Transparent
        TextAlign = 'MiddleRight'
        Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    $headerPanel.Controls.Add($lblCount)
    
    # Handle header panel resize to reposition count label
    $headerPanel.Add_Resize({
            $lblCount = $this.Controls | Where-Object { $_.Name -eq 'lblActiveCount' }
            if ($lblCount) {
                $lblCount.Location = [System.Drawing.Point]::new($this.Width - 170, 27)
            }
        })
    

    
    # Create ListView for active roles
    $listView = New-Object System.Windows.Forms.ListView -Property @{
        Name          = 'lstActive'
        View          = [System.Windows.Forms.View]::Details
        FullRowSelect = $true
        GridLines     = $false
        CheckBoxes    = $true
        MultiSelect   = $true
        Scrollable    = $true
        Dock          = [System.Windows.Forms.DockStyle]::Fill
        Font          = [System.Drawing.Font]::new("Segoe UI", 9)
        BorderStyle   = [System.Windows.Forms.BorderStyle]::None
        BackColor     = [System.Drawing.Color]::White
    }
    
    # Add ListView columns (include a small first column for selection)
    [void]$listView.Columns.Add("", 30)
    [void]$listView.Columns.Add("Type", 90)
    [void]$listView.Columns.Add("Role Name", 220)
    [void]$listView.Columns.Add("Resource", 180)
    [void]$listView.Columns.Add("Scope", 120)
    [void]$listView.Columns.Add("Member Type", 110)
    [void]$listView.Columns.Add("Expires", 120)

    # Prevent checking permanent roles (no EndDateTime)
    $listView.Add_ItemCheck({
        param($sender, $e)
        if ($e.Index -ge 0 -and $e.Index -lt $sender.Items.Count) {
            $item = $sender.Items[$e.Index]
            if ($item -is [System.Windows.Forms.ListViewItem]) {
                $role = $item.Tag
                $isPermanent = -not ($role -and $role.PSObject.Properties['EndDateTime'] -and $role.EndDateTime)
                if ($isPermanent) {
                    # Block checking permanent items
                    $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
                }
            }
        }
    })
    
    # Use default header drawing (no owner-draw)
    
    # Handle column header clicks for select all functionality
    $listView.Add_ColumnClick({
        param($sender, $e)
        if ($e.Column -eq 0) {
            # Toggle all checkboxes for deactivatable (non-permanent) items only
            $allChecked = $true
            $hasDeactivatable = $false
            if ($sender.Items.Count -gt 0) {
                foreach ($item in $sender.Items) {
                    if ($item -is [System.Windows.Forms.ListViewItem]) {
                        $role = $item.Tag
                        $deactivatable = ($role -and $role.PSObject.Properties['EndDateTime'] -and $role.EndDateTime)
                        if ($deactivatable) {
                            $hasDeactivatable = $true
                            if (-not $item.Checked) {
                                $allChecked = $false
                                break
                            }
                        }
                    }
                }
                
                if ($hasDeactivatable) {
                    $newState = -not $allChecked
                    foreach ($item in $sender.Items) {
                        if ($item -is [System.Windows.Forms.ListViewItem]) {
                            $role = $item.Tag
                            $deactivatable = ($role -and $role.PSObject.Properties['EndDateTime'] -and $role.EndDateTime)
                            if ($deactivatable) {
                                $item.Checked = $newState
                            }
                        }
                    }
                }
            }
        }
    })
    
    # Create Select All button (after ListView is created)
    $btnSelectAll = New-Object System.Windows.Forms.Button -Property @{
        Text = "☐ Select All"
        Size = New-Object System.Drawing.Size(100, 25)
        Location = New-Object System.Drawing.Point(10, 10)
        FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
        ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
        Name = "btnActiveSelectAll"
    }
    # Keep a direct reference to the ListView to avoid closure scope issues
    $btnSelectAll.Tag = $listView
    
    # Add Select All click handler (use the ListView stored in Tag)
    $btnSelectAll.Add_Click({
        param($sender, $e)
        $lv = [System.Windows.Forms.ListView]$sender.Tag
        $allChecked = $true
        $hasDeactivatable = $false
        if ($lv -and $lv.Items.Count -gt 0) {
            foreach ($item in $lv.Items) {
                if ($item -is [System.Windows.Forms.ListViewItem]) {
                    $role = $item.Tag
                    $deactivatable = ($role -and $role.PSObject.Properties['EndDateTime'] -and $role.EndDateTime)
                    if ($deactivatable) {
                        $hasDeactivatable = $true
                        if (-not $item.Checked) {
                            $allChecked = $false
                            break
                        }
                    }
                }
            }
            
            if ($hasDeactivatable) {
                $newState = -not $allChecked
                foreach ($item in $lv.Items) {
                    if ($item -is [System.Windows.Forms.ListViewItem]) {
                        $role = $item.Tag
                        $deactivatable = ($role -and $role.PSObject.Properties['EndDateTime'] -and $role.EndDateTime)
                        if ($deactivatable) {
                            $item.Checked = $newState
                        }
                    }
                }
                
                # Update button text directly after change
                $btn = $lv.Parent.Parent.Controls['btnActiveSelectAll']
                if ($btn) {
                    $btn.Text = if ($newState) { "☑ Select All" } else { "☐ Select All" }
                }
            }
        }
    })
    
    # Update Select All button text based on selection state (resolve button by name to avoid scope issues)
    $listView.Add_ItemChecked({
        param($sender, $e)
        if ($sender.Items.Count -gt 0) {
            $allSelected = $true
            $anySelected = $false
            $hasDeactivatable = $false
            foreach ($item in $sender.Items) {
                if ($item -is [System.Windows.Forms.ListViewItem]) {
                    $role = $item.Tag
                    $deactivatable = ($role -and $role.PSObject.Properties['EndDateTime'] -and $role.EndDateTime)
                    if ($deactivatable) {
                        $hasDeactivatable = $true
                        if ($item.Checked) {
                            $anySelected = $true
                        } else {
                            $allSelected = $false
                        }
                    }
                }
            }
            $panelCtrl = $sender.Parent.Parent
            $btn = $panelCtrl.Controls['btnActiveSelectAll']
            if ($btn) {
                if ($hasDeactivatable -and $allSelected) {
                    $btn.Text = "☑ Select All"
                } elseif ($anySelected) {
                    $btn.Text = "☑ Select All"
                } else {
                    $btn.Text = "☐ Select All"
                }
            }
        }
    })
    
    # Add the Select All button to the panel
    $panel.Controls.Add($btnSelectAll)
    
    # Create ListView container
    $listViewContainer = New-Object System.Windows.Forms.Panel -Property @{
        Dock    = [System.Windows.Forms.DockStyle]::Fill
        Padding = New-Object System.Windows.Forms.Padding(0, 75, 0, 0)
    }
    $listViewContainer.Controls.Add($listView)
    $panel.Controls.Add($listViewContainer)
    
    return $panel
}
