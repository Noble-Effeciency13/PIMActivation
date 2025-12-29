function New-PIMEligibleRolesPanel {
    <#
    .SYNOPSIS
        Creates a panel containing eligible PIM roles with a ListView control.
    
    .DESCRIPTION
        Generates a Windows Forms panel with a header section and ListView for displaying
        eligible Privileged Identity Management (PIM) roles. The panel includes:
        - Header with title and role count
        - Multi-column ListView with checkboxes for role selection
        - Responsive layout with proper docking
    
    .OUTPUTS
        System.Windows.Forms.Panel
        Returns a configured panel containing the eligible roles ListView control.
    #>
    [CmdletBinding()]
    param()

    # Create main container panel
    $panel = New-Object System.Windows.Forms.Panel -Property @{
        Name        = 'pnlEligible'
        BackColor   = [System.Drawing.Color]::White
        BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        Dock        = [System.Windows.Forms.DockStyle]::Fill
        Padding     = New-Object System.Windows.Forms.Padding(0)
    }

    # Create header panel with branded background
    $headerPanel = New-Object System.Windows.Forms.Panel -Property @{
        Height    = 70
        BackColor = [System.Drawing.Color]::FromArgb(91, 203, 255)
        Dock      = [System.Windows.Forms.DockStyle]::Top
    }
    $panel.Controls.Add($headerPanel)

    # Title label
    $lblTitle = New-Object System.Windows.Forms.Label -Property @{
        Text      = 'Eligible Roles'
        Location  = [System.Drawing.Point]::new(15, 12)
        Size      = [System.Drawing.Size]::new(200, 25)
        Font      = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        ForeColor = [System.Drawing.Color]::White
        BackColor = [System.Drawing.Color]::Transparent
        Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
    }
    $headerPanel.Controls.Add($lblTitle)

    # Role count label
    $lblCount = New-Object System.Windows.Forms.Label -Property @{
        Name      = 'lblEligibleCount'
        Text      = '0 roles eligible'
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
            $lblCount = $this.Controls | Where-Object { $_.Name -eq 'lblEligibleCount' }
            if ($lblCount) {
                $lblCount.Location = [System.Drawing.Point]::new($this.Width - 170, 27)
            }
        })

    # Create ListView for eligible roles
    $listView = New-Object System.Windows.Forms.ListView -Property @{
        Name          = 'lstEligible'
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

    # Configure ListView columns for policy requirements (include a small first column for selection)
    [void]$listView.Columns.Add("", 30)
    [void]$listView.Columns.Add("Role Name", 220)
    [void]$listView.Columns.Add("Scope", 100)
    [void]$listView.Columns.Add("Member Type", 100)
    [void]$listView.Columns.Add("Max Duration", 100)
    [void]$listView.Columns.Add("MFA", 60)
    [void]$listView.Columns.Add("Auth Context", 120)
    [void]$listView.Columns.Add("Justification", 120)
    [void]$listView.Columns.Add("Ticket", 80)
    [void]$listView.Columns.Add("Approval", 100)
    [void]$listView.Columns.Add("Pending Approval", 120)

    # Handle column header clicks for select all functionality
    $listView.Add_ColumnClick({
        param($sender, $e)
        if ($e.Column -eq 0) {
            # Toggle all checkboxes using the ListView that raised the event
            $allChecked = $true
            if ($sender.Items.Count -gt 0) {
                foreach ($item in $sender.Items) {
                    if ($item -is [System.Windows.Forms.ListViewItem]) {
                        if (-not $item.Checked) {
                            $allChecked = $false
                            break
                        }
                    }
                }
                $newState = -not $allChecked
                foreach ($item in $sender.Items) {
                    if ($item -is [System.Windows.Forms.ListViewItem]) {
                        $item.Checked = $newState
                    }
                }
            }
        }
    })

    # Create Select All button
    $btnSelectAll = New-Object System.Windows.Forms.Button -Property @{
        Text = "☐ Select All"
        Size = New-Object System.Drawing.Size(100, 25)
        Location = New-Object System.Drawing.Point(10, 10)
        FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
        ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
        Name = "btnEligibleSelectAll"
    }
    $btnSelectAll.Tag = $listView

    # Add Select All click handler
    $btnSelectAll.Add_Click({
        param($sender, $e)
        $lv = [System.Windows.Forms.ListView]$sender.Tag
        $allChecked = $true
        if ($lv -and $lv.Items.Count -gt 0) {
            foreach ($item in $lv.Items) {
                if ($item -is [System.Windows.Forms.ListViewItem]) {
                    if (-not $item.Checked) {
                        $allChecked = $false
                        break
                    }
                }
            }
            $newState = -not $allChecked
            foreach ($item in $lv.Items) {
                if ($item -is [System.Windows.Forms.ListViewItem]) {
                    $item.Checked = $newState
                }
            }
            
            # Update button text directly after change
            $btn = $lv.Parent.Parent.Controls['btnEligibleSelectAll']
            if ($btn) {
                $btn.Text = if ($newState) { "☑ Select All" } else { "☐ Select All" }
            }
        }
    })

    # Update Select All button text based on selection state (resolve button by name to avoid scope issues)
    $listView.Add_ItemChecked({
        param($sender, $e)
        if ($sender.Items.Count -gt 0) {
            $allSelected = $true
            $anySelected = $false
            foreach ($item in $sender.Items) {
                if ($item -is [System.Windows.Forms.ListViewItem]) {
                    if ($item.Checked) {
                        $anySelected = $true
                    } else {
                        $allSelected = $false
                    }
                }
            }
            $panelCtrl = $sender.Parent.Parent
            $btn = $panelCtrl.Controls['btnEligibleSelectAll']
            if ($btn) {
                if ($allSelected) {
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

    # Create ListView container with proper spacing
    $listViewContainer = New-Object System.Windows.Forms.Panel -Property @{
        Dock    = [System.Windows.Forms.DockStyle]::Fill
        Padding = New-Object System.Windows.Forms.Padding(0, 75, 0, 0)
    }
    $listViewContainer.Controls.Add($listView)
    $panel.Controls.Add($listViewContainer)

    return $panel
}
