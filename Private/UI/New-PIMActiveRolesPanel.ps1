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
        Name = 'pnlActive'
        BackColor = [System.Drawing.Color]::White
        BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        Dock = [System.Windows.Forms.DockStyle]::Fill
        Padding = New-Object System.Windows.Forms.Padding(0)
    }
    
    # Create header panel with Microsoft blue background
    $headerPanel = New-Object System.Windows.Forms.Panel -Property @{
        Height = 50
        BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        Dock = [System.Windows.Forms.DockStyle]::Top
    }
    $panel.Controls.Add($headerPanel)
    
    # Create title label
    $lblTitle = New-Object System.Windows.Forms.Label -Property @{
        Text = 'Active Roles'
        Location = [System.Drawing.Point]::new(15, 12)
        Size = [System.Drawing.Size]::new(200, 25)
        Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        ForeColor = [System.Drawing.Color]::White
        BackColor = [System.Drawing.Color]::Transparent
        Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
    }
    $headerPanel.Controls.Add($lblTitle)
    
    # Create role count label (right-aligned)
    $lblCount = New-Object System.Windows.Forms.Label -Property @{
        Name = 'lblActiveCount'
        Text = '0 roles active'
        Location = [System.Drawing.Point]::new(0, 27)
        Size = [System.Drawing.Size]::new(150, 15)
        Font = [System.Drawing.Font]::new("Segoe UI", 8)
        ForeColor = [System.Drawing.Color]::White
        BackColor = [System.Drawing.Color]::Transparent
        TextAlign = 'MiddleRight'
        Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
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
        Name = 'lstActive'
        View = [System.Windows.Forms.View]::Details
        FullRowSelect = $true
        GridLines = $false
        CheckBoxes = $true
        MultiSelect = $true
        Scrollable = $true
        Dock = [System.Windows.Forms.DockStyle]::Fill
        Font = [System.Drawing.Font]::new("Segoe UI", 9)
        BorderStyle = [System.Windows.Forms.BorderStyle]::None
        BackColor = [System.Drawing.Color]::White
        OwnerDraw = $true
    }
    
    # Add ListView columns
    [void]$listView.Columns.Add("Type", 70)
    [void]$listView.Columns.Add("Role Name", 220)
    [void]$listView.Columns.Add("Resource", 180)
    [void]$listView.Columns.Add("Scope", 100)
    [void]$listView.Columns.Add("Member Type", 100)
    [void]$listView.Columns.Add("Expires", 100)
    
    # Custom header drawing for modern appearance
    $listView.Add_DrawColumnHeader({
        param($sender, $e)
        
        try {
            # Draw header background
            $headerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(240, 240, 240))
            $e.Graphics.FillRectangle($headerBrush, $e.Bounds)
            $headerBrush.Dispose()
            
            # Draw header borders
            $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 200, 200))
            $e.Graphics.DrawRectangle($borderPen, $e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width - 1, $e.Bounds.Height - 1)
            $borderPen.Dispose()
            
            # Configure text formatting
            $stringFormat = New-Object System.Drawing.StringFormat
            $stringFormat.Alignment = [System.Drawing.StringAlignment]::Near
            $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
            $stringFormat.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
            
            # Draw header text
            $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(32, 31, 30))
            $textBounds = $e.Bounds
            $textBounds.X += 5
            $textBounds.Width -= 10
            
            $rectF = [System.Drawing.RectangleF]::new($textBounds.X, $textBounds.Y, $textBounds.Width, $textBounds.Height)
            $e.Graphics.DrawString($e.Header.Text, $e.Font, $textBrush, $rectF, $stringFormat)
            
            # Cleanup resources
            $textBrush.Dispose()
            $stringFormat.Dispose()
        }
        catch {
            # Silently handle drawing errors to prevent UI disruption
        }
    })
    
    # Use default drawing for items and subitems
    $listView.Add_DrawItem({ param($sender, $e) $e.DrawDefault = $true })
    $listView.Add_DrawSubItem({ param($sender, $e) $e.DrawDefault = $true })
    
    # Enable double buffering for smooth rendering
    $listView.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Instance).SetValue($listView, $true, $null)
    
    # Add hover effect for better user experience
    $listView.Add_MouseMove({
        param($sender, $e)
        $hit = $sender.HitTest($e.Location)
        $sender.Cursor = if ($hit.Item) { [System.Windows.Forms.Cursors]::Hand } else { [System.Windows.Forms.Cursors]::Default }
    })
    
    # Create ListView container with proper spacing
    $listViewContainer = New-Object System.Windows.Forms.Panel -Property @{
        Dock = [System.Windows.Forms.DockStyle]::Fill
        Padding = New-Object System.Windows.Forms.Padding(0, 55, 0, 0)
    }
    $listViewContainer.Controls.Add($listView)
    $panel.Controls.Add($listViewContainer)
    
    return $panel
}
