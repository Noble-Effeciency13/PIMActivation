function New-PIMEligibleRolesPanel {
    <#
    .SYNOPSIS
        Creates a panel containing eligible PIM roles with a ListView control.
    
    .DESCRIPTION
        Generates a Windows Forms panel with a header section and ListView for displaying
        eligible Privileged Identity Management (PIM) roles. The panel includes:
        - Header with title and role count
        - Multi-column ListView with checkboxes for role selection
        - Custom styling and drawing for improved visual appearance
        - Responsive layout with proper docking
    
    .OUTPUTS
        System.Windows.Forms.Panel
        Returns a configured panel containing the eligible roles ListView control.
    
    .NOTES
        The ListView includes columns for Role Name, Scope, Max Duration, MFA requirements,
        Authentication Context, Justification, Ticket requirements, and Approval settings.
    #>
    [CmdletBinding()]
    param()
    
    # Create main container panel
    $panel = New-Object System.Windows.Forms.Panel -Property @{
        Name = 'pnlEligible'
        BackColor = [System.Drawing.Color]::White
        BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        Dock = [System.Windows.Forms.DockStyle]::Fill
        Padding = New-Object System.Windows.Forms.Padding(0)
    }
    
    # Create header panel with branded background
    $headerPanel = New-Object System.Windows.Forms.Panel -Property @{
        Height = 50
        BackColor = [System.Drawing.Color]::FromArgb(91, 203, 255)
        Dock = [System.Windows.Forms.DockStyle]::Top
    }
    $panel.Controls.Add($headerPanel)
    
    # Add title label to header
    $lblTitle = New-Object System.Windows.Forms.Label -Property @{
        Text = 'Eligible Roles'
        Location = [System.Drawing.Point]::new(15, 12)
        Size = [System.Drawing.Size]::new(200, 25)
        Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        ForeColor = [System.Drawing.Color]::FromArgb(0, 78, 146)
        BackColor = [System.Drawing.Color]::Transparent
        Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
    }
    $headerPanel.Controls.Add($lblTitle)
    
    # Add role count label (right-aligned)
    $lblCount = New-Object System.Windows.Forms.Label -Property @{
        Name = 'lblEligibleCount'
        Text = '0 roles available'
        Location = [System.Drawing.Point]::new(0, 27)
        Size = [System.Drawing.Size]::new(150, 15)
        Font = [System.Drawing.Font]::new("Segoe UI", 8)
        ForeColor = [System.Drawing.Color]::FromArgb(0, 78, 146)
        BackColor = [System.Drawing.Color]::Transparent
        TextAlign = 'MiddleRight'
        Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    $headerPanel.Controls.Add($lblCount)
    
    # Position count label dynamically on resize
    $headerPanel.Add_Resize({
        $lblCount = $this.Controls | Where-Object { $_.Name -eq 'lblEligibleCount' }
        if ($lblCount) {
            $lblCount.Location = [System.Drawing.Point]::new($this.Width - 170, 27)
        }
    })
    
    # Create main ListView for eligible roles
    $listView = New-Object System.Windows.Forms.ListView -Property @{
        Name = 'lstEligible'
        Location = [System.Drawing.Point]::new(0, 50)
        Size = [System.Drawing.Size]::new(780, 200)
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
    }
    
    # Configure ListView columns for policy requirements
    [void]$listView.Columns.Add("Role Name", 250)
    [void]$listView.Columns.Add("Scope", 150)
    [void]$listView.Columns.Add("MemberType", 100)
    [void]$listView.Columns.Add("Max Duration", 200)
    [void]$listView.Columns.Add("MFA", 50)
    [void]$listView.Columns.Add("Auth Context", 110)
    [void]$listView.Columns.Add("Justification", 100)
    [void]$listView.Columns.Add("Ticket", 60)
    [void]$listView.Columns.Add("Approval", 70)
    [void]$listView.Columns.Add("Pending Approval", 110)

    # Enable custom drawing for styled headers
    $listView.OwnerDraw = $true
    
    # Custom header drawing for branded appearance
    $listView.Add_DrawColumnHeader({
        param($sender, $e)
        
        try {
            # Draw header background
            $headerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(235, 245, 250))
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
    
    # Use default drawing for list items and subitems
    $listView.Add_DrawItem({
        param($sender, $e)
        $e.DrawDefault = $true
    })
    
    $listView.Add_DrawSubItem({
        param($sender, $e)
        $e.DrawDefault = $true
    })

    # Add interactive cursor feedback
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
