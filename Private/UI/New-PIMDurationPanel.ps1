function New-PIMDurationPanel {
    <#
    .SYNOPSIS
        Creates a duration and justification panel for PIM activation forms.
    
    .DESCRIPTION
        Creates a Windows Forms panel containing controls for setting activation duration 
        (hours/minutes) and providing justification text. Includes validation and character 
        counting for the justification field.
    
    .EXAMPLE
        $durationPanel = New-PIMDurationPanel
        $form.Controls.Add($durationPanel)
        
        Creates and adds a duration panel to a form.
    
    .OUTPUTS
        System.Windows.Forms.Panel
        A panel containing duration controls (hours/minutes dropdowns) and justification textbox.
    
    .NOTES
        - Default duration is set to 8 hours, 0 minutes
        - Minutes are limited to 15-minute intervals (0, 15, 30, 45)
        - Justification is limited to 500 characters with live counter
        - Panel supports anchoring for responsive layouts
    #>
    [CmdletBinding()]
    param()
    
    # Main container panel
    $panel = New-Object System.Windows.Forms.Panel -Property @{
        Name = 'pnlDuration'
        BackColor = [System.Drawing.Color]::White
        BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    
    # Duration controls group
    $grpDuration = New-Object System.Windows.Forms.GroupBox -Property @{
        Text = 'Activation Duration'
        Location = [System.Drawing.Point]::new(10, 10)
        Size = [System.Drawing.Size]::new(300, 100)
        Font = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    }
    
    # Hours selection
    $grpDuration.Controls.AddRange(@(
        (New-Object System.Windows.Forms.Label -Property @{
            Text = 'Hours:'
            Location = [System.Drawing.Point]::new(10, 30)
            Size = [System.Drawing.Size]::new(50, 20)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }),
        ($cmbHours = New-Object System.Windows.Forms.ComboBox -Property @{
            Name = 'cmbHours'
            Location = [System.Drawing.Point]::new(65, 28)
            Size = [System.Drawing.Size]::new(60, 23)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        })
    ))
    
    0..23 | ForEach-Object { [void]$cmbHours.Items.Add($_) }
    $cmbHours.SelectedIndex = 8
    
    # Minutes selection
    $grpDuration.Controls.AddRange(@(
        (New-Object System.Windows.Forms.Label -Property @{
            Text = 'Minutes:'
            Location = [System.Drawing.Point]::new(140, 30)
            Size = [System.Drawing.Size]::new(60, 20)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }),
        ($cmbMinutes = New-Object System.Windows.Forms.ComboBox -Property @{
            Name = 'cmbMinutes'
            Location = [System.Drawing.Point]::new(205, 28)
            Size = [System.Drawing.Size]::new(60, 23)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        })
    ))
    
    0..59 | Where-Object { $_ % 15 -eq 0 } | ForEach-Object { [void]$cmbMinutes.Items.Add($_) }
    $cmbMinutes.SelectedIndex = 0
    
    # Duration info
    $grpDuration.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Name = 'lblDurationInfo'
        Text = 'Maximum allowed duration will be enforced per role'
        Location = [System.Drawing.Point]::new(10, 60)
        Size = [System.Drawing.Size]::new(280, 30)
        Font = [System.Drawing.Font]::new("Segoe UI", 8)
        ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    }))
    
    $panel.Controls.Add($grpDuration)
    
    # Justification group
    $grpJustification = New-Object System.Windows.Forms.GroupBox -Property @{
        Text = 'Justification'
        Location = [System.Drawing.Point]::new(330, 10)
        Size = [System.Drawing.Size]::new(800, 100)
        Font = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    
    # Character counter
    $lblCharCount = New-Object System.Windows.Forms.Label -Property @{
        Name = 'lblCharCount'
        Text = '0 / 500'
        Location = [System.Drawing.Point]::new(720, 5)
        Size = [System.Drawing.Size]::new(70, 20)
        TextAlign = 'MiddleRight'
        Font = [System.Drawing.Font]::new("Segoe UI", 8)
        ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
        Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    
    # Justification textbox
    $txtJustification = New-Object System.Windows.Forms.TextBox -Property @{
        Name = 'txtJustification'
        Location = [System.Drawing.Point]::new(10, 25)
        Size = [System.Drawing.Size]::new(780, 65)
        Multiline = $true
        ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        Font = [System.Drawing.Font]::new("Segoe UI", 9)
        MaxLength = 500
        Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    
    # Character count update handler
    $txtJustification.Add_TextChanged({
        $charCount = $this.Parent.Controls.Find('lblCharCount', $true)[0]
        if ($charCount) {
            $charCount.Text = "$($this.Text.Length) / 500"
            $charCount.ForeColor = if ($this.Text.Length -ge 450) { 
                [System.Drawing.Color]::FromArgb(200, 100, 0) 
            } else { 
                [System.Drawing.Color]::FromArgb(100, 100, 100) 
            }
        }
    })
    
    $grpJustification.Controls.AddRange(@($lblCharCount, $txtJustification))
    $panel.Controls.Add($grpJustification)
    
    return $panel
}