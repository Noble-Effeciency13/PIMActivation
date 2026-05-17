function Show-PIMProfileNameDialog {
    [CmdletBinding()]
    param(
        [string]$Title = 'Save Activation Profile',
        [string]$InitialName = ''
    )

    $result = [PSCustomObject]@{
        Cancelled = $true
        ProfileName = ''
    }

    $form = New-Object System.Windows.Forms.Form -Property @{
        Text            = $Title
        ClientSize      = [System.Drawing.Size]::new(420, 135)
        StartPosition   = 'CenterParent'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        BackColor       = [System.Drawing.Color]::White
        TopMost         = $true
        ShowInTaskbar   = $false
    }

    $label = New-Object System.Windows.Forms.Label -Property @{
        Text     = 'Profile name'
        Location = [System.Drawing.Point]::new(12, 15)
        Size     = [System.Drawing.Size]::new(390, 20)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox -Property @{
        Text     = $InitialName
        Location = [System.Drawing.Point]::new(12, 40)
        Size     = [System.Drawing.Size]::new(390, 23)
        Font     = [System.Drawing.Font]::new('Segoe UI', 9)
    }
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button -Property @{
        Text         = 'OK'
        DialogResult = [System.Windows.Forms.DialogResult]::None
        Location     = [System.Drawing.Point]::new(232, 90)
        Size         = [System.Drawing.Size]::new(80, 30)
        BackColor    = [System.Drawing.Color]::FromArgb(0, 103, 184)
        ForeColor    = [System.Drawing.Color]::White
        FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
        Font         = [System.Drawing.Font]::new('Segoe UI', 9)
        Anchor       = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    $okButton.FlatAppearance.BorderSize = 0
    $okButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($textBox.Text)) {
            Show-TopMostMessageBox -Message 'Profile name is required.' -Title 'Activation Profile' -Icon Warning
            return
        }
        $result.ProfileName = $textBox.Text.Trim()
        $result.Cancelled = $false
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton = New-Object System.Windows.Forms.Button -Property @{
        Text         = 'Cancel'
        DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        Location     = [System.Drawing.Point]::new(322, 90)
        Size         = [System.Drawing.Size]::new(80, 30)
        BackColor    = [System.Drawing.Color]::White
        ForeColor    = [System.Drawing.Color]::FromArgb(32, 31, 30)
        FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
        Font         = [System.Drawing.Font]::new('Segoe UI', 9)
        Anchor       = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    }
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
    $cancelButton.Add_Click({
        $result.Cancelled = $true
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $form.Controls.AddRange(@($okButton, $cancelButton))
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.Add_Shown({ $textBox.Focus() })

    [void]$form.ShowDialog()
    $form.Dispose()
    return $result
}
