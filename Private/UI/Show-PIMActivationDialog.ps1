function Show-PIMActivationDialog {
    <#
    .SYNOPSIS
        Displays a dialog for PIM role activation requirements.
    
    .DESCRIPTION
        Shows a Windows Forms dialog to collect justification and/or ticket information
        required for PIM role activation. Returns user input or cancellation status.
    
    .PARAMETER RequiresJustification
        Specifies that justification text is mandatory for activation.
    
    .PARAMETER RequiresTicket
        Specifies that a ticket number is mandatory for activation.
    
    .PARAMETER OptionalJustification
        Displays justification field as optional with recommended usage note.
    
    .EXAMPLE
        Show-PIMActivationDialog -RequiresJustification
        Shows dialog with required justification field.
    
    .EXAMPLE
        Show-PIMActivationDialog -RequiresTicket -OptionalJustification
        Shows dialog with required ticket field and optional justification.
    
    .OUTPUTS
        PSCustomObject
        Returns object with Justification, TicketNumber, and Cancelled properties.
    
    .NOTES
        Requires System.Windows.Forms assembly for GUI display.
    #>
    [CmdletBinding()]
    param(
        [switch]$RequiresJustification,
        [switch]$RequiresTicket,
        [switch]$OptionalJustification
    )
    
    # Initialize result object
    $result = [PSCustomObject]@{
        Justification = ""
        TicketNumber = ""
        TicketSystem = "ServiceNow"
        Cancelled = $true
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
    
    # Add justification field if required or optional
    if ($RequiresJustification -or $OptionalJustification) {
        $labelText = if ($RequiresJustification) { "Justification (required):" } else { "Justification (optional - recommended):" }
        
        $lblJust = New-Object System.Windows.Forms.Label -Property @{
            Text     = $labelText
            Location = [System.Drawing.Point]::new(10, $y)
            Size     = [System.Drawing.Size]::new(460, 20)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
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
            Text = "Ticket Number *"
            Location = [System.Drawing.Point]::new(10, $y)
            Size = [System.Drawing.Size]::new(120, 23)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($lblTicket)
        
        $txtTicket = New-Object System.Windows.Forms.TextBox -Property @{
            Name = "txtTicket"
            Location = [System.Drawing.Point]::new(130, $y)
            Size = [System.Drawing.Size]::new(280, 23)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($txtTicket)
        
        $y += 30
        
        # Add ticket system dropdown
        $lblTicketSystem = New-Object System.Windows.Forms.Label -Property @{
            Text = "Ticket System"
            Location = [System.Drawing.Point]::new(10, $y)
            Size = [System.Drawing.Size]::new(120, 23)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($lblTicketSystem)
        
        $cmbTicketSystem = New-Object System.Windows.Forms.ComboBox -Property @{
            Name = "cmbTicketSystem"
            Location = [System.Drawing.Point]::new(130, $y)
            Size = [System.Drawing.Size]::new(280, 23)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
            DropDownStyle = 'DropDownList'
        }
        
        # Add common ticket systems
        $ticketSystems = @('ServiceNow', 'Jira', 'Azure DevOps', 'ServiceDesk Plus', 'BMC Remedy', 'Cherwell', 'Other')
        $cmbTicketSystem.Items.AddRange($ticketSystems)
        
        # Try to use saved preference or default to ServiceNow
        $savedSystem = Get-SavedTicketSystem
        if ($savedSystem -and $ticketSystems -contains $savedSystem) {
            $cmbTicketSystem.SelectedItem = $savedSystem
        } else {
            $cmbTicketSystem.SelectedIndex = 0  # Default to ServiceNow
        }
        
        $form.Controls.Add($cmbTicketSystem)
        
        $y += 35
    }
    
    # Add optional justification note
    if ($OptionalJustification -and -not $RequiresJustification) {
        $lblNote = New-Object System.Windows.Forms.Label -Property @{
            Text     = "Note: While justification is optional, providing a clear reason helps with audit trails."
            Location = [System.Drawing.Point]::new(10, $y)
            Size     = [System.Drawing.Size]::new(460, 40)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
            Font     = [System.Drawing.Font]::new("Segoe UI", 8)
        }
        $y += 45
        $form.Controls.Add($lblNote)
    }
    
    # Create OK button with styling and validation
    $okButton = New-Object System.Windows.Forms.Button -Property @{
        Text         = "OK"
        DialogResult = [System.Windows.Forms.DialogResult]::OK
        Location     = [System.Drawing.Point]::new(290, ($y + 20))
        Size         = [System.Drawing.Size]::new(80, 30)
        BackColor    = [System.Drawing.Color]::FromArgb(0, 103, 184)
        ForeColor    = [System.Drawing.Color]::White
        FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
        Font         = [System.Drawing.Font]::new("Segoe UI", 9)
        Cursor       = [System.Windows.Forms.Cursors]::Hand
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
            $_.Cancel = $true
            return
        }
        
        if ($RequiresTicket -and $txtTicket -and [string]::IsNullOrWhiteSpace($txtTicket.Text)) {
            Show-TopMostMessageBox -Message "Ticket number is required for these roles." -Title "Validation Error" -Icon Warning
            $_.Cancel = $true
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
        Location     = [System.Drawing.Point]::new(390, ($y + 20))
        Size         = [System.Drawing.Size]::new(80, 30)
        BackColor    = [System.Drawing.Color]::White
        ForeColor    = [System.Drawing.Color]::FromArgb(32, 31, 30)
        FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
        Font         = [System.Drawing.Font]::new("Segoe UI", 9)
        Cursor       = [System.Windows.Forms.Cursors]::Hand
    }
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
    
    $cancelButton.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245) })
    $cancelButton.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::White })
    
    # Add controls and set form properties
    $form.Controls.AddRange(@($okButton, $cancelButton))
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
    } else {
        $result.Cancelled = $true
    }
    
    $form.Dispose()
    return $result
}

function Get-SavedTicketSystem {
    <#
    .SYNOPSIS
        Retrieves the last used ticket system preference.
    
    .DESCRIPTION
        Gets the saved ticket system preference from the user's local profile.
    
    .OUTPUTS
        String - The saved ticket system name, or $null if not found.
    #>
    [CmdletBinding()]
    param()
    
    $prefsPath = Join-Path $env:LOCALAPPDATA "PIMActivation\preferences.json"
    
    if (Test-Path $prefsPath) {
        try {
            $prefs = Get-Content $prefsPath -Raw | ConvertFrom-Json
            return $prefs.TicketSystem
        }
        catch {
            Write-Verbose "Failed to load ticket system preference: $_"
        }
    }
    
    return $null
}

function Save-TicketSystemPreference {
    <#
    .SYNOPSIS
        Saves the user's ticket system preference.
    
    .PARAMETER System
        The ticket system name to save.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$System
    )
    
    $prefsDir = Join-Path $env:LOCALAPPDATA "PIMActivation"
    $prefsPath = Join-Path $prefsDir "preferences.json"
    
    # Ensure directory exists
    if (-not (Test-Path $prefsDir)) {
        New-Item -ItemType Directory -Path $prefsDir -Force | Out-Null
    }
    
    try {
        # Load existing preferences or create new
        $prefs = if (Test-Path $prefsPath) {
            Get-Content $prefsPath -Raw | ConvertFrom-Json
        } else {
            [PSCustomObject]@{}
        }
        
        # Update ticket system
        $prefs | Add-Member -NotePropertyName TicketSystem -NotePropertyValue $System -Force
        
        # Save
        $prefs | ConvertTo-Json | Set-Content $prefsPath -Force
        Write-Verbose "Saved ticket system preference: $System"
    }
    catch {
        Write-Warning "Failed to save ticket system preference: $_"
    }
}