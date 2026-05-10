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

    .PARAMETER ShowAzureReducedScope
        Displays an optional Azure reduced-scope field for Azure Resource role activations.
    
    .EXAMPLE
        Show-PIMActivationDialog -RequiresJustification
        Shows dialog with required justification field.
    
    .EXAMPLE
        Show-PIMActivationDialog -RequiresTicket -OptionalJustification
        Shows dialog with required ticket field and optional justification.
    
    .OUTPUTS
        PSCustomObject
        Returns object with Justification, TicketNumber, AzureReducedScope, and Cancelled properties.
    
    .NOTES
        Requires System.Windows.Forms assembly for GUI display.
    #>
    [CmdletBinding()]
    param(
        [switch]$RequiresJustification,
        [switch]$RequiresTicket,
        [switch]$OptionalJustification,
        [switch]$ShowAzureReducedScope
    )
    
    # Initialize result object
    $result = [PSCustomObject]@{
        Justification = ""
        TicketNumber  = ""
        TicketSystem  = "ServiceNow"
        AzureReducedScope = ""
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
    $reducedScopeControl = $null
    
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
        $lblReducedScope = New-Object System.Windows.Forms.Label -Property @{
            Text     = "Reduced Scope"
            Location = [System.Drawing.Point]::new(10, $y)
            Size     = [System.Drawing.Size]::new(120, 23)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($lblReducedScope)

        $txtReducedScope = New-Object System.Windows.Forms.TextBox -Property @{
            Name     = "txtAzureReducedScope"
            Location = [System.Drawing.Point]::new(130, $y)
            Size     = [System.Drawing.Size]::new(280, 23)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $form.Controls.Add($txtReducedScope)
        $reducedScopeControl = $txtReducedScope

        $y += 28

        $lblReducedScopeNote = New-Object System.Windows.Forms.Label -Property @{
            Text      = "Example: /subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}. Leave empty to use the role's default scope."
            Location  = [System.Drawing.Point]::new(130, $y)
            Size      = [System.Drawing.Size]::new(340, 34)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
            Font      = [System.Drawing.Font]::new("Segoe UI", 8)
        }
        $form.Controls.Add($lblReducedScopeNote)

        $y += 42
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
    
    $form.ClientSize = [System.Drawing.Size]::new(484, [Math]::Max(311, $y + 75))
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

                if ($ShowAzureReducedScope -and $reducedScopeControl) {
                    $result.AzureReducedScope = $reducedScopeControl.Text.Trim()
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

        if ($ShowAzureReducedScope -and $reducedScopeControl) {
            $result.AzureReducedScope = $reducedScopeControl.Text.Trim()
        }
    }
    else {
        $result.Cancelled = $true
    }
    
    $form.Dispose()
    return $result
}