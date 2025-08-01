function Initialize-PIMForm {
    <#
    .SYNOPSIS
        Initializes and configures the main PIM activation form with role management interface.
    
    .DESCRIPTION
        Creates a comprehensive Windows Forms UI for PIM (Privileged Identity Management) role activation.
        The form includes:
        - Header with title and account switching functionality
        - Split-panel view for active and eligible roles
        - Control panel with activation duration settings and action buttons
        - Keyboard shortcuts for common operations
        - Responsive layout that adapts to window resizing
    
    .PARAMETER SplashForm
        Optional splash screen form object to display loading progress and close after initialization.
        If provided, the splash screen will show progress updates during form creation.
    
    .OUTPUTS
        System.Windows.Forms.Form
        Returns the fully initialized and configured main form ready for display.
    
    .EXAMPLE
        $form = Initialize-PIMForm
        $form.ShowDialog()
        
        Creates and displays the PIM form without a splash screen.
    
    .EXAMPLE
        $splash = Show-LoadingSplash
        $form = Initialize-PIMForm -SplashForm $splash
        $form.ShowDialog()
        
        Creates the PIM form with splash screen progress updates.
    
    .NOTES
        - Form includes keyboard shortcuts: Ctrl+R (Refresh), Ctrl+A (Activate), Ctrl+D (Deactivate), Esc (Close)
        - Requires active PIM services connection via Connect-PIMServices
        - Form automatically loads and displays current role assignments
        - All UI elements follow Microsoft Fluent Design principles with Entra ID color scheme
    #>
    [CmdletBinding()]
    param(
        [PSCustomObject]$SplashForm
    )
    
    try {
        # Keep splash screen alive during form creation
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Update-LoadingStatus -SplashForm $SplashForm -Status "Creating user interface..." -Progress 60
        }
        
        # ===== MAIN FORM CREATION =====
        $form = New-Object System.Windows.Forms.Form -Property @{
            Text = 'PIM Role Activation'
            Size = [System.Drawing.Size]::new(1200, 900)
            MinimumSize = [System.Drawing.Size]::new(800, 600)
            StartPosition = 'CenterScreen'
            BackColor = [System.Drawing.Color]::FromArgb(245, 248, 250)  # Light blue-gray background
            KeyPreview = $true
        }
        
        # Update splash progress
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Update-LoadingStatus -SplashForm $SplashForm -Status "Building controls..." -Progress 65
        }
        
        # ===== HEADER PANEL =====
        # Create header with title and account management
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Update-LoadingStatus -SplashForm $SplashForm -Status "Creating header..." -Progress 70
        }
        
        $headerPanel = New-Object System.Windows.Forms.Panel -Property @{
            Height = 70
            BackColor = [System.Drawing.Color]::White
            BorderStyle = [System.Windows.Forms.BorderStyle]::None
            Dock = [System.Windows.Forms.DockStyle]::Top
        }
        
        # Add blue accent border at bottom of header
        $headerBorder = New-Object System.Windows.Forms.Label -Property @{
            Height = 2
            BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)  # Microsoft blue
            Dock = [System.Windows.Forms.DockStyle]::Bottom
        }
        $headerPanel.Controls.Add($headerBorder)
        
        # Main title label
        $titleLabel = New-Object System.Windows.Forms.Label -Property @{
            Text = 'PIM Role Activation'
            Location = [System.Drawing.Point]::new(20, 18)
            Size = [System.Drawing.Size]::new(400, 35)
            Font = [System.Drawing.Font]::new("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
            ForeColor = [System.Drawing.Color]::FromArgb(0, 78, 146)  # Dark blue
        }
        $headerPanel.Controls.Add($titleLabel)
        
        # Switch Account button with hover effects
        $btnSwitchAccount = New-Object System.Windows.Forms.Button -Property @{
            Name = 'btnSwitchAccount'
            Text = 'Switch Account'
            Size = [System.Drawing.Size]::new(140, 35)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
            Cursor = [System.Windows.Forms.Cursors]::Hand
            Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
            Location = [System.Drawing.Point]::new(1040, 17)
        }
        $btnSwitchAccount.FlatAppearance.BorderSize = 1
        $btnSwitchAccount.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $headerPanel.Controls.Add($btnSwitchAccount)
        
        # Add hover effects for Switch Account button
        $btnSwitchAccount.Add_MouseEnter({ 
            $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $this.ForeColor = [System.Drawing.Color]::White
        })
        $btnSwitchAccount.Add_MouseLeave({ 
            $this.BackColor = [System.Drawing.Color]::White
            $this.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        })
        
        # Current user display label
        $lblCurrentUser = New-Object System.Windows.Forms.Label -Property @{
            Name = 'lblCurrentUser'
            Text = if ($script:CurrentUser -and $script:CurrentUser.UserPrincipalName) { 
                "Signed in as: $($script:CurrentUser.UserPrincipalName)" 
            } else { 
                "Not signed in" 
            }
            Size = [System.Drawing.Size]::new(400, 20)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
            TextAlign = 'MiddleRight'
            Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)  # Medium gray
            Location = [System.Drawing.Point]::new(620, 25)
        }
        $headerPanel.Controls.Add($lblCurrentUser)

        # ===== CONTROL PANEL =====
        # Bottom panel with activation controls and buttons
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Update-LoadingStatus -SplashForm $SplashForm -Status "Creating control panel..." -Progress 75
        }
        
        $controlPanel = New-Object System.Windows.Forms.Panel -Property @{
            Name = 'pnlControls'
            Height = 120
            Dock = [System.Windows.Forms.DockStyle]::Bottom
            BackColor = [System.Drawing.Color]::White
            BorderStyle = [System.Windows.Forms.BorderStyle]::None
            Visible = $true
        }
        
        # Add separator line above control panel
        $controlSeparator = New-Object System.Windows.Forms.Label -Property @{
            Location = [System.Drawing.Point]::new(0, 0)
            Size = [System.Drawing.Size]::new(1200, 1)
            BackColor = [System.Drawing.Color]::FromArgb(229, 229, 229)
            Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
        }
        $controlPanel.Controls.Add($controlSeparator)
        
        # Duration selection group
        $durationGroup = New-Object System.Windows.Forms.GroupBox -Property @{
            Text = 'Activation Duration'
            Location = [System.Drawing.Point]::new(20, 5)
            Size = [System.Drawing.Size]::new(300, 100)
            Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
        }
        $controlPanel.Controls.Add($durationGroup)
        
        # Hours selection
        $lblHours = New-Object System.Windows.Forms.Label -Property @{
            Text = 'Hours:'
            Location = [System.Drawing.Point]::new(10, 25)
            Size = [System.Drawing.Size]::new(45, 20)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $durationGroup.Controls.Add($lblHours)
        
        $cmbHours = New-Object System.Windows.Forms.ComboBox -Property @{
            Name = 'cmbHours'
            Location = [System.Drawing.Point]::new(60, 23)
            Size = [System.Drawing.Size]::new(60, 23)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        0..23 | ForEach-Object { [void]$cmbHours.Items.Add($_) }
        $cmbHours.SelectedIndex = 8  # Default 8 hours
        $durationGroup.Controls.Add($cmbHours)
        
        # Minutes selection
        $lblMinutes = New-Object System.Windows.Forms.Label -Property @{
            Text = 'Minutes:'
            Location = [System.Drawing.Point]::new(130, 25)
            Size = [System.Drawing.Size]::new(55, 20)
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $durationGroup.Controls.Add($lblMinutes)
        
        $cmbMinutes = New-Object System.Windows.Forms.ComboBox -Property @{
            Name = 'cmbMinutes'
            Location = [System.Drawing.Point]::new(190, 23)
            Size = [System.Drawing.Size]::new(60, 23)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            Font = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        @(0, 30) | ForEach-Object { [void]$cmbMinutes.Items.Add($_) }
        $cmbMinutes.SelectedIndex = 0  # Default 0 minutes
        $durationGroup.Controls.Add($cmbMinutes)
        
        # Duration information label
        $lblDurationInfo = New-Object System.Windows.Forms.Label -Property @{
            Name = 'lblDurationInfo'
            Text = 'Max duration enforced per role'
            Location = [System.Drawing.Point]::new(10, 50)
            Size = [System.Drawing.Size]::new(280, 30)
            Font = [System.Drawing.Font]::new("Segoe UI", 8)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
        }
        $durationGroup.Controls.Add($lblDurationInfo)
        
        # Action buttons
        $btnRefresh = New-Object System.Windows.Forms.Button -Property @{
            Name = 'btnRefresh'
            Text = 'Refresh'
            Location = [System.Drawing.Point]::new(350, 40)
            Size = [System.Drawing.Size]::new(100, 35)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font = [System.Drawing.Font]::new("Segoe UI", 10)
            Cursor = [System.Windows.Forms.Cursors]::Hand
            Visible = $true
        }
        $btnRefresh.FlatAppearance.BorderSize = 1
        $btnRefresh.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
        $controlPanel.Controls.Add($btnRefresh)
        
        $btnDeactivate = New-Object System.Windows.Forms.Button -Property @{
            Name = 'btnDeactivate'
            Text = 'Deactivate Roles'
            Location = [System.Drawing.Point]::new(880, 40)
            Size = [System.Drawing.Size]::new(150, 35)
            BackColor = [System.Drawing.Color]::FromArgb(252, 80, 34)  # Entra orange/red
            ForeColor = [System.Drawing.Color]::White
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font = [System.Drawing.Font]::new("Segoe UI", 10)
            Cursor = [System.Windows.Forms.Cursors]::Hand
            Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
            Visible = $true
        }
        $btnDeactivate.FlatAppearance.BorderSize = 0
        $controlPanel.Controls.Add($btnDeactivate)
        
        $btnActivate = New-Object System.Windows.Forms.Button -Property @{
            Name = 'btnActivate'
            Text = 'Activate Roles'
            Location = [System.Drawing.Point]::new(1040, 40)
            Size = [System.Drawing.Size]::new(150, 35)
            BackColor = [System.Drawing.Color]::FromArgb(0, 123, 184)  # Entra blue
            ForeColor = [System.Drawing.Color]::White
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font = [System.Drawing.Font]::new("Segoe UI", 10)
            Cursor = [System.Windows.Forms.Cursors]::Hand
            Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
            Visible = $true
        }
        $btnActivate.FlatAppearance.BorderSize = 0
        $controlPanel.Controls.Add($btnActivate)
        
        # Add button hover effects
        $btnActivate.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 158) })
        $btnActivate.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 184) })
        
        $btnDeactivate.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(218, 72, 31) })
        $btnDeactivate.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(252, 80, 34) })
        
        $btnRefresh.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245) })
        $btnRefresh.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::White })

        # ===== SPLIT CONTAINER FOR ROLE PANELS =====
        # Create resizable split view for active and eligible roles
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Update-LoadingStatus -SplashForm $SplashForm -Status "Setting up role panels..." -Progress 80
        }
        
        $splitContainer = New-Object System.Windows.Forms.SplitContainer
        $splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
        $splitContainer.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $splitContainer.SplitterDistance = 350
        $splitContainer.SplitterWidth = 12  # Padding between panels
        $splitContainer.IsSplitterFixed = $false
        $splitContainer.Panel1MinSize = 150
        $splitContainer.Panel2MinSize = 150
        
        # Position with padding between header and control panel
        $splitContainer.Location = [System.Drawing.Point]::new(15, 70)
        $splitContainer.Size = [System.Drawing.Size]::new(1170, 710)
        $splitContainer.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                                [System.Windows.Forms.AnchorStyles]::Left -bor 
                                [System.Windows.Forms.AnchorStyles]::Right
        $splitContainer.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 250)
        
        # Create and add role panels to split container
        $activePanel = New-PIMActiveRolesPanel
        $eligiblePanel = New-PIMEligibleRolesPanel
        $splitContainer.Panel1.Controls.Add($activePanel)
        $splitContainer.Panel2.Controls.Add($eligiblePanel)
        
        # ===== ASSEMBLE FORM =====
        # Add all components to form in correct order
        $form.Controls.Add($headerPanel)
        $form.Controls.Add($controlPanel)
        $form.Controls.Add($splitContainer)
        
        # Ensure proper layout and visibility
        $controlPanel.Visible = $true
        $controlPanel.BringToFront()
        $form.PerformLayout()
        
        # ===== EVENT HANDLERS =====
        
        # Form Load - ensure proper control positioning
        $form.Add_Load({
            $headerPanel = $this.Controls | Where-Object { $_ -is [System.Windows.Forms.Panel] -and $_.Dock -eq [System.Windows.Forms.DockStyle]::Top } | Select-Object -First 1
            $controlPanel = $this.Controls | Where-Object { $_ -is [System.Windows.Forms.Panel] -and $_.Name -eq 'pnlControls' } | Select-Object -First 1
            $splitContainer = $this.Controls | Where-Object { $_ -is [System.Windows.Forms.SplitContainer] } | Select-Object -First 1
            
            # Position header controls
            if ($headerPanel) {
                $btnSwitchAccount = $headerPanel.Controls | Where-Object { $_.Name -eq 'btnSwitchAccount' } | Select-Object -First 1
                if ($btnSwitchAccount) {
                    $btnSwitchAccount.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 160, 17)
                }
                
                $lblCurrentUser = $headerPanel.Controls | Where-Object { $_.Name -eq 'lblCurrentUser' } | Select-Object -First 1
                if ($lblCurrentUser) {
                    $lblCurrentUser.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 580, 25)
                }
            }
            
            # Position control panel buttons
            if ($controlPanel) {
                $controlPanel.Visible = $true
                
                $btnDeactivate = $controlPanel.Controls | Where-Object { $_.Name -eq 'btnDeactivate' } | Select-Object -First 1
                if ($btnDeactivate) {
                    $btnDeactivate.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 335, 40)
                }
                
                $btnActivate = $controlPanel.Controls | Where-Object { $_.Name -eq 'btnActivate' } | Select-Object -First 1
                if ($btnActivate) {
                    $btnActivate.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 175, 40)
                }
            }
            
            # Resize split container
            if ($splitContainer) {
                $splitContainer.Size = [System.Drawing.Size]::new($this.ClientSize.Width - 30, $this.ClientSize.Height - 190)
            }
        })

        # Form Resize - maintain proper control positioning during window resize
        $form.Add_Resize({
            $headerPanel = $this.Controls | Where-Object { $_ -is [System.Windows.Forms.Panel] -and $_.Dock -eq [System.Windows.Forms.DockStyle]::Top } | Select-Object -First 1
            $controlPanel = $this.Controls | Where-Object { $_ -is [System.Windows.Forms.Panel] -and $_.Name -eq 'pnlControls' } | Select-Object -First 1
            $splitContainer = $this.Controls | Where-Object { $_ -is [System.Windows.Forms.SplitContainer] } | Select-Object -First 1
            
            # Reposition header controls
            if ($headerPanel) {
                $btnSwitchAccount = $headerPanel.Controls | Where-Object { $_.Name -eq 'btnSwitchAccount' } | Select-Object -First 1
                if ($btnSwitchAccount -and -not $btnSwitchAccount.IsDisposed) {
                    $btnSwitchAccount.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 160, 17)
                }
                
                $lblCurrentUser = $headerPanel.Controls | Where-Object { $_.Name -eq 'lblCurrentUser' } | Select-Object -First 1
                if ($lblCurrentUser -and -not $lblCurrentUser.IsDisposed) {
                    $lblCurrentUser.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 580, 25)
                }
            }
            
            # Reposition control panel buttons
            if ($controlPanel) {
                $btnDeactivate = $controlPanel.Controls | Where-Object { $_.Name -eq 'btnDeactivate' } | Select-Object -First 1
                if ($btnDeactivate -and -not $btnDeactivate.IsDisposed) {
                    $btnDeactivate.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 335, 40)
                }
                
                $btnActivate = $controlPanel.Controls | Where-Object { $_.Name -eq 'btnActivate' } | Select-Object -First 1
                if ($btnActivate -and -not $btnActivate.IsDisposed) {
                    $btnActivate.Location = [System.Drawing.Point]::new($this.ClientSize.Width - 175, 40)
                }
            }
            
            # Resize split container with padding
            if ($splitContainer -and -not $splitContainer.IsDisposed) {
                $splitContainer.Size = [System.Drawing.Size]::new($this.ClientSize.Width - 30, $this.ClientSize.Height - 190)
            }
        })
        
        # Switch Account button handler
        $btnSwitchAccount.Add_Click({
            $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                $form,
                "Switching accounts will close this window and restart the application.`n`nContinue?",
                "Switch Account",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($confirmResult -eq 'Yes') {
                try {
                    # Clean up current session
                    Disconnect-PIMServices
                    Clear-AuthenticationCache
                    $form.Close()
                    $script:RestartRequested = $true
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        $form,
                        "Error preparing account switch: $_",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                }
            }
        })
        
        # Activate Roles button handler
        $btnActivate.Add_Click({
            $eligibleListView = $form.Controls.Find("lstEligible", $true)[0]
            
            if ($eligibleListView -and $eligibleListView.CheckedItems.Count -gt 0) {
                # Get selected duration
                $hours = [int]$form.Controls.Find("cmbHours", $true)[0].SelectedItem
                $minutes = [int]$form.Controls.Find("cmbMinutes", $true)[0].SelectedItem
                
                # Store duration for activation handler
                $script:RequestedDuration = @{
                    Hours = $hours
                    Minutes = $minutes
                    TotalMinutes = ($hours * 60) + $minutes
                }
                
                # Execute role activation
                Invoke-PIMRoleActivation -CheckedItems $eligibleListView.CheckedItems -Form $form
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    "Please select at least one eligible role to activate.",
                    "No Selection",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        })
        
        # Deactivate button click handler
        $btnDeactivate.Add_Click({
            Write-Verbose "Deactivate button clicked"
            
            # Get checked items from active roles list
            $activeListView = $Form.Controls.Find('lstActive', $true)[0]
            if ($activeListView -and $activeListView.CheckedItems.Count -gt 0) {
                $checkedItems = @($activeListView.CheckedItems)
                Write-Verbose "Found $($checkedItems.Count) checked active role(s) for deactivation"
                
                # Call deactivation function
                Invoke-PIMRoleDeactivation -CheckedItems $checkedItems -Form $Form
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    'Please select at least one active role to deactivate.',
                    'No Selection',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        })
        
        # Refresh button click handler
        $btnRefresh.Add_Click({
            Write-Verbose "Refresh button clicked"
            
            # Clear role cache to force fresh data retrieval on manual refresh
            Write-Verbose "Clearing role cache to ensure fresh data on manual refresh"
            $script:CachedEligibleRoles = @()
            $script:CachedActiveRoles = @()
            $script:LastRoleFetchTime = $null
            
            # Show operation splash
            $refreshSplash = Show-OperationSplash -Title "Refreshing Roles" -InitialMessage "Updating role information..." -ShowProgressBar $true
            
            try {
                # Get the parent form
                $form = $this.FindForm()
                
                # Refresh both lists
                Update-PIMRolesList -Form $form -RefreshActive -RefreshEligible -SplashForm $refreshSplash
            }
            catch {
                Write-Error "Failed to refresh roles: $_"
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    "Failed to refresh role lists: $($_.Exception.Message)",
                    'Refresh Error',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            finally {
                # Ensure splash is closed
                if ($refreshSplash -and -not $refreshSplash.IsDisposed) {
                    $refreshSplash.Close()
                }
            }
        })

        # Keyboard shortcuts
        $form.Add_KeyDown({
            if ($_.Control) {
                switch ($_.KeyCode) {
                    'R' { $btnRefresh.PerformClick() }  # Ctrl+R: Refresh
                    'A' { $btnActivate.PerformClick() }  # Ctrl+A: Activate
                    'D' { $btnDeactivate.PerformClick() }  # Ctrl+D: Deactivate
                }
            }
            elseif ($_.KeyCode -eq 'Escape') {
                $form.Close()  # Esc: Close form
            }
        })
        $form.KeyPreview = $true
        
        # Update window title with current user
        if ($script:CurrentUser -and $script:CurrentUser.UserPrincipalName) {
            $form.Text = "PIM Role Activation - $($script:CurrentUser.UserPrincipalName)"
        }
        
        # Form cleanup handler
        $form.Add_FormClosing({
            param($sender, $e)
            # Cleanup handled in main application loop
        })
        
        # Store form reference for return
        $formToReturn = $form
        
        # ===== INITIALIZE ROLE DATA =====
        # Load role lists and update splash progress
        try {
            if ($SplashForm -and -not $SplashForm.IsDisposed) {
                Update-LoadingStatus -SplashForm $SplashForm -Status "Loading role data..." -Progress 85
            }
            
            # Load role data with progress updates (this will continue from 85% to 100%)
            $null = Update-PIMRolesList -Form $form -RefreshEligible -RefreshActive -SplashForm $SplashForm
            
            # Complete initialization - this is handled by Update-PIMRolesList now
            # No need to duplicate the final progress update here
        }
        catch {
            # Handle role loading errors gracefully
            if ($SplashForm -and -not $SplashForm.IsDisposed) {
                Close-LoadingSplash -SplashForm $SplashForm
            }
            
            # Add error indicators to role lists
            $activeList = $form.Controls.Find("lstActive", $true)[0]
            if ($activeList) {
                $errorItem = New-Object System.Windows.Forms.ListViewItem
                $errorItem.Text = "Error"
                [void]$errorItem.SubItems.Add("Failed to load active roles")
                [void]$errorItem.SubItems.Add($_.ToString())
                $errorItem.ForeColor = [System.Drawing.Color]::Red
                [void]$activeList.Items.Add($errorItem)
            }
            
            $eligibleList = $form.Controls.Find("lstEligible", $true)[0]
            if ($eligibleList) {
                $errorItem = New-Object System.Windows.Forms.ListViewItem
                $errorItem.Text = "Error loading eligible roles"
                [void]$errorItem.SubItems.Add($_.ToString())
                $errorItem.ForeColor = [System.Drawing.Color]::Red
                [void]$eligibleList.Items.Add($errorItem)
            }
        }
        
        return $formToReturn
    }
    catch {
        # Ensure splash screen cleanup on any error
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Close-LoadingSplash -SplashForm $SplashForm
        }
        
        Write-Error "Failed to initialize form: $_"
        throw
    }
}
