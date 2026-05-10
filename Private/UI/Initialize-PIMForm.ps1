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
    
    .PARAMETER EnableParallelProcessing
        Switch to enable parallel processing of Azure subscriptions during role enumeration.
        Requires PowerShell 7+ and significantly improves performance with multiple subscriptions.
    
    .PARAMETER ThrottleLimit
        Maximum number of parallel operations for Azure subscription processing.
        Default is 6. Only used when EnableParallelProcessing is specified.
    
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
    
    .EXAMPLE
        $form = Initialize-PIMForm -ThrottleLimit 8
        $form.ShowDialog()
        
        Creates the PIM form with parallel Azure subscription processing enabled.
    
    .NOTES
        - Form includes keyboard shortcuts: Ctrl+R (Refresh), Ctrl+A (Activate), Ctrl+D (Deactivate), Esc (Close)
        - Requires active PIM services connection via Connect-PIMServices
        - Form automatically loads and displays current role assignments
        - All UI elements follow Microsoft Fluent Design principles with Entra ID color scheme
    #>
    [CmdletBinding()]
    param(
        [PSCustomObject]$SplashForm,
        
        [switch]$DisableParallelProcessing,
        
        [int]$ThrottleLimit = 10
    )
    
    try {
        # Keep splash screen alive during form creation
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Update-LoadingStatus -SplashForm $SplashForm -Status "Creating user interface..." -Progress 60
        }
        
        # ===== MAIN FORM CREATION =====
        $form = New-Object System.Windows.Forms.Form -Property @{
            Text          = 'PIM Role Activation'
            Size          = [System.Drawing.Size]::new(1200, 900)
            MinimumSize   = [System.Drawing.Size]::new(800, 600)
            StartPosition = 'CenterScreen'
            BackColor     = [System.Drawing.Color]::FromArgb(245, 248, 250)  # Light blue-gray background
            KeyPreview    = $true
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
            Height      = 70
            BackColor   = [System.Drawing.Color]::White
            BorderStyle = [System.Windows.Forms.BorderStyle]::None
            Dock        = [System.Windows.Forms.DockStyle]::Top
        }
        
        # Add blue accent border at bottom of header
        $headerBorder = New-Object System.Windows.Forms.Label -Property @{
            Height    = 2
            BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)  # Microsoft blue
            Dock      = [System.Windows.Forms.DockStyle]::Bottom
        }
        $headerPanel.Controls.Add($headerBorder)
        
        # Main title label
        $titleLabel = New-Object System.Windows.Forms.Label -Property @{
            Text      = 'PIM Role Activation'
            Location  = [System.Drawing.Point]::new(20, 18)
            Size      = [System.Drawing.Size]::new(400, 35)
            Font      = [System.Drawing.Font]::new("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
            ForeColor = [System.Drawing.Color]::FromArgb(0, 78, 146)  # Dark blue
        }
        $headerPanel.Controls.Add($titleLabel)
        
        # Switch Account button with hover effects
        $btnSwitchAccount = New-Object System.Windows.Forms.Button -Property @{
            Name      = 'btnSwitchAccount'
            Text      = 'Switch Account'
            Size      = [System.Drawing.Size]::new(140, 35)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new("Segoe UI", 9)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
            Location  = [System.Drawing.Point]::new(1040, 17)
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
            Name      = 'lblCurrentUser'
            Text      = if ($script:CurrentUser -and $script:CurrentUser.UserPrincipalName) { 
                "Signed in as: $($script:CurrentUser.UserPrincipalName)" 
            }
            else { 
                "Not signed in" 
            }
            Size      = [System.Drawing.Size]::new(400, 20)
            Font      = [System.Drawing.Font]::new("Segoe UI", 9)
            TextAlign = 'MiddleRight'
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)  # Medium gray
            Location  = [System.Drawing.Point]::new(620, 25)
        }
        $headerPanel.Controls.Add($lblCurrentUser)

        # ===== CONTROL PANEL =====
        # Bottom panel with activation controls and buttons
        if ($SplashForm -and -not $SplashForm.IsDisposed) {
            Update-LoadingStatus -SplashForm $SplashForm -Status "Creating control panel..." -Progress 75
        }
        
        $controlPanel = New-Object System.Windows.Forms.Panel -Property @{
            Name        = 'pnlControls'
            Height      = 160
            Dock        = [System.Windows.Forms.DockStyle]::Bottom
            BackColor   = [System.Drawing.Color]::White
            BorderStyle = [System.Windows.Forms.BorderStyle]::None
            Visible     = $true
        }
        
        # Add separator line above control panel
        $controlSeparator = New-Object System.Windows.Forms.Label -Property @{
            Location  = [System.Drawing.Point]::new(0, 0)
            Size      = [System.Drawing.Size]::new(1200, 1)
            BackColor = [System.Drawing.Color]::FromArgb(229, 229, 229)
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
        }
        $controlPanel.Controls.Add($controlSeparator)
        
        # Duration selection group
        $durationGroup = New-Object System.Windows.Forms.GroupBox -Property @{
            Text      = 'Activation Duration'
            Location  = [System.Drawing.Point]::new(20, 5)
            Size      = [System.Drawing.Size]::new(300, 100)
            Font      = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
        }
        $controlPanel.Controls.Add($durationGroup)
        
        # Hours selection
        $lblHours = New-Object System.Windows.Forms.Label -Property @{
            Text     = 'Hours:'
            Location = [System.Drawing.Point]::new(10, 25)
            Size     = [System.Drawing.Size]::new(45, 20)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $durationGroup.Controls.Add($lblHours)
        
        $cmbHours = New-Object System.Windows.Forms.ComboBox -Property @{
            Name          = 'cmbHours'
            Location      = [System.Drawing.Point]::new(60, 23)
            Size          = [System.Drawing.Size]::new(60, 23)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            Font          = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        0..23 | ForEach-Object { [void]$cmbHours.Items.Add($_) }
        $cmbHours.SelectedIndex = 8  # Default 8 hours
        $durationGroup.Controls.Add($cmbHours)
        
        # Minutes selection
        $lblMinutes = New-Object System.Windows.Forms.Label -Property @{
            Text     = 'Minutes:'
            Location = [System.Drawing.Point]::new(130, 25)
            Size     = [System.Drawing.Size]::new(55, 20)
            Font     = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        $durationGroup.Controls.Add($lblMinutes)
        
        $cmbMinutes = New-Object System.Windows.Forms.ComboBox -Property @{
            Name          = 'cmbMinutes'
            Location      = [System.Drawing.Point]::new(190, 23)
            Size          = [System.Drawing.Size]::new(60, 23)
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            Font          = [System.Drawing.Font]::new("Segoe UI", 9)
        }
        @(0, 30) | ForEach-Object { [void]$cmbMinutes.Items.Add($_) }
        $cmbMinutes.SelectedIndex = 0  # Default 0 minutes
        $durationGroup.Controls.Add($cmbMinutes)
        
        # Duration information label
        $lblDurationInfo = New-Object System.Windows.Forms.Label -Property @{
            Name      = 'lblDurationInfo'
            Text      = 'Max duration enforced per role'
            Location  = [System.Drawing.Point]::new(10, 50)
            Size      = [System.Drawing.Size]::new(280, 30)
            Font      = [System.Drawing.Font]::new("Segoe UI", 8)
            ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
        }
        $durationGroup.Controls.Add($lblDurationInfo)
        
        # Action buttons
        $btnRefresh = New-Object System.Windows.Forms.Button -Property @{
            Name      = 'btnRefresh'
            Text      = 'Refresh'
            Location  = [System.Drawing.Point]::new(350, 40)
            Size      = [System.Drawing.Size]::new(100, 35)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new("Segoe UI", 10)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Visible   = $true
        }
        $btnRefresh.FlatAppearance.BorderSize = 1
        $btnRefresh.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
        $controlPanel.Controls.Add($btnRefresh)

        # Full Refresh button – clears all caches and re-fetches everything including policies
        $btnFullRefresh = New-Object System.Windows.Forms.Button -Property @{
            Name      = 'btnFullRefresh'
            Text      = '↻ Full Refresh'
            Location  = [System.Drawing.Point]::new(460, 40)
            Size      = [System.Drawing.Size]::new(130, 35)
            BackColor = [System.Drawing.Color]::White
            ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new("Segoe UI", 10)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Visible   = $true
        }
        $btnFullRefresh.FlatAppearance.BorderSize = 1
        $btnFullRefresh.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $controlPanel.Controls.Add($btnFullRefresh)

        # ── Role type toggle row (right side, second row under De-/Activate buttons) ─
        $lblRoleTypes = New-Object System.Windows.Forms.Label -Property @{
            Text      = 'Include:'
            Location  = [System.Drawing.Point]::new(754, 88)
            AutoSize  = $true
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
        }
        $controlPanel.Controls.Add($lblRoleTypes)

        $chkEntra = New-Object System.Windows.Forms.CheckBox -Property @{
            Name      = 'chkIncludeEntra'
            Text      = 'Entra Roles'
            Location  = [System.Drawing.Point]::new(819, 84)
            AutoSize  = $true
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            Checked   = $script:IncludeEntraRoles
            Cursor    = [System.Windows.Forms.Cursors]::Hand
        }
        $controlPanel.Controls.Add($chkEntra)

        $chkGroups = New-Object System.Windows.Forms.CheckBox -Property @{
            Name      = 'chkIncludeGroups'
            Text      = 'Groups'
            Location  = [System.Drawing.Point]::new(934, 84)
            AutoSize  = $true
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            Checked   = $script:IncludeGroups
            Cursor    = [System.Windows.Forms.Cursors]::Hand
        }
        $controlPanel.Controls.Add($chkGroups)

        $chkAzure = New-Object System.Windows.Forms.CheckBox -Property @{
            Name      = 'chkIncludeAzure'
            Text      = 'Azure Resources'
            Location  = [System.Drawing.Point]::new(1019, 84)
            AutoSize  = $true
            Font      = [System.Drawing.Font]::new('Segoe UI', 9)
            Checked   = $script:IncludeAzureResources
            Cursor    = [System.Windows.Forms.Cursors]::Hand
        }
        $controlPanel.Controls.Add($chkAzure)
        
        $btnDeactivate = New-Object System.Windows.Forms.Button -Property @{
            Name      = 'btnDeactivate'
            Text      = 'Deactivate Roles'
            Location  = [System.Drawing.Point]::new(880, 40)
            Size      = [System.Drawing.Size]::new(150, 35)
            BackColor = [System.Drawing.Color]::FromArgb(252, 80, 34)  # Entra orange/red
            ForeColor = [System.Drawing.Color]::White
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new("Segoe UI", 10)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
            Visible   = $true
        }
        $btnDeactivate.FlatAppearance.BorderSize = 0
        $controlPanel.Controls.Add($btnDeactivate)
        
        $btnActivate = New-Object System.Windows.Forms.Button -Property @{
            Name      = 'btnActivate'
            Text      = 'Activate Roles'
            Location  = [System.Drawing.Point]::new(1040, 40)
            Size      = [System.Drawing.Size]::new(150, 35)
            BackColor = [System.Drawing.Color]::FromArgb(0, 123, 184)  # Entra blue
            ForeColor = [System.Drawing.Color]::White
            FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            Font      = [System.Drawing.Font]::new("Segoe UI", 10)
            Cursor    = [System.Windows.Forms.Cursors]::Hand
            Anchor    = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
            Visible   = $true
        }
        $btnActivate.FlatAppearance.BorderSize = 0
        $controlPanel.Controls.Add($btnActivate)

        $toggleGap = 12
    $toggleTextRightPadding = 6
    $toggleRight = $btnActivate.Location.X + $btnActivate.Width + $toggleTextRightPadding
        $chkAzure.Location = [System.Drawing.Point]::new($toggleRight - $chkAzure.Width, 84)
        $chkGroups.Location = [System.Drawing.Point]::new($chkAzure.Location.X - $toggleGap - $chkGroups.Width, 84)
        $chkEntra.Location = [System.Drawing.Point]::new($chkGroups.Location.X - $toggleGap - $chkEntra.Width, 84)
        $lblRoleTypes.Location = [System.Drawing.Point]::new($chkEntra.Location.X - $toggleGap - $lblRoleTypes.Width, 88)
        
        # Add button hover effects
        $btnActivate.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 158) })
        $btnActivate.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 184) })
        
        $btnDeactivate.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(218, 72, 31) })
        $btnDeactivate.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(252, 80, 34) })
        
        $btnRefresh.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245) })
        $btnRefresh.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::White })

        $btnFullRefresh.Add_MouseEnter({
            $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $this.ForeColor = [System.Drawing.Color]::White
        })
        $btnFullRefresh.Add_MouseLeave({
            $this.BackColor = [System.Drawing.Color]::White
            $this.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        })

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
        $splitContainer.Size = [System.Drawing.Size]::new(1170, 670)
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

                    # Position role-type toggles: second row, right-aligned under De-/Activate buttons
                    $chkAzCtrl  = $controlPanel.Controls | Where-Object { $_.Name -eq 'chkIncludeAzure' }  | Select-Object -First 1
                    $chkGrpCtrl = $controlPanel.Controls | Where-Object { $_.Name -eq 'chkIncludeGroups' } | Select-Object -First 1
                    $chkEntCtrl = $controlPanel.Controls | Where-Object { $_.Name -eq 'chkIncludeEntra' }  | Select-Object -First 1
                    $lblIncCtrl = $controlPanel.Controls | Where-Object { $_.Text -eq 'Include:' }          | Select-Object -First 1
                    if ($btnActivate -and $chkAzCtrl -and $chkGrpCtrl -and $chkEntCtrl -and $lblIncCtrl) {
                        $toggleGap = 12
                        $toggleTextRightPadding = 6
                        $toggleRight = $btnActivate.Location.X + $btnActivate.Width + $toggleTextRightPadding
                        $chkAzCtrl.Location = [System.Drawing.Point]::new($toggleRight - $chkAzCtrl.Width, 84)
                        $chkGrpCtrl.Location = [System.Drawing.Point]::new($chkAzCtrl.Location.X - $toggleGap - $chkGrpCtrl.Width, 84)
                        $chkEntCtrl.Location = [System.Drawing.Point]::new($chkGrpCtrl.Location.X - $toggleGap - $chkEntCtrl.Width, 84)
                        $lblIncCtrl.Location = [System.Drawing.Point]::new($chkEntCtrl.Location.X - $toggleGap - $lblIncCtrl.Width, 88)
                    }
                }
            
                # Resize split container
                if ($splitContainer) {
                    $splitContainer.Size = [System.Drawing.Size]::new($this.ClientSize.Width - 30, $this.ClientSize.Height - 230)
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

                    # Reposition role-type toggles: second row, right-aligned under De-/Activate buttons
                    $chkAzCtrl  = $controlPanel.Controls | Where-Object { $_.Name -eq 'chkIncludeAzure' }  | Select-Object -First 1
                    $chkGrpCtrl = $controlPanel.Controls | Where-Object { $_.Name -eq 'chkIncludeGroups' } | Select-Object -First 1
                    $chkEntCtrl = $controlPanel.Controls | Where-Object { $_.Name -eq 'chkIncludeEntra' }  | Select-Object -First 1
                    $lblIncCtrl = $controlPanel.Controls | Where-Object { $_.Text -eq 'Include:' }          | Select-Object -First 1
                    if ($btnActivate -and $chkAzCtrl -and -not $chkAzCtrl.IsDisposed -and $chkGrpCtrl -and -not $chkGrpCtrl.IsDisposed -and $chkEntCtrl -and -not $chkEntCtrl.IsDisposed -and $lblIncCtrl -and -not $lblIncCtrl.IsDisposed) {
                        $toggleGap = 12
                        $toggleTextRightPadding = 6
                        $toggleRight = $btnActivate.Location.X + $btnActivate.Width + $toggleTextRightPadding
                        $chkAzCtrl.Location = [System.Drawing.Point]::new($toggleRight - $chkAzCtrl.Width, 84)
                        $chkGrpCtrl.Location = [System.Drawing.Point]::new($chkAzCtrl.Location.X - $toggleGap - $chkGrpCtrl.Width, 84)
                        $chkEntCtrl.Location = [System.Drawing.Point]::new($chkGrpCtrl.Location.X - $toggleGap - $chkEntCtrl.Width, 84)
                        $lblIncCtrl.Location = [System.Drawing.Point]::new($chkEntCtrl.Location.X - $toggleGap - $lblIncCtrl.Width, 88)
                    }
                }
            
                # Resize split container with padding
                if ($splitContainer -and -not $splitContainer.IsDisposed) {
                    $splitContainer.Size = [System.Drawing.Size]::new($this.ClientSize.Width - 30, $this.ClientSize.Height - 230)
                }
            })
        
        # Switch Account button handler
        $btnSwitchAccount.Add_Click({
                $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    "Switching accounts will close this window and restart the application.{0}{0}Continue?" -f [Environment]::NewLine,
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
                        Hours        = $hours
                        Minutes      = $minutes
                        TotalMinutes = ($hours * 60) + $minutes
                    }
                
                    # Execute role activation
                    Invoke-PIMRoleActivation -CheckedItems $eligibleListView.CheckedItems -Form $form
                    
                    # Clear checkboxes in eligible roles list after activation
                    foreach ($item in $eligibleListView.Items) {
                        if ($item -is [System.Windows.Forms.ListViewItem]) {
                            $item.Checked = $false
                        }
                    }
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
                    # Filter out permanent roles (no EndDateTime)
                    $checkedItems = @(@($activeListView.CheckedItems) | Where-Object {
                            $_ -is [System.Windows.Forms.ListViewItem] -and $_.Tag -and $_.Tag.PSObject.Properties['EndDateTime'] -and $_.Tag.EndDateTime
                        })
                    $checkedCount = ($checkedItems | Measure-Object).Count
                    Write-Verbose "Found $checkedCount deactivatable active role(s) after filtering permanent roles"
                    
                    if ($checkedCount -gt 0) {
                        # Call deactivation function
                        Invoke-PIMRoleDeactivation -CheckedItems $checkedItems -Form $Form
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show(
                            $form,
                            'No deactivatable roles selected. Permanent roles cannot be deactivated.',
                            'No Deactivatable Selection',
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Information
                        )
                    }
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
            
                # Refresh ACTIVE roles only: clear active role cache to get fresh data
                Write-Verbose "Preparing active-only refresh: clearing active role cache for fresh data"
                $script:CachedActiveRoles = $null
            
                # Show operation splash
                $refreshSplash = Show-OperationSplash -Title "Refreshing Roles" -InitialMessage "Updating role information..." -ShowProgressBar $true
            
                try {
                    # Get the parent form
                    $form = $this.FindForm()
                
                    # Refresh ACTIVE list only; fetch fresh Azure data to show recent activations
                    Update-PIMRolesList -Form $form -RefreshActive -SplashForm $refreshSplash -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
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

        # Full Refresh button - clears all caches and re-fetches everything including policies
        $btnFullRefresh.Add_Click({
                Write-Verbose "Full Refresh button clicked - clearing all caches"
                $form = $this.FindForm()

                # Clear ALL caches so everything is re-fetched from scratch
                $script:CachedEligibleRoles       = $null
                $script:CachedActiveRoles         = $null
                $script:LastRoleFetchTime         = $null
                $script:PolicyCache               = @{}
                $script:AuthenticationContextCache = @{}
                $script:AzureRolesCache           = @()
                $script:AzureRolesCacheTime       = $null
                $script:DirtyAzureSubscriptions   = @()
                $script:DirtyManagementGroups     = @()

                $refreshSplash = Show-OperationSplash -Title "Full Refresh" -InitialMessage "Re-fetching all roles and policies..." -ShowProgressBar $true

                try {
                    Update-PIMRolesList -Form $form -RefreshActive -RefreshEligible -SplashForm $refreshSplash -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
                }
                catch {
                    Write-Error "Full refresh failed: $_"
                    [System.Windows.Forms.MessageBox]::Show(
                        $form,
                        "Full refresh failed: $($_.Exception.Message)",
                        'Refresh Error',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                }
                finally {
                    if ($refreshSplash -and -not $refreshSplash.IsDisposed) {
                        $refreshSplash.Close()
                    }
                }
            })

        # Role type toggle handlers - clear caches and trigger full refresh when changed
        $chkEntra.Add_CheckedChanged({
                $script:IncludeEntraRoles = $this.Checked
                Write-Verbose "Entra Roles toggled: $($script:IncludeEntraRoles)"
                $script:CachedEligibleRoles = $null
                $script:CachedActiveRoles   = $null
                $script:LastRoleFetchTime   = $null
                $f = $this.FindForm()
                $btn = $f.Controls[0].Controls | Where-Object { $_.Name -eq 'pnlControls' } | Select-Object -First 1 | ForEach-Object { $_.Controls | Where-Object { $_.Name -eq 'btnFullRefresh' } | Select-Object -First 1 }
                if (-not $btn) { $btn = $f.Controls | ForEach-Object { $_.Controls } | Where-Object { $_.Name -eq 'btnFullRefresh' } | Select-Object -First 1 }
                if ($btn) { $btn.PerformClick() }
            })

        $chkGroups.Add_CheckedChanged({
                $script:IncludeGroups = $this.Checked
                Write-Verbose "Groups toggled: $($script:IncludeGroups)"
                $script:CachedEligibleRoles = $null
                $script:CachedActiveRoles   = $null
                $script:LastRoleFetchTime   = $null
                $f = $this.FindForm()
                $btn = $f.Controls | ForEach-Object { $_.Controls } | Where-Object { $_.Name -eq 'btnFullRefresh' } | Select-Object -First 1
                if ($btn) { $btn.PerformClick() }
            })

        $chkAzure.Add_CheckedChanged({
                if ($this.Checked) {
                    # Validate Azure modules are available before enabling
                    $azureReady = Initialize-AzureResourceSupport
                    if (-not $azureReady) {
                        # Suppress the event while reverting to avoid re-entrancy
                        $script:_suppressAzureToggle = $true
                        $this.Checked = $false
                        $script:_suppressAzureToggle = $false
                        [System.Windows.Forms.MessageBox]::Show(
                            $this.FindForm(),
                            "Azure Resource support is not available.`nRequired Az PowerShell modules could not be loaded.",
                            'Azure Not Available',
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        )
                        return
                    }
                }
                if ($script:_suppressAzureToggle) { return }
                $script:IncludeAzureResources = $this.Checked
                Write-Verbose "Azure Resources toggled: $($script:IncludeAzureResources)"
                $script:CachedEligibleRoles   = $null
                $script:CachedActiveRoles     = $null
                $script:LastRoleFetchTime     = $null
                $script:AzureRolesCache       = @()
                $script:AzureRolesCacheTime   = $null
                $f = $this.FindForm()
                $btn = $f.Controls | ForEach-Object { $_.Controls } | Where-Object { $_.Name -eq 'btnFullRefresh' } | Select-Object -First 1
                if ($btn) { $btn.PerformClick() }
            })

        # Keyboard shortcuts
        $form.Add_KeyDown({
                if ($_.Control) {
                    $f    = $this
                    $pnl  = $f.Controls | Where-Object { $_.Name -eq 'pnlControls' } | Select-Object -First 1
                    switch ($_.KeyCode) {
                        'R' { ($pnl.Controls | Where-Object { $_.Name -eq 'btnRefresh' }     | Select-Object -First 1)?.PerformClick() }  # Ctrl+R
                        'F' { ($pnl.Controls | Where-Object { $_.Name -eq 'btnFullRefresh' } | Select-Object -First 1)?.PerformClick() }  # Ctrl+F
                        'A' { ($pnl.Controls | Where-Object { $_.Name -eq 'btnActivate' }    | Select-Object -First 1)?.PerformClick() }  # Ctrl+A
                        'D' { ($pnl.Controls | Where-Object { $_.Name -eq 'btnDeactivate' }  | Select-Object -First 1)?.PerformClick() }  # Ctrl+D
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
            $null = Update-PIMRolesList -Form $form -RefreshEligible -RefreshActive -SplashForm $SplashForm -DisableParallelProcessing:$DisableParallelProcessing -ThrottleLimit $ThrottleLimit
            
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
