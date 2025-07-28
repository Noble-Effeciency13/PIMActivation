function Show-OperationSplash {
    <#
    .SYNOPSIS
        Displays a responsive loading splash screen for PIM operations like activation and refresh.
    
    .DESCRIPTION
        Creates a non-blocking loading splash screen that runs in a separate runspace.
        Provides smooth progress animation and status updates for long-running operations.
        Designed for use during role activation and refresh operations.
    
    .PARAMETER Title
        Window title text. Default: "PIM Operation"
    
    .PARAMETER InitialMessage
        Initial status message to display. Default: "Processing..."
    
    .PARAMETER ShowProgressBar
        Whether to show a progress bar. Default: $true
    
    .PARAMETER Width
        Width of the splash window. Default: 450
    
    .PARAMETER Height
        Height of the splash window. Default: 180
    
    .OUTPUTS
        PSCustomObject with UpdateStatus() and Close() methods for controlling the splash screen.
    
    .EXAMPLE
        $splash = Show-OperationSplash -Title "Role Activation" -InitialMessage "Preparing role activation..."
        $splash.UpdateStatus("Activating Global Administrator...", 50)
        $splash.Close()
    
    .EXAMPLE
        $splash = Show-OperationSplash -Title "Refreshing Roles" -InitialMessage "Fetching role data..."
        # Do work...
        $splash.Close()
    #>
    [CmdletBinding()]
    param(
        [string]$Title = "PIM Operation",
        [string]$InitialMessage = "Processing...",
        [bool]$ShowProgressBar = $true,
        [int]$Width = 450,
        [int]$Height = 180
    )

    # Synchronized hashtable for cross-runspace communication
    $syncHash = [hashtable]::Synchronized(@{
        Title = $Title
        Message = $InitialMessage
        Progress = 0
        TargetProgress = 0
        ShouldClose = $false
        Form = $null
        StatusLabel = $null
        ProgressBar = $null
        IsDisposed = $false
        ShowProgressBar = $ShowProgressBar
        Width = $Width
        Height = $Height
    })

    # Create STA runspace for the UI
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
    
    # Create PowerShell instance
    $powershell = [powershell]::Create()
    $powershell.Runspace = $runspace
    
    # UI creation script
    [void]$powershell.AddScript({
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        
        # Main form
        $form = New-Object System.Windows.Forms.Form -Property @{
            Text = $syncHash.Title
            Size = [System.Drawing.Size]::new($syncHash.Width, $syncHash.Height)
            StartPosition = 'CenterScreen'
            FormBorderStyle = 'FixedDialog'
            BackColor = [System.Drawing.Color]::White
            TopMost = $true
            ShowInTaskbar = $true
            MaximizeBox = $false
            MinimizeBox = $false
            ControlBox = $false  # Hide close button during operations
        }

        # Add icon (optional - uses default if not found)
        try {
            $iconPath = Join-Path $PSScriptRoot "Resources\pim-icon.ico"
            if (Test-Path $iconPath) {
                $form.Icon = [System.Drawing.Icon]::new($iconPath)
            }
        } catch {}

        # Header panel with color
        $headerPanel = New-Object System.Windows.Forms.Panel -Property @{
            Location = [System.Drawing.Point]::new(0, 0)
            Size = [System.Drawing.Size]::new($syncHash.Width, 40)
            BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        }
        $form.Controls.Add($headerPanel)

        # Title label in header
        $titleLabel = New-Object System.Windows.Forms.Label -Property @{
            Text = $syncHash.Title
            Font = [System.Drawing.Font]::new("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
            ForeColor = [System.Drawing.Color]::White
            Location = [System.Drawing.Point]::new(15, 10)
            Size = [System.Drawing.Size]::new($syncHash.Width - 30, 25)
            BackColor = [System.Drawing.Color]::Transparent
        }
        $headerPanel.Controls.Add($titleLabel)

        # Status label
        $statusLabel = New-Object System.Windows.Forms.Label -Property @{
            Text = $syncHash.Message
            Font = [System.Drawing.Font]::new("Segoe UI", 10)
            Location = [System.Drawing.Point]::new(15, 55)
            Size = [System.Drawing.Size]::new($syncHash.Width - 30, 40)
            TextAlign = 'MiddleCenter'
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
            Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            AutoEllipsis = $true
        }
        $form.Controls.Add($statusLabel)

        # Progress bar (optional)
        if ($syncHash.ShowProgressBar) {
            $progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{
                Location = [System.Drawing.Point]::new(20, 105)
                Size = [System.Drawing.Size]::new($syncHash.Width - 40, 20)
                Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                Minimum = 0
                Maximum = 100
                Value = 0
                ForeColor = [System.Drawing.Color]::FromArgb(0, 103, 184)
                BackColor = [System.Drawing.Color]::FromArgb(242, 242, 242)
                Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            }
            $form.Controls.Add($progressBar)
            $syncHash.ProgressBar = $progressBar
        }

        # Spinner animation (when no progress bar)
        if (-not $syncHash.ShowProgressBar) {
            $spinnerLabel = New-Object System.Windows.Forms.Label -Property @{
                Text = "⚪⚪⚪"
                Font = [System.Drawing.Font]::new("Segoe UI", 14)
                Location = [System.Drawing.Point]::new(($syncHash.Width / 2) - 30, 100)
                Size = [System.Drawing.Size]::new(60, 30)
                TextAlign = 'MiddleCenter'
                ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            }
            $form.Controls.Add($spinnerLabel)
            
            # Spinner animation timer
            $spinnerStates = @("⚫⚪⚪", "⚪⚫⚪", "⚪⚪⚫", "⚪⚫⚪")
            $spinnerIndex = 0
            $spinnerTimer = New-Object System.Windows.Forms.Timer
            $spinnerTimer.Interval = 200
            $spinnerTimer.Add_Tick({
                $spinnerLabel.Text = $spinnerStates[$spinnerIndex]
                $spinnerIndex = ($spinnerIndex + 1) % $spinnerStates.Count
            })
            $spinnerTimer.Start()
        }
        
        # Store UI references
        $syncHash.Form = $form
        $syncHash.StatusLabel = $statusLabel
        
        $form.Add_FormClosed({ 
            $syncHash.IsDisposed = $true
            if ($spinnerTimer) { 
                $spinnerTimer.Stop()
                $spinnerTimer.Dispose()
            }
        })
        
        # Update timer for smooth animations
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 20
        $timer.Add_Tick({
            # Update status text
            if ($syncHash.StatusLabel.Text -ne $syncHash.Message) {
                $syncHash.StatusLabel.Text = $syncHash.Message
            }
            
            # Animate progress bar if exists
            if ($syncHash.ProgressBar) {
                $currentValue = $syncHash.ProgressBar.Value
                $targetValue = [Math]::Min($syncHash.TargetProgress, 100)
                
                if ($currentValue -ne $targetValue) {
                    # Smooth animation - move 20% of the way to target each frame
                    $diff = $targetValue - $currentValue
                    $step = [Math]::Max(1, [Math]::Abs($diff * 0.2))
                    
                    if ($diff -gt 0) {
                        $syncHash.ProgressBar.Value = [Math]::Min($currentValue + $step, $targetValue)
                    } elseif ($diff -lt 0) {
                        $syncHash.ProgressBar.Value = [Math]::Max($currentValue - $step, $targetValue)
                    }
                }
            }
            
            if ($syncHash.ShouldClose) {
                $timer.Stop()
                $form.Close()
            }
        })
        $timer.Start()
        
        [void]$form.ShowDialog()
        
        # Cleanup
        $timer.Stop()
        $timer.Dispose()
        $syncHash.IsDisposed = $true
    })
    
    # Start splash screen
    $handle = $powershell.BeginInvoke()
    
    # Wait for form to be created
    $maxWait = 50  # 5 seconds max
    $waited = 0
    while (-not $syncHash.Form -and $waited -lt $maxWait) {
        Start-Sleep -Milliseconds 100
        $waited++
    }
    
    # Control object
    $splashControl = [PSCustomObject]@{
        SyncHash = $syncHash
        PowerShell = $powershell
        Runspace = $runspace
        Handle = $handle
    }
    
    $splashControl | Add-Member -MemberType ScriptProperty -Name IsDisposed -Value {
        $this.SyncHash.IsDisposed
    }
    
    $splashControl | Add-Member -MemberType ScriptMethod -Name UpdateStatus -Value {
        param([string]$Status, [int]$Progress = -1)
        if (-not $this.IsDisposed) {
            $this.SyncHash.Message = $Status
            if ($Progress -ge 0 -and $this.SyncHash.ShowProgressBar) {
                $this.SyncHash.TargetProgress = [Math]::Min($Progress, 100)
            }
        }
    }
    
    $splashControl | Add-Member -MemberType ScriptMethod -Name Close -Value {
        if (-not $this.IsDisposed) {
            if ($this.SyncHash.ShowProgressBar) {
                $this.SyncHash.TargetProgress = 100
                Start-Sleep -Milliseconds 300
            }
            
            $this.SyncHash.ShouldClose = $true
            Start-Sleep -Milliseconds 200
            
            if ($this.PowerShell) {
                $this.PowerShell.Stop()
                $this.PowerShell.Dispose()
            }
            if ($this.Runspace) {
                $this.Runspace.Close()
                $this.Runspace.Dispose()
            }
        }
    }
    
    return $splashControl
}