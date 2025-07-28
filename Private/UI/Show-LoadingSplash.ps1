function Show-LoadingSplash {
    <#
    .SYNOPSIS
        Displays a responsive loading splash screen with progress tracking.
    
    .DESCRIPTION
        Creates a non-blocking loading splash screen that runs in a separate runspace.
        Provides smooth progress animation and status updates without freezing the UI.
    
    .PARAMETER Message
        Initial status message to display. Default: "Initializing..."
    
    .PARAMETER Title
        Window title text. Default: "PIM Activation"
    
    .OUTPUTS
        PSCustomObject with UpdateStatus() and Close() methods for controlling the splash screen.
    
    .EXAMPLE
        $splash = Show-LoadingSplash -Message "Loading configuration..."
        $splash.UpdateStatus("Processing items...", 50)
        $splash.Close()
    #>
    [CmdletBinding()]
    param(
        [string]$Message = "Initializing...",
        [string]$Title = "PIM Activation"
    )

    # Synchronized hashtable for cross-runspace communication
    $syncHash = [hashtable]::Synchronized(@{
        Message = $Message
        Progress = 0
        TargetProgress = 0
        ShouldClose = $false
        Form = $null
        StatusLabel = $null
        ProgressBar = $null
        IsDisposed = $false
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
            Text = "PIM Activation"
            Size = [System.Drawing.Size]::new(400, 150)
            StartPosition = 'CenterScreen'
            FormBorderStyle = 'FixedDialog'
            BackColor = [System.Drawing.Color]::White
            TopMost = $true
            ShowInTaskbar = $false
            MaximizeBox = $false
            MinimizeBox = $false
        }

        # Status label
        $statusLabel = New-Object System.Windows.Forms.Label -Property @{
            Text = $syncHash.Message
            Font = [System.Drawing.Font]::new("Segoe UI", 10)
            Location = [System.Drawing.Point]::new(10, 20)
            Size = [System.Drawing.Size]::new(380, 30)
            TextAlign = 'MiddleCenter'
            ForeColor = [System.Drawing.Color]::FromArgb(32, 31, 30)
            Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        }
        $form.Controls.Add($statusLabel)

        # Progress bar
        $progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{
            Location = [System.Drawing.Point]::new(20, 60)
            Size = [System.Drawing.Size]::new(340, 30)
            Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            Minimum = 0
            Maximum = 100
            Value = 0
            ForeColor = [System.Drawing.Color]::FromArgb(0, 103, 184)
            BackColor = [System.Drawing.Color]::FromArgb(242, 242, 242)
            Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        }
        $form.Controls.Add($progressBar)
        
        # Store UI references
        $syncHash.Form = $form
        $syncHash.StatusLabel = $statusLabel
        $syncHash.ProgressBar = $progressBar
        
        $form.Add_FormClosed({ $syncHash.IsDisposed = $true })
        
        # Update timer for smooth animations
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 20
        $timer.Add_Tick({
            # Update status text
            if ($syncHash.StatusLabel.Text -ne $syncHash.Message) {
                $syncHash.StatusLabel.Text = $syncHash.Message
            }
            
            # Animate progress bar
            $currentValue = $syncHash.ProgressBar.Value
            $targetValue = [Math]::Min($syncHash.TargetProgress, 100)
            
            if ($currentValue -ne $targetValue) {
                $diff = $targetValue - $currentValue
                $step = [Math]::Max(1, [Math]::Abs($diff) / 10)
                $newValue = if ($diff -gt 0) { 
                    [Math]::Min($currentValue + $step, $targetValue) 
                } else { 
                    [Math]::Max($currentValue - $step, $targetValue) 
                }
                
                $syncHash.ProgressBar.Value = [int]$newValue
                $syncHash.Progress = [int]$newValue
            }
            
            if ($syncHash.ShouldClose) {
                $timer.Stop()
                $timer.Dispose()
                $form.Hide()
                $form.Close()
                $form.Dispose()
                $syncHash.IsDisposed = $true
                [System.Windows.Forms.Application]::ExitThread()
                return
            }
        })
        $timer.Start()
        
        # Show the form and start message pump
        [void]$form.Show()
        $form.Activate()
        $form.BringToFront()
        $form.TopMost = $true
        
        # Use DoEvents loop instead of Application.Run to avoid blocking
        while (-not $syncHash.ShouldClose -and -not $form.IsDisposed) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 50
        }
        
        # Cleanup when closing
        if (-not $form.IsDisposed) {
            $form.Hide()
            $form.Close()
            $form.Dispose()
        }
        $syncHash.IsDisposed = $true
    })
    
    # Start splash screen
    $handle = $powershell.BeginInvoke()
    
    # Wait a moment for the form to be created
    $maxWait = 50 # 5 seconds maximum
    $waitCount = 0
    while (-not $syncHash.Form -and $waitCount -lt $maxWait) {
        Start-Sleep -Milliseconds 100
        $waitCount++
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
        $this.SyncHash.Message = $Status
        if ($Progress -ge 0) {
            $this.SyncHash.TargetProgress = [Math]::Min($Progress, 100)
        }
    }
    
    $splashControl | Add-Member -MemberType ScriptMethod -Name Close -Value {
        if (-not $this.IsDisposed) {
            $this.SyncHash.TargetProgress = 100
            Start-Sleep -Milliseconds 200

            $this.SyncHash.ShouldClose = $true
            Start-Sleep -Milliseconds 100

            # Wait for the runspace to finish cleanup
            $maxWait = 20  # 2 seconds max
            $waitCount = 0
            while (-not $this.SyncHash.IsDisposed -and $waitCount -lt $maxWait) {
                Start-Sleep -Milliseconds 100
                $waitCount++
            }

            # Force cleanup if needed
            if ($this.PowerShell) {
                try { $this.PowerShell.Stop() } catch {}
                try { $this.PowerShell.Dispose() } catch {}
            }
            if ($this.Runspace) {
                try { $this.Runspace.Close() } catch {}
                try { $this.Runspace.Dispose() } catch {}
            }
        }
    }
    
    Start-Sleep -Milliseconds 100
    return $splashControl
}

function Update-LoadingStatus {
    <#
    .SYNOPSIS
        Updates the splash screen status and progress.
    
    .PARAMETER SplashForm
        The splash screen control object returned by Show-LoadingSplash.
    
    .PARAMETER Status
        New status message to display.
    
    .PARAMETER Progress
        Progress percentage (0-100). Optional.
    
    .EXAMPLE
        Update-LoadingStatus -SplashForm $splash -Status "Processing..." -Progress 75
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$SplashForm,
        
        [Parameter(Mandatory)]
        [string]$Status,

        [int]$Progress = -1
    )

    if ($SplashForm -and $SplashForm.SyncHash -and -not $SplashForm.IsDisposed) {
        $SplashForm.UpdateStatus($Status, $Progress)
    }
}

function Close-LoadingSplash {
    <#
    .SYNOPSIS
        Closes the loading splash screen and cleans up resources.
    
    .PARAMETER SplashForm
        The splash screen control object returned by Show-LoadingSplash.
    
    .EXAMPLE
        Close-LoadingSplash -SplashForm $splash
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$SplashForm
    )

    if ($SplashForm -and -not $SplashForm.IsDisposed) {
        $SplashForm.Close()
    }
}