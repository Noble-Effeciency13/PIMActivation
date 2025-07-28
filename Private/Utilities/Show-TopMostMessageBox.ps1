function Show-TopMostMessageBox {
    <#
    .SYNOPSIS
        Displays a MessageBox that appears on top of all other windows, including operation splash screens.
    
    .DESCRIPTION
        Creates a temporary TopMost form to ensure MessageBox dialogs appear in front of any existing
        operation splash screens or other TopMost windows. The temporary form is automatically cleaned up.
    
    .PARAMETER Message
        The message text to display in the MessageBox.
    
    .PARAMETER Title
        The title of the MessageBox window.
    
    .PARAMETER Buttons
        The buttons to display. Default is OK.
    
    .PARAMETER Icon
        The icon to display. Default is Information.
    
    .EXAMPLE
        Show-TopMostMessageBox -Message "Operation completed successfully!" -Title "Success"
        
    .EXAMPLE
        $result = Show-TopMostMessageBox -Message "Are you sure?" -Title "Confirm" -Buttons YesNo -Icon Question
        if ($result -eq 'Yes') { ... }
    
    .OUTPUTS
        System.Windows.Forms.DialogResult
        Returns the user's response to the MessageBox.
    
    .NOTES
        This function ensures that MessageBox dialogs appear in front of operation splash screens
        and other TopMost windows by creating a temporary TopMost parent form.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter(Mandatory)]
        [string]$Title,
        
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    # Create a temporary form to ensure MessageBox appears on top
    $tempForm = New-Object System.Windows.Forms.Form -Property @{
        TopMost = $true
        WindowState = 'Minimized'
        ShowInTaskbar = $false
        Size = [System.Drawing.Size]::new(0, 0)
        Location = [System.Drawing.Point]::new(-1000, -1000)
    }
    
    try {
        $tempForm.Show()
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            $tempForm,
            $Message,
            $Title,
            $Buttons,
            $Icon
        )
        
        return $result
    }
    finally {
        $tempForm.Close()
        $tempForm.Dispose()
    }
}
