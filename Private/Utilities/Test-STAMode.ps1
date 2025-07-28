function Test-STAMode {
    <#
    .SYNOPSIS
        Tests if PowerShell is running in Single Thread Apartment (STA) mode.
    
    .DESCRIPTION
        Checks if the current PowerShell session is running in STA mode, which is required
        for Windows Forms and some COM objects to function properly.
    
    .EXAMPLE
        Test-STAMode
        Returns $true if running in STA mode, $false if running in MTA mode.
    
    .OUTPUTS
        System.Boolean
        Returns $true if in STA mode, $false otherwise.
    
    .NOTES
        STA (Single Thread Apartment) mode is required for Windows Forms applications.
        PowerShell ISE runs in STA mode by default, while PowerShell console runs in MTA mode.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    [System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
}