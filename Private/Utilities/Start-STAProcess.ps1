function Start-STAProcess {
    <#
    .SYNOPSIS
    Starts a PowerShell 7 process in STA (Single Threaded Apartment) mode to execute PIM activation.

    .DESCRIPTION
    This function creates a new PowerShell 7 process configured for STA mode to run PIM activation
    commands. It converts hashtable parameters to command-line arguments and launches the process
    with appropriate security and execution settings.

    .PARAMETER Parameters
    A hashtable containing the parameters to pass to the Start-PIMActivation command.
    Switch parameters are handled automatically based on their IsPresent property.

    .EXAMPLE
    Start-STAProcess -Parameters @{ RoleName = "Global Administrator"; Justification = "Emergency access" }

    .EXAMPLE
    Start-STAProcess -Parameters @{ RoleName = "User Administrator"; Duration = 2; Force = $true }

    .NOTES
    Requires PowerShell 7+ (pwsh.exe) to be installed and available in PATH.
    The process runs with -ExecutionPolicy Bypass and -NoProfile for compatibility.

    .LINK
    https://aka.ms/powershell
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Hashtable of parameters to pass to Start-PIMActivation")]
        [ValidateNotNull()]
        [hashtable]$Parameters
    )
    
    begin {
        Write-Verbose "Starting STA process initialization"
    }
    
    process {
        Write-Verbose "Processing $($Parameters.Count) parameter(s)"

        # Verify PowerShell 7 availability
        $psExecutable = "pwsh.exe"
        Write-Verbose "Checking for PowerShell 7 executable: $psExecutable"

        $psPath = Get-Command $psExecutable -ErrorAction SilentlyContinue
        if (-not $psPath) {
            $errorMessage = "PowerShell 7 (pwsh.exe) not found. Please install PowerShell 7+ from https://aka.ms/powershell"
            Write-Error $errorMessage
            throw $errorMessage
        }

        Write-Verbose "Found PowerShell 7 at: $($psPath.Source)"

        # Serialize parameters via PSSerializer so no user-supplied string is
        # ever interpolated into the child shell's command line. The child
        # decodes the base64 blob and splats the resulting hashtable.
        $serialized = [System.Management.Automation.PSSerializer]::Serialize($Parameters)
        $paramsB64  = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($serialized))

        $childScript = @"
`$xml    = [System.Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$paramsB64'))
`$params = [System.Management.Automation.PSSerializer]::Deserialize(`$xml)
Import-Module PIMActivation -Force
Start-PIMActivation @params
"@

        $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($childScript))

        # Launch the STA process
        try {
            Write-Verbose "Launching PowerShell 7 process in STA mode"
            Start-Process $psPath.Source -ArgumentList @(
                '-Sta',
                '-ExecutionPolicy', 'Bypass',
                '-NoProfile',
                '-WindowStyle', 'Hidden',
                '-EncodedCommand', $encoded
            ) -NoNewWindow

            Write-Verbose "Successfully launched STA process"
        }
        catch {
            Write-Error "Failed to start PowerShell 7 process: $($_.Exception.Message)"
            throw
        }
    }
}