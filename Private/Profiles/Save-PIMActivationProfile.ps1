function Save-PIMActivationProfile {
    <#
    .SYNOPSIS
        Saves a PIM activation profile for future bulk activation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProfileName,

        [Parameter(Mandatory)]
        [object[]]$SelectedRoles,

        [hashtable]$DefaultDuration,

        [string]$DefaultJustification = 'PowerShell activation'
    )

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        throw 'ProfileName is required.'
    }
    if (-not $SelectedRoles -or $SelectedRoles.Count -eq 0) {
        throw 'At least one role is required to save an activation profile.'
    }

    $profilePath = Get-PIMActivationProfileStorePath
    if (-not (Test-Path -Path $profilePath)) {
        New-Item -Path $profilePath -ItemType Directory -Force | Out-Null
    }

    $safeName = ($ProfileName.Trim() -replace '[\\/:*?"<>|]', '_')
    $filePath = Join-Path $profilePath "$safeName.json"
    $existingProfile = $null
    if (Test-Path -Path $filePath) {
        try {
            $existingProfile = Get-Content -Path $filePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            Write-Verbose "Existing profile '$ProfileName' could not be read and will be overwritten: $($_.Exception.Message)"
        }
    }

    $roleEntries = foreach ($role in $SelectedRoles) {
        $isListViewItem = $role -and $role.GetType().FullName -eq 'System.Windows.Forms.ListViewItem'
        $roleData = if ($isListViewItem -and $role.PSObject.Properties['Tag']) { $role.Tag } else { $role }
        if (-not $roleData) { continue }

        $displayName = if ($roleData -is [string]) {
            $roleData
        }
        elseif ($roleData.PSObject.Properties['DisplayName'] -and $roleData.DisplayName) {
            $roleData.DisplayName
        }
        elseif ($roleData.PSObject.Properties['RoleName'] -and $roleData.RoleName) {
            $roleData.RoleName
        }
        elseif ($role.PSObject.Properties['Text'] -and $role.Text) {
            $role.Text
        }
        else {
            'Unknown Role'
        }

        [PSCustomObject]@{
            Key               = Get-PIMActivationProfileRoleKey -RoleData $roleData
            DisplayName       = $displayName
            Type              = if ($roleData.PSObject.Properties['Type']) { $roleData.Type } else { $null }
            RoleDefinitionId  = if ($roleData.PSObject.Properties['RoleDefinitionId']) { $roleData.RoleDefinitionId } else { $null }
            DirectoryScopeId  = if ($roleData.PSObject.Properties['DirectoryScopeId']) { $roleData.DirectoryScopeId } else { $null }
            FullScope         = if ($roleData.PSObject.Properties['FullScope']) { $roleData.FullScope } else { $null }
            Scope             = if ($roleData.PSObject.Properties['Scope']) { $roleData.Scope } else { $null }
            GroupId           = if ($roleData.PSObject.Properties['GroupId']) { $roleData.GroupId } else { $null }
            AccessId          = if ($roleData.PSObject.Properties['AccessId'] -and $roleData.AccessId) {
                $roleData.AccessId
            }
            elseif ($roleData.PSObject.Properties['Assignment'] -and $roleData.Assignment -and $roleData.Assignment.PSObject.Properties['AccessId'] -and $roleData.Assignment.AccessId) {
                $roleData.Assignment.AccessId
            }
            elseif ($roleData.PSObject.Properties['MemberType'] -and $roleData.MemberType) {
                $roleData.MemberType
            }
            else {
                $null
            }
            ScopeDisplayName  = if ($roleData.PSObject.Properties['ScopeDisplayName']) { $roleData.ScopeDisplayName } else { $null }
        }
    }

    $createdAt = if ($existingProfile -and $existingProfile.PSObject.Properties['CreatedAt']) { $existingProfile.CreatedAt } else { (Get-Date).ToString('o') }
    $profile = [PSCustomObject]@{
        Name                 = $ProfileName.Trim()
        CreatedAt            = $createdAt
        UpdatedAt            = (Get-Date).ToString('o')
        DefaultDuration      = $DefaultDuration
        DefaultJustification = $DefaultJustification
        Roles                = @($roleEntries)
    }

    $profile | ConvertTo-Json -Depth 8 | Set-Content -Path $filePath -Encoding UTF8
    $profile | Add-Member -MemberType NoteProperty -Name FilePath -Value $filePath -Force
    return $profile
}