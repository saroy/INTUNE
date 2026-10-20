Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

function Show-Status([string]$message) {
    $lblStatus.Text = $message
    $lblStatus.Refresh()
}

function Add-Log([string]$message) {
    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $txtLogs.AppendText("[$timestamp] $message`r`n")
    $txtLogs.ScrollToCaret()
}

function Export-ResultsToCSV {
    if (-not $script:lastResults) {
        [System.Windows.Forms.MessageBox]::Show('No import results to export. Run an import first.', 'No Data', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = 'CSV files (*.csv)|*.csv'
    $saveDialog.DefaultExt = '.csv'
    $saveDialog.FileName = "import_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveDialog.ShowDialog() -eq 'OK') {
        try {
            $csv = @()
            $csv += 'Device,Status'
            
            foreach ($device in $script:lastResults.Success) {
                $csv += "$device,added"
            }
            foreach ($device in $script:lastResults.AlreadyExists) {
                $csv += "$device,already exist"
            }
            foreach ($device in $script:lastResults.Failed) {
                $csv += "$device,failed"
            }
            
            $csv | Set-Content -Path $saveDialog.FileName -Encoding UTF8
            Add-Log "✓ Results exported to: $($saveDialog.FileName)"
            [System.Windows.Forms.MessageBox]::Show("Results exported successfully to:`n$($saveDialog.FileName)", 'Export Complete', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            Add-Log "✗ Export failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
}

function Test-GraphModule {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Host 'Microsoft.Graph module not found. Installing...'
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber | Out-Null
    }
    
    # Clean up any previously loaded Microsoft.Graph modules
    Get-Module -Name 'Microsoft.Graph*' | Remove-Module -Force -ErrorAction SilentlyContinue
    
    # Import modules with AllowClobber to handle assembly conflicts
    $modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.DeviceManagement')
    foreach ($module in $modules) {
        try {
            Import-Module $module -AllowClobber -SkipEditionCheck -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Debug "Failed to import $module : $($_.Exception.Message)"
        }
    }
}

function Ensure-GraphModule {
    Test-GraphModule
}

function Invoke-GraphQuerySafe {
    param(
        [scriptblock]$ScriptBlock
    )
    
    try {
        & $ScriptBlock
    }
    catch {
        # If still getting assembly warning, suppress and retry
        if ($_.Exception.Message -like '*Assembly with same name is already loaded*') {
            $WarningPreference = 'SilentlyContinue'
            $result = & $ScriptBlock 2>$null
            $WarningPreference = 'Continue'
            return $result
        }
        throw $_
    }
}

function Connect-ToGraph {
    try {
        if (-not (Get-MgContext)) {
            Show-Status 'Connecting to Microsoft Graph...'
            Connect-MgGraph -Scopes 'Device.ReadWrite.All','Group.ReadWrite.All','Directory.Read.All' -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Debug "Graph connection warning: $($_.Exception.Message)"
    }
}

function Find-GroupByName([string]$groupName) {
    $escaped = $groupName.Replace("'", "''")
    Invoke-GraphQuerySafe { Get-MgGroup -Filter "displayName eq '$escaped'" -ConsistencyLevel eventual -ErrorAction Stop -WarningAction SilentlyContinue } | Select-Object -First 1
}

function Find-DeviceByName([string]$deviceName) {
    $escaped = $deviceName.Replace("'", "''")
    Invoke-GraphQuerySafe { Get-MgDevice -Filter "displayName eq '$escaped' and accountEnabled eq true" -ErrorAction Stop -WarningAction SilentlyContinue } | Select-Object -First 1
}

function Add-DeviceToGroup([string]$groupId, [string]$deviceId) {
    New-MgGroupMember -GroupId $groupId -DirectoryObjectId $deviceId -ErrorAction Stop | Out-Null
}

function Remove-DeviceFromGroup([string]$groupId, [string]$deviceId) {
    Remove-MgGroupMemberByRef -GroupId $groupId -DirectoryObjectId $deviceId -ErrorAction Stop
}

function Search-AndSelectGroup {
    try {
        $groupName = $txtGroup.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($groupName)) {
            [System.Windows.Forms.MessageBox]::Show('Enter a group name to search.', 'Missing Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        Add-Log "Searching for group '$groupName'..."
        Show-Status "Searching for group..."
        Ensure-GraphModule
        Connect-ToGraph

        $escapedName = $groupName.Replace("'", "''")
        $groups = Invoke-GraphQuerySafe { Get-MgGroup -Filter "startswith(displayName, '$escapedName')" -ConsistencyLevel eventual -ErrorAction Stop -WarningAction SilentlyContinue }
        
        if ($null -eq $groups -or $groups.Count -eq 0) {
            Add-Log "✗ No groups found matching '$groupName'"
            [System.Windows.Forms.MessageBox]::Show("No groups found matching '$groupName'.", 'Not Found', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            $lblGroupStatus.Text = '✗ No group selected'
            $lblGroupStatus.ForeColor = [System.Drawing.Color]::Red
            $btnImport.Enabled = $false
            $script:selectedGroup = $null
            Show-Status 'Ready.'
            return
        }

        # Show selection dialog (for 1 or more groups)
        if ($groups.Count -eq 1) {
            Add-Log "✓ Found 1 group: $($groups[0].DisplayName)"
        } else {
            Add-Log "✓ Found $($groups.Count) groups matching '$groupName'"
        }

        $formSelect = New-Object System.Windows.Forms.Form
        $formSelect.Text = 'Select Group'
        $formSelect.Size = New-Object System.Drawing.Size(600, 350)
        $formSelect.StartPosition = 'CenterParent'
        $formSelect.FormBorderStyle = 'FixedDialog'
        $formSelect.MaximizeBox = $false
        $formSelect.MinimizeBox = $false

        $lblSelect = New-Object System.Windows.Forms.Label
        $lblSelect.Location = New-Object System.Drawing.Point(10, 10)
        $lblSelect.Size = New-Object System.Drawing.Size(580, 30)
        if ($groups.Count -eq 1) {
            $lblSelect.Text = "Confirm group selection:"
        } else {
            $lblSelect.Text = "Multiple groups found. Select one:"
        }
        $formSelect.Controls.Add($lblSelect)

        $listGroups = New-Object System.Windows.Forms.ListBox
        $listGroups.Location = New-Object System.Drawing.Point(10, 45)
        $listGroups.Size = New-Object System.Drawing.Size(580, 200)
        $listGroups.HorizontalScrollbar = $true
        $listGroups.SelectionMode = 'One'
        $groupsList = @()
        foreach ($g in $groups) {
            [void]$listGroups.Items.Add("$($g.DisplayName) (ID: $($g.Id.Substring(0, 8))...)")
            $groupsList += $g
        }
        $formSelect.Controls.Add($listGroups)

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Location = New-Object System.Drawing.Point(220, 260)
        $btnOK.Size = New-Object System.Drawing.Size(100, 32)
        $btnOK.Text = 'Select'
        $btnOK.DialogResult = 'OK'
        $formSelect.Controls.Add($btnOK)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(330, 260)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
        $btnCancel.Text = 'Cancel'
        $btnCancel.DialogResult = 'Cancel'
        $formSelect.Controls.Add($btnCancel)

        if ($formSelect.ShowDialog() -eq 'OK' -and $listGroups.SelectedIndex -ge 0) {
            $script:selectedGroup = $groupsList[$listGroups.SelectedIndex]
            Add-Log "✓ Selected group: $($script:selectedGroup.DisplayName)"
            $lblGroupStatus.Text = "✓ $($script:selectedGroup.DisplayName)"
            $lblGroupStatus.ForeColor = [System.Drawing.Color]::Green
            $btnImport.Enabled = $true
            $btnRemove.Enabled = $true
            Show-Status 'Group validated. Ready to import or remove.'
        } else {
            Add-Log 'Group selection cancelled'
            $lblGroupStatus.Text = '✗ No group selected'
            $lblGroupStatus.ForeColor = [System.Drawing.Color]::Red
            $btnImport.Enabled = $false
            $btnRemove.Enabled = $false
            $script:selectedGroup = $null
            Show-Status 'Ready.'
        }
    }
    catch {
        Add-Log "Error searching for group: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        Show-Status 'Error occurred.'
    }
}

function Remove-SelectedDevices {
    try {
        if (-not $script:selectedGroup) {
            Add-Log 'Error: No group selected'
            [System.Windows.Forms.MessageBox]::Show('Please search and select a group first.', 'No Group Selected', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $rawText = $txtDevices.Text
        Add-Log "Debug: Raw input length: $($rawText.Length) characters"
        
        # Split on line breaks and other delimiters, force as array
        [array]$devices = $rawText -split '[\r\n,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 0 }
        
        Add-Log "Debug: Found $($devices.Count) device(s)"
        $devices | ForEach-Object { Add-Log "Debug: Device entry: '$_'" }
        
        if ($devices.Count -eq 0) {
            Add-Log 'Error: No devices provided'
            [System.Windows.Forms.MessageBox]::Show('Paste at least one device name into the list.', 'No Devices', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        Add-Log "Starting removal process..."
        Add-Log "Processing $($devices.Count) device(s) to remove from group: $($script:selectedGroup.DisplayName)..."
        Show-Status "Processing devices..."
        Ensure-GraphModule
        Connect-ToGraph

        $success = @()
        $failed = @()
        $notFound = @()

        $progressBar.Maximum = $devices.Count
        $progressBar.Value = 0

        for ($i = 0; $i -lt $devices.Count; $i++) {
            $deviceName = $devices[$i]
            Show-Status "Processing '$deviceName'..."
            $lblProgress.Text = "$($i + 1) / $($devices.Count)"
            $progressBar.Value = $i
            $form.Refresh()
            
            try {
                $device = Find-DeviceByName -deviceName $deviceName
                if (-not $device) {
                    Add-Log "✗ Device '$deviceName' not found"
                    $notFound += $deviceName
                    continue
                }
                Remove-DeviceFromGroup -groupId $script:selectedGroup.Id -deviceId $device.Id
                Add-Log "✓ Device '$deviceName' removed successfully"
                $success += $deviceName
            }
            catch {
                $errorMsg = $_.Exception.Message
                Add-Log "✗ Device '$deviceName' failed: $errorMsg"
                $failed += $deviceName
            }
        }

        $progressBar.Value = $devices.Count
        $lblProgress.Text = "$($devices.Count) / $($devices.Count)"

        $message = "Removal complete. Removed: $($success.Count). Not Found: $($notFound.Count). Failed: $($failed.Count)."
        Add-Log $message
        Show-Status($message)

        $resultDetails = @()
        if ($success.Count -gt 0) {
            $resultDetails += "Successfully Removed ($($success.Count)):"
            $resultDetails += $success | ForEach-Object { "  ✓ $_" }
        }
        if ($notFound.Count -gt 0) {
            $resultDetails += ""
            $resultDetails += "Not Found in Directory ($($notFound.Count)):"
            $resultDetails += $notFound | ForEach-Object { "  ✗ $_" }
        }
        if ($failed.Count -gt 0) {
            $resultDetails += ""
            $resultDetails += "Failed to Remove ($($failed.Count)):"
            $resultDetails += $failed | ForEach-Object { "  ✗ $_" }
        }

        $detailsText = $resultDetails -join "`n"
        
        # Store results for future reference
        $script:lastResults = @{
            Success = $success
            NotFound = $notFound
            Failed = $failed
            Timestamp = Get-Date
            GroupName = $script:selectedGroup.DisplayName
            Operation = 'Removal'
        }
        
        [System.Windows.Forms.MessageBox]::Show(($message + "`n`n" + $detailsText), 'Removal Results', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    catch {
        Add-Log "Error: $($_.Exception.Message)"
        Show-Status("Error: $($_.Exception.Message)")
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Import-SelectedDevices {
    try {
        if (-not $script:selectedGroup) {
            Add-Log 'Error: No group selected'
            [System.Windows.Forms.MessageBox]::Show('Please search and select a group first.', 'No Group Selected', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $rawText = $txtDevices.Text
        Add-Log "Debug: Raw input length: $($rawText.Length) characters"
        
        # Split on line breaks and other delimiters, force as array
        [array]$devices = $rawText -split '[\r\n,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 0 }
        
        Add-Log "Debug: Found $($devices.Count) device(s)"
        $devices | ForEach-Object { Add-Log "Debug: Device entry: '$_'" }
        
        if ($devices.Count -eq 0) {
            Add-Log 'Error: No devices provided'
            [System.Windows.Forms.MessageBox]::Show('Paste at least one device name into the list.', 'No Devices', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        Add-Log "Starting import process..."
        Add-Log "Processing $($devices.Count) device(s) to group: $($script:selectedGroup.DisplayName)..."
        Show-Status "Processing devices..."
        Ensure-GraphModule
        Connect-ToGraph

        $success = @()
        $failed = @()
        $alreadyExists = @()

        $progressBar.Maximum = $devices.Count
        $progressBar.Value = 0

        for ($i = 0; $i -lt $devices.Count; $i++) {
            $deviceName = $devices[$i]
            Show-Status "Processing '$deviceName'..."
            $lblProgress.Text = "$($i + 1) / $($devices.Count)"
            $progressBar.Value = $i
            $form.Refresh()
            
            try {
                $device = Find-DeviceByName -deviceName $deviceName
                if (-not $device) {
                    Add-Log "✗ Device '$deviceName' not found or is not enabled in Entra"
                    $failed += $deviceName
                    continue
                }
                Add-DeviceToGroup -groupId $script:selectedGroup.Id -deviceId $device.Id
                Add-Log "✓ Device '$deviceName' added successfully"
                $success += $deviceName
            }
            catch {
                $errorMsg = $_.Exception.Message
                if ($errorMsg -like "*already exist*" -or $errorMsg -like "*members*") {
                    Add-Log "⚠ Device '$deviceName' already exists in the group"
                    $alreadyExists += $deviceName
                } else {
                    Add-Log "✗ Device '$deviceName' failed: $errorMsg"
                    $failed += $deviceName
                }
            }
        }

        $progressBar.Value = $devices.Count
        $lblProgress.Text = "$($devices.Count) / $($devices.Count)"

        $message = "Import complete. Added: $($success.Count). Already Exists: $($alreadyExists.Count). Failed: $($failed.Count)."
        Add-Log $message
        Show-Status($message)

        $resultDetails = @()
        if ($success.Count -gt 0) {
            $resultDetails += "Successfully Added ($($success.Count)):"
            $resultDetails += $success | ForEach-Object { "  ✓ $_" }
        }
        if ($alreadyExists.Count -gt 0) {
            $resultDetails += ""
            $resultDetails += "Already Exists in Group ($($alreadyExists.Count)):"
            $resultDetails += $alreadyExists | ForEach-Object { "  ⚠ $_" }
        }
        if ($failed.Count -gt 0) {
            $resultDetails += ""
            $resultDetails += "Failed ($($failed.Count)):"
            $resultDetails += $failed | ForEach-Object { "  ✗ $_" }
        }

        $detailsText = $resultDetails -join "`n"
        
        # Store results for CSV export
        $script:lastResults = @{
            Success = $success
            AlreadyExists = $alreadyExists
            Failed = $failed
            Timestamp = Get-Date
            GroupName = $script:selectedGroup.DisplayName
        }
        
        [System.Windows.Forms.MessageBox]::Show(($message + "`n`n" + $detailsText), 'Import Results', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    catch {
        Add-Log "Error: $($_.Exception.Message)"
        Show-Status("Error: $($_.Exception.Message)")
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Entra Device Group Importer'
$form.Size = New-Object System.Drawing.Size(950,700)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true

$lblGroup = New-Object System.Windows.Forms.Label
$lblGroup.Location = New-Object System.Drawing.Point(12,18)
$lblGroup.Size = New-Object System.Drawing.Size(180,20)
$lblGroup.Text = 'Entra Group Display Name:'
$form.Controls.Add($lblGroup)

$txtGroup = New-Object System.Windows.Forms.TextBox
$txtGroup.Location = New-Object System.Drawing.Point(200,14)
$txtGroup.Size = New-Object System.Drawing.Size(280,24)
$form.Controls.Add($txtGroup)

$btnSearchGroup = New-Object System.Windows.Forms.Button
$btnSearchGroup.Location = New-Object System.Drawing.Point(490,14)
$btnSearchGroup.Size = New-Object System.Drawing.Size(100,24)
$btnSearchGroup.Text = 'Search'
$btnSearchGroup.Add_Click({ Search-AndSelectGroup })
$form.Controls.Add($btnSearchGroup)

$lblGroupStatus = New-Object System.Windows.Forms.Label
$lblGroupStatus.Location = New-Object System.Drawing.Point(600,14)
$lblGroupStatus.Size = New-Object System.Drawing.Size(330,24)
$lblGroupStatus.Text = '✗ No group selected'
$lblGroupStatus.ForeColor = [System.Drawing.Color]::Red
$lblGroupStatus.TextAlign = 'MiddleLeft'
$form.Controls.Add($lblGroupStatus)

$lblDevices = New-Object System.Windows.Forms.Label
$lblDevices.Location = New-Object System.Drawing.Point(12,58)
$lblDevices.Size = New-Object System.Drawing.Size(180,20)
$lblDevices.Text = 'Device names (one per line):'
$form.Controls.Add($lblDevices)

$txtDevices = New-Object System.Windows.Forms.TextBox
$txtDevices.Location = New-Object System.Drawing.Point(15,82)
$txtDevices.Size = New-Object System.Drawing.Size(900,190)
$txtDevices.Multiline = $true
$txtDevices.ScrollBars = 'Vertical'
$form.Controls.Add($txtDevices)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15,280)
$progressBar.Size = New-Object System.Drawing.Size(900,20)
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(15,305)
$lblProgress.Size = New-Object System.Drawing.Size(900,20)
$lblProgress.Text = '0 / 0'
$form.Controls.Add($lblProgress)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Location = New-Object System.Drawing.Point(15,330)
$btnImport.Size = New-Object System.Drawing.Size(130,32)
$btnImport.Text = 'Import to Group'
$btnImport.Enabled = $false
$btnImport.Add_Click({ Import-SelectedDevices })
$form.Controls.Add($btnImport)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(160,330)
$btnClear.Size = New-Object System.Drawing.Size(100,32)
$btnClear.Text = 'Clear List'
$btnClear.Add_Click({ $txtDevices.Clear(); $progressBar.Value = 0; $lblProgress.Text = '0 / 0'; Show-Status('Ready.') })
$form.Controls.Add($btnClear)

$btnExportCSV = New-Object System.Windows.Forms.Button
$btnExportCSV.Location = New-Object System.Drawing.Point(265,330)
$btnExportCSV.Size = New-Object System.Drawing.Size(120,32)
$btnExportCSV.Text = 'Export to CSV'
$btnExportCSV.Add_Click({ Export-ResultsToCSV })
$form.Controls.Add($btnExportCSV)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Location = New-Object System.Drawing.Point(400,330)
$btnRemove.Size = New-Object System.Drawing.Size(130,32)
$btnRemove.Text = 'Remove from Group'
$btnRemove.Enabled = $false
$btnRemove.Add_Click({ Remove-SelectedDevices })
$form.Controls.Add($btnRemove)

$lblLogs = New-Object System.Windows.Forms.Label
$lblLogs.Location = New-Object System.Drawing.Point(15,375)
$lblLogs.Size = New-Object System.Drawing.Size(50,20)
$lblLogs.Text = 'Logs:'
$form.Controls.Add($lblLogs)

$txtLogs = New-Object System.Windows.Forms.TextBox
$txtLogs.Location = New-Object System.Drawing.Point(15,395)
$txtLogs.Size = New-Object System.Drawing.Size(900,240)
$txtLogs.Multiline = $true
$txtLogs.ScrollBars = 'Vertical'
$txtLogs.ReadOnly = $true
$txtLogs.Font = New-Object System.Drawing.Font('Courier New', 9)
$form.Controls.Add($txtLogs)

$btnClearLogs = New-Object System.Windows.Forms.Button
$btnClearLogs.Location = New-Object System.Drawing.Point(815,330)
$btnClearLogs.Size = New-Object System.Drawing.Size(100,32)
$btnClearLogs.Text = 'Clear Logs'
$btnClearLogs.Add_Click({ $txtLogs.Clear() })
$form.Controls.Add($btnClearLogs)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12,645)
$lblStatus.Size = New-Object System.Drawing.Size(900,24)
$lblStatus.Text = 'Ready.'
$form.Controls.Add($lblStatus)

$script:selectedGroup = $null
$script:lastResults = $null

$form.Add_Shown({ $txtGroup.Focus() })
[void]$form.ShowDialog()
