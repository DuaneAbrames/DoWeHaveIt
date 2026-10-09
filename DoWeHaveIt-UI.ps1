[CmdletBinding()]
Param(
    [switch]$TimeMachine = $false,
    [switch]$Music = $false
)

. (Join-Path $PSScriptRoot 'DoWeHaveIt-Common.ps1')
$script:InfoLabelBaseColor = $null
$script:lastSuccessfulRunAt = $null
$script:displayedRunTimestamp = $null
$script:currentPublishedRunId = $null
$script:displayedRunId = $null
$script:isViewingHistoricalRun = $false
$script:lastStatusMessage = $null
$script:pollIntervalMilliseconds = 300000
$script:uiIdleThresholdMilliseconds = 3000000
$script:lastUiActivityAt = [datetime]::MinValue
$script:pendingGridRefresh = $null
$script:currentDataTable = $null

function Get-RunMetadataById {
	param(
		[Parameter(Mandatory = $true)]
		[int64]$RunId,
		[string]$DatasetType = $script:activeDatasetType
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	$rows = ConvertFrom-DataTable (Invoke-AppQuery @"
SELECT
	run_id,
	dataset_type,
	started_at,
	completed_at,
	status,
	total_rows
FROM inventory_runs
WHERE run_id = $RunId
	AND dataset_type = $(Escape-SqlString $normalizedDatasetType);
"@)

	if ($rows.Count -eq 0) {
		return $null
	}

	$row = $rows[0]
	return [pscustomobject]@{
		RunId = [int64]$row.run_id
		DatasetType = [string]$row.dataset_type
		StartedAt = if ($row.started_at -eq [DBNull]::Value) { $null } else { [datetime]$row.started_at }
		CompletedAt = if ($row.completed_at -eq [DBNull]::Value) { $null } else { [datetime]$row.completed_at }
		Status = [string]$row.status
		RowCount = [int]$row.total_rows
	}
}

function Get-RunItems {
	param(
		[Parameter(Mandatory = $true)]
		[int64]$RunId
	)

	return @(ConvertFrom-DataTable (Invoke-AppQuery @"
SELECT
	name,
	location
FROM inventory_items
WHERE run_id = $RunId
ORDER BY name, location;
"@))
}

function Get-RunSources {
	param(
		[Parameter(Mandatory = $true)]
		[int64]$RunId
	)

	return @(ConvertFrom-DataTable (Invoke-AppQuery @"
SELECT
	source_name AS SourceName,
	row_count AS RowCount
FROM inventory_run_sources
WHERE run_id = $RunId
ORDER BY display_order;
"@))
}

function Get-CompletedRunsForPicker {
	param(
		[string]$DatasetType = $script:activeDatasetType
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	return @(ConvertFrom-DataTable (Invoke-AppQuery @"
SELECT
	run_id,
	dataset_type,
	completed_at,
	total_rows
FROM inventory_runs
WHERE status = 'completed'
	AND dataset_type = $(Escape-SqlString $normalizedDatasetType)
ORDER BY completed_at DESC, run_id DESC;
"@))
}

function New-RunDataTable {
	$table = New-Object System.Data.DataTable
	[void]$table.Columns.Add('Name')
	[void]$table.Columns.Add('Location')
	return ,$table
}

function Convert-ItemsToDataTable {
	param(
		$Items
	)

	$table = New-RunDataTable
	foreach ($item in @($Items)) {
		if ($null -eq $item) {
			continue
		}

		$row = $table.NewRow()
		$name = $null
		$location = $null

		if ($item.PSObject.Properties['name']) {
			$name = $item.name
		} elseif ($item.PSObject.Properties['Name']) {
			$name = $item.Name
		}

		if ($item.PSObject.Properties['location']) {
			$location = $item.location
		} elseif ($item.PSObject.Properties['Location']) {
			$location = $item.Location
		}

		$row['Name'] = [string]$name
		$row['Location'] = [string]$location
		$table.Rows.Add($row)
	}
	return ,$table
}

function Get-StatusAgeMinutes {
	param(
		[Nullable[datetime]]$Timestamp
	)

	if (-not $Timestamp.HasValue) {
		return $null
	}

	return [int][Math]::Floor(((Get-Date) - $Timestamp.Value).TotalMinutes)
}

function Update-InfoLabelPosition {
	param(
		[System.Windows.Forms.Label]$Label,
		[System.Windows.Forms.Form]$Form
	)

	try {
		if ($null -eq $Label -or $null -eq $Form) {
			return
		}
		$pref = $Label.PreferredSize
		$Label.Location = New-Object System.Drawing.Point(($Form.ClientSize.Width - 20 - $pref.Width), $Label.Location.Y)
	} catch {}
}

function Set-InfoLabel {
	param(
		[System.Windows.Forms.Label]$Label,
		[int]$RowCount,
		$DisplayTimestamp,
		$HeartbeatTimestamp,
		[System.Windows.Forms.Form]$Form,
		[string]$DatasetType = $script:activeDatasetType,
		[switch]$HistoricalRun,
		[string]$StatusMessage
	)

	if ($null -eq $Label) {
		return
	}

	if ($null -eq $script:InfoLabelBaseColor) {
		$script:InfoLabelBaseColor = $Label.ForeColor
	}

	$displayDate = $null
	if ($null -ne $DisplayTimestamp -and $DisplayTimestamp -ne [DBNull]::Value) {
		$displayDate = [datetime]$DisplayTimestamp
	}

	$heartbeatDate = $null
	if ($null -ne $HeartbeatTimestamp -and $HeartbeatTimestamp -ne [DBNull]::Value) {
		$heartbeatDate = [datetime]$HeartbeatTimestamp
	}

	$displayText = if ($null -ne $displayDate) {
		$displayDate.ToString('MM-dd-yyyy HH:mm')
	} else {
		'No data'
	}

	$prefix = if ($HistoricalRun) { 'Snapshot' } else { 'OK' }
	$datasetLabel = Get-DatasetDisplayName $DatasetType
	$text = "$RowCount r - $datasetLabel - $prefix - $displayText"
	if (-not [string]::IsNullOrWhiteSpace($StatusMessage)) {
		$text = "$text - $StatusMessage"
	}
	$Label.Text = $text

	$isStale = $false
	if ($null -ne $heartbeatDate) {
		$isStale = ((Get-Date) - $heartbeatDate).TotalMinutes -gt 90
	}

	if ($isStale) {
		$Label.ForeColor = [System.Drawing.Color]::Firebrick
	} else {
		$Label.ForeColor = $script:InfoLabelBaseColor
	}

	if ($null -ne $Form) {
		Update-InfoLabelPosition -Label $Label -Form $Form
	}
}

function Get-ControlCenterY {
	param([System.Windows.Forms.Control]$ctl)
	return [int][Math]::Round($ctl.Location.Y + ($ctl.Height / 2.0))
}

function Align-TopRowControls {
	param(
		[System.Windows.Forms.Control[]]$Controls
	)

	$usableControls = @($Controls | Where-Object { $null -ne $_ -and $_.Visible })
	if ($usableControls.Count -eq 0) {
		return
	}

	$centers = foreach ($control in $usableControls) {
		Get-ControlCenterY -ctl $control
	}
	$centerY = [int][Math]::Round((($centers | Measure-Object -Average).Average))
	foreach ($control in $usableControls) {
		$newY = [int]($centerY - [Math]::Round($control.Height / 2.0))
		$control.Location = New-Object System.Drawing.Point($control.Location.X, $newY)
	}
}

function Escape-ForRowFilter {
	param([string]$text)

	$sb = New-Object System.Text.StringBuilder
	foreach ($ch in $text.ToCharArray()) {
		switch ($ch) {
			"'" { [void]$sb.Append("''") }
			"[" { [void]$sb.Append("[[]") }
			"]" { [void]$sb.Append("[]]") }
			"%" { [void]$sb.Append("[%]") }
			"*" { [void]$sb.Append("[*]") }
			"_" { [void]$sb.Append("[_]") }
			"?" { [void]$sb.Append("[?]") }
			"#" { [void]$sb.Append("[#]") }
			Default { [void]$sb.Append($ch) }
		}
	}
	return $sb.ToString()
}

function Get-ResolvedItemPath {
	param(
		[string]$name,
		[string]$location
	)

	$folder = $null
	$fileToSelect = $null
	$explorable = $true

	if ([string]::IsNullOrWhiteSpace($location) -or [string]::IsNullOrWhiteSpace($name)) {
		return [pscustomobject]@{ Explorable = $false; Folder = $null; File = $null; ItemPath = $null }
	}

	if ($location -in @('DVD Profiler', 'Disc List')) {
		$explorable = $false
	} elseif ($location -match '^(Movies|Troma|TV Shows|101 Horror Movies Mega Pack)$') {
		$folder = Join-Path (Join-Path $mediaDir $location) $name
	} elseif ($location -eq 'GoG') {
		$folder = Join-Path $GogDir $name
	} elseif ($location -eq 'SteamUnlocked') {
		$fileToSelect = Join-Path $SteamUnlockedDir $name
		$folder = Split-Path -Parent $fileToSelect
	} elseif ($location -eq 'Completed') {
		$fileToSelect = Join-Path $CompletedDir $name
		$folder = Split-Path -Parent $fileToSelect
	} elseif ($location -like 'SwitchRoms*') {
		$sub = $location -replace '^SwitchRoms[\\/]*', ''
		$folder = Join-Path $SwitchDir $sub
		$fileToSelect = Join-Path $folder $name
	} else {
		$folder = Join-Path $mediaDir $location
		$fileToSelect = Join-Path $folder $name
	}

	$itemPath = if ($fileToSelect) { $fileToSelect } else { $folder }
	return [pscustomobject]@{ Explorable = $explorable; Folder = $folder; File = $fileToSelect; ItemPath = $itemPath }
}

function Show-SourceLogDialog {
	param(
		[int64]$RunId,
		[System.Windows.Forms.Form]$Owner,
		[System.Drawing.Icon]$Icon
	)

	$logRows = Get-RunSources -RunId $RunId
	if ($logRows.Count -eq 0) {
		[System.Windows.Forms.MessageBox]::Show(
			"No source counts were found for run $RunId.",
			'Data Sources',
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Information
		) | Out-Null
		return
	}

	$logTable = New-Object System.Data.DataTable
	[void]$logTable.Columns.Add('Source')
	[void]$logTable.Columns.Add('Rows')
	foreach ($logRow in $logRows) {
		$dtRow = $logTable.NewRow()
		$dtRow['Source'] = $logRow.SourceName
		$dtRow['Rows'] = $logRow.RowCount
		$logTable.Rows.Add($dtRow)
	}

	$logForm = New-Object System.Windows.Forms.Form
	$logForm.Text = 'Data Sources'
	$logForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
	$logForm.ClientSize = New-Object System.Drawing.Size(700, 500)
	$logForm.MinimumSize = New-Object System.Drawing.Size(500, 300)
	$logForm.MaximizeBox = $true
	$logForm.MinimizeBox = $false
	$logForm.ShowInTaskbar = $false
	$logForm.Icon = $Icon

	$logGrid = New-Object System.Windows.Forms.DataGridView
	$logGrid.Location = New-Object System.Drawing.Point(12, 12)
	$logGrid.Size = New-Object System.Drawing.Size(676, 442)
	$logGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
		[System.Windows.Forms.AnchorStyles]::Bottom -bor
		[System.Windows.Forms.AnchorStyles]::Left -bor
		[System.Windows.Forms.AnchorStyles]::Right
	$logGrid.DataSource = $logTable
	$logGrid.AutoGenerateColumns = $true
	$logGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
	$logGrid.BackgroundColor = [System.Drawing.SystemColors]::Window
	$logGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$logGrid.RowHeadersVisible = $false
	$logGrid.ReadOnly = $true
	$logGrid.EnableHeadersVisualStyles = $false
	$logGrid.AllowUserToAddRows = $false
	$logGrid.AllowUserToDeleteRows = $false
	$logGrid.AllowUserToResizeRows = $false
	$logGrid.AllowUserToResizeColumns = $false
	$logGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
	$logGrid.MultiSelect = $false
	$logGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
	$logGrid.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
	$logGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(52, 73, 94)
	$logGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
	$logGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(52, 73, 94)
	$logGrid.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
	$logGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font($logGrid.Font, [System.Drawing.FontStyle]::Bold)
	$logGrid.DefaultCellStyle.SelectionBackColor = $logGrid.DefaultCellStyle.BackColor
	$logGrid.DefaultCellStyle.SelectionForeColor = $logGrid.DefaultCellStyle.ForeColor
	$logGrid.Add_DataBindingComplete({
		param($incomingSender, $e)
		foreach ($column in $incomingSender.Columns) {
			$column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
		}
	})
	$logGrid.Add_RowPrePaint({
		param($incomingSender, $e)
		if ($e.RowIndex -lt 0) { return }
		$isLastRow = $e.RowIndex -eq ($incomingSender.Rows.Count - 1)
		if ($isLastRow) {
			$row = $incomingSender.Rows[$e.RowIndex]
			$row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 244, 214)
			$row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(90, 60, 0)
			$row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 230, 170)
			$row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(90, 60, 0)
			$row.DefaultCellStyle.Font = New-Object System.Drawing.Font($incomingSender.Font, [System.Drawing.FontStyle]::Bold)
		}
	})

	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = 'OK'
	$okButton.Size = New-Object System.Drawing.Size(90, 28)
	$okButton.Location = New-Object System.Drawing.Point(598, 462)
	$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

	$logForm.AcceptButton = $okButton
	$logForm.CancelButton = $okButton
	$logForm.Controls.AddRange(@($logGrid, $okButton))

	if ($Owner) {
		[void]$logForm.ShowDialog($Owner)
	} else {
		[void]$logForm.ShowDialog()
	}
}

function Show-RunPickerDialog {
	param(
		[System.Windows.Forms.Form]$Owner,
		[System.Drawing.Icon]$Icon,
		[int64]$CurrentPublishedRunId,
		[string]$DatasetType = $script:activeDatasetType
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	$runs = Get-CompletedRunsForPicker -DatasetType $normalizedDatasetType
	if ($runs.Count -eq 0) {
		[System.Windows.Forms.MessageBox]::Show(
			"No completed $(Get-DatasetDisplayName $normalizedDatasetType) snapshots are available yet.",
			'Time Machine',
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Information
		) | Out-Null
		return $null
	}

	$table = New-Object System.Data.DataTable
	[void]$table.Columns.Add('Run ID', [long])
	[void]$table.Columns.Add('When')
	[void]$table.Columns.Add('Rows', [int])
	[void]$table.Columns.Add('Age')
	[void]$table.Columns.Add('Mode')

	foreach ($run in $runs) {
		$dtRow = $table.NewRow()
		$completedAt = [datetime]$run.completed_at
		$age = New-TimeSpan -Start $completedAt -End (Get-Date)
		$ageText = if ($age.TotalDays -ge 1) {
			'{0}d {1}h' -f [int]$age.TotalDays, $age.Hours
		} elseif ($age.TotalHours -ge 1) {
			'{0}h {1}m' -f [int]$age.TotalHours, $age.Minutes
		} else {
			'{0}m' -f [int]$age.TotalMinutes
		}

		$dtRow['Run ID'] = [int64]$run.run_id
		$dtRow['When'] = $completedAt.ToString('MM-dd-yyyy HH:mm')
		$dtRow['Rows'] = [int]$run.total_rows
		$dtRow['Age'] = $ageText
		$dtRow['Mode'] = if ([int64]$run.run_id -eq $CurrentPublishedRunId) { 'Current' } else { 'Snapshot' }
		$table.Rows.Add($dtRow)
	}

	$pickerForm = New-Object System.Windows.Forms.Form
	$pickerForm.Text = "Time Machine - $(Get-DatasetDisplayName $normalizedDatasetType)"
	$pickerForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
	$pickerForm.ClientSize = New-Object System.Drawing.Size(760, 520)
	$pickerForm.MinimumSize = New-Object System.Drawing.Size(620, 350)
	$pickerForm.ShowInTaskbar = $false
	$pickerForm.Icon = $Icon

	$grid = New-Object System.Windows.Forms.DataGridView
	$grid.Location = New-Object System.Drawing.Point(12, 12)
	$grid.Size = New-Object System.Drawing.Size(736, 438)
	$grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
		[System.Windows.Forms.AnchorStyles]::Bottom -bor
		[System.Windows.Forms.AnchorStyles]::Left -bor
		[System.Windows.Forms.AnchorStyles]::Right
	$grid.DataSource = $table
	$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
	$grid.RowHeadersVisible = $false
	$grid.ReadOnly = $true
	$grid.MultiSelect = $false
	$grid.AllowUserToAddRows = $false
	$grid.AllowUserToDeleteRows = $false
	$grid.AllowUserToResizeRows = $false
	$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

	$selectedRunId = $null

	$useButton = New-Object System.Windows.Forms.Button
	$useButton.Text = 'Use Snapshot'
	$useButton.Size = New-Object System.Drawing.Size(120, 30)
	$useButton.Location = New-Object System.Drawing.Point(498, 468)
	$useButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

	$currentButton = New-Object System.Windows.Forms.Button
	$currentButton.Text = 'Current'
	$currentButton.Size = New-Object System.Drawing.Size(90, 30)
	$currentButton.Location = New-Object System.Drawing.Point(370, 468)
	$currentButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = 'Cancel'
	$cancelButton.Size = New-Object System.Drawing.Size(90, 30)
	$cancelButton.Location = New-Object System.Drawing.Point(630, 468)
	$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

	$applySelection = {
		if ($grid.CurrentRow -and $grid.CurrentRow.Index -ge 0) {
			$selectedRunId = [int64]$grid.CurrentRow.Cells['Run ID'].Value
			$pickerForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$pickerForm.Close()
		}
	}

	$useButton.Add_Click($applySelection)
	$grid.Add_CellDoubleClick({
		param($sender, $e)
		if ($e.RowIndex -ge 0) {
			& $applySelection
		}
	})

	$currentButton.Add_Click({
		$selectedRunId = $CurrentPublishedRunId
		$pickerForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$pickerForm.Close()
	})

	$pickerForm.AcceptButton = $useButton
	$pickerForm.CancelButton = $cancelButton
	$pickerForm.Controls.AddRange(@($grid, $useButton, $currentButton, $cancelButton))
	[void]$pickerForm.ShowDialog($Owner)

	return $selectedRunId
}

function Show-ConnectionFailure {
	param(
		[string]$Message,
		[System.Windows.Forms.Form]$Owner
	)

	[System.Windows.Forms.MessageBox]::Show(
		$Message,
		'Database Connection',
		[System.Windows.Forms.MessageBoxButtons]::OK,
		[System.Windows.Forms.MessageBoxIcon]::Error
	) | Out-Null
}

function Start-UiMode {
	if ($PSVersionTable.Platform -eq 'Unix') {
		throw 'UI mode is only supported on Windows. Run the worker with -createList from Linux.'
	}

	Add-Type -AssemblyName System.Windows.Forms
	[System.Windows.Forms.Application]::EnableVisualStyles()

	try {
		Initialize-AppDatabase
	} catch {
		Show-ConnectionFailure -Message $_.Exception.Message -Owner $null
		return
	}

	$state = Get-CurrentPublishedRunMetadata -DatasetType $script:activeDatasetType
	if (-not $state -or -not $state.RunId) {
		$items = @()
		$script:currentDataTable = Convert-ItemsToDataTable -Items $items
		$script:lastSuccessfulRunAt = $null
		$script:currentPublishedRunId = $null
		$script:displayedRunId = $null
		$script:displayedRunTimestamp = $null
	} else {
		$items = Get-RunItems -RunId $state.RunId
		$script:currentDataTable = Convert-ItemsToDataTable -Items $items
		$script:lastSuccessfulRunAt = $state.LastSuccessfulRunAt
		$script:currentPublishedRunId = $state.RunId
		$script:displayedRunId = $state.RunId
		$script:displayedRunTimestamp = $state.LastSuccessfulRunAt
	}
	$script:isViewingHistoricalRun = $false

	$FormStartWidth = if ($TimeMachine) { 1065 } else { 940 }

	$Form = New-Object system.Windows.Forms.Form
	$Form.ClientSize = New-Object System.Drawing.Point($FormStartWidth, 400)
	$Form.Text = 'Do We Have It?'
	$Icon = New-Object system.drawing.icon ("$scriptDir/Help.256.ico")
	$Form.Icon = $Icon
	$Form.TopMost = $false
	$Form.MinimumSize = New-Object System.Drawing.Size(740, 250)

	$FilterTextBox = New-Object system.Windows.Forms.TextBox
	$FilterTextBox.Multiline = $false
	$FilterTextBox.Width = 334
	$FilterTextBox.Height = 20
	$FilterTextBox.Location = New-Object System.Drawing.Point(50, 5)
	$FilterTextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

	$FilterBoxLabel = New-Object system.Windows.Forms.Label
	$FilterBoxLabel.Text = 'Filter:'
	$FilterBoxLabel.AutoSize = $true
	$FilterBoxLabel.Location = New-Object System.Drawing.Point(10, 8)
	$FilterBoxLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

	$InfoLabel = New-Object system.Windows.Forms.Label
	$InfoLabel.AutoSize = $true
	$InfoLabel.Location = New-Object System.Drawing.Point(($Form.Width - 228), 9)
	$InfoLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
	$InfoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$InfoLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$InfoLabel.Cursor = [System.Windows.Forms.Cursors]::Hand

	$DataGridView1 = New-Object system.Windows.Forms.DataGridView
	$DataGridView1.Width = $FormStartWidth - 40
	$DataGridView1.Height = 340
	$DataGridView1.Location = New-Object System.Drawing.Point(20, 50)
	$DataGridView1.DataSource = $script:currentDataTable
	$DataGridView1.AutoSizeColumnsMode = 'Fill'
	$DataGridView1.RowHeadersVisible = $false
	$DataGridView1.ReadOnly = $true
	$DataGridView1.AllowUserToResizeColumns = $false
	$DataGridView1.AllowUserToResizeRows = $false
	$DataGridView1.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
	$DataGridView1.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
	$DataGridView1.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableWithoutHeaderText

	$gridContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
	$miCopy = New-Object System.Windows.Forms.ToolStripMenuItem 'Copy'
	$miCopyPath = New-Object System.Windows.Forms.ToolStripMenuItem 'Copy Path'
	$miSearchGoogle = New-Object System.Windows.Forms.ToolStripMenuItem 'Search with Google'
	[void]$gridContextMenu.Items.Add($miCopy)
	[void]$gridContextMenu.Items.Add($miCopyPath)
	[void]$gridContextMenu.Items.Add($miSearchGoogle)
	$DataGridView1.ContextMenuStrip = $gridContextMenu

	$PasteButton = New-Object system.Windows.Forms.Button
	$PasteButton.Text = 'Paste'
	$PasteButton.Width = 78
	$PasteButton.Height = 21
	$PasteButton.Location = New-Object System.Drawing.Point(392, 6)
	$PasteButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

	$ClearButton = New-Object system.Windows.Forms.Button
	$ClearButton.Text = 'Clear'
	$ClearButton.Width = 78
	$ClearButton.Height = 21
	$ClearButton.Location = New-Object System.Drawing.Point(477, 6)
	$ClearButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

	$MusicCheckbox = New-Object System.Windows.Forms.CheckBox
	$MusicCheckbox.Text = 'Music'
	$MusicCheckbox.AutoSize = $true
	$MusicCheckbox.Location = New-Object System.Drawing.Point(565, 8)
	$MusicCheckbox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
	$MusicCheckbox.Checked = ($script:activeDatasetType -eq $MusicDatasetType)

	$TimeMachineButton = $null
	if ($TimeMachine) {
		$TimeMachineButton = New-Object system.Windows.Forms.Button
		$TimeMachineButton.Text = 'Time Machine'
		$TimeMachineButton.Width = 105
		$TimeMachineButton.Height = 21
		$TimeMachineButton.Location = New-Object System.Drawing.Point(640, 6)
		$TimeMachineButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
	}

	$PollTimer = New-Object System.Windows.Forms.Timer
	$PollTimer.Interval = $script:pollIntervalMilliseconds

	$UiIdleTimer = New-Object System.Windows.Forms.Timer
	$UiIdleTimer.Interval = 1000

	$UpdateInfoDisplay = {
		$view = $DataGridView1.DataSource
		$rowCount = if ($view -is [System.Data.DataView]) { $view.Count } elseif ($script:currentDataTable) { $script:currentDataTable.Rows.Count } else { 0 }
		Set-InfoLabel -Label $InfoLabel -RowCount $rowCount -DisplayTimestamp $script:displayedRunTimestamp -HeartbeatTimestamp $script:lastSuccessfulRunAt -Form $Form -DatasetType $script:activeDatasetType -HistoricalRun:$script:isViewingHistoricalRun -StatusMessage $script:lastStatusMessage
	}

	$MarkUiActivity = {
		$script:lastUiActivityAt = Get-Date
	}

	$IsUiIdle = {
		if ($script:lastUiActivityAt -eq [datetime]::MinValue) {
			return $true
		}

		return (((Get-Date) - $script:lastUiActivityAt).TotalMilliseconds -ge $script:uiIdleThresholdMilliseconds)
	}

	$ClearPendingGridRefresh = {
		$script:pendingGridRefresh = $null
	}

	$ApplyPendingGridRefresh = {
		if ($null -eq $script:pendingGridRefresh) {
			return $false
		}

		$pending = $script:pendingGridRefresh
		if ($pending.DatasetType -ne $script:activeDatasetType) {
			$script:pendingGridRefresh = $null
			return $false
		}

		$script:pendingGridRefresh = $null

		if ($pending.ClearGrid) {
			$script:displayedRunId = $null
			$script:displayedRunTimestamp = $null
			$script:currentDataTable = Convert-ItemsToDataTable -Items @()
			$DataGridView1.DataSource = $script:currentDataTable
			& $ApplyFilter
			return $true
		}

		if ($pending.RunId -and ($script:displayedRunId -ne $pending.RunId)) {
			& $LoadRunIntoGrid -RunId $pending.RunId -DisplayTimestamp $pending.LastSuccessfulRunAt
			return $true
		}

		$script:displayedRunTimestamp = $pending.LastSuccessfulRunAt
		& $UpdateInfoDisplay
		return $true
	}

	$QueuePendingGridRefresh = {
		param(
			$LatestMetadata,
			[switch]$ClearGrid
		)

		$script:pendingGridRefresh = [pscustomobject]@{
			ClearGrid = $ClearGrid.IsPresent
			DatasetType = $script:activeDatasetType
			RunId = if ($LatestMetadata) { $LatestMetadata.RunId } else { $null }
			LastSuccessfulRunAt = if ($LatestMetadata) { $LatestMetadata.LastSuccessfulRunAt } else { $null }
		}
	}

	$RegisterIdleTracking = {
		param([System.Windows.Forms.Control]$Control)

		if (-not $Control) {
			return
		}

		$Control.Add_Click({ & $MarkUiActivity })
		$Control.Add_MouseDown({ & $MarkUiActivity })
		$Control.Add_MouseWheel({ & $MarkUiActivity })
		$Control.Add_KeyDown({ & $MarkUiActivity })

		if ($Control -is [System.Windows.Forms.TextBoxBase]) {
			$Control.Add_TextChanged({ & $MarkUiActivity })
		}

		if ($Control -is [System.Windows.Forms.DataGridView]) {
			$Control.Add_Scroll({ & $MarkUiActivity })
		}

		$Control.Add_ControlAdded({
			param($sender, $e)
			if ($e.Control) {
				& $RegisterIdleTracking -Control $e.Control
			}
		})

		foreach ($child in $Control.Controls) {
			& $RegisterIdleTracking -Control $child
		}
	}

	$ApplyFilter = {
		$dv = New-Object System.Data.DataView($script:currentDataTable)
		$terms = $FilterTextBox.Text -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
		if ($terms.Count -eq 0) {
			$dv.RowFilter = $null
		} else {
			$escapedTerms = $terms | ForEach-Object { Escape-ForRowFilter $_ }
			$dv.RowFilter = ($escapedTerms | ForEach-Object { "(Name LIKE '%$_%' OR Location LIKE '%$_%')" }) -join ' AND '
		}
		$DataGridView1.DataSource = $dv
		& $UpdateInfoDisplay
	}

	$LoadRunIntoGrid = {
		param(
			[int64]$RunId,
			$DisplayTimestamp,
			[switch]$HistoricalView
		)

		$itemsForRun = Get-RunItems -RunId $RunId
		$newTable = Convert-ItemsToDataTable -Items $itemsForRun
		$script:currentDataTable = $newTable
		$script:displayedRunId = $RunId
		$script:displayedRunTimestamp = if ($null -ne $DisplayTimestamp -and $DisplayTimestamp -ne [DBNull]::Value) { [datetime]$DisplayTimestamp } else { $null }
		$script:isViewingHistoricalRun = $HistoricalView.IsPresent
		$DataGridView1.DataSource = $script:currentDataTable
		& $ApplyFilter
		$FilterTextBox.Focus()
		$FilterTextBox.SelectionStart = $FilterTextBox.Text.Length
	}

	$ReloadCurrentRun = {
		& $ClearPendingGridRefresh
		try {
			$latest = Get-CurrentPublishedRunMetadata -DatasetType $script:activeDatasetType
			if (-not $latest -or -not $latest.RunId) {
				$script:lastSuccessfulRunAt = $null
				$script:currentPublishedRunId = $null
				$script:displayedRunId = $null
				$script:displayedRunTimestamp = $null
				$script:currentDataTable = Convert-ItemsToDataTable -Items @()
				$DataGridView1.DataSource = $script:currentDataTable
				& $ApplyFilter
				return
			}

			$script:lastSuccessfulRunAt = $latest.LastSuccessfulRunAt
			$script:currentPublishedRunId = $latest.RunId
			if (-not $script:isViewingHistoricalRun) {
				& $LoadRunIntoGrid -RunId $latest.RunId -DisplayTimestamp $latest.LastSuccessfulRunAt
			} else {
				& $ApplyFilter
			}
			$script:lastStatusMessage = $null
		} catch {
			$script:lastStatusMessage = 'DB unavailable'
			& $ApplyFilter
		}
	}

	$miCopy.Add_Click({
		try {
			if ($DataGridView1.GetCellCount([System.Windows.Forms.DataGridViewElementStates]::Selected) -gt 0) {
				$dataObj = $DataGridView1.GetClipboardContent()
				if ($dataObj) {
					[System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)
				}
			} elseif ($DataGridView1.CurrentCell -and $null -ne $DataGridView1.CurrentCell.Value) {
				[System.Windows.Forms.Clipboard]::SetText([string]$DataGridView1.CurrentCell.Value)
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				'Could not copy selection to the clipboard.',
				'Copy',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	$DataGridView1.Add_CellMouseDown({
		param($incomingSender, $e)
		if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
			$DataGridView1.CurrentCell = $DataGridView1.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
			$DataGridView1.Rows[$e.RowIndex].Selected = $true
		}
	})

	$DataGridView1.Add_CellDoubleClick({
		param($incomingSender, $e)
		if ($e.RowIndex -lt 0) { return }
		$row = $DataGridView1.Rows[$e.RowIndex]
		$name = [string]$row.Cells['Name'].Value
		$location = [string]$row.Cells['Location'].Value
		if ([string]::IsNullOrWhiteSpace($location) -or [string]::IsNullOrWhiteSpace($name)) { return }

		$res = Get-ResolvedItemPath -name $name -location $location
		if (-not $res.Explorable) {
			[System.Windows.Forms.MessageBox]::Show(
				"Cannot explore items from '$location'.",
				'Open Location',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			) | Out-Null
			return
		}

		try {
			if ($res.File -and (Test-Path -LiteralPath $res.File)) {
				Start-Process -FilePath 'explorer.exe' -ArgumentList ("/select,""$($res.File)""")
			} elseif ($res.Folder -and (Test-Path -LiteralPath $res.Folder)) {
				Start-Process -FilePath 'explorer.exe' -ArgumentList ("""$($res.Folder)""")
			} else {
				[System.Windows.Forms.MessageBox]::Show(
					"Path not found for '$name' (Location: $location).",
					'Open Location',
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Warning
				) | Out-Null
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				'Could not open Explorer.',
				'Open Location',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	$miCopyPath.Add_Click({
		try {
			if ($DataGridView1.CurrentCell -and $DataGridView1.CurrentCell.RowIndex -ge 0) {
				$row = $DataGridView1.Rows[$DataGridView1.CurrentCell.RowIndex]
				$res = Get-ResolvedItemPath -name ([string]$row.Cells['Name'].Value) -location ([string]$row.Cells['Location'].Value)
				if (-not $res.Explorable -or [string]::IsNullOrEmpty($res.ItemPath)) {
					[System.Windows.Forms.MessageBox]::Show(
						"Cannot determine a filesystem path for '$([string]$row.Cells['Location'].Value)'.",
						'Copy Path',
						[System.Windows.Forms.MessageBoxButtons]::OK,
						[System.Windows.Forms.MessageBoxIcon]::Information
					) | Out-Null
					return
				}
				[System.Windows.Forms.Clipboard]::SetText($res.ItemPath)
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				'Could not copy the path to the clipboard.',
				'Copy Path',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	$miSearchGoogle.Add_Click({
		try {
			if ($DataGridView1.CurrentCell -and $DataGridView1.CurrentCell.RowIndex -ge 0) {
				$row = $DataGridView1.Rows[$DataGridView1.CurrentCell.RowIndex]
				$name = [string]$row.Cells['Name'].Value
				if (-not [string]::IsNullOrWhiteSpace($name)) {
					$url = 'https://www.google.com/search?q=' + [System.Uri]::EscapeDataString($name)
					Start-Process $url | Out-Null
				}
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				'Could not open browser for Google search.',
				'Search with Google',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	$DataGridView1.Add_KeyDown({
		param($incomingSender, $e)
		if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
			try {
				if ($DataGridView1.GetCellCount([System.Windows.Forms.DataGridViewElementStates]::Selected) -gt 0) {
					$dataObj = $DataGridView1.GetClipboardContent()
					if ($dataObj) {
						[System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)
					}
				} elseif ($DataGridView1.CurrentCell -and $null -ne $DataGridView1.CurrentCell.Value) {
					[System.Windows.Forms.Clipboard]::SetText([string]$DataGridView1.CurrentCell.Value)
				}
			} catch {
				[System.Windows.Forms.MessageBox]::Show(
					'Could not copy selection to the clipboard.',
					'Copy',
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
			}
			$e.Handled = $true
		}
	})

	$FilterTextBox.Add_TextChanged({
		& $MarkUiActivity
		& $ApplyFilter
	})

	$InfoLabel.Add_Click({
		try {
			if ($script:displayedRunId) {
				Show-SourceLogDialog -RunId $script:displayedRunId -Owner $Form -Icon $Icon
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				"Could not load source counts.`r`n$($_.Exception.Message)",
				'Data Sources',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	$PasteButton.Add_Click({
		try {
			if (-not [System.Windows.Forms.Clipboard]::ContainsText()) {
				[System.Windows.Forms.MessageBox]::Show(
					'Clipboard does not contain text.',
					'Paste',
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Information
				) | Out-Null
				return
			}

			$text = [System.Windows.Forms.Clipboard]::GetText().Trim()
			if ([string]::IsNullOrWhiteSpace($text)) {
				[System.Windows.Forms.MessageBox]::Show(
					'Clipboard text is empty.',
					'Paste',
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Information
				) | Out-Null
				return
			}

			if ($text.Length -gt 50) {
				[System.Windows.Forms.MessageBox]::Show(
					'Clipboard text is longer than 50 characters (limit: 50).',
					'Paste',
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Warning
				) | Out-Null
				return
			}

			$FilterTextBox.Text = $text
			$FilterTextBox.Focus()
			$FilterTextBox.SelectionStart = $FilterTextBox.Text.Length
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				'Could not read the clipboard.',
				'Paste',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	$ClearButton.Add_Click({
		$FilterTextBox.Text = ''
		$FilterTextBox.Focus()
		$FilterTextBox.SelectionStart = 0
	})

	$MusicCheckbox.Add_CheckedChanged({
		try {
			& $MarkUiActivity
			& $ClearPendingGridRefresh
			$script:activeDatasetType = if ($MusicCheckbox.Checked) { $MusicDatasetType } else { $RegularDatasetType }
			$script:isViewingHistoricalRun = $false
			$script:lastStatusMessage = $null
			$FilterTextBox.Text = ''
			& $ReloadCurrentRun
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				"Could not switch datasets.`r`n$($_.Exception.Message)",
				'Dataset Switch',
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	if ($TimeMachineButton) {
		$TimeMachineButton.Add_Click({
			try {
				& $MarkUiActivity
				& $ClearPendingGridRefresh
				$selectedRunId = Show-RunPickerDialog -Owner $Form -Icon $Icon -CurrentPublishedRunId $script:currentPublishedRunId -DatasetType $script:activeDatasetType
				if (-not $selectedRunId) {
					return
				}

				if ($selectedRunId -eq $script:currentPublishedRunId) {
					$script:isViewingHistoricalRun = $false
					$script:lastStatusMessage = $null
					& $ReloadCurrentRun
					return
				}

				$selectedMeta = Get-RunMetadataById -RunId $selectedRunId -DatasetType $script:activeDatasetType
				if (-not $selectedMeta) {
					return
				}

				$script:lastStatusMessage = $null
				& $LoadRunIntoGrid -RunId $selectedRunId -DisplayTimestamp $selectedMeta.CompletedAt -HistoricalView
			} catch {
				[System.Windows.Forms.MessageBox]::Show(
					"Could not load snapshot.`r`n$($_.Exception.Message)",
					'Time Machine',
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
			}
		})
	}

	$PollTimer.Add_Tick({
		if ($script:isViewingHistoricalRun) {
			try {
				$latest = Get-CurrentPublishedRunMetadata -DatasetType $script:activeDatasetType
				if ($latest) {
					$script:lastSuccessfulRunAt = $latest.LastSuccessfulRunAt
					$script:currentPublishedRunId = $latest.RunId
					$script:lastStatusMessage = $null
				}
			} catch {
				$script:lastStatusMessage = 'DB unavailable'
			}
			& $UpdateInfoDisplay
			return
		}

		try {
			$latest = Get-CurrentPublishedRunMetadata -DatasetType $script:activeDatasetType
			if (-not $latest -or -not $latest.RunId) {
				$script:lastSuccessfulRunAt = $null
				$script:currentPublishedRunId = $null
				$script:lastStatusMessage = $null
				if (& $IsUiIdle) {
					$script:displayedRunId = $null
					$script:displayedRunTimestamp = $null
					$script:currentDataTable = Convert-ItemsToDataTable -Items @()
					$DataGridView1.DataSource = $script:currentDataTable
					& $ClearPendingGridRefresh
					& $ApplyFilter
				} else {
					& $QueuePendingGridRefresh -ClearGrid
					& $UpdateInfoDisplay
				}
				return
			}

			$script:lastSuccessfulRunAt = $latest.LastSuccessfulRunAt
			$script:currentPublishedRunId = $latest.RunId
			if ($script:displayedRunId -ne $latest.RunId) {
				$script:lastStatusMessage = $null
				if (& $IsUiIdle) {
					& $ClearPendingGridRefresh
					& $LoadRunIntoGrid -RunId $latest.RunId -DisplayTimestamp $latest.LastSuccessfulRunAt
				} else {
					& $QueuePendingGridRefresh -LatestMetadata $latest
					& $UpdateInfoDisplay
				}
			} else {
				$script:displayedRunTimestamp = $latest.LastSuccessfulRunAt
				$script:lastStatusMessage = $null
				& $UpdateInfoDisplay
			}
		} catch {
			$script:lastStatusMessage = 'DB unavailable'
			& $UpdateInfoDisplay
		}
	})

	$UiIdleTimer.Add_Tick({
		if (& $IsUiIdle) {
			[void](& $ApplyPendingGridRefresh)
		}
	})

	$topRowControls = @($FilterBoxLabel, $FilterTextBox, $PasteButton, $ClearButton, $MusicCheckbox, $InfoLabel)
	if ($TimeMachineButton) {
		$topRowControls += $TimeMachineButton
	}

	$Form.Add_Resize({
		$newWidth = $this.Width
		$newHeight = $this.Height
		$DataGridView1.Width = $newWidth - 56
		$DataGridView1.Height = $newHeight - 99
		Align-TopRowControls -Controls $topRowControls
		Update-InfoLabelPosition -Label $InfoLabel -Form $this
	})

	$controls = @($FilterTextBox, $PasteButton, $ClearButton, $MusicCheckbox, $FilterBoxLabel, $InfoLabel, $DataGridView1)
	if ($TimeMachineButton) {
		$controls += $TimeMachineButton
	}
	$Form.Controls.AddRange($controls)
	& $RegisterIdleTracking -Control $Form

	$Form.Add_Shown({
		& $MarkUiActivity
		Align-TopRowControls -Controls $topRowControls
		Update-InfoLabelPosition -Label $InfoLabel -Form $Form
		& $ApplyFilter
		$PollTimer.Start()
		$UiIdleTimer.Start()
	})

	$Form.Add_FormClosing({
		param($sender, $e)
		if ($e.CloseReason -ne [System.Windows.Forms.CloseReason]::UserClosing) {
			return
		}

		$result = [System.Windows.Forms.MessageBox]::Show(
			'Quit Do We Have It?',
			'Confirm Exit',
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Question,
			[System.Windows.Forms.MessageBoxDefaultButton]::Button2
		)

		if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
			$e.Cancel = $true
		}
	})

	$Form.Add_FormClosed({
		$PollTimer.Stop()
		$UiIdleTimer.Stop()
		Close-AppConnection
	})

	[void]$Form.ShowDialog()
}

try {
    Start-UiMode
} finally {
    Close-AppConnection
}
