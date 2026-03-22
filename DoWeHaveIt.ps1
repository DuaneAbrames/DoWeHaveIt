[CmdletBinding()]
Param ([switch]$createList = $false
)
if ($PSVersionTable.Platform -eq "Unix") {
	$mediaDir = '/mnt/user/media/'
	$scriptDir = '/mnt/user/bigone/scripts/DoWeHaveIt/'
	$discListDir = '/mnt/user/bigone/DISC LIST/'
	$GogDir = '/mnt/user/bigone/Bittorrent/Completed/GoG/'
	$CompletedDir = '/mnt/user/bigone/Bittorrent/Completed/'
	$SteamUnlockedDir = '/mnt/user/bigone/Bittorrent/Completed/SteamUnlocked/'
	$SwitchDir = '/mnt/remotes/WHELP_SwitchRoms/'
	$createList = $true
} else {
    $mediaDir = 'R:/'
	$scriptDir = 'S:/scripts/DoWeHaveIt/'
	$discListDir = 'S:/DISC LIST/'
	$GogDir = 'S:/BitTorrent/Completed/Gog/'
	$CompletedDir = 'S:/BitTorrent/Completed/'
	$SteamUnlockedDir = 'S:/BitTorrent/Completed/SteamUnlocked/'
	$SwitchDir = '\\whelp\SwitchRoms\'
}
$outputCSV = "$($scriptDir)DoWeHaveIt.csv"
$makeListLog = "$($scriptDir)makelist.csv"
$ExcludedThings = ('___Duplicates','.trash')
if ($createList) {
	$makeListRows = New-Object System.Collections.Generic.List[object]

	function Write-CreateListMessage {
		param(
			[string]$Message
		)
		Write-Host $Message
	}

	function Write-SourceRowCount {
		param(
			[string]$Label,
			[int]$StartCount,
			[array]$OutputItems
		)
		$rowsAdded = $OutputItems.Count - $StartCount
		$makeListRows.Add([pscustomobject]@{
			Source = $Label
			Rows = $rowsAdded
		}) | Out-Null
		Write-CreateListMessage "$Label - $rowsAdded"
	}

	Write-Host 'Checking directories under media:'
	$dirsToInclude = ('Movies','TV Shows')
	$output = @()
	foreach ($d in $dirsToInclude) {
		$startCount = $output.Count
		foreach ($i in (Get-ChildItem "$mediaDir$d" -Directory)) {
			#Write-Host "    - $i
			if ($ExcludedThings -notcontains $i.name) {
				$output += new-object PSObject -property @{name=$i.name;location=$d}
			}
		}
		Write-SourceRowCount -Label $d -StartCount $startCount -OutputItems $output
	}

	$startCount = $output.Count
	$dl = import-csv "$($discListDir)disc list.csv"

	#This pulls the first column name from the CSV, which is the title of the movie/game/etc. The second column is usually "Location" but we want to be flexible in case it changes.
	$firstColumn = ($dl[0].PSObject.Properties | Select-Object -First 1).Name
	$xx = $dl | Select-Object -ExpandProperty $firstColumn

	foreach ($i in $xx) {
		if ($ExcludedThings -notcontains $i) {
			$output += new-object PSObject -property @{name=$i;location='Disc List'}
		}
	}
	Write-SourceRowCount -Label 'Disc List' -StartCount $startCount -OutputItems $output

	$startCount = $output.Count
	[xml]$dvdProfiler = get-content "$($discListDir)Collection.xml"
	foreach ($i in $dvdProfiler.Collection.DVD.title) {
		if ($ExcludedThings -notcontains $i) {
			$output += new-object PSObject -property @{name=$i;location='DVD Profiler'}
		}
	}
	Write-SourceRowCount -Label 'DVD Profiler' -StartCount $startCount -OutputItems $output
	## Episodes Section ##
	$VideoExtensions = @('.mkv','.mp4','.avi','.mov','.m4v','.wmv','.ts','.m2ts','.mpg','.mpeg','.flv','.webm')
	$dirsToInclude = ('TV Shows')
	foreach ($d in $dirsToInclude) {
		$startCount = $output.Count
		foreach ($i in (Get-ChildItem "$mediaDir$d" -Recurse -file | Where-Object{$VideoExtensions -contains $_.extension})) {
			$dir = Split-Path -leaf (Split-Path $i.fullname)
			if ($ExcludedThings -notcontains $i.name -and ($dir -ilike "Season *" -or $dir -ieq "Specials")) {
				$output += new-object PSObject -property @{name=$i.name;location=$i.DirectoryName.Replace($mediaDir,'')}
			}
		}
		Write-SourceRowCount -Label "$d (Episodes)" -StartCount $startCount -OutputItems $output
	}
	#Here will Be GoG, Switch, Etc.
	$dirsToInclude = ('Base','Update','DLC')
	foreach ($d in $dirsToInclude) {
		$startCount = $output.Count
		foreach ($i in (Get-ChildItem "$SwitchDir$d" -Recurse -file)) {
			if ($ExcludedThings -notcontains $i.name) {
				$output += new-object PSObject -property @{name=$i.name;location="SwitchRoms/$($i.DirectoryName.Replace($SwitchDir,''))"}
			}
		}
		Write-SourceRowCount -Label "$d (Switch)" -StartCount $startCount -OutputItems $output
	}
	$startCount = $output.Count
	foreach ($i in (Get-ChildItem "$CompletedDir" -File '*.rar')) {

		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="Completed"}
		}
	}
	Write-SourceRowCount -Label 'Completed (RAR Files)' -StartCount $startCount -OutputItems $output

	$startCount = $output.Count
	foreach ($i in (Get-ChildItem "$CompletedDir" -File '*.zip')) {

		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="Completed"}
		}
	}
	Write-SourceRowCount -Label 'Completed (ZIP Files)' -StartCount $startCount -OutputItems $output

	$startCount = $output.Count
	foreach ($i in (Get-ChildItem "$CompletedDir" -File '*.iso')) {

		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="Completed"}
		}
	}
	Write-SourceRowCount -Label 'Completed (ISO Files)' -StartCount $startCount -OutputItems $output

	$startCount = $output.Count
	foreach ($i in (Get-ChildItem "$SteamUnlockedDir" -File)) {
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="SteamUnlocked"}
		}
	}
	Write-SourceRowCount -Label 'SteamUnlocked' -StartCount $startCount -OutputItems $output

	$startCount = $output.Count
	foreach ($i in (Get-ChildItem "$GogDir" -Directory -Exclude "Logs")) {
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="GoG"}
		}
	}
	Write-SourceRowCount -Label 'GoG' -StartCount $startCount -OutputItems $output
	$output = $output | Where-Object{$_.name -notlike ""} | sort-object name,location |  Select-Object  name,location
	$output | Export-csv -notypeinformation $outputCSV
	$makeListRows.Add([pscustomobject]@{
		Source = 'TOTAL'
		Rows = $output.count
	}) | Out-Null
	Write-CreateListMessage "$($output.count) rows exported."
	$makeListRows | Export-Csv -NoTypeInformation -Path $makeListLog
} else {
	$csv = Import-CSV  $outputCSV |  Select-Object  name,location
	$script:lastUpdated = (Get-Item $outputCSV).LastWriteTime
	$dt = New-Object System.Data.DataTable
	$nameColumn = $dt.columns.add('Name')
	$locationColumn = $dt.columns.add('Location')
	foreach ($i in $csv) {
		$r = $dt.NewRow()
		$r.$locationColumn = $i.location
		$r.$nameColumn = $i.name
		$dt.Rows.Add($r)
	}
	Add-Type -AssemblyName System.Windows.Forms
	[System.Windows.Forms.Application]::EnableVisualStyles()

	$FormStartWidth = 860

	$Form                            = New-Object system.Windows.Forms.Form
	$Form.ClientSize                 = New-Object System.Drawing.Point($FormStartWidth,400)
	$Form.text                       = "Do We Have It?"
	$Icon                            = New-Object system.drawing.icon ("$scriptDir/Help.256.ico")
	$Form.Icon                       = $Icon
	$Form.TopMost                    = $false
	$form.MinimumSize = New-Object System.Drawing.Size(710, 250)

	$FilterTextBox                   = New-Object system.Windows.Forms.TextBox
	$FilterTextBox.multiline         = $false
	$FilterTextBox.width             = 334
	$FilterTextBox.height            = 20
	$FilterTextBox.location          = New-Object System.Drawing.Point(50,5)
	$FilterTextBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

	$FilterBoxLabel                  = New-Object system.Windows.Forms.Label
	$FilterBoxLabel.text             = "Filter:"
	$FilterBoxLabel.AutoSize         = $true
	$FilterBoxLabel.width            = 25
	$FilterBoxLabel.height           = 10
	$FilterBoxLabel.location         = New-Object System.Drawing.Point(10,8)
	$FilterBoxLabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

	$InfoLabel                  = New-Object system.Windows.Forms.Label
	$InfoLabel.text             = $(get-date (get-item $outputCSV).lastwritetime -format "MM-dd-yyyy HH:mm")
	$InfoLabel.AutoSize         = $true
	$InfoLabel.width            = 200
	$InfoLabel.height           = 10
	$InfoLabel.location         = New-Object System.Drawing.Point(($Form.Width - 228),9)
	$InfoLabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
	$InfoLabel.TextAlign        = [System.Drawing.ContentAlignment]::MiddleRight
	$InfoLabel.Anchor           = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
	$InfoLabel.Cursor           = [System.Windows.Forms.Cursors]::Hand

	function Update-InfoLabelPosition {
		param(
			[System.Windows.Forms.Label] $label,
			[System.Windows.Forms.Form] $form
		)
		try {
			if ($null -eq $label -or $null -eq $form) { return }
			$pref = $label.PreferredSize
			$label.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 20 - $pref.Width), $label.Location.Y)
		} catch {}
	}
	# Initial alignment 20px from the right edge
	Update-InfoLabelPosition -label $InfoLabel -form $Form

	function Set-InfoLabel {
		param(
			[System.Windows.Forms.Label] $label,
			[int] $rowCount,
			[datetime] $lastUpdated,
			[System.Windows.Forms.Form] $form
		)
		if ($null -eq $label) { return }
		$label.Text = "$rowCount Rows - " + (Get-Date $lastUpdated -Format "MM-dd-yyyy HH:mm")
		if ($null -ne $form) {
			Update-InfoLabelPosition -label $label -form $form
		}
	}

	function Show-MakeListLogDialog {
		param(
			[string]$LogPath,
			[System.Windows.Forms.Form]$Owner
		)

		if (-not (Test-Path -LiteralPath $LogPath)) {
			[System.Windows.Forms.MessageBox]::Show(
				"Could not find makelist csv: $LogPath",
				"Create List Log",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			) | Out-Null
			return
		}

		$logRows = Import-Csv -LiteralPath $LogPath
		$logTable = New-Object System.Data.DataTable
		$null = $logTable.Columns.Add('Source')
		$null = $logTable.Columns.Add('Rows')
		foreach ($logRow in $logRows) {
			$dtRow = $logTable.NewRow()
			$dtRow['Source'] = $logRow.Source
			$dtRow['Rows'] = $logRow.Rows
			$logTable.Rows.Add($dtRow)
		}

		$logForm = New-Object System.Windows.Forms.Form
		$logForm.Text = 'Create List Log'
		$logForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
		$logForm.ClientSize = New-Object System.Drawing.Size(700, 500)
		$logForm.MinimumSize = New-Object System.Drawing.Size(500, 300)
		$logForm.MaximizeBox = $true
		$logForm.MinimizeBox = $false
		$logForm.ShowInTaskbar = $false
		$logForm.Icon = $Icon

		$logGrid = New-Object System.Windows.Forms.DataGridView
		$logGrid.Location = New-Object System.Drawing.Point(12,12)
		$logGrid.Size = New-Object System.Drawing.Size(676,442)
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
		$logGrid.AllowUserToAddRows = $false
		$logGrid.AllowUserToDeleteRows = $false
		$logGrid.AllowUserToResizeRows = $false
		$logGrid.AllowUserToResizeColumns = $false
		$logGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
		$logGrid.MultiSelect = $false
		$logGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
		$logGrid.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
		$logGrid.DefaultCellStyle.SelectionBackColor = $logGrid.DefaultCellStyle.BackColor
		$logGrid.DefaultCellStyle.SelectionForeColor = $logGrid.DefaultCellStyle.ForeColor
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
		$okButton.Size = New-Object System.Drawing.Size(90,28)
		$okButton.Location = New-Object System.Drawing.Point(598,462)
		$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

		$logForm.AcceptButton = $okButton
		$logForm.CancelButton = $okButton
		$logForm.Controls.AddRange(@($logGrid,$okButton))

		if ($null -ne $Owner) {
			[void]$logForm.ShowDialog($Owner)
		} else {
			[void]$logForm.ShowDialog()
		}
	}

	function Get-InfoLabelTimestamp {
		param(
			[string]$LabelText
		)

		if ([string]::IsNullOrWhiteSpace($LabelText) -or $LabelText -notmatch ' - ') {
			return $null
		}

		$timestampText = ($LabelText -split ' - ', 2)[1].Trim()
		try {
			return [datetime]::ParseExact($timestampText, 'MM-dd-yyyy HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
		} catch {
			return $null
		}
	}

	function Get-MakeListTimestamp {
		param(
			[string]$LogPath
		)

		if (-not (Test-Path -LiteralPath $LogPath)) {
			return $null
		}

		try {
			return (Get-Item -LiteralPath $LogPath).LastWriteTime
		} catch {
			return $null
		}
	}

	function Get-MinuteStamp {
		param(
			[datetime]$DateValue
		)

		return Get-Date $DateValue -Format 'yyyy-MM-dd HH:mm'
	}

	function Get-ControlCenterY {
		param([System.Windows.Forms.Control]$ctl)
		return [int][Math]::Round($ctl.Location.Y + ($ctl.Height / 2.0))
	}

	function Align-TopRowControls {
		param(
			[System.Windows.Forms.Control]$lbl,
			[System.Windows.Forms.Control]$txt,
			[System.Windows.Forms.Control]$btnPaste,
			[System.Windows.Forms.Control]$btnRefresh,
			[System.Windows.Forms.Control]$info
		)
		$centers = @(
			(Get-ControlCenterY -ctl $lbl),
			(Get-ControlCenterY -ctl $txt),
			(Get-ControlCenterY -ctl $btnPaste),
			(Get-ControlCenterY -ctl $btnRefresh),
			(Get-ControlCenterY -ctl $info)
		)
		$avg = (($centers | Measure-Object -Average).Average)
		$centerY = [int][Math]::Round($avg)
		foreach ($c in @($lbl,$txt,$btnPaste,$btnRefresh,$info)) {
			$h = $c.Height
			$newY = [int]($centerY - [Math]::Round($h/2.0))
			$c.Location = New-Object System.Drawing.Point($c.Location.X, $newY)
		}
	}

    # Initial align will be performed after all controls are created/added

	$DataGridView1                   = New-Object system.Windows.Forms.DataGridView
	$DataGridView1.width             = 820  #Form - 56
	$DataGridView1.height            = 340  #form - 99
	$DataGridView1.location          = New-Object System.Drawing.Point(20,50)
	$DataGridView1.DataSource = $dt
	$DataGridView1.AutoSizeColumnsMode = "Fill"
	$DataGridView1.RowHeadersVisible = $false
	$DataGridView1.ReadOnly = $true
	$DataGridView1.AllowUserToResizeColumns = $false
	$DataGridView1.AllowUserToResizeRows = $false
	$DataGridView1.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
	$DataGridView1.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
	# Enable copying from the grid via Ctrl+C and context menu
	$DataGridView1.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableWithoutHeaderText

	# Add a context menu with Copy action
	$gridContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
	$miCopy = New-Object System.Windows.Forms.ToolStripMenuItem 'Copy'
	$miCopyPath = New-Object System.Windows.Forms.ToolStripMenuItem 'Copy Path'
	$miSearchGoogle = New-Object System.Windows.Forms.ToolStripMenuItem 'Search with Google'
	$null = $gridContextMenu.Items.Add($miCopy)
	$null = $gridContextMenu.Items.Add($miCopyPath)
	$null = $gridContextMenu.Items.Add($miSearchGoogle)
	$DataGridView1.ContextMenuStrip = $gridContextMenu

	# Handle Copy menu click
	$miCopy.Add_Click({
		try {
			if ($DataGridView1.GetCellCount([System.Windows.Forms.DataGridViewElementStates]::Selected) -gt 0) {
				$dataObj = $DataGridView1.GetClipboardContent()
				if ($null -ne $dataObj) {
					[System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)
				}
			} elseif ($null -ne $DataGridView1.CurrentCell -and $null -ne $DataGridView1.CurrentCell.Value) {
				[System.Windows.Forms.Clipboard]::SetText([string]$DataGridView1.CurrentCell.Value)
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				"Could not copy selection to the clipboard.",
				"Copy",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	# Ensure right-click selects the row under the cursor
	$DataGridView1.Add_CellMouseDown({
		param($incomingSender,$e)
		if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
			$DataGridView1.CurrentCell = $DataGridView1.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
			$DataGridView1.Rows[$e.RowIndex].Selected = $true
		}
	})

	# Helper to resolve the full path for an item based on Name/Location
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
		if ($location -in @('DVD Profiler','Disc List')) {
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
			$sub = $location -replace '^SwitchRoms[\\/]*',''
			$folder = Join-Path $SwitchDir $sub
			$fileToSelect = Join-Path $folder $name
		} else {
			# Assume relative to media root (e.g., TV Shows\Show\Season 1)
			$folder = Join-Path $mediaDir $location
			$fileToSelect = Join-Path $folder $name
		}
		$itemPath = if ($fileToSelect) { $fileToSelect } else { $folder }
		[pscustomobject]@{ Explorable = $explorable; Folder = $folder; File = $fileToSelect; ItemPath = $itemPath }
	}

	# Open location on double-click
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
				"Open Location",
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
					"Open Location",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Warning
				) | Out-Null
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				"Could not open Explorer.",
				"Open Location",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	# Handle Copy Path menu click
	$miCopyPath.Add_Click({
		try {
			if ($DataGridView1.CurrentCell -and $DataGridView1.CurrentCell.RowIndex -ge 0) {
				$row = $DataGridView1.Rows[$DataGridView1.CurrentCell.RowIndex]
				$name = [string]$row.Cells['Name'].Value
				$location = [string]$row.Cells['Location'].Value
				$res = Get-ResolvedItemPath -name $name -location $location
				if (-not $res.Explorable -or [string]::IsNullOrEmpty($res.ItemPath)) {
					[System.Windows.Forms.MessageBox]::Show(
						"Cannot determine a filesystem path for '$location'.",
						"Copy Path",
						[System.Windows.Forms.MessageBoxButtons]::OK,
						[System.Windows.Forms.MessageBoxIcon]::Information
					) | Out-Null
					return
				}
				[System.Windows.Forms.Clipboard]::SetText($res.ItemPath)
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				"Could not copy the path to the clipboard.",
				"Copy Path",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	# Handle Search with Google menu click
	$miSearchGoogle.Add_Click({
		try {
			if ($DataGridView1.CurrentCell -and $DataGridView1.CurrentCell.RowIndex -ge 0) {
				$row = $DataGridView1.Rows[$DataGridView1.CurrentCell.RowIndex]
				$name = [string]$row.Cells['Name'].Value
				if (-not [string]::IsNullOrWhiteSpace($name)) {
					$q = [System.Uri]::EscapeDataString($name)
					$url = "https://www.google.com/search?q=$q"
					Start-Process $url | Out-Null
				}
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show(
				"Could not open browser for Google search.",
				"Search with Google",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

	# Handle Ctrl+C in the grid to copy selection
	$DataGridView1.Add_KeyDown({
		param($incomingSender,$e)
		if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
			try {
				if ($DataGridView1.GetCellCount([System.Windows.Forms.DataGridViewElementStates]::Selected) -gt 0) {
					$dataObj = $DataGridView1.GetClipboardContent()
					if ($null -ne $dataObj) {
						[System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)
					}
				} elseif ($null -ne $DataGridView1.CurrentCell -and $null -ne $DataGridView1.CurrentCell.Value) {
					[System.Windows.Forms.Clipboard]::SetText([string]$DataGridView1.CurrentCell.Value)
				}
			} catch {
				[System.Windows.Forms.MessageBox]::Show(
					"Could not copy selection to the clipboard.",
					"Copy",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
			}
			$e.Handled = $true
		}
	})

	$PasteButton                         = New-Object system.Windows.Forms.Button
	$PasteButton.text                    = "Paste"
	$PasteButton.width                   = 78
	$PasteButton.height                  = 21
	$PasteButton.location                = New-Object System.Drawing.Point(392,6)
	$PasteButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

	$RefreshButton                         = New-Object system.Windows.Forms.Button
	$RefreshButton.text                    = "Refresh"
	$RefreshButton.width                   = 78
	$RefreshButton.height                  = 21
	$RefreshButton.location                = New-Object System.Drawing.Point(477,6)
	$RefreshButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

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
		$sb.ToString()
	}

	$FilterTextBox.Add_TextChanged({
		$dv = New-Object System.Data.DataView($dt)
		$terms = $FilterTextBox.Text -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
		if ($terms.Count -eq 0) {
			$dv.RowFilter = $null
		} else {
			$escapedTerms = $terms | ForEach-Object { Escape-ForRowFilter $_ }
			$dv.RowFilter = ($escapedTerms | ForEach-Object { "(Name LIKE '%$_%' OR Location LIKE '%$_%')" }) -join ' AND '
		}
		$DataGridView1.DataSource = $dv
		Set-InfoLabel -label $InfoLabel -rowCount $dv.Count -lastUpdated $script:lastUpdated -form $Form
	})

	# Set initial info label text
	Set-InfoLabel -label $InfoLabel -rowCount $dt.Rows.Count -lastUpdated $script:lastUpdated -form $Form

	function Refresh-GridData {
		$FilterTextBox.Text = ''
		$InfoLabel.text = '                Updating     '
		Start-sleep -Milliseconds 100
		$dt.clear()

		$csv = Import-CSV  $outputCSV | Select-Object name,location
		$script:lastUpdated = (Get-Item $outputCSV).LastWriteTime
		foreach ($i in $csv) {
			$r = $dt.NewRow()
			$r.$locationColumn = $i.location
			$r.$nameColumn = $i.name
			$dt.Rows.Add($r)
		}
		$DataGridView1.DataSource = $null
		$DataGridView1.DataSource = $dt
		Set-InfoLabel -label $InfoLabel -rowCount $dt.Rows.Count -lastUpdated $script:lastUpdated -form $Form
		Align-TopRowControls -lbl $FilterBoxLabel -txt $FilterTextBox -btnPaste $PasteButton -btnRefresh $RefreshButton -info $InfoLabel
		Update-InfoLabelPosition -label $InfoLabel -form $Form
		$FilterTextBox.Focus()
	}

	$InfoLabel.Add_Click({
		$infoTimestamp = Get-InfoLabelTimestamp -LabelText $InfoLabel.Text
		$makeListTimestamp = Get-MakeListTimestamp -LogPath $makeListLog
		if ($null -ne $infoTimestamp -and $null -ne $makeListTimestamp) {
			if ((Get-MinuteStamp -DateValue $infoTimestamp) -ne (Get-MinuteStamp -DateValue $makeListTimestamp)) {
				Refresh-GridData
			}
		}
		Show-MakeListLogDialog -LogPath $makeListLog -Owner $Form
	})

	$RefreshButton.Add_Click(
	{
		Refresh-GridData
	})

	# PasteButton click
$PasteButton.Add_Click({
    try {
        if (-not [System.Windows.Forms.Clipboard]::ContainsText()) {
            [System.Windows.Forms.MessageBox]::Show(
                "Clipboard does not contain text.",
                "Paste",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $text = [System.Windows.Forms.Clipboard]::GetText().Trim()

        if ([string]::IsNullOrWhiteSpace($text)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Clipboard text is empty.",
                "Paste",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        if ($text.Length -gt 50) {
            [System.Windows.Forms.MessageBox]::Show(
                "Clipboard text is longer than 50 characters (limit: 50).",
                "Paste",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        $FilterTextBox.Text = $text
        $FilterTextBox.Focus()
        $FilterTextBox.SelectionStart = $FilterTextBox.Text.Length
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not read the clipboard.",
            "Paste",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})


	$Form.Add_Resize(
	{
		$newWidth = $this.Width
        $newHeight = $this.Height
		$DataGridView1.width = $newWidth - 56
		$DataGridView1.height = $newHeight - 99
		Align-TopRowControls -lbl $FilterBoxLabel -txt $FilterTextBox -btnPaste $PasteButton -btnRefresh $RefreshButton -info $InfoLabel
		Update-InfoLabelPosition -label $InfoLabel -form $this
		#Write-Host $form.width $form.height
	} )

    $Form.controls.AddRange(@($FilterTextBox,$PasteButton,$RefreshButton,$FilterBoxLabel,$InfoLabel,$DataGridView1))
    # Align once the form is shown to ensure final sizes/positions
    $Form.Add_Shown({
        Align-TopRowControls -lbl $FilterBoxLabel -txt $FilterTextBox -btnPaste $PasteButton -btnRefresh $RefreshButton -info $InfoLabel
        Update-InfoLabelPosition -label $InfoLabel -form $Form
    })
    #Write-Host $form.width $form.height
    [void]$Form.ShowDialog()

	#$list |  Select-Object  name,location |  out-gridview -OutputMode Single
}
