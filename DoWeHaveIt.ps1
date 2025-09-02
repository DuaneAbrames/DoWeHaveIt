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
$ExcludedThings = ('___Duplicates')
if ($createList) {
	Write-Host 'Checking directories under media:'
	$dirsToInclude = ('Movies','Troma','TV Shows','101 Horror Movies Mega Pack')
	$output = @()
	foreach ($d in $dirsToInclude) {
		Write-Host "  $d"
		foreach ($i in (gci "$mediaDir$d" -Directory)) {
			#Write-Host "    - $i
			if ($ExcludedThings -notcontains $i.name) {
				$output += new-object PSObject -property @{name=$i.name;location=$d}
			}
		}
	}
	Write-Host 'Disc List (Import)'
	$dl = import-csv "$($discListDir)disc list.csv"
	foreach ($i in $dl) {
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location='Disc List'}
		}
	}
	
	Write-Host 'DVD Profiler (Import)'
	[xml]$dvdProfiler = get-content "$($discListDir)Collection.xml"
	foreach ($i in $dvdProfiler.Collection.DVD.title) {
		if ($ExcludedThings -notcontains $i) {
			$output += new-object PSObject -property @{name=$i;location='DVD Profiler'}
		}
	}
	## Episodes Section ##
	$VideoExtensions = @('.mkv','.mp4','.avi','.mov','.m4v','.wmv','.ts','.m2ts','.mpg','.mpeg','.flv','.webm')
	$dirsToInclude = ('TV Shows')
	foreach ($d in $dirsToInclude) {
		Write-Host "$d (Episodes)"
		foreach ($i in (gci "$mediaDir$d" -Recurse -file | ?{$VideoExtensions -contains $_.extension})) {
			$dir = Split-Path -leaf (Split-Path $i.fullname)
			if ($ExcludedThings -notcontains $i.name -and ($dir -ilike "Season *" -or $dir -ieq "Specials")) {
				$output += new-object PSObject -property @{name=$i.name;location=$i.DirectoryName.Replace($mediaDir,'')}
			}
		}
	}
	#Here will Be GoG, Switch, Etc.
	$dirsToInclude = ('Base','Update','DLC')
	foreach ($d in $dirsToInclude) {
		Write-Host "$d (Switch)"
		foreach ($i in (gci "$SwitchDir$d" -Recurse -file)) {
			if ($ExcludedThings -notcontains $i.name) {
				$output += new-object PSObject -property @{name=$i.name;location="SwitchRoms/$($i.DirectoryName.Replace($SwitchDir,''))"}
			}
		}
	}
	Write-Host "Completed (RAR Files)"
	foreach ($i in (gci "$CompletedDir" -File '*.rar')) {
		
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="Completed"}
		}
	}
	Write-Host "Completed (ZIP Files)"
	foreach ($i in (gci "$CompletedDir" -File '*.zip')) {
		
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="Completed"}
		}
	}
	Write-Host "Completed (ISO Files)"
	foreach ($i in (gci "$CompletedDir" -File '*.iso')) {
		
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="Completed"}
		}
	}
	Write-Host "SteamUnlocked"
	foreach ($i in (gci "$SteamUnlockedDir" -File)) {
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="SteamUnlocked"}
		}
	}
	Write-Host "GoG"
	foreach ($i in (gci "$GogDir" -Directory -Exclude "Logs")) {
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="GoG"}
		}
	}
	$output = $output | sort-object name,location | select name,location 
	$output | Export-csv -notypeinformation "$($scriptDir)DoWeHaveIt.csv"
	Write-Host "$($output.count) rows exported."
} else {
	$csv = Import-CSV  "$($scriptDir)DoWeHaveIt.csv" | select name,location 
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

	$Form                            = New-Object system.Windows.Forms.Form
	$Form.ClientSize                 = New-Object System.Drawing.Point(860,400)
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
	$InfoLabel.text             = "List updated at: $(get-date (get-item S:\scripts\DoWeHaveIt.csv).lastwritetime -format "dd-MM-yyyy HH:mm")"
	$InfoLabel.AutoSize         = $true
	$InfoLabel.width            = 200
	$InfoLabel.height           = 10
	$InfoLabel.location         = New-Object System.Drawing.Point(550,8)
	$InfoLabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

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
		param($sender,$e)
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
		param($sender, $e)
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
				Start-Process -FilePath 'explorer.exe' -ArgumentList ("/select,`"$res.File`"")
			} elseif ($res.Folder -and (Test-Path -LiteralPath $res.Folder)) {
				Start-Process -FilePath 'explorer.exe' -ArgumentList ("`"$res.Folder`"")
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
		param($sender,$e)
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
	
	$FilterTextBox.Add_TextChanged( 
		{
			$dv = New-Object System.Data.DataView($dt)
			# Escape special characters for DataView RowFilter LIKE patterns safely
			$search = [string]$FilterTextBox.Text
			$sb = New-Object System.Text.StringBuilder
			foreach ($ch in $search.ToCharArray()) {
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
			$escaped = $sb.ToString()
			$DV.RowFilter = "Name LIKE '*$escaped*'"
			$DataGridView1.DataSource = $dv
		}
	)

	$RefreshButton.Add_Click(
	{
		#Write-Host "Click!"
		$FilterTextBox.Text = ''
		$InfoLabel.text             = '                Updating     ' 
		Start-sleep -Milliseconds 100
		#write-host $csv.count
		$dt.clear()
		
		$csv = Import-CSV  "$($scriptDir)DoWeHaveIt.csv" | select name,location
		#$csv = $csv | sort-object name
		foreach ($i in $csv) { 
			$r = $dt.NewRow()
			$r.$locationColumn = $i.location
			$r.$nameColumn = $i.name
			$dt.Rows.Add($r) 
		}
		$DataGridView1.DataSource = $null
		$DataGridView1.DataSource = $dt
		$InfoLabel.text             = "List updated at: $(get-date (get-item S:\scripts\DoWeHaveIt.csv).lastwritetime -format "dd-MM-yyyy HH:mm")"
		$FilterTextBox.Focus()
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
		$InfoLabel.location         = New-Object System.Drawing.Point(($newWidth - 226),11)
		#Write-Host $form.width $form.height
	} )
	
	$Form.controls.AddRange(@($FilterTextBox,$PasteButton,$RefreshButton,$FilterBoxLabel,$InfoLabel,$DataGridView1))
	#Write-Host $form.width $form.height
	[void]$Form.ShowDialog()
	
	#$list | select name,location |  out-gridview -OutputMode Single
}
