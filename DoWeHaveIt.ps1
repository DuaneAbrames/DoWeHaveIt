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
	Write-Host 'Checking directories'
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
	Write-Host ' Disc List (Import)'
	$dl = import-csv "$($discListDir)disc list.csv"
	foreach ($i in $dl) {
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location='Disc List'}
		}
	}
	
	Write-Host ' DVD Profiler (Import)'
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
		Write-Host "  $d (Episodes)"
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
		Write-Host "  $d (Switch)"
		foreach ($i in (gci "$SwitchDir$d" -Recurse -file)) {
			if ($ExcludedThings -notcontains $i.name) {
				$output += new-object PSObject -property @{name=$i.name;location="SwitchRoms\$($i.DirectoryName.Replace($SwitchDir,''))"}
			}
		}
	}
	Write-Host "  Completed (RAR Files)"
	foreach ($i in (gci "$CompletedDir" -File '*.rar')) {
		
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="Completed"}
		}
	}
	Write-Host "  SteamUnlocked"
	foreach ($i in (gci "$SteamUnlockedDir" -File)) {
		if ($ExcludedThings -notcontains $i.name) {
			$output += new-object PSObject -property @{name=$i.name;location="SteamUnlocked"}
		}
	}
	Write-Host "  GoG"
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
	$Icon                            = New-Object system.drawing.icon ("$scriptDir\Help.256.ico")
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
	$InfoLabel.location         = New-Object System.Drawing.Point(650,8)
	$InfoLabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

	$DataGridView1                   = New-Object system.Windows.Forms.DataGridView
	$DataGridView1.width             = 820  #Form - 56
	$DataGridView1.height            = 340  #form - 99
	$DataGridView1.location          = New-Object System.Drawing.Point(20,50)
	$DataGridView1.DataSource = $dt
	$DataGridView1.AutoSizeColumnsMode = "Fill" 
	$DataGridView1.RowHeadersVisible = $false
	$DataGridView1.ReadOnly = $true
	
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
			$DV.RowFilter = "Name LIKE '*$($FilterTextBox.text)*'"
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
		$csv += Import-CSV  "$($scriptDir)DoWeHaveEpisodes.csv" | select name,location
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
	
	$PasteButton.Add_Click(
	{
		
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