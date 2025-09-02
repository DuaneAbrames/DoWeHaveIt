[CmdletBinding()]
Param ([switch]$createList = $false
)
if ($PSVersionTable.Platform -eq "Unix") {
	$mediaDir = '/mnt/user/media/'
	$scriptDir = '/mnt/user/bigone/scripts/'
	$discListDir = '/mnt/user/bigone/DISC LIST/'
	$GogDir = '/mnt/user/bigone/Bittorrent/Completed/GoG'
	$SteamUnlokedDire = '/mnt/usr/bigone/BitTorrent/Completed/SteamUnlocked/'
	$SwitchDir = '/mnt/remotes/WHELP_SwitchRoms/'
	$createList = $true
} else {
    $mediaDir = 'R:/'
	$scriptDir = 'S:/scripts/'
	$discListDir = 'S:/DISC LIST/'
	$GogDir = 'S:/BitTorrent/Completed/Gog/'
	$SteamUnlokedDire = 'S:/BitTorrent/Completed/SteamUnlocked/'
	$SwitchDir = '\\whelp\SwitchRoms\'
}

if ($createList) {
	Write-Host 'Checking directories'
	$dirsToInclude = ('Movies','Troma','TV Shows','101 Horror Movies Mega Pack')
	$output = @()
	foreach ($d in $dirsToInclude) {
		Write-Host "  $d"
		foreach ($i in (gci "$mediaDir$d" -Directory)) {
			#Write-Host "    - $i"
			$output += new-object PSObject -property @{name=$i.name;location=$d}
		}
	}
	Write-Host 'Importing Disc List'
	$dl = import-csv "$($discListDir)disc list.csv"
	foreach ($i in $dl) {
		$output += new-object PSObject -property @{name=$i.name;location='Disc List'}
	}
	
	Write-Host 'Importing DVD Profiler'
	[xml]$dvdProfiler = get-content "$($discListDir)Collection.xml"
	foreach ($i in $dvdProfiler.Collection.DVD.title) {
		$output += new-object PSObject -property @{name=$i;location='DVD Profiler'}
	}
	$output = $output | sort-object name,location | select name,location 
	$output | Export-csv -notypeinformation "$($scriptDir)DoWeHaveIt.csv"
	Write-Host "$($output.count) rows exported."
	## New Episode code
	$dirsToInclude = ('TV Shows')
	$output2 = @()
	foreach ($d in $dirsToInclude) {
		Write-Host "  $d (Episodes)"
		foreach ($i in (gci "$mediaDir$d" -Recurse -file)) {
			$i.basename
			$output2 += new-object PSObject -property @{name=$i.name;location=$i.DirectoryName.Replace($mediaDir,'')}
		}
	}
	#Here will Be GoG, Switch, Etc.
	$dirsToInclude = ('Base','Update','DLC')
	foreach ($d in $dirsToInclude) {
		Write-Host "  $d (Episodes)"
		foreach ($i in (gci "$SwitchDir$d" -Recurse -file)) {
			$i.basename
			$output2 += new-object PSObject -property @{name=$i.name;location="SwitchRoms\$($i.DirectoryName.Replace($SwitchDir,''))"}
		}
	}
	foreach ($i in (gci "$GogDir" -Directory -Exclude "Logs")) {
		$i.basename
		$output2 += new-object PSObject -property @{name=$i.name;location="GoG\$($i.Name.Replace($GogDir,''))"}
	}
	$output2 | export-csv -NoTypeInformation "$($scriptDir)DoWeHaveEpisodes.csv"
	Write-Host "$($output2.count) rows exported."
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
	
	
	$EpisodeCheckBox = New-Object System.Windows.Forms.CheckBox
    $EpisodeCheckBox.Location = New-Object System.Drawing.Size(500, 8)
    $EpisodeCheckBox.Size = New-Object System.Drawing.Size(100, 20)
    $EpisodeCheckBox.Text = "Episodes?"
    $EpisodeCheckBox.Checked = $false
    	
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
	
	$Button1                         = New-Object system.Windows.Forms.Button
	$Button1.text                    = "Refresh"
	$Button1.width                   = 78
	$Button1.height                  = 21
	$Button1.location                = New-Object System.Drawing.Point(392,6)
	$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

    $EpisodeCheckBox.Add_CheckStateChanged(
	{
		#Write-Host "Click!"
		$FilterTextBox.Text = ''
		$InfoLabel.text             = '                Updating     ' 
		Start-sleep -Milliseconds 100
		#write-host $csv.count
		$dt.clear()
		$csv = Import-CSV  "$($scriptDir)DoWeHaveIt.csv" | select name,location
		if ($EpisodeCheckBox.Checked) {
			$csv += Import-CSV  "$($scriptDir)DoWeHaveEpisodes.csv" | select name,location
		}
		$csv = $csv | sort-object name
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
	
	$FilterTextBox.Add_TextChanged( 
		{
			$dv = New-Object System.Data.DataView($dt)
			$DV.RowFilter = "Name LIKE '*$($FilterTextBox.text)*'"
			$DataGridView1.DataSource = $dv
		}
	)

	$Button1.Add_Click(
	{
		#Write-Host "Click!"
		$FilterTextBox.Text = ''
		$InfoLabel.text             = '                Updating     ' 
		Start-sleep -Milliseconds 100
		#write-host $csv.count
		$dt.clear()
		
		$csv = Import-CSV  "$($scriptDir)DoWeHaveIt.csv" | select name,location
		if ($EpisodeCheckBox.Checked) {
			$csv += Import-CSV  "$($scriptDir)DoWeHaveEpisodes.csv" | select name,location
		}
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
	
	$Form.Add_Resize(
	{
		$newWidth = $this.Width
        $newHeight = $this.Height
		$DataGridView1.width = $newWidth - 56
		$DataGridView1.height = $newHeight - 99
		$InfoLabel.location         = New-Object System.Drawing.Point(($newWidth - 226),11)
		#Write-Host $form.width $form.height
	} )
	
	$Form.controls.AddRange(@($FilterTextBox,$Button1,$FilterBoxLabel,$EpisodeCheckBox,$InfoLabel,$DataGridView1))
	#Write-Host $form.width $form.height
	[void]$Form.ShowDialog()
	
	#$list | select name,location |  out-gridview -OutputMode Single
}