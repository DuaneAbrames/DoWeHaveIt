[CmdletBinding()]
Param(
    [switch]$Music = $false
)

. (Join-Path $PSScriptRoot 'DoWeHaveIt-Common.ps1')

if ($PSVersionTable.Platform -eq 'Unix') {
    renice 10 -p $PID
}

function Write-CreateListMessage {
	param(
		[string]$Message
	)
	Write-Host $Message
}

function Get-LatestSuccessfulRunAge {
	param(
		[string]$DatasetType
	)

	$meta = Get-CurrentPublishedRunMetadata -DatasetType $DatasetType
	if (-not $meta -or -not $meta.LastSuccessfulRunAt) {
		return $null
	}

	return (Get-Date) - $meta.LastSuccessfulRunAt
}

function Test-MusicScanNeeded {
	$musicAge = Get-LatestSuccessfulRunAge -DatasetType $MusicDatasetType
	if ($null -eq $musicAge) {
		return $true
	}

	return $musicAge.TotalHours -gt 24
}

function Get-InventorySnapshot {
	Write-CreateListMessage 'Checking directories under media:'
	$output = @()
	$makeListRows = New-Object System.Collections.Generic.List[object]

	function Add-SourceRowCount {
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

	$dirsToInclude = @('Movies', 'TV Shows')
	foreach ($d in $dirsToInclude) {
		$startCount = $output.Count
		foreach ($i in (Get-ChildItem "$mediaDir$d" -Directory)) {
			if ($ExcludedThings -notcontains $i.Name) {
				$output += New-Object PSObject -Property @{ name = $i.Name; location = $d }
			}
		}
		Add-SourceRowCount -Label $d -StartCount $startCount -OutputItems $output
	}

	$startCount = $output.Count
	$dl = Import-Csv "$discListDir/disc list.csv"
	$firstColumn = ($dl[0].PSObject.Properties | Select-Object -First 1).Name
	$discTitles = $dl | Select-Object -ExpandProperty $firstColumn
	foreach ($i in $discTitles) {
		if ($ExcludedThings -notcontains $i) {
			$output += New-Object PSObject -Property @{ name = $i; location = 'Disc List' }
		}
	}
	Add-SourceRowCount -Label 'Disc List' -StartCount $startCount -OutputItems $output

	$startCount = $output.Count
	[xml]$dvdProfiler = Get-Content "$discListDir/Collection.xml"
	foreach ($i in $dvdProfiler.Collection.DVD.title) {
		if ($ExcludedThings -notcontains $i) {
			$output += New-Object PSObject -Property @{ name = $i; location = 'DVD Profiler' }
		}
	}
	Add-SourceRowCount -Label 'DVD Profiler' -StartCount $startCount -OutputItems $output

	$VideoExtensions = @('.mkv', '.mp4', '.avi', '.mov', '.m4v', '.wmv', '.ts', '.m2ts', '.mpg', '.mpeg', '.flv', '.webm')
	foreach ($d in @('TV Shows')) {
		$startCount = $output.Count
		foreach ($i in (Get-ChildItem "$mediaDir$d" -Recurse -File | Where-Object { $VideoExtensions -contains $_.Extension })) {
			$dir = Split-Path -Leaf (Split-Path $i.FullName)
			if ($ExcludedThings -notcontains $i.Name -and ($dir -ilike 'Season *' -or $dir -ieq 'Specials')) {
				$output += New-Object PSObject -Property @{ name = $i.Name; location = $i.DirectoryName.Replace($mediaDir, '') }
			}
		}
		Add-SourceRowCount -Label "$d (Episodes)" -StartCount $startCount -OutputItems $output
	}

	foreach ($d in @('Base', 'Update', 'DLC')) {
		$startCount = $output.Count
		foreach ($i in (Get-ChildItem "$SwitchDir$d" -Recurse -File)) {
			if ($ExcludedThings -notcontains $i.Name) {
				$output += New-Object PSObject -Property @{ name = $i.Name; location = "SwitchRoms/$($i.DirectoryName.Replace($SwitchDir, ''))" }
			}
		}
		Add-SourceRowCount -Label "$d (Switch)" -StartCount $startCount -OutputItems $output
	}

	foreach ($pattern in @(
		@{ Filter = '*.rar'; Label = 'Completed (RAR Files)' },
		@{ Filter = '*.zip'; Label = 'Completed (ZIP Files)' },
		@{ Filter = '*.iso'; Label = 'Completed (ISO Files)' }
	)) {
		$startCount = $output.Count
		foreach ($i in (Get-ChildItem $CompletedDir -File $pattern.Filter)) {
			if ($ExcludedThings -notcontains $i.Name) {
				$output += New-Object PSObject -Property @{ name = $i.Name; location = 'Completed' }
			}
		}
		Add-SourceRowCount -Label $pattern.Label -StartCount $startCount -OutputItems $output
	}

	$startCount = $output.Count
	foreach ($i in (Get-ChildItem $SteamUnlockedDir -File)) {
		if ($ExcludedThings -notcontains $i.Name) {
			$output += New-Object PSObject -Property @{ name = $i.Name; location = 'SteamUnlocked' }
		}
	}
	Add-SourceRowCount -Label 'SteamUnlocked' -StartCount $startCount -OutputItems $output

	$startCount = $output.Count
	foreach ($i in (Get-ChildItem $GogDir -Directory -Exclude 'Logs')) {
		if ($ExcludedThings -notcontains $i.Name) {
			$output += New-Object PSObject -Property @{ name = $i.Name; location = 'GoG' }
		}
	}
	Add-SourceRowCount -Label 'GoG' -StartCount $startCount -OutputItems $output

	$output = $output |
		Where-Object { $_.name -notlike '' } |
		Sort-Object name, location |
		Select-Object name, location

	$makeListRows.Add([pscustomobject]@{
		Source = 'TOTAL'
		Rows = $output.Count
	}) | Out-Null

	Write-CreateListMessage "$($output.Count) rows exported."

	$itemsArray = @($output)
	$sourcesArray = @($makeListRows.ToArray())
	return (New-Object PSObject -Property ([ordered]@{
		Items = $itemsArray
		Sources = $sourcesArray
	}))
}

function Get-MusicInventorySnapshot {
	$musicRoot = Join-Path $mediaDir $MusicDirName
	if (-not (Test-Path -LiteralPath $musicRoot)) {
		throw "Music directory not found: $musicRoot"
	}

	Write-CreateListMessage "Checking music files under ${MusicDirName}:"
	$output = @()
	foreach ($file in (Get-ChildItem $musicRoot -Recurse -File | Where-Object { $MusicExtensions -contains $_.Extension.ToLowerInvariant() })) {
		if ($ExcludedThings -contains $file.Name) {
			continue
		}

		$relativeDirectory = $file.DirectoryName.Replace($mediaDir, '').TrimStart('\', '/')
		if ([string]::IsNullOrWhiteSpace($relativeDirectory)) {
			$relativeDirectory = $MusicDirName
		}

		$output += New-Object PSObject -Property @{
			name = $file.Name
			location = ($relativeDirectory -replace '\\', '/')
		}
	}

	$output = $output |
		Sort-Object name, location |
		Select-Object name, location

	$sourceCount = $output.Count
	Write-CreateListMessage "Music Files - $sourceCount"
	Write-CreateListMessage "$sourceCount music rows exported."

	$sourcesArray = @(
		[pscustomobject]@{ Source = 'Music Files'; Rows = $sourceCount },
		[pscustomobject]@{ Source = 'TOTAL'; Rows = $sourceCount }
	)

	return (New-Object PSObject -Property ([ordered]@{
		Items = @($output)
		Sources = $sourcesArray
	}))
}

function Start-WorkerMode {
	Initialize-AppDatabase

	if ($Music) {
		$script:activeDatasetType = $MusicDatasetType
		$musicSnapshot = Get-MusicInventorySnapshot
		$null = Publish-InventorySnapshot -Items $musicSnapshot.Items -Sources $musicSnapshot.Sources -DatasetType $MusicDatasetType
		return
	}

	$script:activeDatasetType = $RegularDatasetType
	$regularSnapshot = Get-InventorySnapshot
	$null = Publish-InventorySnapshot -Items $regularSnapshot.Items -Sources $regularSnapshot.Sources -DatasetType $RegularDatasetType

	if (Test-MusicScanNeeded) {
		Write-CreateListMessage 'Latest music scan is older than 24 hours. Running music snapshot.'
		$musicSnapshot = Get-MusicInventorySnapshot
		$null = Publish-InventorySnapshot -Items $musicSnapshot.Items -Sources $musicSnapshot.Sources -DatasetType $MusicDatasetType
	} else {
		Write-CreateListMessage 'Skipping music snapshot. Latest music scan is less than 24 hours old.'
	}
}

try {
    Start-WorkerMode
} finally {
    Close-AppConnection
}
