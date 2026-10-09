# Dot-source into the UI or worker script scope.
if ($PSVersionTable.Platform -eq 'Unix') {
	$mediaDir = '/mnt/user/media/'
	$scriptDir = '/mnt/user/bigone/scripts/DoWeHaveIt/'
	$discListDir = '/mnt/user/bigone/DISC LIST/'
	$GogDir = '/mnt/user/bigone/Bittorrent/Completed/GoG/'
	$CompletedDir = '/mnt/user/bigone/Bittorrent/Completed/'
	$SteamUnlockedDir = '/mnt/user/bigone/Bittorrent/Completed/SteamUnlocked/'
	$SwitchDir = '/mnt/remotes/WHELP_SwitchRoms/'
} else {
	$mediaDir = 'R:/'
	$scriptDir = 'S:/scripts/DoWeHaveIt/'
	$discListDir = 'S:/DISC LIST/'
	$GogDir = 'S:/BitTorrent/Completed/Gog/'
	$CompletedDir = 'S:/BitTorrent/Completed/'
	$SteamUnlockedDir = 'S:/BitTorrent/Completed/SteamUnlocked/'
	$SwitchDir = '\\whelp\SwitchRoms\'
}


$outputCSV = "$scriptDir/DoWeHaveIt.csv"
$makeListLog = "$scriptDir/makelist.csv"
$databaseInfoPath = "$scriptDir/database-info.txt"
$ExcludedThings = ('___Duplicates', '.trash')
$MusicDirName = 'mp3'
$MusicExtensions = @('.mp3', '.flac', '.m4a', '.aac', '.ogg', '.wav', '.wma', '.opus', '.ape', '.alac')
$RegularDatasetType = 'regular'
$MusicDatasetType = 'music'
$script:dbConfig = $null
$script:activeDatasetType = if ($Music) { $MusicDatasetType } else { $RegularDatasetType }

function Get-DatasetStateId {
	param(
		[string]$DatasetType
	)

	switch ($DatasetType) {
		$MusicDatasetType { return 2 }
		default { return 1 }
	}
}

function Get-NormalizedDatasetType {
	param(
		[string]$DatasetType
	)

	if ([string]::IsNullOrWhiteSpace($DatasetType)) {
		return $RegularDatasetType
	}

	switch ($DatasetType.Trim().ToLowerInvariant()) {
		'music' { return $MusicDatasetType }
		default { return $RegularDatasetType }
	}
}

function Get-DatasetDisplayName {
	param(
		[string]$DatasetType
	)

	switch (Get-NormalizedDatasetType $DatasetType) {
		$MusicDatasetType { return 'Music' }
		default { return 'Media' }
	}
}

function Escape-SqlString {
	param(
		[AllowNull()][string]$Value
	)

	if ($null -eq $Value) {
		return 'NULL'
	}

	return "'" + ($Value -replace '\\', '\\\\' -replace "'", "''") + "'"
}

function ConvertTo-MySqlDateTimeLiteral {
	param(
		$Value
	)

	if ($null -eq $Value) {
		return 'NULL'
	}

	if ($Value -is [System.Nullable[datetime]]) {
		if (-not $Value.HasValue) {
			return 'NULL'
		}
		return "'" + $Value.Value.ToString('yyyy-MM-dd HH:mm:ss') + "'"
	}

	return "'" + ([datetime]$Value).ToString('yyyy-MM-dd HH:mm:ss') + "'"
}

function ConvertFrom-DataTable {
	param(
		[Parameter(Mandatory = $true)]
		[object]$InputObject
	)

	if ($null -eq $InputObject) {
		return @()
	}

	if ($InputObject -is [System.Data.DataTable]) {
		$rows = @()
		foreach ($row in $InputObject.Rows) {
			$item = [ordered]@{}
			foreach ($column in $InputObject.Columns) {
				$item[$column.ColumnName] = $row[$column.ColumnName]
			}
			$rows += [pscustomobject]$item
		}
		return $rows
	}

	if ($InputObject -is [System.Data.DataRow]) {
		$item = [ordered]@{}
		foreach ($column in $InputObject.Table.Columns) {
			$item[$column.ColumnName] = $InputObject[$column.ColumnName]
		}
		return @([pscustomobject]$item)
	}

	return @($InputObject)
}

function Get-DatabaseConfig {
	if ($script:dbConfig) {
		return $script:dbConfig
	}

	if (-not (Test-Path -LiteralPath $databaseInfoPath)) {
		throw "Database config file not found: $databaseInfoPath"
	}

	$config = @{}
	foreach ($line in Get-Content -LiteralPath $databaseInfoPath) {
		if ($line -match '^\s*([^:]+):\s*(.*)\s*$') {
			$config[$matches[1].Trim()] = $matches[2].Trim()
		}
	}

	foreach ($requiredKey in @('host', 'port', 'database', 'db_user', 'db_pass')) {
		if (-not $config.ContainsKey($requiredKey) -or [string]::IsNullOrWhiteSpace($config[$requiredKey])) {
			throw "Missing database config key '$requiredKey' in $databaseInfoPath"
		}
	}

	$script:dbConfig = [pscustomobject]@{
		Host = $config['host']
		Port = [int]$config['port']
		Database = $config['database']
		User = $config['db_user']
		Password = $config['db_pass']
	}

	return $script:dbConfig
}

function Import-AppModules {
	if (-not (Get-Module -Name SimplySql)) {
		$module = Get-Module -ListAvailable SimplySql | Sort-Object Version -Descending | Select-Object -First 1
		if ($module) {
			Import-Module $module.Path -ErrorAction Stop
			return
		}

		$candidateRoots = @()
		foreach ($entry in ($env:PSModulePath -split [IO.Path]::PathSeparator)) {
			if (-not [string]::IsNullOrWhiteSpace($entry) -and (Test-Path -LiteralPath $entry)) {
				$candidateRoots += $entry
			}
		}

		$candidateManifest = Get-ChildItem -Path $candidateRoots -Filter 'SimplySql.psd1' -Recurse -ErrorAction SilentlyContinue |
			Sort-Object FullName -Descending |
			Select-Object -First 1

		if ($candidateManifest) {
			Import-Module $candidateManifest.FullName -ErrorAction Stop
			return
		}

		Import-Module SimplySql -ErrorAction Stop
	}
}

function Open-AppConnection {
	Import-AppModules
	$config = Get-DatabaseConfig
	$currentConnection = Get-SqlConnection -ErrorAction SilentlyContinue
	if ($currentConnection) {
		return
	}

	$credential = New-Object System.Management.Automation.PSCredential(
		$config.User,
		(ConvertTo-SecureString $config.Password -AsPlainText -Force)
	)

	Open-MySqlConnection -Server $config.Host -Port $config.Port -Database $config.Database -Credential $credential | Out-Null
}

function Close-AppConnection {
	try {
		Close-SqlConnection -ErrorAction SilentlyContinue | Out-Null
	} catch {}
}

function Invoke-AppQuery {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Sql
	)

	Open-AppConnection
	return (Invoke-SqlQuery -Query $Sql -ErrorAction Stop)
}

function Invoke-AppScalarValue {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Sql
	)

	Open-AppConnection
	return (Invoke-SqlScalar -Query $Sql -ErrorAction Stop)
}

function Invoke-AppUpdate {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Sql
	)

	Open-AppConnection
	Invoke-SqlUpdate -Query $Sql -ErrorAction Stop | Out-Null
}

function Start-AppTransaction {
	Open-AppConnection
	Start-SqlTransaction -ErrorAction Stop | Out-Null
}

function Complete-AppTransaction {
	Complete-SqlTransaction -ErrorAction Stop | Out-Null
}

function Undo-AppTransaction {
	try {
		Undo-SqlTransaction -ErrorAction SilentlyContinue | Out-Null
	} catch {}
}

function Test-ColumnExists {
	param(
		[string]$TableName,
		[string]$ColumnName
	)

	$config = Get-DatabaseConfig
	$result = Invoke-AppScalarValue @"
SELECT COUNT(*)
FROM information_schema.columns
WHERE table_schema = $(Escape-SqlString $config.Database)
	AND table_name = $(Escape-SqlString $TableName)
	AND column_name = $(Escape-SqlString $ColumnName);
"@

	return ([int]$result) -gt 0
}

function Test-IndexExists {
	param(
		[string]$TableName,
		[string]$IndexName
	)

	$config = Get-DatabaseConfig
	$result = Invoke-AppScalarValue @"
SELECT COUNT(*)
FROM information_schema.statistics
WHERE table_schema = $(Escape-SqlString $config.Database)
	AND table_name = $(Escape-SqlString $TableName)
	AND index_name = $(Escape-SqlString $IndexName);
"@

	return ([int]$result) -gt 0
}

function Initialize-DatabaseSchema {
	Invoke-AppUpdate @"
CREATE TABLE IF NOT EXISTS inventory_runs (
	run_id BIGINT NOT NULL AUTO_INCREMENT,
	dataset_type VARCHAR(32) NOT NULL DEFAULT 'regular',
	started_at DATETIME NOT NULL,
	completed_at DATETIME NULL,
	status VARCHAR(32) NOT NULL,
	total_rows INT NOT NULL DEFAULT 0,
	worker_host VARCHAR(255) NULL,
	error_message TEXT NULL,
	retention_class VARCHAR(32) NULL,
	PRIMARY KEY (run_id),
	KEY idx_inventory_runs_dataset (dataset_type),
	KEY idx_inventory_runs_completed (completed_at),
	KEY idx_inventory_runs_status (status)
);
"@

	if (-not (Test-ColumnExists -TableName 'inventory_runs' -ColumnName 'dataset_type')) {
		Invoke-AppUpdate "ALTER TABLE inventory_runs ADD COLUMN dataset_type VARCHAR(32) NOT NULL DEFAULT 'regular' AFTER run_id;"
	}
	if (-not (Test-IndexExists -TableName 'inventory_runs' -IndexName 'idx_inventory_runs_dataset')) {
		Invoke-AppUpdate "ALTER TABLE inventory_runs ADD INDEX idx_inventory_runs_dataset (dataset_type);"
	}

	Invoke-AppUpdate @"
CREATE TABLE IF NOT EXISTS inventory_run_sources (
	run_id BIGINT NOT NULL,
	source_name VARCHAR(255) NOT NULL,
	row_count INT NOT NULL,
	display_order INT NOT NULL,
	PRIMARY KEY (run_id, display_order),
	KEY idx_inventory_run_sources_run (run_id)
);
"@

	Invoke-AppUpdate @"
CREATE TABLE IF NOT EXISTS inventory_items (
	run_id BIGINT NOT NULL,
	name VARCHAR(1024) NOT NULL,
	location VARCHAR(1024) NOT NULL,
	KEY idx_inventory_items_run_name (run_id, name(255)),
	KEY idx_inventory_items_run_location (run_id, location(255))
);
"@

	Invoke-AppUpdate @"
CREATE TABLE IF NOT EXISTS inventory_state (
	state_id TINYINT NOT NULL,
	current_run_id BIGINT NULL,
	last_successful_run_at DATETIME NULL,
	last_row_count INT NOT NULL DEFAULT 0,
	PRIMARY KEY (state_id)
);
"@

	Invoke-AppUpdate @"
INSERT INTO inventory_state (state_id, current_run_id, last_successful_run_at, last_row_count)
SELECT 1, NULL, NULL, 0
WHERE NOT EXISTS (
	SELECT 1
	FROM inventory_state
	WHERE state_id = 1
);
"@

	Invoke-AppUpdate @"
INSERT INTO inventory_state (state_id, current_run_id, last_successful_run_at, last_row_count)
SELECT 2, NULL, NULL, 0
WHERE NOT EXISTS (
	SELECT 1
	FROM inventory_state
	WHERE state_id = 2
);
"@
}

function Get-SystemTimestamp {
	return Get-Date
}

function Get-CurrentPublishedRunMetadata {
	param(
		[string]$DatasetType = $script:activeDatasetType
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	$stateId = Get-DatasetStateId $normalizedDatasetType
	$rows = ConvertFrom-DataTable (Invoke-AppQuery @"
SELECT
	s.current_run_id,
	s.last_successful_run_at,
	s.last_row_count
FROM inventory_state s
WHERE s.state_id = $stateId;
"@)

	if ($rows.Count -eq 0) {
		return $null
	}

	$row = $rows[0]
	return [pscustomobject]@{
		DatasetType = $normalizedDatasetType
		RunId = if ($null -eq $row.current_run_id -or $row.current_run_id -eq [DBNull]::Value) { $null } else { [int64]$row.current_run_id }
		LastSuccessfulRunAt = if ($null -eq $row.last_successful_run_at -or $row.last_successful_run_at -eq [DBNull]::Value) { $null } else { [datetime]$row.last_successful_run_at }
		RowCount = if ($null -eq $row.last_row_count -or $row.last_row_count -eq [DBNull]::Value) { 0 } else { [int]$row.last_row_count }
	}
}

function Get-PreviousCompletedRunId {
	param(
		[int64]$CurrentRunId,
		[string]$DatasetType = $script:activeDatasetType
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	$whereClause = if ($CurrentRunId) { "AND run_id <> $CurrentRunId" } else { '' }
	$result = Invoke-AppScalarValue @"
SELECT run_id
FROM inventory_runs
WHERE status = 'completed'
	AND dataset_type = $(Escape-SqlString $normalizedDatasetType)
$whereClause
ORDER BY completed_at DESC, run_id DESC
LIMIT 1;
"@

	if ($null -eq $result -or $result -eq [DBNull]::Value -or [string]::IsNullOrWhiteSpace([string]$result)) {
		return $null
	}

	return [int64]$result
}

function Insert-InventoryItemsBatch {
	param(
		[int64]$RunId,
		[object[]]$Items
	)

	if ($Items.Count -eq 0) {
		return
	}

	$batchSize = 500
	for ($index = 0; $index -lt $Items.Count; $index += $batchSize) {
		$endIndex = [Math]::Min($index + $batchSize - 1, $Items.Count - 1)
		$values = New-Object System.Collections.Generic.List[string]
		for ($i = $index; $i -le $endIndex; $i++) {
			$item = $Items[$i]
			$values.Add("($RunId, $(Escape-SqlString ([string]$item.name)), $(Escape-SqlString ([string]$item.location)))") | Out-Null
		}

		Invoke-AppUpdate @"
INSERT INTO inventory_items (run_id, name, location)
VALUES
$($values -join ",`n");
"@
	}
}

function Insert-InventorySources {
	param(
		[int64]$RunId,
		[object[]]$Sources
	)

	if ($Sources.Count -eq 0) {
		return
	}

	$values = New-Object System.Collections.Generic.List[string]
	for ($i = 0; $i -lt $Sources.Count; $i++) {
		$source = $Sources[$i]
		$values.Add("($RunId, $(Escape-SqlString ([string]$source.Source)), $([int]$source.Rows), $($i + 1))") | Out-Null
	}

	Invoke-AppUpdate @"
INSERT INTO inventory_run_sources (run_id, source_name, row_count, display_order)
VALUES
$($values -join ",`n");
"@
}

function Get-KeeperRunIds {
	param(
		[string]$DatasetType = $script:activeDatasetType
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	$currentMeta = Get-CurrentPublishedRunMetadata -DatasetType $normalizedDatasetType
	$currentRunId = if ($currentMeta) { $currentMeta.RunId } else { $null }
	$previousRunId = Get-PreviousCompletedRunId -CurrentRunId $currentRunId -DatasetType $normalizedDatasetType

	$allRuns = ConvertFrom-DataTable (Invoke-AppQuery @"
SELECT
	run_id,
	dataset_type,
	completed_at,
	status
FROM inventory_runs
WHERE dataset_type = $(Escape-SqlString $normalizedDatasetType)
ORDER BY completed_at ASC, run_id ASC;
"@)

	$keepers = New-Object 'System.Collections.Generic.HashSet[long]'
	if ($currentRunId) { [void]$keepers.Add([int64]$currentRunId) }
	if ($previousRunId) { [void]$keepers.Add([int64]$previousRunId) }

	$now = Get-Date
	$dailyCutoff = $now.Date.AddDays(-14)
	$monthlyCutoff = $now.AddMonths(-2)
	$dailyBuckets = @{}
	$weeklyBuckets = @{}
	$monthlyBuckets = @{}

	foreach ($run in $allRuns) {
		$runId = [int64]$run.run_id
		$status = [string]$run.status
		if ($status -ne 'completed') {
			[void]$keepers.Add($runId)
			continue
		}

		if ($run.completed_at -eq [DBNull]::Value -or $null -eq $run.completed_at) {
			[void]$keepers.Add($runId)
			continue
		}

		$completedAt = [datetime]$run.completed_at
		if ($completedAt -ge $dailyCutoff) {
			$key = $completedAt.ToString('yyyy-MM-dd')
			if (-not $dailyBuckets.ContainsKey($key)) {
				$dailyBuckets[$key] = $runId
				[void]$keepers.Add($runId)
			}
			continue
		}

		if ($completedAt -ge $monthlyCutoff) {
			$weekYear = [System.Globalization.ISOWeek]::GetYear($completedAt)
			$weekNumber = [System.Globalization.ISOWeek]::GetWeekOfYear($completedAt)
			$key = '{0:D4}-W{1:D2}' -f $weekYear, $weekNumber
			if (-not $weeklyBuckets.ContainsKey($key)) {
				$weeklyBuckets[$key] = $runId
				[void]$keepers.Add($runId)
			}
			continue
		}

		$key = $completedAt.ToString('yyyy-MM')
		if (-not $monthlyBuckets.ContainsKey($key)) {
			$monthlyBuckets[$key] = $runId
			[void]$keepers.Add($runId)
		}
	}

	return ,$keepers
}

function Invoke-SnapshotCleanup {
	param(
		[string]$DatasetType = $script:activeDatasetType
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	$keepers = Get-KeeperRunIds -DatasetType $normalizedDatasetType
	$completedRuns = ConvertFrom-DataTable (Invoke-AppQuery @"
SELECT run_id
FROM inventory_runs
WHERE status = 'completed'
	AND dataset_type = $(Escape-SqlString $normalizedDatasetType);
"@)

	$deleteIds = @()
	foreach ($run in $completedRuns) {
		$runId = [int64]$run.run_id
		if (-not $keepers.Contains($runId)) {
			$deleteIds += $runId
		}
	}

	if ($deleteIds.Count -eq 0) {
		return
	}

	$idList = $deleteIds -join ','
	Invoke-AppUpdate "DELETE FROM inventory_items WHERE run_id IN ($idList);"
	Invoke-AppUpdate "DELETE FROM inventory_run_sources WHERE run_id IN ($idList);"
	Invoke-AppUpdate "DELETE FROM inventory_runs WHERE run_id IN ($idList);"
}

function Publish-InventorySnapshot {
	param(
		[Parameter(Mandatory = $true)]
		[object[]]$Items,
		[Parameter(Mandatory = $true)]
		[object[]]$Sources,
		[string]$DatasetType = $script:activeDatasetType,
		[datetime]$StartedAt = (Get-Date),
		[Nullable[datetime]]$CompletedAt = $null,
		[switch]$IsBootstrap
	)

	$normalizedDatasetType = Get-NormalizedDatasetType $DatasetType
	$stateId = Get-DatasetStateId $normalizedDatasetType
	$workerHost = [System.Environment]::MachineName
	$runId = $null
	$startedAtSql = ConvertTo-MySqlDateTimeLiteral $StartedAt
	$retentionClassSql = if ($IsBootstrap) { Escape-SqlString 'bootstrap' } else { 'NULL' }
	$transactionCompleted = $false

	try {
		Start-AppTransaction

		Invoke-AppUpdate @"
INSERT INTO inventory_runs (dataset_type, started_at, completed_at, status, total_rows, worker_host, error_message, retention_class)
VALUES (
	$(Escape-SqlString $normalizedDatasetType),
	$startedAtSql,
	NULL,
	'running',
	0,
	$(Escape-SqlString $workerHost),
	NULL,
	$retentionClassSql
);
"@

		$runId = [int64](Invoke-AppScalarValue 'SELECT LAST_INSERT_ID();')
		Insert-InventoryItemsBatch -RunId $runId -Items $Items
		Insert-InventorySources -RunId $runId -Sources $Sources

		$completedStamp = if ($CompletedAt.HasValue) { $CompletedAt.Value } else { Get-Date }
		$completedAtSql = ConvertTo-MySqlDateTimeLiteral $completedStamp
		Invoke-AppUpdate @"
UPDATE inventory_runs
SET
	completed_at = $completedAtSql,
	status = 'completed',
	total_rows = $($Items.Count),
	error_message = NULL
WHERE run_id = $runId;
"@

		Invoke-AppUpdate @"
UPDATE inventory_state
SET
	current_run_id = $runId,
	last_successful_run_at = $completedAtSql,
	last_row_count = $($Items.Count)
WHERE state_id = $stateId;
"@

		Complete-AppTransaction
		$transactionCompleted = $true
	} catch {
		if (-not $transactionCompleted) {
			Undo-AppTransaction
		}
		if ($runId -and -not $transactionCompleted) {
			try {
				Invoke-AppUpdate @"
UPDATE inventory_runs
SET
	status = 'failed',
	error_message = $(Escape-SqlString $_.Exception.Message)
WHERE run_id = $runId;
"@
			} catch {}
		}
		throw
	}

	Invoke-SnapshotCleanup -DatasetType $normalizedDatasetType
	return $runId
}

function Initialize-BootstrapData {
	$currentMeta = Get-CurrentPublishedRunMetadata -DatasetType $RegularDatasetType
	if ($currentMeta -and $currentMeta.RunId) {
		return
	}

	if (-not (Test-Path -LiteralPath $outputCSV)) {
		return
	}

	$csvItems = Import-Csv -LiteralPath $outputCSV | Select-Object name, location
	$sources = @()
	if (Test-Path -LiteralPath $makeListLog) {
		$sources = Import-Csv -LiteralPath $makeListLog | Select-Object Source, Rows
	} else {
		$sources = @([pscustomobject]@{
			Source = 'TOTAL'
			Rows = $csvItems.Count
		})
	}

	$completedAt = (Get-Item -LiteralPath $outputCSV).LastWriteTime
	Publish-InventorySnapshot -Items $csvItems -Sources $sources -DatasetType $RegularDatasetType -StartedAt $completedAt -CompletedAt $completedAt -IsBootstrap | Out-Null
}

function Initialize-AppDatabase {
	Initialize-DatabaseSchema
	Initialize-BootstrapData
}
