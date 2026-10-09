# Structural refactor

The compatibility launcher selects the UI on Windows or the worker for `-createList` and Unix, forwarding dataset and UI options. Both components run directly. Common is dot-sourced into each component's script scope to preserve `$script:` state and UI event-handler behavior. All original function bodies are unchanged.

## Changes beyond relocation

- The launcher delegates using `&` and sibling paths resolved with `$PSScriptRoot`; each component owns its startup call and existing `try/finally` connection cleanup.
- UI accepts `-TimeMachine` and `-Music`; worker accepts `-Music` and always starts worker mode. The original launcher retains all three switches.
- Unix automatic worker selection moved to the launcher. Unix `renice` moved to the worker so it still runs once before worker startup, including direct invocation.
- Configuration/state initialization is divided by ownership. Common initializes database configuration and active dataset; UI initializes its display/event state. Path values and constants are unchanged.

No enumeration, SQL/schema/publication/retention, import, UI, or function-body changes were made. The original unused `Get-SystemTimestamp` helper is retained in Common.

## Function ownership

| Original function | Destination |
| --- | --- |
| `Write-CreateListMessage` | `DoWeHaveIt-Worker.ps1` |
| `Get-DatasetStateId` | `DoWeHaveIt-Common.ps1` |
| `Get-NormalizedDatasetType` | `DoWeHaveIt-Common.ps1` |
| `Get-DatasetDisplayName` | `DoWeHaveIt-Common.ps1` |
| `Escape-SqlString` | `DoWeHaveIt-Common.ps1` |
| `ConvertTo-MySqlDateTimeLiteral` | `DoWeHaveIt-Common.ps1` |
| `ConvertFrom-DataTable` | `DoWeHaveIt-Common.ps1` |
| `Get-DatabaseConfig` | `DoWeHaveIt-Common.ps1` |
| `Import-AppModules` | `DoWeHaveIt-Common.ps1` |
| `Open-AppConnection` | `DoWeHaveIt-Common.ps1` |
| `Close-AppConnection` | `DoWeHaveIt-Common.ps1` |
| `Invoke-AppQuery` | `DoWeHaveIt-Common.ps1` |
| `Invoke-AppScalarValue` | `DoWeHaveIt-Common.ps1` |
| `Invoke-AppUpdate` | `DoWeHaveIt-Common.ps1` |
| `Start-AppTransaction` | `DoWeHaveIt-Common.ps1` |
| `Complete-AppTransaction` | `DoWeHaveIt-Common.ps1` |
| `Undo-AppTransaction` | `DoWeHaveIt-Common.ps1` |
| `Test-ColumnExists` | `DoWeHaveIt-Common.ps1` |
| `Test-IndexExists` | `DoWeHaveIt-Common.ps1` |
| `Initialize-DatabaseSchema` | `DoWeHaveIt-Common.ps1` |
| `Get-SystemTimestamp` | `DoWeHaveIt-Common.ps1` |
| `Get-LatestSuccessfulRunAge` | `DoWeHaveIt-Worker.ps1` |
| `Test-MusicScanNeeded` | `DoWeHaveIt-Worker.ps1` |
| `Get-CurrentPublishedRunMetadata` | `DoWeHaveIt-Common.ps1` |
| `Get-RunMetadataById` | `DoWeHaveIt-UI.ps1` |
| `Get-RunItems` | `DoWeHaveIt-UI.ps1` |
| `Get-RunSources` | `DoWeHaveIt-UI.ps1` |
| `Get-CompletedRunsForPicker` | `DoWeHaveIt-UI.ps1` |
| `Get-PreviousCompletedRunId` | `DoWeHaveIt-Common.ps1` |
| `New-RunDataTable` | `DoWeHaveIt-UI.ps1` |
| `Convert-ItemsToDataTable` | `DoWeHaveIt-UI.ps1` |
| `Get-InventorySnapshot` | `DoWeHaveIt-Worker.ps1` |
| `Get-MusicInventorySnapshot` | `DoWeHaveIt-Worker.ps1` |
| `Insert-InventoryItemsBatch` | `DoWeHaveIt-Common.ps1` |
| `Insert-InventorySources` | `DoWeHaveIt-Common.ps1` |
| `Get-KeeperRunIds` | `DoWeHaveIt-Common.ps1` |
| `Invoke-SnapshotCleanup` | `DoWeHaveIt-Common.ps1` |
| `Publish-InventorySnapshot` | `DoWeHaveIt-Common.ps1` |
| `Initialize-BootstrapData` | `DoWeHaveIt-Common.ps1` |
| `Initialize-AppDatabase` | `DoWeHaveIt-Common.ps1` |
| `Get-StatusAgeMinutes` | `DoWeHaveIt-UI.ps1` |
| `Update-InfoLabelPosition` | `DoWeHaveIt-UI.ps1` |
| `Set-InfoLabel` | `DoWeHaveIt-UI.ps1` |
| `Get-ControlCenterY` | `DoWeHaveIt-UI.ps1` |
| `Align-TopRowControls` | `DoWeHaveIt-UI.ps1` |
| `Escape-ForRowFilter` | `DoWeHaveIt-UI.ps1` |
| `Get-ResolvedItemPath` | `DoWeHaveIt-UI.ps1` |
| `Show-SourceLogDialog` | `DoWeHaveIt-UI.ps1` |
| `Show-RunPickerDialog` | `DoWeHaveIt-UI.ps1` |
| `Show-ConnectionFailure` | `DoWeHaveIt-UI.ps1` |
| `Start-UiMode` | `DoWeHaveIt-UI.ps1` |
| `Start-WorkerMode` | `DoWeHaveIt-Worker.ps1` |

## Manual checks (pending on the actual application environment)

Inventory operations can update the database; perform them as intentional application operations or against an isolated copy.

- Windows: launch the compatibility launcher without arguments and UI directly. Check filtering, keyboard shortcuts, clipboard, Explorer, source counts, music selection, display formatting, polling, idle/deferred refresh, and closing the form.
- Windows: launch each UI entry point with `-TimeMachine`, also with `-Music`; check historical run selection and returning to the current snapshot.
- Regular updates: run the launcher with `-createList` and worker directly without arguments. Compare item/source rows, ordering, and current run metadata; verify the existing music age decision and connection cleanup on success/failure.
- Music updates: run `DoWeHaveIt.ps1 -createList -Music` and `DoWeHaveIt-Worker.ps1 -Music`; compare rows, sources, and publication metadata.
- Linux PowerShell 7: run the launcher without arguments, then with `-Music`; run the worker directly in both modes. Verify Unix paths, priority adjustment, worker selection, and no Windows Forms initialization.
- From another working directory, launch using an absolute script path. Verify sibling loading and unchanged application data paths.

## Validation completed

- PowerShell 7 and Windows PowerShell 5.1: all four scripts parse successfully.
- AST comparison against the pre-refactor working file: all 52 original top-level functions occur exactly once and retain identical function text, including inventory scans, SQL, transaction/publication order, cleanup, and UI event handlers.
- On each PowerShell version, isolated copies with startup/connection stubs passed 16 launcher routing combinations (simulated Windows/Unix, all switch combinations) and 16 direct component startup/cleanup cases (both datasets, success/failure). Shared database initialization functions were available in each component; worker startup did not load UI functions. These tests execute no database operations, inventory scans, or Forms initialization.
- Sibling loading was exercised from a working directory different from the isolated scripts' directory.
- Live Windows Forms behavior, real scans/database publication, and execution on an actual Linux host remain pending. No production inventory or database update was performed.
