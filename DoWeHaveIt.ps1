[CmdletBinding()]
Param(
    [switch]$createList = $false,
    [switch]$TimeMachine = $false,
    [switch]$Music = $false
)

if ($createList -or $PSVersionTable.Platform -eq 'Unix') {
    & (Join-Path $PSScriptRoot 'DoWeHaveIt-Worker.ps1') -Music:$Music
} else {
    & (Join-Path $PSScriptRoot 'DoWeHaveIt-UI.ps1') -TimeMachine:$TimeMachine -Music:$Music
}
