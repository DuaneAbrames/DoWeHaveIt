[CmdletBinding()]
param([switch]$Music)

# Container-only policy: filesystem failures must stop the scan before publication.
$ErrorActionPreference = 'Stop'
& /config/DoWeHaveIt-Worker.ps1 -Music:$Music
