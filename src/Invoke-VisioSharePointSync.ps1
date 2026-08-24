[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationPath,

    [ValidateSet('Validate', 'DryRun')]
    [string]$Mode = 'Validate'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Public scaffold entry point. Intentionally no Execute mode exists. Validate and
# DryRun only read the supplied JSON file and perform in-memory checks.
$corePath = Join-Path -Path $PSScriptRoot -ChildPath 'VisioSharePointSync.Core.ps1'

try {
    . $corePath
    $result = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $ConfigurationPath -Mode $Mode
    Format-VisioSharePointSyncAssessment -Result $result -Mode $Mode | Write-Output
    exit (Get-VisioSharePointSyncExitCode -Result $result)
}
catch {
    # Do not print exception details: they can contain configuration fragments.
    Write-Output 'Visio-SharePoint-Sync (Geruest)'
    Write-Output "Modus: $Mode"
    Write-Output 'Status: UNGUELTIG'
    Write-Output 'Exitcode: 1'
    Write-Output 'Fehler:'
    Write-Output '  - Interner Fehler bei der reinen Konfigurationspruefung.'
    exit 1
}
