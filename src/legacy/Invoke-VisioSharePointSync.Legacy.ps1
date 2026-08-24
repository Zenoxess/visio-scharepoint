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

# LEGACY ENTRY POINT: retained only for compatibility with the first
# Validate/DryRun scaffold. New development and production automation must use
# the explicitly separated entry points in src\development and src\production.
$srcRoot = Split-Path -Path $PSScriptRoot -Parent
$productionRoot = Join-Path -Path $srcRoot -ChildPath 'production'
$corePath = Join-Path -Path $productionRoot -ChildPath 'VisioSharePointSync.Core.ps1'

try {
    . $corePath
    $result = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $ConfigurationPath -Mode $Mode
    $formattedResult = @(Format-VisioSharePointSyncAssessment -Result $result -Mode $Mode)
    if ($formattedResult.Count -gt 0) {
        $formattedResult[0] = 'Visio-SharePoint-Sync (LEGACY-Geruest)'
    }
    $formattedResult | Write-Output
    exit (Get-VisioSharePointSyncExitCode -Result $result)
}
catch {
    # Do not print exception details: they can contain configuration fragments.
    Write-Output 'Visio-SharePoint-Sync (LEGACY-Geruest)'
    Write-Output "Modus: $Mode"
    Write-Output 'Status: UNGUELTIG'
    Write-Output 'Exitcode: 1'
    Write-Output 'Fehler:'
    Write-Output '  - Interner Fehler bei der reinen Konfigurationspruefung.'
    exit 1
}
