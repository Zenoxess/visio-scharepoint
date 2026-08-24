[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RuntimeConfigurationPath,

    [ValidateSet('Validate', 'Simulate')]
    [string]$Mode = 'Validate'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Development and acceptance entry point. Production automation must use the
# dedicated script under src\production and never this multi-mode entry point.
$bootstrapRunId = [guid]::NewGuid().ToString('D')
$bootstrapStartedUtc = [DateTime]::UtcNow.ToString('o')

try {
    $srcRoot = Split-Path -Path $PSScriptRoot -Parent
    $productionRoot = Join-Path -Path $srcRoot -ChildPath 'production'
    $corePath = Join-Path -Path $productionRoot -ChildPath 'VisioSharePointSync.Core.ps1'
    $runtimePath = Join-Path -Path $productionRoot -ChildPath 'VisioSharePointSync.Pipeline.Runtime.ps1'
    $simulationPath = Join-Path -Path $PSScriptRoot -ChildPath 'VisioSharePointSync.Pipeline.Simulation.ps1'
    . $corePath
    . $runtimePath
    . $simulationPath

    $pipelineParameters = @{
        ConfigurationPath        = $ConfigurationPath
        RuntimeConfigurationPath = $RuntimeConfigurationPath
        Mode                     = $Mode
    }
    $pipelineResult = Invoke-VisioSharePointSyncPipeline @pipelineParameters

    Format-VssPipelineResult -Result $pipelineResult | Write-Output
    exit ([int]$pipelineResult.ExitCode)
}
catch {
    # Do not depend on partially loaded runtime functions in the bootstrap catch.
    Write-Output 'Visio-SharePoint-Pipeline'
    Write-Output "RunId: $bootstrapRunId"
    Write-Output "Modus: $Mode"
    Write-Output 'Status: INTERNAL_ERROR'
    Write-Output 'Exitcode: 1'
    Write-Output "StartedUtc: $bootstrapStartedUtc"
    Write-Output ("FinishedUtc: {0}" -f [DateTime]::UtcNow.ToString('o'))
    Write-Output 'Stufen:'
    Write-Output '  - keine'
    Write-Output 'Dateien:'
    Write-Output '  - keine'
    Write-Output 'Fehler:'
    Write-Output '  - Interner Fehler beim Laden des Pipeline-Geruests.'
    Write-Output 'Platzhalter:'
    Write-Output '  - keine'
    exit 1
}
