[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RuntimeConfigurationPath,

    [switch]$AllowExternalSideEffects
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Dedicated entry point for the later unattended production run. It exposes no
# mode selection, loads no simulation adapters and always uses the fail-closed
# Execute path from the shared runtime implementation.
$bootstrapRunId = [guid]::NewGuid().ToString('D')
$bootstrapStartedUtc = [DateTime]::UtcNow.ToString('o')

try {
    $corePath = Join-Path -Path $PSScriptRoot -ChildPath 'VisioSharePointSync.Core.ps1'
    $runtimePath = Join-Path -Path $PSScriptRoot -ChildPath 'VisioSharePointSync.Pipeline.Runtime.ps1'
    . $corePath
    . $runtimePath

    $pipelineParameters = @{
        ConfigurationPath        = $ConfigurationPath
        RuntimeConfigurationPath = $RuntimeConfigurationPath
        Mode                     = 'Execute'
        AllowExternalSideEffects = $AllowExternalSideEffects
    }
    $pipelineResult = Invoke-VisioSharePointSyncPipeline @pipelineParameters
    Format-VssPipelineResult -Result $pipelineResult | Write-Output
    exit ([int]$pipelineResult.ExitCode)
}
catch {
    # Do not depend on partially loaded runtime functions in the bootstrap catch.
    Write-Output 'Visio-SharePoint-Pipeline'
    Write-Output "RunId: $bootstrapRunId"
    Write-Output 'Modus: Execute'
    Write-Output 'Status: INTERNAL_ERROR'
    Write-Output 'Exitcode: 1'
    Write-Output "StartedUtc: $bootstrapStartedUtc"
    Write-Output ("FinishedUtc: {0}" -f [DateTime]::UtcNow.ToString('o'))
    Write-Output 'Stufen:'
    Write-Output '  - keine'
    Write-Output 'Dateien:'
    Write-Output '  - keine'
    Write-Output 'Fehler:'
    Write-Output '  - Der produktive Pipeline-Einstieg konnte nicht gestartet werden.'
    Write-Output 'Platzhalter:'
    Write-Output '  - keine'
    exit 1
}
