[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
$pipelinePath = Join-Path $projectRoot 'src\Invoke-VisioSharePointSync.Pipeline.ps1'
$pipelineRuntimePath = Join-Path $projectRoot 'src\VisioSharePointSync.Pipeline.Runtime.ps1'
$pipelineSimulationPath = Join-Path $projectRoot 'src\VisioSharePointSync.Pipeline.Simulation.ps1'
$productionPath = Join-Path $projectRoot 'src\Invoke-VisioSharePointSync.Production.ps1'
$corePath = Join-Path $projectRoot 'src\VisioSharePointSync.Core.ps1'
$legacyTestPath = Join-Path $PSScriptRoot 'Test-Scaffold.ps1'
$graphFixturePath = Join-Path $PSScriptRoot 'fixtures\complete.config.json'
$restFixturePath = Join-Path $PSScriptRoot 'fixtures\server-rest.config.json'
$incompleteFixturePath = Join-Path $projectRoot 'config\sync.example.json'
$runtimeExamplePath = Join-Path $projectRoot 'config\runtime.example.json'

$script:PassedCount = 0
$script:FailedCount = 0
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false, $true)

function Assert-VssTrue {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-VssEqual {
    param([AllowNull()][object]$Expected, [AllowNull()][object]$Actual, [string]$Message)
    if ($Expected -ne $Actual) {
        throw ("{0} Erwartet: [{1}], erhalten: [{2}]." -f $Message, $Expected, $Actual)
    }
}

function Assert-VssMatches {
    param([AllowEmptyString()][string]$Actual, [string]$Pattern, [string]$Message)
    if ($Actual -notmatch $Pattern) { throw $Message }
}

function Invoke-VssTest {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        $script:PassedCount++
        Write-Output "PASS: $Name"
    }
    catch {
        $script:FailedCount++
        Write-Output "FAIL: $Name"
        Write-Output ("  {0}" -f $_.Exception.Message)
    }
}

function New-VssTestSandbox {
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ('vss-pipeline-tests-' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($path)
    return $path
}

function Remove-VssTestSandbox {
    param([string]$Path)
    if (-not [System.IO.Directory]::Exists($Path)) { return }
    $tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd([char]92)
    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([char]92)
    if (-not $fullPath.StartsWith($tempRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Test-Sandbox liegt nicht sicher unter dem Temp-Verzeichnis.'
    }
    [System.IO.Directory]::Delete($fullPath, $true)
}

function Read-VssJsonClone {
    param([string]$Path)
    $text = [System.IO.File]::ReadAllText($Path, $script:Utf8NoBom)
    return ($text | ConvertFrom-Json)
}

function Write-VssJsonFile {
    param([object]$Value, [string]$Path)
    $json = $Value | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($Path, $json, $script:Utf8NoBom)
}

function New-VssPrimaryConfiguration {
    param(
        [string]$FixturePath,
        [string]$SandboxPath,
        [string]$Prefix
    )
    $configuration = Read-VssJsonClone $FixturePath
    $configuration.Operations.StatePath = Join-Path $SandboxPath ($Prefix + '-state.json')
    $configuration.Operations.StagingPath = Join-Path $SandboxPath ($Prefix + '-staging')
    $configuration.Operations.LogPath = Join-Path $SandboxPath ($Prefix + '-logs')
    return $configuration
}

function New-VssRuntimeConfiguration {
    param(
        [string]$SandboxPath,
        [string]$Prefix,
        [bool]$ExecutionEnabled
    )
    $runtime = Read-VssJsonClone $runtimeExamplePath
    $runtime.Execution.Enabled = $ExecutionEnabled
    $runtime.Conversion.AdapterKind = 'VisioCom'
    $runtime.Conversion.AdapterVersion = 'test-visio-com-1.0'
    $runtime.Conversion.ExternalExecutablePath = $null
    $runtime.Operations.QuarantinePath = Join-Path $SandboxPath ($Prefix + '-quarantine')
    return $runtime
}

function Invoke-VssPipelineCli {
    param(
        [string]$ConfigurationPath,
        [string]$RuntimeConfigurationPath,
        [string]$EntryPointPath = $pipelinePath,
        [ValidateSet('Validate', 'Simulate', 'Execute')][string]$Mode = 'Validate',
        [switch]$AllowExternalSideEffects
    )
    $hostExecutable = (Get-Process -Id $PID).Path
    $argumentList = @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', $EntryPointPath,
        '-ConfigurationPath', $ConfigurationPath,
        '-RuntimeConfigurationPath', $RuntimeConfigurationPath,
        '-Mode', $Mode
    )
    if ($AllowExternalSideEffects) { $argumentList += '-AllowExternalSideEffects' }
    $captured = @(& $hostExecutable @argumentList 2>&1)
    $exitCode = $LASTEXITCODE
    $lines = @($captured | ForEach-Object { [string]$_ })
    return [pscustomobject]@{
        ExitCode = $exitCode
        Lines    = $lines
        Output   = ($lines -join [Environment]::NewLine)
    }
}

function Invoke-VssProductionCli {
    param(
        [string]$ConfigurationPath,
        [string]$RuntimeConfigurationPath,
        [string]$EntryPointPath = $productionPath,
        [switch]$AllowExternalSideEffects
    )
    $hostExecutable = (Get-Process -Id $PID).Path
    $argumentList = @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', $EntryPointPath,
        '-ConfigurationPath', $ConfigurationPath,
        '-RuntimeConfigurationPath', $RuntimeConfigurationPath
    )
    if ($AllowExternalSideEffects) { $argumentList += '-AllowExternalSideEffects' }
    $captured = @(& $hostExecutable @argumentList 2>&1)
    $exitCode = $LASTEXITCODE
    $lines = @($captured | ForEach-Object { [string]$_ })
    return [pscustomobject]@{
        ExitCode = $exitCode
        Lines    = $lines
        Output   = ($lines -join [Environment]::NewLine)
    }
}

function Invoke-VssPowerShellFile {
    param([string]$Path)
    $hostExecutable = (Get-Process -Id $PID).Path
    $captured = @(& $hostExecutable -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Path 2>&1)
    $exitCode = $LASTEXITCODE
    $lines = @($captured | ForEach-Object { [string]$_ })
    return [pscustomobject]@{
        ExitCode = $exitCode
        Output   = ($lines -join [Environment]::NewLine)
    }
}

function Get-VssDirectorySnapshot {
    param([string]$Path)
    if (-not [System.IO.Directory]::Exists($Path)) { return @() }
    $root = [System.IO.Path]::GetFullPath($Path).TrimEnd([char]92)
    foreach ($item in @(Get-ChildItem -LiteralPath $root -Force -Recurse | Sort-Object FullName)) {
        $relativePath = $item.FullName.Substring($root.Length).TrimStart([char]92)
        if ($item.PSIsContainer) {
            Write-Output ('D|{0}' -f $relativePath)
        }
        else {
            Write-Output ('F|{0}|{1}|{2}' -f $relativePath, $item.Length, $item.LastWriteTimeUtc.Ticks)
        }
    }
}

function Get-VssPipelineAst {
    $source = @(
        [System.IO.File]::ReadAllText($pipelineRuntimePath, $script:Utf8NoBom),
        [System.IO.File]::ReadAllText($pipelineSimulationPath, $script:Utf8NoBom)
    ) -join ([Environment]::NewLine + [Environment]::NewLine)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput(
        $source,
        [ref]$tokens,
        [ref]$parseErrors
    )
    return [pscustomobject]@{ Ast = $ast; Errors = @($parseErrors) }
}

function Get-VssFunctionAst {
    param(
        [Parameter(Mandatory = $true)][System.Management.Automation.Language.Ast]$Ast,
        [Parameter(Mandatory = $true)][string]$Name
    )
    return $Ast.Find({
        param($node)
        return (
            ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
            ([string]::Equals($node.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase))
        )
    }, $true)
}

function Import-VssPipelineFunctionDefinitions {
    param([System.Management.Automation.Language.ScriptBlockAst]$Ast)
    $coreTokens = $null
    $coreParseErrors = $null
    $coreAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $corePath,
        [ref]$coreTokens,
        [ref]$coreParseErrors
    )
    if (@($coreParseErrors).Count -gt 0) { throw 'Core-Skript besitzt Parserfehler.' }
    $definitions = @(
        $coreAst.EndBlock.Statements |
            Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] } |
            ForEach-Object { $_.Extent.Text }
    ) + @(
        $Ast.EndBlock.Statements |
            Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] } |
            ForEach-Object { $_.Extent.Text }
    )
    if ($definitions.Count -eq 0) { throw 'Pipeline-Skript enthaelt keine top-level Funktionen.' }
    $moduleName = 'VssPipelineTestHelpers_' + [guid]::NewGuid().ToString('N')
    $modulePreamble = @'
$script:VssPipelinePlaceholderValue = '__PLACEHOLDER_REQUIRED__'
$script:VssPipelineRuntimeTopLevelKeys = @('SchemaVersion', 'Execution', 'Source', 'Conversion', 'Upload', 'Operations')
$script:VssPipelineRuntimeSectionKeys = @{
    Execution  = @('Enabled')
    Source     = @('Recurse', 'ExcludeHidden', 'ExcludeSystem', 'ExcludeReparsePoints', 'ExcludePatterns')
    Conversion = @('AdapterKind', 'AdapterVersion', 'ExternalExecutablePath')
    Upload     = @('SessionThresholdBytes', 'ChunkSizeBytes')
    Operations = @('QuarantinePath')
}
'@
    $moduleSource = $modulePreamble + [Environment]::NewLine + ($definitions -join [Environment]::NewLine) + [Environment]::NewLine + 'Export-ModuleMember -Function *'
    $module = New-Module -Name $moduleName -ScriptBlock ([scriptblock]::Create($moduleSource))
    Import-Module $module -Force
    return $module
}

function Assert-VssObservableResultContract {
    param([string]$Output, [string]$ExpectedMode, [int]$ExpectedExitCode)
    Assert-VssMatches $Output '(?m)^Visio-SharePoint-Pipeline\s*$' 'CLI-Kopf fehlt oder ist instabil.'
    Assert-VssMatches $Output '(?mi)^RunId:\s*[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\s*$' 'RunId fehlt oder ist keine GUID.'
    Assert-VssMatches $Output ("(?m)^Modus:\s*{0}\s*$" -f [regex]::Escape($ExpectedMode)) 'Modus ist im CLI-Ergebnis nicht beobachtbar.'
    Assert-VssMatches $Output '(?m)^Status:\s*\S+\s*$' 'Status fehlt im CLI-Ergebnis.'
    Assert-VssMatches $Output ("(?m)^Exitcode:\s*{0}\s*$" -f $ExpectedExitCode) 'Exitcode fehlt im formatierten CLI-Ergebnis.'
    Assert-VssMatches $Output '(?m)^StartedUtc:\s*\S+\s*$' 'StartedUtc fehlt im CLI-Ergebnis.'
    Assert-VssMatches $Output '(?m)^FinishedUtc:\s*\S+\s*$' 'FinishedUtc fehlt im CLI-Ergebnis.'
    foreach ($sectionName in @('Stufen', 'Dateien', 'Platzhalter')) {
        Assert-VssMatches $Output ("(?m)^{0}:\s*$" -f [regex]::Escape($sectionName)) "Ergebnisabschnitt $sectionName fehlt."
    }
}

$sandbox = New-VssTestSandbox
try {
    $graphConfiguration = New-VssPrimaryConfiguration -FixturePath $graphFixturePath -SandboxPath $sandbox -Prefix 'graph'
    $restConfiguration = New-VssPrimaryConfiguration -FixturePath $restFixturePath -SandboxPath $sandbox -Prefix 'rest'
    $executeConfiguration = New-VssPrimaryConfiguration -FixturePath $graphFixturePath -SandboxPath $sandbox -Prefix 'execute'
    $runtimeDisabled = New-VssRuntimeConfiguration -SandboxPath $sandbox -Prefix 'disabled' -ExecutionEnabled $false
    $runtimeEnabled = New-VssRuntimeConfiguration -SandboxPath $sandbox -Prefix 'enabled' -ExecutionEnabled $true

    $graphConfigurationPath = Join-Path $sandbox 'graph.config.json'
    $restConfigurationPath = Join-Path $sandbox 'rest.config.json'
    $executeConfigurationPath = Join-Path $sandbox 'execute.config.json'
    $runtimeDisabledPath = Join-Path $sandbox 'runtime-disabled.json'
    $runtimeEnabledPath = Join-Path $sandbox 'runtime-enabled.json'
    Write-VssJsonFile $graphConfiguration $graphConfigurationPath
    Write-VssJsonFile $restConfiguration $restConfigurationPath
    Write-VssJsonFile $executeConfiguration $executeConfigurationPath
    Write-VssJsonFile $runtimeDisabled $runtimeDisabledPath
    Write-VssJsonFile $runtimeEnabled $runtimeEnabledPath

    Invoke-VssTest -Name 'Pipelinequelle, Stufen und Platzhalter bleiben statisch auffindbar' -Body {
        Assert-VssTrue ([System.IO.File]::Exists($pipelinePath)) 'Pipeline-CLI fehlt.'
        Assert-VssTrue ([System.IO.File]::Exists($pipelineRuntimePath)) 'Gemeinsame Pipeline-Runtime fehlt.'
        Assert-VssTrue ([System.IO.File]::Exists($pipelineSimulationPath)) 'Getrennte Pipeline-Simulation fehlt.'
        Assert-VssTrue ([System.IO.File]::Exists($productionPath)) 'Produktiv-CLI fehlt.'
        $source = @(
            [System.IO.File]::ReadAllText($pipelinePath, $script:Utf8NoBom),
            [System.IO.File]::ReadAllText($pipelineRuntimePath, $script:Utf8NoBom),
            [System.IO.File]::ReadAllText($pipelineSimulationPath, $script:Utf8NoBom)
        ) -join ([Environment]::NewLine + [Environment]::NewLine)
        $stageIds = @(
            'ValidateConfiguration', 'ValidateRuntimeConfiguration', 'AcquireSingleRunLock',
            'InventorySource', 'StageStableSource', 'ConvertVisioToPdf', 'VerifySourceUnchanged',
            'EnsureSharePointFolders', 'UploadOrUpdatePdf', 'ReconcileLocalAndRemoteState',
            'PersistSyncState', 'WriteStructuredLog', 'CleanupStaging',
            'ReleaseSingleRunLock', 'SummarizeRun'
        )
        foreach ($stageId in $stageIds) {
            Assert-VssTrue $source.Contains($stageId) "Stufen-ID $stageId ist im Quelltext nicht auffindbar."
        }

        $placeholderIds = @(
            'VSS-CNV-001', 'VSS-AUTH-GRAPH-001', 'VSS-SPREST-LARGE-001',
            'VSS-SPREST-PATH-001', 'VSS-RUNTIME-QUARANTINE-001',
            'VSS-INTEGRATION-001', 'VSS-GOLIVE-001'
        )
        foreach ($placeholderId in $placeholderIds) {
            Assert-VssTrue $source.Contains($placeholderId) "Platzhalter-ID $placeholderId ist im Quelltext nicht auffindbar."
            Assert-VssTrue $source.Contains("# >>> PLACEHOLDER [$placeholderId] BEGIN") "BEGIN-Marker fuer $placeholderId fehlt."
            Assert-VssTrue $source.Contains("# <<< PLACEHOLDER [$placeholderId] END") "END-Marker fuer $placeholderId fehlt."
        }
        Assert-VssTrue $source.Contains('__PLACEHOLDER_REQUIRED__') 'Der kanonische Platzhalter-Sentinel fehlt.'

        $parsed = Get-VssPipelineAst
        Assert-VssEqual 0 @($parsed.Errors).Count 'Pipeline-Skript besitzt Parserfehler.'
        Assert-VssTrue ($source -notmatch '(?i)-ComObject\b|\b(New-Object\s+Visio\.Application|Remove-PnP|Move-PnP|Recycle-PnP)') 'Direkte COM-, Delete-, Move- oder Archive-Operation gefunden.'
        $httpCalls = @($parsed.Ast.FindAll({
            param($node)
            return (($node -is [System.Management.Automation.Language.CommandAst]) -and
                ([string]$node.GetCommandName() -in @('Invoke-RestMethod', 'Invoke-WebRequest')))
        }, $true))
        foreach ($httpCall in $httpCalls) {
            $ancestor = $httpCall.Parent
            while ($null -ne $ancestor -and -not ($ancestor -is [System.Management.Automation.Language.FunctionDefinitionAst])) { $ancestor = $ancestor.Parent }
            Assert-VssTrue ($null -ne $ancestor -and $ancestor.Name -in @('Send-VssGraphUploadSessionChunks', 'Invoke-VssGraphTransport', 'Get-VssRestRequestDigest', 'Invoke-VssRestTransport')) "Direkter HTTP-Aufruf ausserhalb der vorgesehenen Transportgrenze: $($httpCall.Extent.Text)"
        }
        $simulationFunction = Get-VssFunctionAst -Ast $parsed.Ast -Name 'Invoke-VssPipelineSimulation'
        Assert-VssTrue ($null -ne $simulationFunction) 'Invoke-VssPipelineSimulation fehlt.'
        $simulationSource = $simulationFunction.Extent.Text
        $forbiddenSimulationCalls = '(?i)\b(Get-ChildItem|Invoke-WebRequest|Invoke-RestMethod|Invoke-VssPipelineInventory|Copy-VssPipelineStableFile|Enter-VssPipelineLock|Save-VssPipelineState|Write-VssPipelineLog|Invoke-VssGraphTransport|Invoke-VssRestTransport)\b|ComObject|System\.IO\.File\]::Exists'
        Assert-VssTrue ($simulationSource -notmatch $forbiddenSimulationCalls) 'Simulation enthaelt einen verbotenen externen oder schreibenden Aufruf.'
    }

    Invoke-VssTest -Name 'Produktiveinstieg laedt keine Simulation und bleibt bis zur Freigabe fail-closed' -Body {
        $productionSource = [System.IO.File]::ReadAllText($productionPath, $script:Utf8NoBom)
        $productionTokens = $null
        $productionParseErrors = $null
        $productionAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $productionPath,
            [ref]$productionTokens,
            [ref]$productionParseErrors
        )
        Assert-VssEqual 0 @($productionParseErrors).Count 'Produktiv-CLI besitzt Parserfehler.'
        $productionParameterNames = @(
            $productionAst.ParamBlock.Parameters |
                ForEach-Object { $_.Name.VariablePath.UserPath }
        )
        Assert-VssEqual 'ConfigurationPath|RuntimeConfigurationPath|AllowExternalSideEffects' ($productionParameterNames -join '|') 'Produktiv-CLI besitzt unerwartete Test- oder Modusparameter.'
        foreach ($forbiddenText in @(
            'VisioSharePointSync.Pipeline.Simulation.ps1',
            'Invoke-VssPipelineSimulation',
            'New-VssPipelineSimulationAdapterSet',
            'memory://',
            'FakePdf',
            'tests\fixtures'
        )) {
            Assert-VssTrue (-not $productionSource.Contains($forbiddenText)) "Produktiv-CLI enthaelt Simulations- oder Testbezug: $forbiddenText"
        }
        Assert-VssTrue $productionSource.Contains('VisioSharePointSync.Pipeline.Runtime.ps1') 'Produktiv-CLI laedt die gemeinsame Runtime nicht.'
        Assert-VssTrue (-not $productionSource.Contains('Invoke-VisioSharePointSync.Pipeline.ps1')) 'Produktiv-CLI delegiert noch an den gemischten Pipeline-Einstieg.'

        $before = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        $result = Invoke-VssProductionCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $runtimeDisabledPath
        $after = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        Assert-VssEqual 3 $result.ExitCode 'Produktiv-CLI ohne doppelte Freigabe muss Exitcode 3 liefern.'
        Assert-VssEqual $before $after 'Blockierter Produktivaufruf hat Dateien oder Verzeichnisse veraendert.'
        Assert-VssObservableResultContract -Output $result.Output -ExpectedMode Execute -ExpectedExitCode 3

        $beforeApprovedAttempt = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        $approvedAttempt = Invoke-VssProductionCli -ConfigurationPath $executeConfigurationPath -RuntimeConfigurationPath $runtimeEnabledPath -AllowExternalSideEffects
        $afterApprovedAttempt = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        Assert-VssEqual 3 $approvedAttempt.ExitCode 'Offene Adapter- und Go-live-Freigaben muessen den Produktiv-CLI trotz beider Laufzeit-Gates blockieren.'
        Assert-VssEqual $beforeApprovedAttempt $afterApprovedAttempt 'Readiness-blockierter Produktivaufruf hat Dateien oder Verzeichnisse veraendert.'
        Assert-VssObservableResultContract -Output $approvedAttempt.Output -ExpectedMode Execute -ExpectedExitCode 3

        $isolatedProductionPath = Join-Path $sandbox 'Invoke-VisioSharePointSync.Production.Isolated.ps1'
        Copy-Item -LiteralPath $productionPath -Destination $isolatedProductionPath
        $bootstrapFailure = Invoke-VssProductionCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $runtimeDisabledPath -EntryPointPath $isolatedProductionPath
        Assert-VssEqual 1 $bootstrapFailure.ExitCode 'Fehlende Runtime neben dem Produktiv-CLI muss Exitcode 1 liefern.'
        Assert-VssObservableResultContract -Output $bootstrapFailure.Output -ExpectedMode Execute -ExpectedExitCode 1

        $isolatedPipelinePath = Join-Path $sandbox 'Invoke-VisioSharePointSync.Pipeline.Isolated.ps1'
        Copy-Item -LiteralPath $pipelinePath -Destination $isolatedPipelinePath
        $pipelineBootstrapFailure = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $runtimeDisabledPath -EntryPointPath $isolatedPipelinePath -Mode Validate
        Assert-VssEqual 1 $pipelineBootstrapFailure.ExitCode 'Fehlende Runtime neben dem Pipeline-CLI muss Exitcode 1 liefern.'
        Assert-VssObservableResultContract -Output $pipelineBootstrapFailure.Output -ExpectedMode Validate -ExpectedExitCode 1
    }

    Invoke-VssTest -Name 'Valide Runtime liefert stabilen beobachtbaren CLI-Vertrag' -Body {
        $before = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        $result = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $runtimeDisabledPath -Mode Validate
        $after = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        Assert-VssEqual 0 $result.ExitCode 'Valide Konfigurationen muessen in Validate erfolgreich sein.'
        Assert-VssEqual $before $after 'Validate hat Dateien oder Verzeichnisse veraendert.'
        Assert-VssObservableResultContract -Output $result.Output -ExpectedMode Validate -ExpectedExitCode 0
    }

    Invoke-VssTest -Name 'Runtime-Platzhalter bleiben sichtbar; Typfehler und unbekannte Schluessel werden abgewiesen' -Body {
        $placeholderResult = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $runtimeExamplePath -Mode Validate
        Assert-VssEqual 0 $placeholderResult.ExitCode 'Runtime-Platzhalter duerfen Validate nicht blockieren.'
        Assert-VssObservableResultContract -Output $placeholderResult.Output -ExpectedMode Validate -ExpectedExitCode 0
        foreach ($placeholderId in @('VSS-CNV-001', 'VSS-RUNTIME-QUARANTINE-001')) {
            Assert-VssTrue $placeholderResult.Output.Contains($placeholderId) "Runtime-Platzhalter $placeholderId fehlt in der Diagnose."
        }

        $wrongType = New-VssRuntimeConfiguration -SandboxPath $sandbox -Prefix 'wrong-type' -ExecutionEnabled $false
        $wrongType.Execution.Enabled = 'false'
        $wrongTypePath = Join-Path $sandbox 'runtime-wrong-type.json'
        Write-VssJsonFile $wrongType $wrongTypePath
        $wrongTypeResult = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $wrongTypePath -Mode Validate
        Assert-VssEqual 2 $wrongTypeResult.ExitCode 'String statt Boolean muss Exitcode 2 liefern.'
        Assert-VssMatches $wrongTypeResult.Output '(?m)^Fehler:\s*$' 'Runtime-Typfehler ist im CLI-Ergebnis nicht beobachtbar.'

        $unknown = New-VssRuntimeConfiguration -SandboxPath $sandbox -Prefix 'unknown' -ExecutionEnabled $false
        $unknown | Add-Member -NotePropertyName 'UnexpectedRuntimeSetting-USER-CONTROLLED' -NotePropertyValue 'RUNTIME-VALUE-SENTINEL'
        $unknownPath = Join-Path $sandbox 'runtime-unknown.json'
        Write-VssJsonFile $unknown $unknownPath
        $unknownResult = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $unknownPath -Mode Validate
        Assert-VssEqual 2 $unknownResult.ExitCode 'Unbekannter Runtime-Schluessel muss Exitcode 2 liefern.'
        Assert-VssMatches $unknownResult.Output '(?m)^Fehler:\s*$' 'Unknown-Key-Fehler ist im CLI-Ergebnis nicht beobachtbar.'
        Assert-VssTrue ($unknownResult.Output -notmatch 'UnexpectedRuntimeSetting|USER-CONTROLLED|RUNTIME-VALUE-SENTINEL') 'Runtime-Diagnose gibt benutzerkontrollierte Schluessel oder Werte aus.'

        $secretLike = New-VssRuntimeConfiguration -SandboxPath $sandbox -Prefix 'secret-like' -ExecutionEnabled $false
        $secretLike | Add-Member -NotePropertyName 'ClientSecret' -NotePropertyValue 'SENTINEL-CLIENT-SECRET'
        $secretLikePath = Join-Path $sandbox 'runtime-secret-like.json'
        Write-VssJsonFile $secretLike $secretLikePath
        $secretLikeResult = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $secretLikePath -Mode Validate
        Assert-VssEqual 2 $secretLikeResult.ExitCode 'Geheimnisartiger Runtime-Schluessel muss Exitcode 2 liefern.'
        Assert-VssTrue (-not $secretLikeResult.Output.Contains('SENTINEL-CLIENT-SECRET')) 'Sentinel-Secret wurde ausgegeben.'

        $duplicatePath = Join-Path $sandbox 'runtime-duplicate.json'
        $duplicateJson = '{"SchemaVersion":"1.0","SchemaVersion":"1.0","Execution":{"Enabled":false},"Source":{"Recurse":true,"ExcludeHidden":true,"ExcludeSystem":true,"ExcludeReparsePoints":true,"ExcludePatterns":["~$*","*.tmp"]},"Conversion":{"AdapterKind":"VisioCom","AdapterVersion":"test","ExternalExecutablePath":null},"Upload":{"SessionThresholdBytes":10485760,"ChunkSizeBytes":10485760},"Operations":{"QuarantinePath":"C:\\VssTest\\Quarantine"}}'
        [System.IO.File]::WriteAllText($duplicatePath, $duplicateJson, $script:Utf8NoBom)
        $duplicateResult = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $duplicatePath -Mode Validate
        Assert-VssEqual 2 $duplicateResult.ExitCode 'Doppelte Runtime-Schluessel muessen Exitcode 2 liefern.'

        $bomPath = Join-Path $sandbox 'runtime-bom.json'
        $runtimeText = [System.IO.File]::ReadAllText($runtimeDisabledPath, $script:Utf8NoBom)
        $bomEncoding = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($bomPath, $runtimeText, $bomEncoding)
        $bomResult = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $bomPath -Mode Validate
        Assert-VssEqual 2 $bomResult.ExitCode 'UTF-8-BOM muss fuer Runtime-JSON abgewiesen werden.'
    }

    Invoke-VssTest -Name 'Graph-Simulation plant rein im Speicher und bleibt nebenwirkungsfrei' -Body {
        $before = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        $result = Invoke-VssPipelineCli -ConfigurationPath $graphConfigurationPath -RuntimeConfigurationPath $runtimeDisabledPath -Mode Simulate -AllowExternalSideEffects
        $after = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        Assert-VssEqual 0 $result.ExitCode 'Graph-Simulation muss erfolgreich sein.'
        Assert-VssEqual $before $after 'Graph-Simulation hat Dateien oder Verzeichnisse veraendert.'
        Assert-VssObservableResultContract -Output $result.Output -ExpectedMode Simulate -ExpectedExitCode 0
        Assert-VssMatches $result.Output '(?i)MicrosoftGraphV1' 'Graph-Transport ist im Simulationsplan nicht sichtbar.'
        Assert-VssMatches $result.Output '(?i)SmallUpload' 'Graph-Simulation plant den erwarteten kleinen Upload nicht.'
        Assert-VssMatches $result.Output '(?i)\bPUT\b' 'Graph-Simulation weist keine PUT-Methode aus.'
        Assert-VssMatches $result.Output '(?i)PdfFile' 'Graph-Simulation weist keinen PDF-Bodyvertrag aus.'
        Assert-VssMatches $result.Output '(?i)SIMULATED' 'Simulierter Stufen-/Dateistatus fehlt.'
        Assert-VssMatches $result.Output '(?i)ExternalSideEffect\s*[:=]\s*(True|true|\$true)' 'RequestPlan kennzeichnet die nur geplante externe Wirkung nicht.'
    }

    Invoke-VssTest -Name 'REST-Simulation plant rein im Speicher und bleibt nebenwirkungsfrei' -Body {
        $before = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        $result = Invoke-VssPipelineCli -ConfigurationPath $restConfigurationPath -RuntimeConfigurationPath $runtimeDisabledPath -Mode Simulate
        $after = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        Assert-VssEqual 0 $result.ExitCode 'REST-Simulation muss erfolgreich sein.'
        Assert-VssEqual $before $after 'REST-Simulation hat Dateien oder Verzeichnisse veraendert.'
        Assert-VssObservableResultContract -Output $result.Output -ExpectedMode Simulate -ExpectedExitCode 0
        Assert-VssMatches $result.Output '(?i)SharePointRest' 'REST-Transport ist im Simulationsplan nicht sichtbar.'
        Assert-VssMatches $result.Output '(?i)CreateIfAbsent' 'Sichere REST-Create-Operation fehlt im Simulationsplan.'
        Assert-VssMatches $result.Output '(?i)\bPOST\b' 'REST-Simulation weist keine POST-Methode aus.'
        Assert-VssMatches $result.Output '(?i)PdfFile' 'REST-Simulation weist keinen PDF-Bodyvertrag aus.'
        Assert-VssMatches $result.Output '(?i)SIMULATED' 'Simulierter REST-Dateistatus fehlt.'
    }

    Invoke-VssTest -Name 'Execute verlangt beide Gates und blockiert vor jeder Nebenwirkung' -Body {
        $before = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine

        $disabledResult = Invoke-VssPipelineCli -ConfigurationPath $executeConfigurationPath -RuntimeConfigurationPath $runtimeDisabledPath -Mode Execute -AllowExternalSideEffects
        Assert-VssEqual 3 $disabledResult.ExitCode 'Execution.Enabled=false muss Execute mit Exitcode 3 blockieren.'
        Assert-VssObservableResultContract -Output $disabledResult.Output -ExpectedMode Execute -ExpectedExitCode 3
        Assert-VssMatches $disabledResult.Output '(?i)Execution\.Enabled' 'Diagnose fuer das Runtime-Execute-Gate fehlt.'

        $switchMissingResult = Invoke-VssPipelineCli -ConfigurationPath $executeConfigurationPath -RuntimeConfigurationPath $runtimeEnabledPath -Mode Execute
        Assert-VssEqual 3 $switchMissingResult.ExitCode 'Fehlender AllowExternalSideEffects-Schalter muss Exitcode 3 liefern.'
        Assert-VssMatches $switchMissingResult.Output '(?i)AllowExternalSideEffects' 'Diagnose fuer das CLI-Execute-Gate fehlt.'

        $readinessResult = Invoke-VssPipelineCli -ConfigurationPath $executeConfigurationPath -RuntimeConfigurationPath $runtimeEnabledPath -Mode Execute -AllowExternalSideEffects
        Assert-VssEqual 3 $readinessResult.ExitCode 'Nicht implementierte produktive Adapter muessen Readiness-Exitcode 3 liefern.'
        foreach ($placeholderId in @('VSS-CNV-001', 'VSS-AUTH-GRAPH-001', 'VSS-INTEGRATION-001', 'VSS-GOLIVE-001')) {
            Assert-VssTrue $readinessResult.Output.Contains($placeholderId) "Readiness-Platzhalter $placeholderId fehlt."
        }

        $restReadinessResult = Invoke-VssPipelineCli -ConfigurationPath $restConfigurationPath -RuntimeConfigurationPath $runtimeEnabledPath -Mode Execute -AllowExternalSideEffects
        Assert-VssEqual 3 $restReadinessResult.ExitCode 'Nicht freigegebener REST-Execute-Pfad muss Exitcode 3 liefern.'
        foreach ($placeholderId in @('VSS-SPREST-LARGE-001', 'VSS-SPREST-PATH-001', 'VSS-INTEGRATION-001', 'VSS-GOLIVE-001')) {
            Assert-VssTrue $restReadinessResult.Output.Contains($placeholderId) "REST-Readiness-Platzhalter $placeholderId fehlt."
        }

        $after = @(Get-VssDirectorySnapshot $sandbox) -join [Environment]::NewLine
        Assert-VssEqual $before $after 'Blockiertes Execute hat vor Abschluss der Gates Nebenwirkungen erzeugt.'
    }

    Invoke-VssTest -Name 'Ungueltige Fachkonfiguration bleibt Exitcode 2 statt Readiness 3' -Body {
        $result = Invoke-VssPipelineCli -ConfigurationPath $incompleteFixturePath -RuntimeConfigurationPath $runtimeEnabledPath -Mode Execute -AllowExternalSideEffects
        Assert-VssEqual 2 $result.ExitCode 'Fachkonfigurationsfehler muessen vor Execute-Readiness mit Exitcode 2 enden.'
        Assert-VssObservableResultContract -Output $result.Output -ExpectedMode Execute -ExpectedExitCode 2
    }

    Invoke-VssTest -Name 'Zielpfad- und RequestPlan-Helper lassen sich isoliert und sicher pruefen' -Body {
        $parsed = Get-VssPipelineAst
        Assert-VssEqual 0 @($parsed.Errors).Count 'Pipeline-Skript besitzt Parserfehler.'
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            Assert-VssEqual 'Folder/Sub/Diagram.pdf' (Get-VssPipelineTargetPath -RelativeSourcePath 'Folder\Sub\Diagram.vsdx') 'Backslash-Zielmapping ist falsch.'
            Assert-VssEqual 'Folder/Sub/Diagram.pdf' (Get-VssPipelineTargetPath -RelativeSourcePath 'Folder/Sub/Diagram.vsdm') 'Slash-Zielmapping ist falsch.'
            Assert-VssEqual 'Diagram.pdf' (Get-VssPipelineTargetPath -RelativeSourcePath 'Diagram.vsd') 'Mapping einer Datei im Root ist falsch.'

            foreach ($invalidPath in @('', '/root/file.vsdx', 'C:\root\file.vsdx', '..\evil.vsdx', 'Folder\..\evil.vsdx', 'Folder\.\evil.vsdx', 'Folder\\evil.vsdx')) {
                $didThrow = $false
                try { [void](Get-VssPipelineTargetPath -RelativeSourcePath $invalidPath -ErrorAction Stop) }
                catch { $didThrow = $true }
                Assert-VssTrue $didThrow "Unsicherer relativer Quellpfad wurde akzeptiert: [$invalidPath]"
            }

            $graphConfigObject = Read-VssJsonClone $graphConfigurationPath
            $restConfigObject = Read-VssJsonClone $restConfigurationPath
            $runtimeObject = Read-VssJsonClone $runtimeDisabledPath
            $contractResult = New-VssPipelineResult -RunId ([guid]::NewGuid().ToString('D')) -Mode Validate -Status VALID -ExitCode 0 -StartedUtc ([DateTime]::UtcNow.ToString('o')) -FinishedUtc ([DateTime]::UtcNow.ToString('o'))
            Assert-VssEqual 'RunId|Mode|Status|ExitCode|StartedUtc|FinishedUtc|Stages|Files|Errors|Warnings|Placeholders' (@($contractResult.PSObject.Properties.Name) -join '|') 'Oeffentlicher Ergebnisfeldvertrag ist instabil.'
            $smallGraphPlan = New-VssGraphRequestPlan -Configuration $graphConfigObject -RuntimeConfiguration $runtimeObject -TargetRelativePath 'Folder/Diagram.pdf' -ContentLength 1024
            $largeGraphPlan = New-VssGraphRequestPlan -Configuration $graphConfigObject -RuntimeConfiguration $runtimeObject -TargetRelativePath 'Folder/Diagram.pdf' -ContentLength ([long]$runtimeObject.Upload.SessionThresholdBytes)
            $folderResolutionPlan = New-VssGraphFolderResolutionPlan -Configuration $graphConfigObject -TargetRelativePath 'Folder/Sub/Diagram.pdf'
            $restPlan = New-VssRestRequestPlan -Configuration $restConfigObject -RuntimeConfiguration $runtimeObject -TargetRelativePath 'Folder/Diagram.pdf' -ContentLength 1024
            $requestPlanContract = 'Transport|Operation|Method|Uri|Headers|BodyKind|ExternalSideEffect'
            foreach ($plan in @($smallGraphPlan, $largeGraphPlan, $restPlan)) {
                Assert-VssEqual $requestPlanContract (@($plan.PSObject.Properties.Name) -join '|') 'RequestPlan-Feldvertrag ist instabil.'
                Assert-VssTrue ([bool]$plan.ExternalSideEffect) 'Plan-Helper kennzeichnet die spaetere externe Wirkung nicht.'
            }
            Assert-VssEqual 'SmallUpload' $smallGraphPlan.Operation 'Graph-Schwellwert unterhalb der Grenze ist falsch.'
            Assert-VssEqual 'PUT' $smallGraphPlan.Method 'Kleiner Graph-Upload verwendet nicht PUT.'
            Assert-VssEqual 'CreateUploadSession' $largeGraphPlan.Operation 'Graph-Schwellwert an der Grenze ist falsch.'
            Assert-VssEqual 'POST' $largeGraphPlan.Method 'Upload-Session verwendet nicht POST.'
            Assert-VssEqual 2 @($folderResolutionPlan.Steps).Count 'Graph-Ordneraufloesung behaelt den relativen Unterbaum nicht bei.'
            Assert-VssEqual 'fail' $folderResolutionPlan.Steps[0].CreateBody['@microsoft.graph.conflictBehavior'] 'Graph-Ordneranlage ist nicht fail-closed.'
            Assert-VssEqual 'CreateIfAbsent' $restPlan.Operation 'REST-Create-Operation ist falsch.'
            Assert-VssEqual 'POST' $restPlan.Method 'REST-Upload verwendet nicht POST.'
        }
        finally {
            Remove-Module $module -Force -ErrorAction SilentlyContinue
        }
    }

    Invoke-VssTest -Name 'In-Memory-Adapter durchlaufen die echte Orchestrierung und speichern State erst nach Commit' -Body {
        $parsed = Get-VssPipelineAst
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            $configuration = Read-VssJsonClone $graphConfigurationPath
            $runtime = Read-VssJsonClone $runtimeDisabledPath
            $runtime.Execution.Enabled = $true
            $adapter = New-VssPipelineSimulationAdapterSet
            $run = Invoke-VssPipelineExecution -Configuration $configuration -RuntimeConfiguration $runtime -RunId ([guid]::NewGuid().ToString('D')) -Mode Execute -AdapterSet $adapter -AllowExternalSideEffects
            Assert-VssEqual 0 $run.ExitCode 'Vollstaendiges Fake-Adapterset muss erfolgreich durchlaufen.'
            Assert-VssEqual 'SUCCEEDED' $run.Status 'Intern freigegebener Fake-Execute-Lauf besitzt falschen Status.'
            Assert-VssEqual 1 @($run.Files).Count 'Fake-Lauf muss genau einen deterministischen Kandidaten verarbeiten.'
            Assert-VssEqual 'UPLOADED' $run.Files[0].Status 'Deterministische Fake-Datei wurde nicht uebertragen.'
            Assert-VssEqual 1 $adapter.Memory.CommitCount 'Fake-Remote erhielt nicht genau einen Commit.'
            Assert-VssEqual 1 $adapter.Memory.SaveCount 'State wurde nicht genau einmal gespeichert.'

            $trace = @($adapter.Memory.Trace)
            $expectedTrace = @(
                'GoLive.FakeApproval', 'Lock.Acquire', 'State.Load', 'Inventory', 'Stage:Simulation\Example.vsdx',
                'Convert:Simulation\Example.vsdx', 'Remote.Inspect:Simulation/Example.pdf',
                'Verify:Simulation\Example.vsdx', 'Remote.EnsureFolders:Simulation/Example.pdf',
                'Remote.Commit:simulation/example.vsdx', 'State.Save:simulation/example.vsdx',
                'Log:UploadOrUpdatePdf', 'State.ReportOrphans', 'Log:SummarizeRun',
                'Staging.Cleanup', 'Lock.Release'
            )
            Assert-VssEqual ($expectedTrace -join '|') ($trace -join '|') 'Adapter-Stagefolge ist nicht deterministisch oder Commit/State-Reihenfolge ist falsch.'
            Assert-VssTrue ($trace.IndexOf('Remote.Commit:simulation/example.vsdx') -lt $trace.IndexOf('State.Save:simulation/example.vsdx')) 'State wurde vor dem bestaetigten Commit gespeichert.'

            $stateItem = @($adapter.Memory.State.Items)[0]
            $stateFields = 'SourceKey|SourceRelativePath|SourceSha256|ConversionFingerprint|PdfSha256|TargetRelativePath|ApiKind|RemoteItemId|ETag|LastSuccessfulCommitUtc|Provenance'
            Assert-VssEqual $stateFields (@($stateItem.PSObject.Properties.Name) -join '|') 'State-Mindestfeldvertrag ist instabil.'
            Assert-VssEqual 'SIM-ITEM-0001' $stateItem.RemoteItemId 'Fake-Remote-ID fehlt im State.'
            Assert-VssEqual '"SIM-ETAG-0001"' $stateItem.ETag 'Fake-eTag fehlt im State.'

            $commitCountBefore = [int]$adapter.Memory.CommitCount
            $saveCountBefore = [int]$adapter.Memory.SaveCount
            $secondRun = Invoke-VssPipelineExecution -Configuration $configuration -RuntimeConfiguration $runtime -RunId ([guid]::NewGuid().ToString('D')) -Mode Execute -AdapterSet $adapter -AllowExternalSideEffects
            Assert-VssEqual 'SKIPPED_UNCHANGED' $secondRun.Files[0].Status 'Unveraenderte Datei wurde nicht uebersprungen.'
            Assert-VssEqual $commitCountBefore $adapter.Memory.CommitCount 'Unveraenderte Datei wurde erneut committed.'
            Assert-VssEqual $saveCountBefore $adapter.Memory.SaveCount 'Unveraenderte Datei loeste einen State-Save aus.'
        }
        finally { Remove-Module $module -Force -ErrorAction SilentlyContinue }
    }

    Invoke-VssTest -Name 'Einzeldateifehler werden gesammelt; naechster Kandidat und finally laufen weiter' -Body {
        $parsed = Get-VssPipelineAst
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            $configuration = Read-VssJsonClone $graphConfigurationPath
            $runtime = Read-VssJsonClone $runtimeDisabledPath
            $adapter = New-VssPipelineSimulationAdapterSet
            $hashA = Get-VssPipelineSha256Text -Text 'FAIL-A'
            $hashB = Get-VssPipelineSha256Text -Text 'GOOD-B'
            $memory = $adapter.Memory
            $adapter.Inventory = {
                param($request)
                [void]$memory.Trace.Add('Inventory')
                return [pscustomobject]@{
                    Complete = $true
                    Errors = @()
                    Candidates = @(
                        [pscustomobject]@{ SourcePath='memory://A-Fail.vsdx'; RelativePath='A-Fail.vsdx'; Length=1; LastWriteTimeUtc='2030-01-02T03:04:05Z'; SyntheticSha256=$hashA },
                        [pscustomobject]@{ SourcePath='memory://B-Good.vsdx'; RelativePath='B-Good.vsdx'; Length=1; LastWriteTimeUtc='2030-01-02T03:04:05Z'; SyntheticSha256=$hashB }
                    )
                }
            }.GetNewClosure()
            $baseConvert = $adapter.Convert
            $adapter.Convert = {
                param($request)
                if ([string]$request.StagedSource.RelativePath -ceq 'A-Fail.vsdx') { throw 'EXPECTED-FAKE-CONVERSION-FAILURE' }
                return & $baseConvert $request
            }.GetNewClosure()

            $run = Invoke-VssPipelineExecution -Configuration $configuration -RuntimeConfiguration $runtime -RunId ([guid]::NewGuid().ToString('D')) -Mode Simulation -AdapterSet $adapter
            Assert-VssEqual 4 $run.ExitCode 'Einzeldateifehler muss Teilerfolg/Exitcode 4 ergeben.'
            Assert-VssEqual 'PARTIAL' $run.Status 'Einzeldateifehler besitzt falschen Laufstatus.'
            Assert-VssEqual 2 @($run.Files).Count 'Nach dem ersten Dateifehler wurde nicht weitergearbeitet.'
            Assert-VssEqual 'FAILED' $run.Files[0].Status 'Erster Fake-Kandidat sollte fehlschlagen.'
            Assert-VssEqual 'UPLOADED' $run.Files[1].Status 'Zweiter Fake-Kandidat wurde nach Fehler nicht committed.'
            Assert-VssEqual 1 $adapter.Memory.CommitCount 'Nur der erfolgreiche Kandidat darf committed werden.'
            Assert-VssEqual 1 $adapter.Memory.SaveCount 'Nur der bestaetigte Commit darf State speichern.'
            Assert-VssTrue (@($adapter.Memory.Trace).Contains('Staging.Cleanup')) 'Cleanup fehlt nach Teilerfolg.'
            Assert-VssTrue (@($adapter.Memory.Trace).Contains('Lock.Release')) 'Lock-Freigabe fehlt nach Teilerfolg.'
        }
        finally { Remove-Module $module -Force -ErrorAction SilentlyContinue }
    }

    Invoke-VssTest -Name 'Unvollstaendige oder beschaedigte Inventar-/State-Lage loest keine Orphan- oder Remote-Aktion aus' -Body {
        $parsed = Get-VssPipelineAst
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            $configuration = Read-VssJsonClone $graphConfigurationPath
            $runtime = Read-VssJsonClone $runtimeDisabledPath
            $adapter = New-VssPipelineSimulationAdapterSet
            $memory = $adapter.Memory
            $baseInventory = $adapter.Inventory
            $adapter.Inventory = {
                param($request)
                $result = & $baseInventory $request
                $result.Complete = $false
                $result.Errors = @([pscustomobject]@{ RelativePath='Denied'; Operation='EnumerateFiles'; Result='ACCESS_FAILED' })
                return $result
            }.GetNewClosure()
            $run = Invoke-VssPipelineExecution -Configuration $configuration -RuntimeConfiguration $runtime -RunId ([guid]::NewGuid().ToString('D')) -Mode Simulation -AdapterSet $adapter
            Assert-VssEqual 4 $run.ExitCode 'Unvollstaendige Inventur muss trotz Weiterverarbeitung als Teilerfolg enden.'
            Assert-VssEqual 'UPLOADED' $run.Files[0].Status 'Erreichbarer Kandidat wurde bei unvollstaendiger Inventur nicht weiterverarbeitet.'
            Assert-VssTrue (-not @($adapter.Memory.Trace).Contains('State.ReportOrphans')) 'Unvollstaendige Inventur loeste Orphan-Auswertung aus.'

            $corruptAdapter = New-VssPipelineSimulationAdapterSet
            $corruptMemory = $corruptAdapter.Memory
            $corruptAdapter.LoadState = {
                param($request)
                [void]$corruptMemory.Trace.Add('State.Load')
                return [pscustomobject]@{ SchemaVersion='BROKEN'; Items=@() }
            }.GetNewClosure()
            $corruptRun = Invoke-VssPipelineExecution -Configuration $configuration -RuntimeConfiguration $runtime -RunId ([guid]::NewGuid().ToString('D')) -Mode Simulation -AdapterSet $corruptAdapter
            Assert-VssEqual 4 $corruptRun.ExitCode 'Beschaedigter State muss operativ blockieren.'
            Assert-VssTrue (-not (@($corruptAdapter.Memory.Trace) -match '^Remote\.')) 'Beschaedigter State erlaubte einen Remote-Aufruf.'
            Assert-VssTrue (@($corruptAdapter.Memory.Trace).Contains('Lock.Release')) 'Lock wurde nach beschaedigtem State nicht freigegeben.'
        }
        finally { Remove-Module $module -Force -ErrorAction SilentlyContinue }
    }

    Invoke-VssTest -Name 'Kollisionen, unbekannte Ziele und eTags ergeben sichere Commit-Entscheidungen' -Body {
        $parsed = Get-VssPipelineAst
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            foreach ($stage in @(Get-VssPipelineStageCatalog)) {
                foreach ($field in @('Name', 'EffectKind', 'Implementation', 'PlaceholderId', 'Handler')) {
                    Assert-VssTrue (@($stage.PSObject.Properties.Name).Contains($field)) "Stage-Vertragsfeld $field fehlt bei $($stage.Id)."
                }
            }
            $collisionPlans = @(Resolve-VssPipelineTargetCollisions -Candidates @(
                [pscustomobject]@{ RelativePath='Folder\A.vsd' },
                [pscustomobject]@{ RelativePath='folder\a.vsdx' }
            ))
            Assert-VssTrue ([bool]$collisionPlans[0].HasCollision -and [bool]$collisionPlans[1].HasCollision) 'Case-insensitive PDF-Zielkollision wurde nicht erkannt.'

            $absent = [pscustomobject]@{ Exists=$false; RemoteItemId=$null; ETag=$null; TargetRelativePath='Folder/A.pdf'; Provenance=$null }
            $create = Resolve-VssPipelineRemoteWriteDecision -SourceKey 'folder/a.vsdx' -TargetRelativePath 'Folder/A.pdf' -StateItem $null -RemoteTarget $absent
            Assert-VssEqual 'Create' $create.Action 'Fehlendes Ziel muss einen sicheren Create-Vertrag ergeben.'

            $unknown = [pscustomobject]@{ Exists=$true; RemoteItemId='UNKNOWN'; ETag='"u"'; TargetRelativePath='Folder/A.pdf'; Provenance=$null }
            $conflict = Resolve-VssPipelineRemoteWriteDecision -SourceKey 'folder/a.vsdx' -TargetRelativePath 'Folder/A.pdf' -StateItem $null -RemoteTarget $unknown
            Assert-VssEqual 'Conflict' $conflict.Action 'Unbekanntes vorhandenes Ziel wurde nicht blockiert.'
            Assert-VssEqual 'UnknownExistingTarget' $conflict.Reason 'Konfliktgrund fuer unbekanntes Ziel ist instabil.'

            $stateItem = [pscustomobject]@{ SourceKey='folder/a.vsdx'; TargetRelativePath='Folder/A.pdf'; RemoteItemId='KNOWN-ID'; ETag='"known"' }
            $known = [pscustomobject]@{ Exists=$true; RemoteItemId='KNOWN-ID'; ETag='"known"'; TargetRelativePath='folder/a.pdf'; Provenance='folder/a.vsdx' }
            $update = Resolve-VssPipelineRemoteWriteDecision -SourceKey 'folder/a.vsdx' -TargetRelativePath 'Folder/A.pdf' -StateItem $stateItem -RemoteTarget $known
            Assert-VssEqual 'Update' $update.Action 'Bekanntes dienstverwaltetes Ziel wurde nicht als Update erkannt.'
            Assert-VssEqual '"known"' $update.IfMatch 'Gespeicherter eTag wurde nicht als If-Match weitergegeben.'
            $known.ETag = '"foreign-change"'
            $etagConflict = Resolve-VssPipelineRemoteWriteDecision -SourceKey 'folder/a.vsdx' -TargetRelativePath 'Folder/A.pdf' -StateItem $stateItem -RemoteTarget $known
            Assert-VssEqual 'ETagConflict' $etagConflict.Reason 'Fremdaenderung wurde nicht als eTag-Konflikt erkannt.'

            $graphConfigObject = Read-VssJsonClone $graphConfigurationPath
            $restConfigObject = Read-VssJsonClone $restConfigurationPath
            $runtimeObject = Read-VssJsonClone $runtimeDisabledPath
            $graphCreatePlan = New-VssGraphRequestPlan -Configuration $graphConfigObject -RuntimeConfiguration $runtimeObject -TargetRelativePath 'Folder/A.pdf' -ContentLength 1024 -CommitDirective $create
            Assert-VssEqual '*' $graphCreatePlan.Headers['If-None-Match'] 'Graph-Create besitzt keine Create-Vorbedingung.'
            $known.ETag = '"known"'
            $graphUpdatePlan = New-VssGraphRequestPlan -Configuration $graphConfigObject -RuntimeConfiguration $runtimeObject -TargetRelativePath 'Folder/A.pdf' -ContentLength 1024 -CommitDirective $update
            Assert-VssEqual '"known"' $graphUpdatePlan.Headers['If-Match'] 'Graph-Update besitzt keine eTag-Vorbedingung.'
            Assert-VssTrue $graphUpdatePlan.Uri.Contains('KNOWN-ID') 'Graph-Update ist nicht an die bekannte Remote-ID gebunden.'
            $restCreatePlan = New-VssRestRequestPlan -Configuration $restConfigObject -RuntimeConfiguration $runtimeObject -TargetRelativePath 'Folder/A.pdf' -ContentLength 1024 -CommitDirective $create
            Assert-VssTrue $restCreatePlan.Uri.Contains('overwrite=false') 'REST-Create koennte ein unbekanntes Ziel ueberschreiben.'
            $specialPathBlocked = $false
            try { [void](New-VssRestRequestPlan -Configuration $restConfigObject -RuntimeConfiguration $runtimeObject -TargetRelativePath 'Folder/A#1.pdf' -ContentLength 1 -CommitDirective $create) }
            catch { $specialPathBlocked = $_.Exception.Message.Contains('VSS-SPREST-PATH-001') }
            Assert-VssTrue $specialPathBlocked 'Nicht verifizierter REST-ResourcePath-Fall wurde nicht mit stabiler ID blockiert.'
        }
        finally { Remove-Module $module -Force -ErrorAction SilentlyContinue }
    }

    Invoke-VssTest -Name 'Retry nutzt Fake-Uhr, respektiert Retry-After und wiederholt keine dauerhaften 4xx' -Body {
        $parsed = Get-VssPipelineAst
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            $delays = New-Object System.Collections.ArrayList
            $attempts429 = [pscustomobject]@{ Count = 0 }
            $exception429 = New-Object System.Exception('HTTP 429')
            $exception429.Data['StatusCode'] = 429
            $exception429.Data['RetryAfter'] = 7
            $operation429 = {
                param($attempt)
                $attempts429.Count++
                if ($attempts429.Count -eq 1) { throw $exception429 }
                return 'OK-429'
            }.GetNewClosure()
            $delayAction = { param([double]$seconds) [void]$delays.Add($seconds) }.GetNewClosure()
            $result429 = Invoke-VssPipelineWithRetry -Operation $operation429 -MaxRetryCount 3 -OperationName 'Fake429' -DelayAction $delayAction -JitterProvider { 0.0 }
            Assert-VssEqual 'OK-429' $result429 '429-Retry lieferte kein Ergebnis.'
            Assert-VssEqual 2 $attempts429.Count '429 wurde nicht exakt einmal wiederholt.'
            Assert-VssEqual 7 ([int]$delays[0]) 'Retry-After wurde nicht verwendet.'

            foreach ($statusCode in @(400, 401, 403, 404, 409, 412)) {
                $attempts4xx = [pscustomobject]@{ Count = 0 }
                $exception4xx = New-Object System.Exception("HTTP $statusCode")
                $exception4xx.Data['StatusCode'] = $statusCode
                $operation4xx = {
                    param($attempt)
                    $attempts4xx.Count++
                    throw $exception4xx
                }.GetNewClosure()
                $didThrow = $false
                try { [void](Invoke-VssPipelineWithRetry -Operation $operation4xx -MaxRetryCount 3 -OperationName 'Fake4xx' -DelayAction { throw 'DELAY_NOT_ALLOWED' } -JitterProvider { 0.0 }) }
                catch { $didThrow = $true }
                Assert-VssTrue $didThrow "HTTP $statusCode wurde nicht weitergegeben."
                Assert-VssEqual 1 $attempts4xx.Count "HTTP $statusCode wurde faelschlich wiederholt."
            }

            $delays503 = New-Object System.Collections.ArrayList
            $attempts503 = [pscustomobject]@{ Count = 0 }
            $exception503 = New-Object System.Exception('HTTP 503')
            $exception503.Data['StatusCode'] = 503
            $operation503 = {
                param($attempt)
                $attempts503.Count++
                if ($attempts503.Count -eq 1) { throw $exception503 }
                return 'OK-503'
            }.GetNewClosure()
            [void](Invoke-VssPipelineWithRetry -Operation $operation503 -MaxRetryCount 1 -OperationName 'Fake503' -DelayAction { param($seconds) [void]$delays503.Add([double]$seconds) }.GetNewClosure() -JitterProvider { 0.25 })
            Assert-VssEqual 2 $attempts503.Count '503 wurde nicht begrenzt wiederholt.'
            Assert-VssEqual 2.25 ([double]$delays503[0]) 'Exponentieller Backoff mit Fake-Jitter ist falsch.'

            $programAttempts = [pscustomobject]@{ Count = 0 }
            $programError = { param($attempt) $programAttempts.Count++; throw [System.InvalidOperationException]::new('PROGRAMMFEHLER') }.GetNewClosure()
            try { [void](Invoke-VssPipelineWithRetry -Operation $programError -MaxRetryCount 3 -OperationName 'ProgramError' -DelayAction { throw 'DELAY_NOT_ALLOWED' }) } catch { }
            Assert-VssEqual 1 $programAttempts.Count 'Programmierfehler ohne HTTP-Response wurde faelschlich wiederholt.'
        }
        finally { Remove-Module $module -Force -ErrorAction SilentlyContinue }
    }

    Invoke-VssTest -Name 'Graph-Upload-Session sendet teilbare 10-MiB-Fragmente mit korrekten Bereichen' -Body {
        $parsed = Get-VssPipelineAst
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            $contentPath = Join-Path $sandbox 'fake-large.pdf'
            $content = New-Object byte[] (10485760 + 1)
            [System.IO.File]::WriteAllBytes($contentPath, $content)
            $ranges = New-Object System.Collections.ArrayList
            $callState = [pscustomobject]@{ Count = 0 }
            $fakeChunkTransport = {
                param($uploadUrl, $headers, $payload, $attempt)
                [void]$attempt
                Assert-VssEqual 'https://upload.example.invalid/SENTINEL-PREAUTH' $uploadUrl 'Fake-Upload-URL wurde veraendert.'
                [void]$ranges.Add([string]$headers['Content-Range'])
                $callState.Count++
                if ($callState.Count -eq 1) {
                    Assert-VssEqual 10485760 $payload.Length 'Erstes Graph-Fragment ist nicht 10 MiB gross.'
                    return [pscustomobject]@{ nextExpectedRanges=@('10485760-') }
                }
                Assert-VssEqual 1 $payload.Length 'Letztes Graph-Fragment besitzt falsche Restgroesse.'
                return [pscustomobject]@{ id='SIM-LARGE-ID'; eTag='"large"' }
            }.GetNewClosure()
            $response = Send-VssGraphUploadSessionChunks -UploadUrl 'https://upload.example.invalid/SENTINEL-PREAUTH' -ContentPath $contentPath -ChunkSizeBytes 10485760 -MaxRetryCount 0 -AllowExternalSideEffects -ChunkTransport $fakeChunkTransport
            Assert-VssEqual 2 $callState.Count 'Graph-Zustandsautomat sendete nicht genau zwei Fragmente.'
            Assert-VssEqual 'bytes 0-10485759/10485761' $ranges[0] 'Erster Content-Range ist falsch.'
            Assert-VssEqual 'bytes 10485760-10485760/10485761' $ranges[1] 'Letzter Content-Range ist falsch.'
            Assert-VssEqual 'SIM-LARGE-ID' $response.id 'Finale Graph-Antwort wurde nicht weitergereicht.'
        }
        finally { Remove-Module $module -Force -ErrorAction SilentlyContinue }
    }

    Invoke-VssTest -Name 'State-Schreiben ist atomar und Logs verwenden eine feste redigierende Positivliste' -Body {
        $parsed = Get-VssPipelineAst
        $module = Import-VssPipelineFunctionDefinitions -Ast $parsed.Ast
        try {
            $sourcePath = 'Folder\Atomic.vsdx'
            $sourceKey = Get-VssPipelineSourceKey -RelativeSourcePath $sourcePath
            $hash = Get-VssPipelineSha256Text -Text 'atomic'
            $state = [pscustomobject][ordered]@{
                SchemaVersion='1.0'
                Items=@([pscustomobject][ordered]@{
                    SourceKey=$sourceKey; SourceRelativePath=$sourcePath; SourceSha256=$hash
                    ConversionFingerprint=$hash; PdfSha256=$hash; TargetRelativePath='Folder/Atomic.pdf'
                    ApiKind='MicrosoftGraphV1'; RemoteItemId='ATOMIC-ID'; ETag='"atomic-1"'
                    LastSuccessfulCommitUtc='2030-01-02T03:04:06Z'; Provenance=$sourceKey
                })
            }
            $statePath = Join-Path $sandbox 'atomic-state.json'
            Save-VssPipelineState -State $state -LiteralPath $statePath
            $state.Items[0].ETag = '"atomic-2"'
            Save-VssPipelineState -State $state -LiteralPath $statePath
            $loaded = Read-VssPipelineState -LiteralPath $statePath
            Assert-VssEqual '"atomic-2"' $loaded.Items[0].ETag 'Atomarer Replace enthielt nicht den letzten Commit-State.'
            Assert-VssEqual 0 @(Get-ChildItem -LiteralPath $sandbox -Filter 'atomic-state.json.*.tmp' -File).Count 'Temp-State-Datei blieb nach Replace liegen.'

            $event = [pscustomobject][ordered]@{
                RunId=[guid]::NewGuid().ToString('D'); StageId='UploadOrUpdatePdf'
                RelativePath='C:\SECRET\absolute.vsdx'; Result='Bearer SENTINEL-TOKEN https://upload.example.invalid/?sig=SENTINEL'
                DurationMs=12; RetryCount=1; HttpStatus=503; RequestId='bad request id with spaces'
                Uri='https://upload.example.invalid/?sig=SENTINEL'; Message='SENTINEL-MESSAGE'; Authorization='Bearer SENTINEL-TOKEN'
            }
            $safe = ConvertTo-VssPipelineRedactedLogObject -InputObject $event
            $safeJson = ConvertTo-Json -InputObject $safe -Compress
            Assert-VssEqual 'RunId|StageId|RelativePath|Result|DurationMs|RetryCount|HttpStatus|RequestId' (@($safe.PSObject.Properties.Name) -join '|') 'Log-Positivliste ist instabil.'
            Assert-VssTrue ($safeJson -notmatch 'SENTINEL|upload\.example|SECRET|Authorization|Message|Uri') 'Log-Redaktion liess Secret, Upload-URL oder absoluten Pfad durch.'
            Assert-VssTrue ($safeJson -match '\[REDACTED\]') 'Redigierte Werte sind nicht sichtbar markiert.'
        }
        finally { Remove-Module $module -Force -ErrorAction SilentlyContinue }
    }

    Invoke-VssTest -Name 'Bestehende Scaffold-Tests bleiben als Regression gruen' -Body {
        $legacy = Invoke-VssPowerShellFile -Path $legacyTestPath
        Assert-VssEqual 0 $legacy.ExitCode 'Bestehende Test-Scaffold.ps1 ist fehlgeschlagen.'
        Assert-VssMatches $legacy.Output 'Ergebnis:\s*[0-9]+ bestanden, 0 fehlgeschlagen\.' 'Bestehende Tests melden Fehlschlaege oder keinen stabilen Abschluss.'
    }
}
finally {
    Remove-VssTestSandbox $sandbox
}

Write-Output ''
Write-Output ("Ergebnis Pipeline: {0} bestanden, {1} fehlgeschlagen." -f $script:PassedCount, $script:FailedCount)
if ($script:FailedCount -gt 0) { exit 1 }
exit 0
