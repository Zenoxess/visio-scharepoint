[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
$corePath = Join-Path $projectRoot 'src\production\VisioSharePointSync.Core.ps1'
$adapterPath = Join-Path $projectRoot 'src\legacy\VisioSharePointSync.AdapterStubs.Legacy.ps1'
$cliPath = Join-Path $projectRoot 'src\legacy\Invoke-VisioSharePointSync.Legacy.ps1'
$decisionRegisterPath = Join-Path $projectRoot 'docs\OFFENE-FRAGEN.md'
$completePath = Join-Path $PSScriptRoot 'fixtures\complete.config.json'
$invalidPath = Join-Path $PSScriptRoot 'fixtures\invalid.config.json'
$secretPath = Join-Path $PSScriptRoot 'fixtures\secret-key.config.json'
$managedIdentityPath = Join-Path $PSScriptRoot 'fixtures\managed-identity.config.json'
$serverRestPath = Join-Path $PSScriptRoot 'fixtures\server-rest.config.json'
$incompletePath = Join-Path $projectRoot 'config\sync.example.json'

. $corePath
. $adapterPath

$script:PassedCount = 0
$script:FailedCount = 0

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
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ('vss-tests-' + [guid]::NewGuid().ToString('N'))
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

function Write-VssUtf8NoBomText {
    param([string]$Path, [string]$Text)
    $encoding = New-Object System.Text.UTF8Encoding($false, $true)
    [System.IO.File]::WriteAllText($Path, $Text, $encoding)
}

function Get-VssConfigurationClone {
    param([string]$Path = $completePath)
    return (Read-VssUtf8FileStrict -LiteralPath $Path | ConvertFrom-Json)
}

function Write-VssConfiguration {
    param([object]$Configuration, [string]$Path)
    $json = $Configuration | ConvertTo-Json -Depth 12
    Write-VssUtf8NoBomText -Path $Path -Text $json
}

function Set-VssTestConfigurationValue {
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowNull()][object]$Value
    )
    $segments = @($Path.Split('.'))
    $container = $Configuration
    for ($index = 0; $index -lt ($segments.Count - 1); $index++) {
        $container = $container.PSObject.Properties[$segments[$index]].Value
    }
    $container.PSObject.Properties[$segments[$segments.Count - 1]].Value = $Value
}

function Invoke-VssCli {
    param(
        [string]$ConfigurationPath,
        [AllowNull()][string]$Mode,
        [AllowNull()][string]$EntryPointPath
    )
    if ([string]::IsNullOrWhiteSpace($EntryPointPath)) { $EntryPointPath = $cliPath }
    $hostExecutable = (Get-Process -Id $PID).Path
    if ([string]::IsNullOrWhiteSpace($Mode)) {
        $output = @(& $hostExecutable -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $EntryPointPath -ConfigurationPath $ConfigurationPath 2>&1)
    }
    else {
        $output = @(& $hostExecutable -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $EntryPointPath -ConfigurationPath $ConfigurationPath -Mode $Mode 2>&1)
    }
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($output -join [Environment]::NewLine) }
}

Invoke-VssTest -Name 'Vollstaendige Konfiguration und Ergebnisvertrag' -Body {
    $result = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $completePath -Mode Validate
    Assert-VssTrue $result.IsValid 'Die vollstaendige Fixture muss gueltig sein.'
    Assert-VssEqual 0 @($result.Errors).Count 'Vollstaendige Fixture hat Fehler.'
    Assert-VssEqual 0 @($result.UnresolvedDecisionIds).Count 'Vollstaendige Fixture hat offene Entscheidungen.'
    Assert-VssEqual 0 (Get-VisioSharePointSyncExitCode $result) 'Vollstaendige Fixture muss Exitcode 0 abbilden.'
    $propertyContract = @($result.PSObject.Properties.Name) -join '|'
    Assert-VssEqual 'IsValid|Errors|Warnings|UnresolvedDecisionIds|PlannedStages' $propertyContract 'Ergebnisfelder wurden veraendert.'
}

Invoke-VssTest -Name 'Managed Identity und Server REST sind ohne Entra-Dummywerte gueltig' -Body {
    $managedResult = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $managedIdentityPath
    $serverResult = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $serverRestPath
    Assert-VssTrue $managedResult.IsValid 'ManagedIdentity ohne Zertifikatsfelder muss gueltig sein.'
    Assert-VssTrue $serverResult.IsValid 'Server REST mit WindowsIntegrated muss ohne Entra-Felder gueltig sein.'
    Assert-VssEqual 0 (Get-VisioSharePointSyncExitCode $managedResult) 'ManagedIdentity Exitcode ist falsch.'
    Assert-VssEqual 0 (Get-VisioSharePointSyncExitCode $serverResult) 'Server REST Exitcode ist falsch.'
}

Invoke-VssTest -Name 'API-Zielzweige und Authentifizierungsarten bilden geschlossene Varianten' -Body {
    $onlineWindows = Get-VssConfigurationClone -Path $managedIdentityPath
    $onlineWindows.SharePoint.AuthenticationKind = 'WindowsIntegrated'
    Assert-VssTrue (-not (Test-VssConfigurationObject $onlineWindows).IsValid) 'Online+Graph+WindowsIntegrated wurde akzeptiert.'

    $serverManaged = Get-VssConfigurationClone -Path $serverRestPath
    $serverManaged.SharePoint.AuthenticationKind = 'ManagedIdentity'
    Assert-VssTrue (-not (Test-VssConfigurationObject $serverManaged).IsValid) 'Server+REST+ManagedIdentity wurde akzeptiert.'

    $managedWithCertificate = Get-VssConfigurationClone -Path $managedIdentityPath
    $managedWithCertificate.SharePoint.CertificateThumbprint = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    Assert-VssTrue (-not (Test-VssConfigurationObject $managedWithCertificate).IsValid) 'ManagedIdentity mit Zertifikatsfeld wurde akzeptiert.'

    $workloadMissingReference = Get-VssConfigurationClone -Path $managedIdentityPath
    $workloadMissingReference.SharePoint.AuthenticationKind = 'WorkloadIdentity'
    $workloadMissingReference.SharePoint.TenantId = '11111111-1111-4111-8111-111111111111'
    $workloadMissingReference.SharePoint.ClientId = '22222222-2222-4222-8222-222222222222'
    $workloadMissing = Test-VssConfigurationObject $workloadMissingReference
    Assert-VssTrue (-not $workloadMissing.IsValid) 'WorkloadIdentity ohne Dateireferenz wurde akzeptiert.'
    Assert-VssTrue (@($workloadMissing.UnresolvedDecisionIds) -contains 'SEC-001') 'Fehlende Workload-Dateireferenz meldet SEC-001 nicht.'

    $workload = Get-VssConfigurationClone -Path $managedIdentityPath
    $workload.SharePoint.AuthenticationKind = 'WorkloadIdentity'
    $workload.SharePoint.TenantId = '11111111-1111-4111-8111-111111111111'
    $workload.SharePoint.ClientId = '22222222-2222-4222-8222-222222222222'
    $workload.SharePoint.WorkloadIdentityFilePath = 'C:\ProgramData\VisioSharePointSync\federated-token.txt'
    Assert-VssTrue (Test-VssConfigurationObject $workload).IsValid 'Vollstaendige WorkloadIdentity wurde abgewiesen.'

    $workload.SharePoint.WorkloadIdentityFilePath = '..\token.txt'
    Assert-VssTrue (-not (Test-VssConfigurationObject $workload).IsValid) 'Relative Workload-Dateireferenz wurde akzeptiert.'

    $workload.SharePoint.WorkloadIdentityFilePath = $workload.Operations.StagingPath + '\federated-token.txt'
    Assert-VssTrue (-not (Test-VssConfigurationObject $workload).IsValid) 'Workload-Dateireferenz innerhalb von Staging wurde akzeptiert.'

    $delegated = Get-VssConfigurationClone
    $delegated.SharePoint.AuthenticationKind = 'Delegated'
    $delegated.SharePoint.CertificateStoreLocation = $null
    $delegated.SharePoint.CertificateThumbprint = $null
    Assert-VssTrue (-not (Test-VssConfigurationObject $delegated).IsValid) 'Delegated wurde fuer das unbeaufsichtigte Geruest akzeptiert.'

    $graphWithRestTarget = Get-VssConfigurationClone
    $graphWithRestTarget.SharePoint.SiteUrl = 'https://sharepoint.example.invalid/sites/engineering'
    Assert-VssTrue (-not (Test-VssConfigurationObject $graphWithRestTarget).IsValid) 'Graph-Konfiguration mit REST-Zielarm wurde akzeptiert.'

    $restWithGraphTarget = Get-VssConfigurationClone -Path $serverRestPath
    $restWithGraphTarget.SharePoint.SiteId = 'unexpected-id'
    Assert-VssTrue (-not (Test-VssConfigurationObject $restWithGraphTarget).IsValid) 'REST-Konfiguration mit Graph-Zielarm wurde akzeptiert.'

    $onlineRest = Get-VssConfigurationClone
    $onlineRest.SharePoint.ApiKind = 'SharePointRest'
    $onlineRest.SharePoint.SiteId = $null
    $onlineRest.SharePoint.DriveId = $null
    $onlineRest.SharePoint.TargetFolderId = $null
    $onlineRest.SharePoint.SiteUrl = 'https://sharepoint.example.invalid/sites/engineering'
    $onlineRest.SharePoint.LibraryName = 'Documents'
    $onlineRest.SharePoint.TargetFolderPath = 'Published/Visio'
    Assert-VssTrue (Test-VssConfigurationObject $onlineRest).IsValid 'Bewusst unterstuetzter Online+REST-Zweig wurde abgewiesen.'
}

Invoke-VssTest -Name 'Ungueltige und unvollstaendige Konfigurationen bleiben erwartbar' -Body {
    $invalid = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $invalidPath
    $first = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $incompletePath
    $second = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $incompletePath
    Assert-VssTrue (-not $invalid.IsValid) 'Ungueltige Fixture wurde akzeptiert.'
    Assert-VssTrue ((@($invalid.Errors) -join '|') -match 'SharePointServer kann nicht mit MicrosoftGraphV1') 'Server+Graph-Konflikt fehlt.'
    Assert-VssEqual 2 (Get-VisioSharePointSyncExitCode $invalid) 'Ungueltige Fixture muss Exitcode 2 liefern.'
    Assert-VssEqual 0 @($first.Errors).Count 'Offene Beispielwerte duerfen keine Strukturfehler erzeugen.'
    Assert-VssEqual (@($first.UnresolvedDecisionIds) -join '|') (@($second.UnresolvedDecisionIds) -join '|') 'Offene IDs sind nicht stabil.'
    foreach ($blocker in @(Get-VssBlockerDecisionIds)) {
        Assert-VssTrue (@($first.UnresolvedDecisionIds) -contains $blocker) 'Ein Blocker fehlt im unvollstaendigen Ergebnis.'
    }
}

Invoke-VssTest -Name 'Ungueltige Auswahlwerte nennen ihre konkrete Frage-ID' -Body {
    $selectionCases = @(
        [pscustomobject]@{ Path = 'SchemaVersion'; Value = '9.9'; QuestionId = 'ENV-004' },
        [pscustomobject]@{ Path = 'Source.Extensions'; Value = @('.unsafe'); QuestionId = 'SRC-003' },
        [pscustomobject]@{ Path = 'Conversion.Provider'; Value = 'Unsupported'; QuestionId = 'CNV-001' },
        [pscustomobject]@{ Path = 'Conversion.OutputFormat'; Value = 'XPS'; QuestionId = 'CNV-003' },
        [pscustomobject]@{ Path = 'Conversion.PageRange'; Value = 'SomePages'; QuestionId = 'CNV-003' },
        [pscustomobject]@{ Path = 'Conversion.Intent'; Value = 'Archive'; QuestionId = 'CNV-003' },
        [pscustomobject]@{ Path = 'SharePoint.Platform'; Value = 'OtherPlatform'; QuestionId = 'SP-001' },
        [pscustomobject]@{ Path = 'SharePoint.ApiKind'; Value = 'OtherApi'; QuestionId = 'SP-001' },
        [pscustomobject]@{ Path = 'SharePoint.AuthenticationKind'; Value = 'OtherAuth'; QuestionId = 'SEC-001' },
        [pscustomobject]@{ Path = 'SharePoint.CertificateStoreLocation'; Value = 'OtherStore'; QuestionId = 'SEC-003' },
        [pscustomobject]@{ Path = 'Sync.Direction'; Value = 'OtherDirection'; QuestionId = 'SYN-001' },
        [pscustomobject]@{ Path = 'Sync.ConflictPolicy'; Value = 'OtherConflict'; QuestionId = 'SYN-003' },
        [pscustomobject]@{ Path = 'Sync.RenamePolicy'; Value = 'OtherRename'; QuestionId = 'SYN-005' },
        [pscustomobject]@{ Path = 'Sync.DeletePolicy'; Value = 'OtherDelete'; QuestionId = 'SYN-005' },
        [pscustomobject]@{ Path = 'Sync.FingerprintMode'; Value = 'OtherFingerprint'; QuestionId = 'SYN-002' }
    )
    foreach ($selectionCase in $selectionCases) {
        $configuration = Get-VssConfigurationClone
        Set-VssTestConfigurationValue -Configuration $configuration -Path $selectionCase.Path -Value $selectionCase.Value
        $result = Test-VssConfigurationObject $configuration
        $diagnostics = (@($result.Errors) + @($result.UnresolvedDecisionIds)) -join '|'
        Assert-VssTrue ($diagnostics.Contains($selectionCase.QuestionId)) 'Auswahlfehler enthaelt die verknuepfte Frage-ID nicht.'
    }
}

Invoke-VssTest -Name 'Blocker-Gate stimmt exakt mit dem Markdown-Register ueberein' -Body {
    $expected = @(
        'ENV-001', 'ENV-002', 'ENV-003', 'SRC-001', 'SRC-002',
        'CNV-001', 'CNV-002', 'SP-001', 'SP-002', 'SEC-001',
        'SEC-002', 'SYN-001', 'SYN-002', 'SYN-003', 'SYN-005',
        'OPS-001', 'ACC-001', 'ACC-003'
    )
    Assert-VssEqual ($expected -join '|') (@(Get-VssBlockerDecisionIds) -join '|') 'Hardcodierte Blockerliste ist falsch.'

    $markdown = [System.IO.File]::ReadAllText($decisionRegisterPath, (New-Object System.Text.UTF8Encoding($false, $true)))
    $matches = [regex]::Matches($markdown, '(?m)^\|\s*([A-Z]+-[0-9]{3})\s*\|\s*Blocker\s*\|')
    $markdownIds = @($matches | ForEach-Object { $_.Groups[1].Value })
    Assert-VssEqual 18 $markdownIds.Count 'Markdown muss exakt 18 Blocker enthalten.'
    Assert-VssEqual ($expected -join '|') ($markdownIds -join '|') 'Blocker in Code und Markdown weichen ab.'

    $allMatches = [regex]::Matches($markdown, '(?m)^\|\s*([A-Z]+-[0-9]{3})\s*\|')
    $allMarkdownIds = @($allMatches | ForEach-Object { $_.Groups[1].Value })
    Assert-VssEqual 38 $allMarkdownIds.Count 'Markdown muss exakt 38 registrierte Fragen enthalten.'
    $referencedIds = @($script:VssConfigurationFields | ForEach-Object { $_.QuestionId }) + @(Get-VssBlockerDecisionIds)
    foreach ($referencedId in @($referencedIds | Select-Object -Unique)) {
        Assert-VssTrue (Test-VssStringArrayContainsOrdinal -Values $allMarkdownIds -Expected $referencedId) 'Eine im Validator referenzierte Frage-ID fehlt im Markdown.'
    }

    $sandbox = New-VssTestSandbox
    try {
        $configuration = Get-VssConfigurationClone
        $configuration.ResolvedDecisionIds = @($configuration.ResolvedDecisionIds | Where-Object { $_ -ne 'ACC-003' })
        $partialPath = Join-Path $sandbox 'partial.json'
        Write-VssConfiguration $configuration $partialPath
        $partial = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $partialPath
        Assert-VssEqual 'ACC-003' (@($partial.UnresolvedDecisionIds) -join '|') 'Fehlender Blocker wurde nicht exakt gegated.'
        Assert-VssEqual 2 (Get-VisioSharePointSyncExitCode $partial) 'Teilweise Blockerfreigabe muss Exitcode 2 liefern.'

        $configuration.ResolvedDecisionIds += 'ACC-999'
        $unknownIdPath = Join-Path $sandbox 'unknown-id.json'
        Write-VssConfiguration $configuration $unknownIdPath
        $unknownId = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $unknownIdPath
        Assert-VssTrue (@($unknownId.Errors).Count -gt 0) 'Unbekannte Decision-ID wurde akzeptiert.'
        Assert-VssTrue ((@($unknownId.Errors) -join '|') -notmatch 'ACC-999') 'Benutzerkontrollierte Decision-ID wurde ausgegeben.'

        $configuration = Get-VssConfigurationClone
        $configuration.ResolvedDecisionIds += 'SEC-003'
        $nonBlockerPath = Join-Path $sandbox 'non-blocker-id.json'
        Write-VssConfiguration $configuration $nonBlockerPath
        $nonBlocker = Invoke-VisioSharePointSyncAssessment $nonBlockerPath
        Assert-VssTrue (@($nonBlocker.Errors).Count -gt 0) 'Bekannte Nicht-Blocker-ID wurde als Freigabe-Token akzeptiert.'
        Assert-VssTrue ((@($nonBlocker.Errors) -join '|') -notmatch 'SEC-003') 'Nicht zugelassene Decision-ID wurde ausgegeben.'

        $configuration = Get-VssConfigurationClone
        $configuration.ResolvedDecisionIds = @($configuration.ResolvedDecisionIds | ForEach-Object {
            if ($_ -ceq 'ENV-001') { 'env-001' } else { $_ }
        })
        $wrongCasePath = Join-Path $sandbox 'wrong-case-id.json'
        Write-VssConfiguration $configuration $wrongCasePath
        $wrongCase = Invoke-VisioSharePointSyncAssessment $wrongCasePath
        Assert-VssTrue (@($wrongCase.Errors).Count -gt 0) 'Falsch geschriebene Blocker-ID wurde akzeptiert.'
        Assert-VssTrue (@($wrongCase.UnresolvedDecisionIds | Where-Object { $_ -ceq 'ENV-001' }).Count -eq 1) 'Ordinaler Blocker-Gate meldet ENV-001 nicht.'
        Assert-VssTrue ((@($wrongCase.Errors) -join '|') -notmatch 'env-001') 'Benutzerkontrollierte Blocker-ID wurde ausgegeben.'
    }
    finally { Remove-VssTestSandbox $sandbox }
}

Invoke-VssTest -Name 'BOM-loses striktes UTF-8 erhaelt Muenchen korrekt' -Body {
    $sandbox = New-VssTestSandbox
    try {
        $configuration = Get-VssConfigurationClone
        $munich = 'M' + [char]0x00FC + 'nchen'
        $configuration.Source.FileNamePattern = '^' + $munich + '-.*\.vsdx$'
        $utf8Path = Join-Path $sandbox 'unicode.json'
        Write-VssConfiguration $configuration $utf8Path
        $bytes = [System.IO.File]::ReadAllBytes($utf8Path)
        Assert-VssTrue (-not (($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB) -and ($bytes[2] -eq 0xBF))) 'Testdatei besitzt unerwartet einen BOM.'
        Assert-VssTrue ((Read-VssUtf8FileStrict $utf8Path).Contains($munich)) 'UTF-8-Umlaut wurde beim Lesen veraendert.'
        Assert-VssTrue (Invoke-VisioSharePointSyncAssessment $utf8Path).IsValid 'BOM-lose UTF-8-Konfiguration mit Umlaut ist ungueltig.'

        $json = $configuration | ConvertTo-Json -Depth 12
        $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
        $bomPath = Join-Path $sandbox 'bom.json'
        [System.IO.File]::WriteAllBytes($bomPath, [byte[]](@(0xEF, 0xBB, 0xBF) + @($utf8.GetBytes($json))))
        Assert-VssEqual 2 (Get-VisioSharePointSyncExitCode (Invoke-VisioSharePointSyncAssessment $bomPath)) 'BOM muss erwartbar abgewiesen werden.'

        $invalidUtf8Path = Join-Path $sandbox 'invalid-utf8.json'
        [System.IO.File]::WriteAllBytes($invalidUtf8Path, [byte[]](0x7B, 0x22, 0xC3, 0x28, 0x22, 0x3A, 0x31, 0x7D))
        Assert-VssEqual 2 (Get-VisioSharePointSyncExitCode (Invoke-VisioSharePointSyncAssessment $invalidUtf8Path)) 'Ungueltiges UTF-8 muss Exitcode 2 liefern.'
    }
    finally { Remove-VssTestSandbox $sandbox }
}

Invoke-VssTest -Name 'Secret- und Unknown-Key-Diagnostik ist geschlossen und redigiert' -Body {
    $secret = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $secretPath
    $secretText = @($secret.Errors) -join [Environment]::NewLine
    Assert-VssTrue (@($secret.Errors).Count -ge 1) 'Secret-Key wurde nicht abgewiesen.'
    Assert-VssTrue ($secretText -notmatch 'ClientSecret|TOP-SECRET|SENTINEL-SECRET-VALUE') 'Secret-Key oder Wert wurde in Diagnose ausgegeben.'
    foreach ($keyName in @('Authorization', 'UploadUrl', 'ClientAssertion', 'PfxPassphrase', 'ClientSecret-TOP-SECRET')) {
        Assert-VssTrue (Test-VssSecretLikeKeyName $keyName) 'Erweiterter Secret-Key wurde nicht erkannt.'
    }

    $sandbox = New-VssTestSandbox
    try {
        $controlledKey = 'UnexpectedSetting-USER-CONTROLLED'
        $controlledValue = 'UNKNOWN-SENTINEL-VALUE'
        foreach ($location in @('Root', 'Section')) {
            $configuration = Get-VssConfigurationClone
            if ($location -eq 'Root') { $configuration | Add-Member -NotePropertyName $controlledKey -NotePropertyValue $controlledValue }
            else { $configuration.Source | Add-Member -NotePropertyName $controlledKey -NotePropertyValue $controlledValue }
            $path = Join-Path $sandbox ("unknown-$location.json")
            Write-VssConfiguration $configuration $path
            $result = Invoke-VisioSharePointSyncAssessment $path
            $formatted = @(Format-VisioSharePointSyncAssessment $result) -join [Environment]::NewLine
            Assert-VssTrue (@($result.Errors).Count -gt 0) 'Unbekannter Key wurde akzeptiert.'
            Assert-VssTrue ($formatted -notmatch 'UnexpectedSetting|USER-CONTROLLED|UNKNOWN-SENTINEL-VALUE') 'Unknown-Key-Diagnose leakt Benutzertext.'
        }

        $wrongCaseKeyConfiguration = Get-VssConfigurationClone
        $schemaValue = $wrongCaseKeyConfiguration.SchemaVersion
        $wrongCaseKeyConfiguration.PSObject.Properties.Remove('SchemaVersion')
        $wrongCaseKeyConfiguration | Add-Member -NotePropertyName 'schemaVersion' -NotePropertyValue $schemaValue
        $wrongCaseKeyPath = Join-Path $sandbox 'wrong-case-key.json'
        Write-VssConfiguration $wrongCaseKeyConfiguration $wrongCaseKeyPath
        Assert-VssTrue (@((Invoke-VisioSharePointSyncAssessment $wrongCaseKeyPath).Errors).Count -gt 0) 'Falsch geschriebener Schema-Key wurde akzeptiert.'
    }
    finally { Remove-VssTestSandbox $sandbox }
}

Invoke-VssTest -Name 'UNC- und lokale Pfade werden rein syntaktisch gehaertet' -Body {
    foreach ($validUnc in @('\\server\share', '\\fileserver.example.invalid\drawings\sub folder')) {
        Assert-VssTrue (Test-VssUncRootPathSyntax $validUnc) 'Gueltiger UNC-Pfad wurde abgewiesen.'
    }
    foreach ($invalidUnc in @(
        'Z:\Drawings', 'relative\drawings', '\\', '\\server', '\\?\C:\Windows',
        '\\.\GLOBALROOT\Device', '\\server\share\..\other', '\\server\share\bad.',
        '\\server\share\file:ads', '\\server\share\CON', '\\server\share\CLOCK$',
        '\\server\share\CONIN$', ('\\server\share\COM' + [char]0x00B9), '\\server\share\*'
    )) {
        Assert-VssTrue (-not (Test-VssUncRootPathSyntax $invalidUnc)) 'Unsicherer UNC-Pfad wurde akzeptiert.'
    }

    foreach ($validLocal in @('C:\ProgramData\Vss\state.json', 'D:\Vss\staging')) {
        Assert-VssTrue (Test-VssAbsoluteLocalPathSyntax $validLocal) 'Gueltiger lokaler Pfad wurde abgewiesen.'
    }
    foreach ($invalidLocal in @(
        'state.json', 'C:state.json', 'C:\', '\\server\share', '\\?\C:\Temp\state.json',
        'C:\data\..\state.json', 'C:\temp\*', 'C:\Temp\CON', 'C:\Temp\state.json:ads'
    )) {
        Assert-VssTrue (-not (Test-VssAbsoluteLocalPathSyntax $invalidLocal)) 'Unsicherer lokaler Pfad wurde akzeptiert.'
    }
    Assert-VssTrue (Test-VssLocalPathOverlap @('C:\Data', 'c:\data\child')) 'Verschachtelte Pfade wurden nicht erkannt.'
    Assert-VssTrue (-not (Test-VssLocalPathOverlap @('C:\Data', 'C:\Database'))) 'Segmentverschiedene Geschwister wurden als Ueberlappung erkannt.'

    $configuration = Get-VssConfigurationClone
    $configuration.Source.RootPath = '\\?\C:\Windows'
    $configuration.Operations.StatePath = 'C:\Vss\staging\state.json'
    $configuration.Operations.StagingPath = 'C:\Vss\staging'
    $configuration.Operations.LogPath = 'C:\Vss\logs'
    $result = Test-VssConfigurationObject $configuration
    Assert-VssTrue (@($result.Errors).Count -ge 2) 'UNC-/Overlap-Fehler wurden nicht gemeinsam erkannt.'

    $coreSource = [System.IO.File]::ReadAllText($corePath, (New-Object System.Text.UTF8Encoding($false, $true)))
    Assert-VssTrue ($coreSource -notmatch '\bResolve-Path\b|\bTest-Path\b') 'Produktivvalidator enthaelt aufloesende Pfadzugriffe.'
}

Invoke-VssTest -Name 'Regex-Timeout wird begrenzt und produktiv kompiliert' -Body {
    foreach ($timeout in @(50, 5000)) {
        $regex = New-VssSourceRegex -Pattern '^PPSI-.*\.vsdx$' -TimeoutMilliseconds $timeout
        Assert-VssEqual $timeout ([int]$regex.MatchTimeout.TotalMilliseconds) 'Regex-Timeout wurde nicht uebernommen.'
    }
    foreach ($badFactoryTimeout in @(49, 5001)) {
        $factoryRejected = $false
        try { [void](New-VssSourceRegex -Pattern '^PPSI-.*\.vsdx$' -TimeoutMilliseconds $badFactoryTimeout -ErrorAction Stop) }
        catch { $factoryRejected = $true }
        Assert-VssTrue $factoryRejected 'Regex-Factory hat einen Timeout ausserhalb 50..5000 akzeptiert.'
    }
    $sandbox = New-VssTestSandbox
    try {
        foreach ($badTimeout in @(49, 5001, '500', $null)) {
            $configuration = Get-VssConfigurationClone
            $configuration.Source.RegexTimeoutMilliseconds = $badTimeout
            $path = Join-Path $sandbox ('timeout-' + [guid]::NewGuid().ToString('N') + '.json')
            Write-VssConfiguration $configuration $path
            Assert-VssTrue (-not (Invoke-VisioSharePointSyncAssessment $path).IsValid) 'Ungueltiger Regex-Timeout wurde akzeptiert.'
        }
        $configuration = Get-VssConfigurationClone
        $configuration.Source.FileNamePattern = '['
        $badPatternPath = Join-Path $sandbox 'bad-pattern.json'
        Write-VssConfiguration $configuration $badPatternPath
        Assert-VssTrue (@((Invoke-VisioSharePointSyncAssessment $badPatternPath).Errors).Count -gt 0) 'Ungueltiger Regex wurde akzeptiert.'

        $timedRegex = New-VssSourceRegex -Pattern '^(a+)+$' -TimeoutMilliseconds 50
        $didTimeout = $false
        try { [void]$timedRegex.IsMatch(('a' * 20000) + '!') }
        catch [System.Text.RegularExpressions.RegexMatchTimeoutException] { $didTimeout = $true }
        Assert-VssTrue $didTimeout 'Katastrophaler Regex wurde nicht durch Timeout begrenzt.'
    }
    finally { Remove-VssTestSandbox $sandbox }
}

Invoke-VssTest -Name 'SharePoint-URL ist HTTPS-only und frei von Query und Fragment' -Body {
    Assert-VssTrue (Test-VssAbsoluteHttpUrlSyntax 'https://sharepoint.example.invalid/sites/engineering') 'Gueltige HTTPS-SiteUrl wurde abgewiesen.'
    foreach ($invalidUrl in @(
        'http://sharepoint.example.invalid/sites/engineering',
        ' https://sharepoint.example.invalid/sites/engineering ',
        'https://sharepoint.example.invalid/sites/engineering?x=1',
        'https://sharepoint.example.invalid/sites/engineering#fragment',
        'https://user@sharepoint.example.invalid/sites/engineering',
        'https://sharepoint.example.invalid/sites/a/../admin',
        'https://sharepoint.example.invalid/sites/a/%2e%2e/admin',
        'https://sharepoint.example.invalid/sites/a/%2E./admin',
        'https://sharepoint.example.invalid/sites/a%2f..%2fadmin'
    )) {
        Assert-VssTrue (-not (Test-VssAbsoluteHttpUrlSyntax $invalidUrl)) 'Unsichere SiteUrl wurde akzeptiert.'
    }
}

Invoke-VssTest -Name 'Graph-IDs bleiben sichere undurchsichtige Einzelsegmente' -Body {
    Assert-VssTrue (Test-VssGraphSiteIdSyntax 'example.sharepoint.com,33333333-3333-4333-8333-333333333333,44444444-4444-4444-8444-444444444444') 'Gueltige Graph-SiteId wurde abgewiesen.'
    foreach ($validOpaqueId in @('b!dummy-drive-id', '01DUMMYTARGETFOLDERID', 'opaque_id.with-allowed~chars')) {
        Assert-VssTrue (Test-VssGraphOpaqueIdSyntax $validOpaqueId) 'Gueltige opaque Graph-ID wurde abgewiesen.'
    }

    $invalidCases = @(
        [pscustomobject]@{ Path = 'SharePoint.SiteId'; Value = '../evil?x#fragment' },
        [pscustomobject]@{ Path = 'SharePoint.SiteId'; Value = ' example.sharepoint.com,33333333-3333-4333-8333-333333333333,44444444-4444-4444-8444-444444444444' },
        [pscustomobject]@{ Path = 'SharePoint.DriveId'; Value = '../../drive' },
        [pscustomobject]@{ Path = 'SharePoint.DriveId'; Value = 'drive?query' },
        [pscustomobject]@{ Path = 'SharePoint.TargetFolderId'; Value = 'folder/child' },
        [pscustomobject]@{ Path = 'SharePoint.TargetFolderId'; Value = ('folder' + [char]10 + 'child') },
        [pscustomobject]@{ Path = 'SharePoint.TargetFolderId'; Value = 'folder child' },
        [pscustomobject]@{ Path = 'SharePoint.TargetFolderId'; Value = '..' }
    )
    foreach ($invalidCase in $invalidCases) {
        $configuration = Get-VssConfigurationClone
        Set-VssTestConfigurationValue -Configuration $configuration -Path $invalidCase.Path -Value $invalidCase.Value
        $result = Test-VssConfigurationObject $configuration
        Assert-VssTrue (-not $result.IsValid) 'Unsichere Graph-ID wurde akzeptiert.'
        Assert-VssTrue ((@($result.Errors) -join '|').Contains('SP-002')) 'Graph-ID-Fehler enthaelt SP-002 nicht.'
    }
}

Invoke-VssTest -Name 'Doppelte JSON-Schluessel werden vor ConvertFrom-Json redigiert abgewiesen' -Body {
    Assert-VssTrue (Test-VssJsonHasDuplicateObjectKeys '{"SchemaVersion":"9.9","SchemaVersion":"1.0"}') 'Exakt doppelter JSON-Key wurde nicht erkannt.'
    Assert-VssTrue (Test-VssJsonHasDuplicateObjectKeys '{"Source":{"RootPath":"one","rootpath":"two"}}') 'Case-insensitiv doppelter JSON-Key wurde nicht erkannt.'
    Assert-VssTrue (Test-VssJsonHasDuplicateObjectKeys '{"\u0053chemaVersion":"9.9","SchemaVersion":"1.0"}') 'Unicode-escaped doppelter JSON-Key wurde nicht erkannt.'
    Assert-VssTrue (-not (Test-VssJsonHasDuplicateObjectKeys '{"A":1,"Nested":{"A":2}}')) 'Gleicher Key in getrennten Scopes wurde faelschlich abgewiesen.'

    $sandbox = New-VssTestSandbox
    try {
        $configurationText = Read-VssUtf8FileStrict -LiteralPath $completePath
        $duplicateText = $configurationText.Replace(
            '"SchemaVersion": "1.0",',
            '"Duplicate-SENTINEL": 1, "duplicate-sentinel": 2, "SchemaVersion": "1.0",'
        )
        $duplicatePath = Join-Path $sandbox 'duplicate.json'
        Write-VssUtf8NoBomText -Path $duplicatePath -Text $duplicateText
        $result = Invoke-VisioSharePointSyncAssessment $duplicatePath
        $formatted = @(Format-VisioSharePointSyncAssessment $result) -join [Environment]::NewLine
        Assert-VssTrue (-not $result.IsValid) 'Doppelte JSON-Keys wurden im produktiven Lesepfad akzeptiert.'
        Assert-VssEqual 2 (Get-VisioSharePointSyncExitCode $result) 'Doppelte JSON-Keys muessen Exitcode 2 liefern.'
        Assert-VssTrue ($formatted -notmatch 'Duplicate-SENTINEL|duplicate-sentinel') 'Duplicate-Key-Diagnose leakt Benutzertext.'
    }
    finally { Remove-VssTestSandbox $sandbox }
}

Invoke-VssTest -Name 'Graph- und REST-Stubs sind getrennt und DryRun ruft keinen auf' -Body {
    foreach ($stubName in @('Invoke-VssMicrosoftGraphAdapter', 'Invoke-VssSharePointRestAdapter')) {
        Assert-VssTrue ($null -ne (Get-Command $stubName -ErrorAction SilentlyContinue)) 'Adapter-Stub fehlt.'
        $threw = $false
        try { & $stubName -PdfArtifact ([pscustomobject]@{}) -Configuration ([pscustomobject]@{}) }
        catch { $threw = $true }
        Assert-VssTrue $threw 'Adapter-Stub ist unerwartet ausfuehrbar.'
    }
    Assert-VssTrue ($null -eq (Get-Command 'Invoke-VssSharePointAdapter' -ErrorAction SilentlyContinue)) 'Alter mehrdeutiger SharePoint-Stub existiert noch.'

    $sandbox = New-VssTestSandbox
    try {
        $configuration = Get-VssConfigurationClone
        $configuration.Operations.StatePath = Join-Path $sandbox 'outputs\state.json'
        $configuration.Operations.StagingPath = Join-Path $sandbox 'staging'
        $configuration.Operations.LogPath = Join-Path $sandbox 'logs'
        $path = Join-Path $sandbox 'dryrun.json'
        Write-VssConfiguration $configuration $path

        $script:GraphCalls = 0
        $script:RestCalls = 0
        function Invoke-VssMicrosoftGraphAdapter { param($PdfArtifact, $Configuration) $script:GraphCalls++ }
        function Invoke-VssSharePointRestAdapter { param($PdfArtifact, $Configuration) $script:RestCalls++ }
        $result = Invoke-VisioSharePointSyncAssessment -ConfigurationPath $path -Mode DryRun
        Assert-VssTrue $result.IsValid 'Gueltiger DryRun ist fehlgeschlagen.'
        Assert-VssEqual 0 $script:GraphCalls 'DryRun hat den Graph-Adapter aufgerufen.'
        Assert-VssEqual 0 $script:RestCalls 'DryRun hat den REST-Adapter aufgerufen.'
        Assert-VssTrue (-not [System.IO.File]::Exists($configuration.Operations.StatePath)) 'DryRun hat State geschrieben.'
        Assert-VssTrue (-not [System.IO.Directory]::Exists($configuration.Operations.StagingPath)) 'DryRun hat Staging angelegt.'
        Assert-VssTrue (-not [System.IO.Directory]::Exists($configuration.Operations.LogPath)) 'DryRun hat Logs angelegt.'
        Assert-VssTrue (@($result.PlannedStages) -contains 'EnsureSharePointFolders') 'DryRun-Plan ist unvollstaendig.'
        Assert-VssTrue (@($result.PlannedStages) -contains 'SummarizeRun') 'DryRun-Zusammenfassung fehlt.'
    }
    finally { Remove-VssTestSandbox $sandbox }
}

Invoke-VssTest -Name 'CLI-Exitcodes 0, 2 und interner Catch 1 bleiben stabil' -Body {
    $validCli = Invoke-VssCli $completePath $null $null
    $incompleteCli = Invoke-VssCli $incompletePath Validate $null
    $invalidCli = Invoke-VssCli $invalidPath Validate $null
    $secretCli = Invoke-VssCli $secretPath Validate $null
    $dryRunCli = Invoke-VssCli $completePath DryRun $null
    Assert-VssEqual 0 $validCli.ExitCode 'Vollstaendige CLI-Config muss Exit 0 liefern.'
    Assert-VssTrue ($validCli.Output -match 'Modus: Validate') 'Default-Modus ist nicht Validate.'
    Assert-VssTrue ($validCli.Output -match 'LEGACY-Geruest') 'Alter CLI-Einstieg ist nicht sichtbar als Legacy gekennzeichnet.'
    Assert-VssEqual 2 $incompleteCli.ExitCode 'Unvollstaendige CLI-Config muss Exit 2 liefern.'
    Assert-VssEqual 2 $invalidCli.ExitCode 'Ungueltige CLI-Config muss Exit 2 liefern.'
    Assert-VssEqual 2 $secretCli.ExitCode 'Secret-Config muss Exit 2 liefern.'
    Assert-VssTrue ($secretCli.Output -notmatch 'ClientSecret|TOP-SECRET|SENTINEL-SECRET-VALUE') 'CLI leakt Secret-Key oder Wert.'
    Assert-VssEqual 0 $dryRunCli.ExitCode 'Gueltiger DryRun muss Exit 0 liefern.'

    $sandbox = New-VssTestSandbox
    try {
        $missing = Join-Path $sandbox 'missing.json'
        $malformed = Join-Path $sandbox 'malformed.json'
        Write-VssUtf8NoBomText $malformed '{ invalid json'
        Assert-VssEqual 2 (Invoke-VssCli $missing Validate $null).ExitCode 'Fehlende Config muss Exit 2 liefern.'
        Assert-VssEqual 2 (Invoke-VssCli $malformed Validate $null).ExitCode 'Malformed JSON muss Exit 2 liefern.'
        $missingResult = Invoke-VisioSharePointSyncAssessment $missing
        Assert-VssEqual 18 @($missingResult.UnresolvedDecisionIds).Count 'Fehlende Config muss alle Blocker melden.'

        $brokenCli = Join-Path $sandbox 'Invoke-VisioSharePointSync.ps1'
        [System.IO.File]::Copy($cliPath, $brokenCli)
        $internalFailure = Invoke-VssCli $completePath Validate $brokenCli
        Assert-VssEqual 1 $internalFailure.ExitCode 'Interner CLI-Fehler muss Exit 1 liefern.'
        Assert-VssTrue ($internalFailure.Output -match 'Interner Fehler') 'CLI-Catch ist nicht menschenlesbar.'
    }
    finally { Remove-VssTestSandbox $sandbox }
}

Write-Output ''
Write-Output ("Ergebnis: {0} bestanden, {1} fehlgeschlagen." -f $script:PassedCount, $script:FailedCount)
if ($script:FailedCount -gt 0) { exit 1 }
exit 0
