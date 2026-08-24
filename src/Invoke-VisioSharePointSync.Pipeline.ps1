[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RuntimeConfigurationPath,

    [ValidateSet('Validate', 'Simulate', 'Execute')]
    [string]$Mode = 'Validate',

    [switch]$AllowExternalSideEffects
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# This entry point is deliberately separate from Invoke-VisioSharePointSync.ps1.
# Validate and Simulate are side-effect free. Execute has two explicit gates and
# a static readiness check; the stable placeholders below therefore block every
# current Execute request before a lock, UNC traversal, COM activation or HTTP
# request can occur.

$script:VssPipelinePlaceholderValue = '__PLACEHOLDER_REQUIRED__'
$script:VssPipelineRuntimeTopLevelKeys = @('SchemaVersion', 'Execution', 'Source', 'Conversion', 'Upload', 'Operations')
$script:VssPipelineRuntimeSectionKeys = @{
    Execution  = @('Enabled')
    Source     = @('Recurse', 'ExcludeHidden', 'ExcludeSystem', 'ExcludeReparsePoints', 'ExcludePatterns')
    Conversion = @('AdapterKind', 'AdapterVersion', 'ExternalExecutablePath')
    Upload     = @('SessionThresholdBytes', 'ChunkSizeBytes')
    Operations = @('QuarantinePath')
}

function Test-VssPipelineObjectContainer {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $false }
    if ($Value -is [System.Collections.IDictionary]) { return $true }
    return ($Value -is [System.Management.Automation.PSCustomObject])
}

function Get-VssPipelineDirectPropertyNames {
    param([Parameter(Mandatory = $true)][object]$InputObject)

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) { Write-Output ([string]$key) }
        return
    }
    foreach ($property in $InputObject.PSObject.Properties) { Write-Output $property.Name }
}

function Get-VssPipelinePropertyValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return [pscustomobject]@{ Found = $false; Value = $null }
    }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            if ([string]::Equals([string]$key, $Name, [System.StringComparison]::Ordinal)) {
                return [pscustomobject]@{ Found = $true; Value = $InputObject[$key] }
            }
        }
        return [pscustomobject]@{ Found = $false; Value = $null }
    }
    foreach ($property in $InputObject.PSObject.Properties) {
        if ([string]::Equals($property.Name, $Name, [System.StringComparison]::Ordinal)) {
            return [pscustomobject]@{ Found = $true; Value = $property.Value }
        }
    }
    return [pscustomobject]@{ Found = $false; Value = $null }
}

function Test-VssPipelineContainsOrdinal {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Values,
        [Parameter(Mandatory = $true)][string]$Expected
    )

    foreach ($value in $Values) {
        if ([string]::Equals($value, $Expected, [System.StringComparison]::Ordinal)) { return $true }
    }
    return $false
}

function Test-VssPipelineIntegerValue {
    param([AllowNull()][object]$Value)

    return (
        ($Value -is [byte]) -or ($Value -is [sbyte]) -or
        ($Value -is [int16]) -or ($Value -is [uint16]) -or
        ($Value -is [int32]) -or ($Value -is [uint32]) -or
        ($Value -is [int64]) -or ($Value -is [uint64])
    )
}

function Test-VssPipelineStringArrayValue {
    param([AllowNull()][object]$Value)

    if (($Value -is [string]) -or ($Value -is [System.Collections.IDictionary]) -or
        -not ($Value -is [System.Collections.IEnumerable])) { return $false }
    foreach ($item in @($Value)) {
        if (-not ($item -is [string])) { return $false }
    }
    return $true
}

function Get-VssPipelineClosedSchemaErrors {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$AllowedKeys
    )

    $errors = @()
    if (-not (Test-VssPipelineObjectContainer -Value $InputObject)) {
        $errors += "$Path muss ein JSON-Objekt sein."
        return $errors
    }
    foreach ($name in @(Get-VssPipelineDirectPropertyNames -InputObject $InputObject)) {
        if (-not (Test-VssPipelineContainsOrdinal -Values $AllowedKeys -Expected $name)) {
            # Do not echo user-controlled property names. A name can itself carry
            # secret material, so unknown-key diagnostics stay deliberately generic.
            $errors += "$Path enthaelt mindestens einen unbekannten Schluessel."
            break
        }
    }
    foreach ($requiredName in $AllowedKeys) {
        $property = Get-VssPipelinePropertyValue -InputObject $InputObject -Name $requiredName
        if (-not $property.Found) { $errors += "$Path.$requiredName fehlt." }
    }
    return $errors
}

function New-VssPipelineRuntimeAssessment {
    param(
        [AllowNull()][object]$Configuration,
        [string[]]$Errors = @(),
        [string[]]$Warnings = @(),
        [string[]]$Placeholders = @()
    )

    $errorItems = @($Errors)
    return [pscustomobject]@{
        IsValid       = [bool]($errorItems.Count -eq 0)
        Configuration = $Configuration
        Errors        = $errorItems
        Warnings      = @($Warnings)
        Placeholders  = @($Placeholders | Select-Object -Unique)
    }
}

function Test-VssPipelineRuntimeConfiguration {
    [CmdletBinding()]
    param([AllowNull()][object]$Configuration)

    $errors = @()
    $warnings = @()
    $placeholders = @()

    if (-not (Test-VssPipelineObjectContainer -Value $Configuration)) {
        return New-VssPipelineRuntimeAssessment -Configuration $Configuration -Errors @('Die Wurzel der Laufzeitkonfiguration muss ein JSON-Objekt sein.')
    }

    if (Test-VssContainsProhibitedKey -InputObject $Configuration) {
        # Der konkrete Schluessel wird absichtlich nicht wiedergegeben: Schon ein
        # frei gewaehlter JSON-Schluessel kann geheimes Material enthalten.
        $errors += 'Die Laufzeitkonfiguration enthaelt mindestens ein nicht zulaessiges geheimnisartiges Feld.'
    }

    $errors += @(Get-VssPipelineClosedSchemaErrors -InputObject $Configuration -Path 'Runtime' -AllowedKeys $script:VssPipelineRuntimeTopLevelKeys)
    foreach ($sectionName in @('Execution', 'Source', 'Conversion', 'Upload', 'Operations')) {
        $sectionResult = Get-VssPipelinePropertyValue -InputObject $Configuration -Name $sectionName
        if ($sectionResult.Found) {
            $errors += @(Get-VssPipelineClosedSchemaErrors -InputObject $sectionResult.Value -Path "Runtime.$sectionName" -AllowedKeys $script:VssPipelineRuntimeSectionKeys[$sectionName])
        }
    }

    $schemaVersion = Get-VssPipelinePropertyValue -InputObject $Configuration -Name 'SchemaVersion'
    if ($schemaVersion.Found -and (-not ($schemaVersion.Value -is [string]) -or $schemaVersion.Value -cne '1.0')) {
        $errors += 'Runtime.SchemaVersion muss exakt "1.0" sein.'
    }

    $execution = Get-VssPipelinePropertyValue -InputObject $Configuration -Name 'Execution'
    if ($execution.Found -and (Test-VssPipelineObjectContainer -Value $execution.Value)) {
        $enabled = Get-VssPipelinePropertyValue -InputObject $execution.Value -Name 'Enabled'
        if ($enabled.Found -and -not ($enabled.Value -is [bool])) {
            $errors += 'Runtime.Execution.Enabled muss ein Boolean sein.'
        }
    }

    $source = Get-VssPipelinePropertyValue -InputObject $Configuration -Name 'Source'
    if ($source.Found -and (Test-VssPipelineObjectContainer -Value $source.Value)) {
        foreach ($booleanName in @('Recurse', 'ExcludeHidden', 'ExcludeSystem', 'ExcludeReparsePoints')) {
            $booleanResult = Get-VssPipelinePropertyValue -InputObject $source.Value -Name $booleanName
            if ($booleanResult.Found -and -not ($booleanResult.Value -is [bool])) {
                $errors += "Runtime.Source.$booleanName muss ein Boolean sein."
            }
        }
        $excludePatterns = Get-VssPipelinePropertyValue -InputObject $source.Value -Name 'ExcludePatterns'
        if ($excludePatterns.Found) {
            if (-not (Test-VssPipelineStringArrayValue -Value $excludePatterns.Value)) {
                $errors += 'Runtime.Source.ExcludePatterns muss ein String-Array sein.'
            }
            else {
                foreach ($pattern in @($excludePatterns.Value)) {
                    if ([string]::IsNullOrWhiteSpace($pattern)) {
                        $errors += 'Runtime.Source.ExcludePatterns darf keine leeren Eintraege enthalten.'
                        break
                    }
                    try { [void](New-Object System.Management.Automation.WildcardPattern($pattern, [System.Management.Automation.WildcardOptions]::IgnoreCase)) }
                    catch {
                        $errors += 'Runtime.Source.ExcludePatterns enthaelt ein ungueltiges Wildcard-Muster.'
                        break
                    }
                }
            }
        }
    }

    $conversion = Get-VssPipelinePropertyValue -InputObject $Configuration -Name 'Conversion'
    if ($conversion.Found -and (Test-VssPipelineObjectContainer -Value $conversion.Value)) {
        $adapterKind = Get-VssPipelinePropertyValue -InputObject $conversion.Value -Name 'AdapterKind'
        $adapterVersion = Get-VssPipelinePropertyValue -InputObject $conversion.Value -Name 'AdapterVersion'
        $externalExecutablePath = Get-VssPipelinePropertyValue -InputObject $conversion.Value -Name 'ExternalExecutablePath'

        if ($adapterKind.Found) {
            $allowedAdapterKinds = @($script:VssPipelinePlaceholderValue, 'VisioCom', 'ExternalConverter')
            if (-not ($adapterKind.Value -is [string]) -or
                -not (Test-VssPipelineContainsOrdinal -Values $allowedAdapterKinds -Expected ([string]$adapterKind.Value))) {
                $errors += 'Runtime.Conversion.AdapterKind ist nicht unterstuetzt.'
            }
            elseif ($adapterKind.Value -ceq $script:VssPipelinePlaceholderValue) {
                $placeholders += 'VSS-CNV-001'
            }
        }
        if ($adapterVersion.Found) {
            if (-not ($adapterVersion.Value -is [string]) -or [string]::IsNullOrWhiteSpace($adapterVersion.Value)) {
                $errors += 'Runtime.Conversion.AdapterVersion muss ein nicht leerer String sein.'
            }
            elseif ($adapterVersion.Value -ceq $script:VssPipelinePlaceholderValue) {
                $placeholders += 'VSS-CNV-001'
            }
        }
        if ($externalExecutablePath.Found -and $null -ne $externalExecutablePath.Value -and
            -not ($externalExecutablePath.Value -is [string])) {
            $errors += 'Runtime.Conversion.ExternalExecutablePath muss null oder ein String sein.'
        }
        if ($adapterKind.Found -and ($adapterKind.Value -ceq 'VisioCom') -and
            $externalExecutablePath.Found -and $null -ne $externalExecutablePath.Value) {
            $errors += 'Runtime.Conversion.ExternalExecutablePath muss bei VisioCom null sein.'
        }
        if ($adapterKind.Found -and ($adapterKind.Value -ceq 'ExternalConverter')) {
            if (-not $externalExecutablePath.Found -or $null -eq $externalExecutablePath.Value -or
                (($externalExecutablePath.Value -is [string]) -and
                 ($externalExecutablePath.Value -ceq $script:VssPipelinePlaceholderValue))) {
                $placeholders += 'VSS-CNV-001'
            }
            elseif (($externalExecutablePath.Value -is [string]) -and
                -not (Test-VssAbsoluteLocalPathSyntax -Value $externalExecutablePath.Value)) {
                $errors += 'Runtime.Conversion.ExternalExecutablePath muss ein sicherer absoluter lokaler Nicht-Root-Pfad sein.'
            }
        }
    }

    $upload = Get-VssPipelinePropertyValue -InputObject $Configuration -Name 'Upload'
    if ($upload.Found -and (Test-VssPipelineObjectContainer -Value $upload.Value)) {
        foreach ($integerName in @('SessionThresholdBytes', 'ChunkSizeBytes')) {
            $integerResult = Get-VssPipelinePropertyValue -InputObject $upload.Value -Name $integerName
            if ($integerResult.Found -and
                (-not (Test-VssPipelineIntegerValue -Value $integerResult.Value) -or $integerResult.Value -le 0)) {
                $errors += "Runtime.Upload.$integerName muss eine positive Ganzzahl sein."
            }
        }
        $chunkSize = Get-VssPipelinePropertyValue -InputObject $upload.Value -Name 'ChunkSizeBytes'
        if ($chunkSize.Found -and (Test-VssPipelineIntegerValue -Value $chunkSize.Value) -and $chunkSize.Value -gt 0) {
            if (($chunkSize.Value % 327680) -ne 0) {
                $errors += 'Runtime.Upload.ChunkSizeBytes muss fuer Graph ein Vielfaches von 327680 Bytes sein.'
            }
            if ($chunkSize.Value -gt 62914560) {
                $errors += 'Runtime.Upload.ChunkSizeBytes darf 60 MiB nicht ueberschreiten.'
            }
        }
    }

    $operations = Get-VssPipelinePropertyValue -InputObject $Configuration -Name 'Operations'
    if ($operations.Found -and (Test-VssPipelineObjectContainer -Value $operations.Value)) {
        $quarantinePath = Get-VssPipelinePropertyValue -InputObject $operations.Value -Name 'QuarantinePath'
        if ($quarantinePath.Found) {
            if (-not ($quarantinePath.Value -is [string]) -or [string]::IsNullOrWhiteSpace($quarantinePath.Value)) {
                $errors += 'Runtime.Operations.QuarantinePath muss ein nicht leerer String sein.'
            }
            elseif ($quarantinePath.Value -ceq $script:VssPipelinePlaceholderValue) {
                $placeholders += 'VSS-RUNTIME-QUARANTINE-001'
            }
            elseif (-not (Test-VssAbsoluteLocalPathSyntax -Value $quarantinePath.Value)) {
                $errors += 'Runtime.Operations.QuarantinePath muss ein sicherer absoluter lokaler Nicht-Root-Pfad sein.'
            }
        }
    }

    return New-VssPipelineRuntimeAssessment -Configuration $Configuration -Errors $errors -Warnings $warnings -Placeholders $placeholders
}

function Read-VssPipelineRuntimeConfiguration {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$LiteralPath)

    if (-not [System.IO.File]::Exists($LiteralPath)) {
        return New-VssPipelineRuntimeAssessment -Configuration $null -Errors @('Die Laufzeitkonfigurationsdatei wurde nicht gefunden.')
    }
    try { $jsonText = Read-VssUtf8FileStrict -LiteralPath $LiteralPath }
    catch {
        return New-VssPipelineRuntimeAssessment -Configuration $null -Errors @('Die Laufzeitkonfiguration konnte nicht als BOM-loses gueltiges UTF-8 gelesen werden.')
    }
    try {
        if (Test-VssJsonHasDuplicateObjectKeys -JsonText $jsonText) {
            return New-VssPipelineRuntimeAssessment -Configuration $null -Errors @('Die Laufzeitkonfiguration enthaelt doppelte JSON-Schluessel.')
        }
        $configuration = ConvertFrom-Json -InputObject $jsonText -ErrorAction Stop
    }
    catch {
        return New-VssPipelineRuntimeAssessment -Configuration $null -Errors @('Die Laufzeitkonfiguration enthaelt kein gueltiges JSON.')
    }
    return Test-VssPipelineRuntimeConfiguration -Configuration $configuration
}

function New-VssPipelineResult {
    param(
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][ValidateSet('Validate', 'Simulate', 'Execute')][string]$Mode,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][ValidateRange(0, 4)][int]$ExitCode,
        [Parameter(Mandatory = $true)][string]$StartedUtc,
        [Parameter(Mandatory = $true)][string]$FinishedUtc,
        [object[]]$Stages = @(),
        [object[]]$Files = @(),
        [string[]]$Errors = @(),
        [string[]]$Warnings = @(),
        [string[]]$Placeholders = @()
    )

    # The public result contract intentionally has exactly these eleven fields.
    return [pscustomobject][ordered]@{
        RunId        = $RunId
        Mode         = $Mode
        Status       = $Status
        ExitCode     = $ExitCode
        StartedUtc   = $StartedUtc
        FinishedUtc  = $FinishedUtc
        Stages       = @($Stages)
        Files        = @($Files)
        Errors       = @($Errors)
        Warnings     = @($Warnings)
        Placeholders = @($Placeholders | Select-Object -Unique)
    }
}

function New-VssPipelineStageResult {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$EffectKind,
        [Parameter(Mandatory = $true)][string]$Detail,
        [ValidateSet('Ready', 'Simulation', 'Placeholder')][string]$Implementation = 'Ready',
        [AllowNull()][string]$PlaceholderId = $null,
        [AllowNull()][string]$Handler = $null
    )

    return [pscustomobject][ordered]@{
        Id             = $Id
        Name           = $Id
        Status         = $Status
        EffectKind     = $EffectKind
        Implementation = $Implementation
        PlaceholderId  = $PlaceholderId
        Handler        = $Handler
        Detail         = $Detail
    }
}

function Get-VssPipelineStageCatalog {
    [CmdletBinding()]
    param()

    # `Id` bleibt als kompatibler Alias erhalten. Der formale Stage-Vertrag
    # besteht aus Name, EffectKind, Implementation, PlaceholderId und Handler.
    return @(
        [pscustomobject][ordered]@{ Name = 'ValidateConfiguration';        Id = 'ValidateConfiguration';        EffectKind = 'None';                 Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'Read-VssPipelinePrimaryConfigurationAssessment' }
        [pscustomobject][ordered]@{ Name = 'ValidateRuntimeConfiguration'; Id = 'ValidateRuntimeConfiguration'; EffectKind = 'None';                 Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'Read-VssPipelineRuntimeConfiguration' }
        [pscustomobject][ordered]@{ Name = 'AcquireSingleRunLock';         Id = 'AcquireSingleRunLock';         EffectKind = 'LocalSynchronization'; Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'AcquireLock' }
        [pscustomobject][ordered]@{ Name = 'InventorySource';              Id = 'InventorySource';              EffectKind = 'NetworkRead';          Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'Inventory' }
        [pscustomobject][ordered]@{ Name = 'StageStableSource';            Id = 'StageStableSource';            EffectKind = 'LocalWrite';           Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'StageSource' }
        [pscustomobject][ordered]@{ Name = 'ConvertVisioToPdf';            Id = 'ConvertVisioToPdf';            EffectKind = 'ComOrProcess';         Implementation = 'Placeholder'; PlaceholderId = 'VSS-CNV-001';         Handler = 'Convert' }
        [pscustomobject][ordered]@{ Name = 'VerifySourceUnchanged';        Id = 'VerifySourceUnchanged';        EffectKind = 'NetworkRead';          Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'VerifySource' }
        [pscustomobject][ordered]@{ Name = 'EnsureSharePointFolders';      Id = 'EnsureSharePointFolders';      EffectKind = 'RemoteWrite';          Implementation = 'Placeholder'; PlaceholderId = 'VSS-INTEGRATION-001'; Handler = 'EnsureFolders' }
        [pscustomobject][ordered]@{ Name = 'UploadOrUpdatePdf';            Id = 'UploadOrUpdatePdf';            EffectKind = 'RemoteWrite';          Implementation = 'Placeholder'; PlaceholderId = 'VSS-INTEGRATION-001'; Handler = 'Upload' }
        [pscustomobject][ordered]@{ Name = 'ReconcileLocalAndRemoteState'; Id = 'ReconcileLocalAndRemoteState'; EffectKind = 'LocalRead';            Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'ReportOrphans' }
        [pscustomobject][ordered]@{ Name = 'PersistSyncState';             Id = 'PersistSyncState';             EffectKind = 'LocalWrite';           Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'SaveState' }
        [pscustomobject][ordered]@{ Name = 'WriteStructuredLog';           Id = 'WriteStructuredLog';           EffectKind = 'LocalWrite';           Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'WriteLog' }
        [pscustomobject][ordered]@{ Name = 'CleanupStaging';               Id = 'CleanupStaging';               EffectKind = 'LocalDelete';          Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'Cleanup' }
        [pscustomobject][ordered]@{ Name = 'ReleaseSingleRunLock';         Id = 'ReleaseSingleRunLock';         EffectKind = 'Synchronization';      Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'ReleaseLock' }
        [pscustomobject][ordered]@{ Name = 'SummarizeRun';                 Id = 'SummarizeRun';                 EffectKind = 'None';                 Implementation = 'Ready';       PlaceholderId = $null;                 Handler = 'Internal' }
    )
}

function Get-VssPipelineTargetPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RelativeSourcePath)

    if ($RelativeSourcePath -ne $RelativeSourcePath.Trim()) { throw 'Der relative Quellpfad enthaelt aeusseren Leerraum.' }
    if ([System.IO.Path]::IsPathRooted($RelativeSourcePath) -or $RelativeSourcePath.StartsWith('/') -or
        $RelativeSourcePath.StartsWith('\') -or $RelativeSourcePath.Contains(':')) {
        throw 'Der Quellpfad muss relativ sein.'
    }
    $normalized = $RelativeSourcePath.Replace([char]92, [char]47)
    $segments = @($normalized.Split([char]47))
    if ($segments.Count -eq 0) { throw 'Der relative Quellpfad ist leer.' }
    foreach ($segment in $segments) {
        if (-not (Test-VssSafeWindowsSegment -Segment $segment -AdditionalInvalidChars ([char[]]'[]'))) {
            throw 'Der relative Quellpfad enthaelt ein unsicheres Segment.'
        }
    }
    $fileName = $segments[$segments.Count - 1]
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    if ([string]::IsNullOrWhiteSpace($baseName)) { throw 'Der Quellpfad enthaelt keinen gueltigen Dateinamen.' }
    $segments[$segments.Count - 1] = $baseName + '.pdf'
    return ($segments -join '/')
}

function ConvertTo-VssPipelineEscapedUriPath {
    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RelativePath)

    $segments = @($RelativePath.Replace([char]92, [char]47).Split([char]47))
    $escaped = @()
    foreach ($segment in $segments) {
        if ([string]::IsNullOrWhiteSpace($segment) -or $segment -eq '.' -or $segment -eq '..') {
            throw 'Ein URI-Pfadsegment ist ungueltig.'
        }
        $escaped += [System.Uri]::EscapeDataString($segment)
    }
    return ($escaped -join '/')
}

function New-VssGraphRequestPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TargetRelativePath,
        [Parameter(Mandatory = $true)][ValidateRange(0, [long]::MaxValue)][long]$ContentLength,
        [AllowNull()][object]$CommitDirective = $null
    )

    $driveId = [string]$Configuration.SharePoint.DriveId
    $targetFolderId = [string]$Configuration.SharePoint.TargetFolderId
    if (-not (Test-VssGraphOpaqueIdSyntax -Value $driveId) -or
        -not (Test-VssGraphOpaqueIdSyntax -Value $targetFolderId)) {
        throw 'Fuer den Graph-Requestplan fehlen sichere Drive- oder Zielordner-IDs.'
    }
    $escapedDriveId = [System.Uri]::EscapeDataString($driveId)
    $escapedFolderId = [System.Uri]::EscapeDataString($targetFolderId)
    $escapedTargetPath = ConvertTo-VssPipelineEscapedUriPath -RelativePath $TargetRelativePath
    $threshold = [long]$RuntimeConfiguration.Upload.SessionThresholdBytes
    $action = 'Create'
    if ($null -ne $CommitDirective) { $action = [string]$CommitDirective.Action }
    if (($action -cne 'Create') -and ($action -cne 'Update')) {
        throw 'Der Graph-Commitvertrag erlaubt nur Create oder Update.'
    }

    $headers = [ordered]@{ 'Content-Type' = 'application/pdf' }
    if ($action -ceq 'Create') {
        # Die vorgelagerte Remote-Pruefung muss bereits "nicht vorhanden"
        # ergeben haben. If-None-Match schuetzt zusaetzlich gegen das Rennen bis
        # zum Commit; die Integrations-Allowlist bleibt trotzdem Pflicht.
        $headers['If-None-Match'] = '*'
        $itemReference = "items/$escapedFolderId`:/$escapedTargetPath`:"
    }
    else {
        $remoteItemId = [string]$CommitDirective.RemoteItemId
        $ifMatch = [string]$CommitDirective.IfMatch
        if (-not (Test-VssGraphOpaqueIdSyntax -Value $remoteItemId) -or [string]::IsNullOrWhiteSpace($ifMatch)) {
            throw 'Ein Graph-Update benoetigt eine bekannte Remote-ID und einen eTag.'
        }
        $headers['If-Match'] = $ifMatch
        $itemReference = 'items/' + [System.Uri]::EscapeDataString($remoteItemId)
    }

    if ($ContentLength -lt $threshold) {
        return [pscustomobject][ordered]@{
            Transport          = 'MicrosoftGraphV1'
            Operation          = 'SmallUpload'
            Method             = 'PUT'
            Uri                = "https://graph.microsoft.com/v1.0/drives/$escapedDriveId/$itemReference/content"
            Headers            = $headers
            BodyKind           = 'PdfFile'
            ExternalSideEffect = $true
        }
    }
    return [pscustomobject][ordered]@{
        Transport          = 'MicrosoftGraphV1'
        Operation          = 'CreateUploadSession'
        Method             = 'POST'
        Uri                = "https://graph.microsoft.com/v1.0/drives/$escapedDriveId/$itemReference/createUploadSession"
        Headers            = $headers
        BodyKind           = 'UploadSessionThenPdfChunks'
        ExternalSideEffect = $true
    }
}

function New-VssGraphFolderResolutionPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TargetRelativePath
    )

    $driveId = [string]$Configuration.SharePoint.DriveId
    $rootFolderId = [string]$Configuration.SharePoint.TargetFolderId
    if (-not (Test-VssGraphOpaqueIdSyntax -Value $driveId) -or -not (Test-VssGraphOpaqueIdSyntax -Value $rootFolderId)) {
        throw 'Fuer die Graph-Ordneraufloesung fehlen sichere Drive- oder Wurzelordner-IDs.'
    }
    $normalized = $TargetRelativePath.Replace([char]92, [char]47)
    $segments = @($normalized.Split([char]47))
    $directorySegments = @()
    if ($segments.Count -gt 1) { $directorySegments = @($segments[0..($segments.Count - 2)]) }
    $escapedDriveId = [System.Uri]::EscapeDataString($driveId)
    $escapedRootId = [System.Uri]::EscapeDataString($rootFolderId)
    $accumulated = @()
    $steps = @()
    foreach ($segment in $directorySegments) {
        if (-not (Test-VssSafeWindowsSegment -Segment $segment -AdditionalInvalidChars ([char[]]'[]'))) {
            throw 'Der Graph-Zielordner enthaelt ein unsicheres Segment.'
        }
        $accumulated += $segment
        $escapedPath = ConvertTo-VssPipelineEscapedUriPath -RelativePath ($accumulated -join '/')
        $steps += [pscustomobject][ordered]@{
            Segment       = $segment
            ResolveMethod = 'GET'
            ResolveUri    = "https://graph.microsoft.com/v1.0/drives/$escapedDriveId/items/$escapedRootId`:/$escapedPath"
            CreateMethod  = 'POST'
            CreateUri     = "https://graph.microsoft.com/v1.0/drives/$escapedDriveId/items/{PARENT_ITEM_ID}/children"
            CreateBody    = [ordered]@{ name = $segment; folder = [ordered]@{}; '@microsoft.graph.conflictBehavior' = 'fail' }
        }
    }
    return [pscustomobject][ordered]@{
        Transport       = 'MicrosoftGraphV1'
        InitialParentId = $rootFolderId
        Steps           = @($steps)
    }
}

function ConvertTo-VssPipelineODataLiteral {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)
    return $Value.Replace("'", "''")
}

function New-VssRestRequestPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TargetRelativePath,
        [Parameter(Mandatory = $true)][ValidateRange(0, [long]::MaxValue)][long]$ContentLength,
        [AllowNull()][object]$CommitDirective = $null
    )

    # SharePoint REST has no approved large-file protocol in this scaffold. The
    # planner still describes that branch without attempting a request; Execute
    # readiness is blocked by VSS-SPREST-LARGE-001.
    $siteUrl = [string]$Configuration.SharePoint.SiteUrl
    if (-not (Test-VssAbsoluteHttpUrlSyntax -Value $siteUrl)) { throw 'Fuer den REST-Requestplan fehlt eine sichere SiteUrl.' }

    $normalizedTarget = $TargetRelativePath.Replace([char]92, [char]47)
    if ($normalizedTarget.Contains('%') -or $normalizedTarget.Contains('#')) {
        # >>> PLACEHOLDER [VSS-SPREST-PATH-001] BEGIN
        throw [System.NotSupportedException]::new(
            '[VSS-SPREST-PATH-001] Die ResourcePath-Semantik fuer Prozent- und Rautezeichen ist fuer die konkrete SharePoint-Version noch nicht verifiziert.'
        )
        # <<< PLACEHOLDER [VSS-SPREST-PATH-001] END
    }
    $targetSegments = @($normalizedTarget.Split([char]47))
    $fileName = $targetSegments[$targetSegments.Count - 1]
    $targetDirectories = @()
    if ($targetSegments.Count -gt 1) { $targetDirectories = @($targetSegments[0..($targetSegments.Count - 2)]) }

    $siteUri = [uri]$siteUrl
    $folderSegments = @()
    if (-not [string]::IsNullOrWhiteSpace($Configuration.SharePoint.LibraryName)) {
        $folderSegments += [string]$Configuration.SharePoint.LibraryName
    }
    if (-not [string]::IsNullOrWhiteSpace($Configuration.SharePoint.TargetFolderPath)) {
        $folderSegments += @(([string]$Configuration.SharePoint.TargetFolderPath).Split([char]47))
    }
    $folderSegments += $targetDirectories
    $serverRelativeFolder = ($siteUri.AbsolutePath.TrimEnd('/') + '/' + ($folderSegments -join '/')).Replace('//', '/')
    $folderLiteral = ConvertTo-VssPipelineODataLiteral -Value $serverRelativeFolder
    $fileLiteral = ConvertTo-VssPipelineODataLiteral -Value $fileName
    $action = 'Create'
    if ($null -ne $CommitDirective) { $action = [string]$CommitDirective.Action }
    if (($action -cne 'Create') -and ($action -cne 'Update')) {
        throw 'Der REST-Commitvertrag erlaubt nur Create oder Update.'
    }
    $overwrite = if ($action -ceq 'Update') { 'true' } else { 'false' }
    $apiRelative = "_api/web/GetFolderByServerRelativeUrl('$folderLiteral')/Files/add(url='$fileLiteral',overwrite=$overwrite)"
    $uri = $siteUrl.TrimEnd('/') + '/' + $apiRelative

    $operation = if ($action -ceq 'Update') { 'UpdateKnownItem' } else { 'CreateIfAbsent' }
    $bodyKind = 'PdfFile'
    if ($ContentLength -ge [long]$RuntimeConfiguration.Upload.SessionThresholdBytes) {
        $operation = 'LargeUploadPlaceholder'
        $bodyKind = 'Placeholder:VSS-SPREST-LARGE-001'
    }
    $headers = [ordered]@{ Accept = 'application/json;odata=nometadata'; 'Content-Type' = 'application/pdf'; 'X-RequestDigest' = '[REQUIRED_AT_EXECUTION]' }
    if ($action -ceq 'Create') { $headers['If-None-Match'] = '*' }
    else {
        $ifMatch = [string]$CommitDirective.IfMatch
        if ([string]::IsNullOrWhiteSpace($ifMatch)) { throw 'Ein REST-Update benoetigt einen bekannten eTag.' }
        $headers['If-Match'] = $ifMatch
    }
    return [pscustomobject][ordered]@{
        Transport          = 'SharePointRest'
        Operation          = $operation
        Method             = 'POST'
        Uri                = [System.Uri]::EscapeUriString($uri)
        Headers            = $headers
        BodyKind           = $bodyKind
        ExternalSideEffect = $true
    }
}

function New-VssPipelineRequestPlan {
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration,
        [Parameter(Mandatory = $true)][string]$TargetRelativePath,
        [Parameter(Mandatory = $true)][long]$ContentLength,
        [AllowNull()][object]$CommitDirective = $null
    )

    switch -CaseSensitive ([string]$Configuration.SharePoint.ApiKind) {
        'MicrosoftGraphV1' {
            return New-VssGraphRequestPlan -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration -TargetRelativePath $TargetRelativePath -ContentLength $ContentLength -CommitDirective $CommitDirective
        }
        'SharePointRest' {
            return New-VssRestRequestPlan -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration -TargetRelativePath $TargetRelativePath -ContentLength $ContentLength -CommitDirective $CommitDirective
        }
        default { throw 'Es ist kein unterstuetzter SharePoint-Transport konfiguriert.' }
    }
}

function Invoke-VssPipelineSimulation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration,
        [string]$RunId = '00000000-0000-4000-8000-000000000001'
    )

    # Der gleiche Orchestrator wie bei Execute wird verwendet; nur das hier im
    # Code erzeugte Adapterset ersetzt ausnahmslos Lock, Dateisystem, Uhr, COM
    # und Netzwerk durch deterministische In-Memory-Handler.
    $adapterSet = New-VssPipelineSimulationAdapterSet
    return Invoke-VssPipelineExecution -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration -RunId $RunId -Mode Simulation -AdapterSet $adapterSet
}

function Test-VssPipelineExcludedRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Patterns
    )

    $normalized = $RelativePath.Replace([char]92, [char]47)
    foreach ($patternText in $Patterns) {
        $pattern = New-Object System.Management.Automation.WildcardPattern($patternText, [System.Management.Automation.WildcardOptions]::IgnoreCase)
        if ($pattern.IsMatch($normalized)) { return $true }
    }
    return $false
}

function Test-VssPipelineFileAttribute {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileAttributes]$Attributes,
        [Parameter(Mandatory = $true)][System.IO.FileAttributes]$Expected
    )
    return (($Attributes -band $Expected) -ne 0)
}

function Invoke-VssPipelineInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration
    )

    $rootPath = [string]$Configuration.Source.RootPath
    $root = New-Object System.IO.DirectoryInfo($rootPath)
    if (-not $root.Exists) { throw 'Die konfigurierte Quellwurzel ist nicht erreichbar.' }

    $sourceRegex = New-VssSourceRegex -Pattern ([string]$Configuration.Source.FileNamePattern) -TimeoutMilliseconds ([int]$Configuration.Source.RegexTimeoutMilliseconds)
    $extensions = @($Configuration.Source.Extensions)
    $patterns = @($RuntimeConfiguration.Source.ExcludePatterns)
    $candidates = @()
    $inventoryErrors = @()
    $complete = $true
    $pending = New-Object 'System.Collections.Generic.Stack[System.IO.DirectoryInfo]'
    $pending.Push($root)
    $rootPrefix = $root.FullName.TrimEnd([char]92)

    while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        $relativeDirectory = $directory.FullName.Substring($rootPrefix.Length).TrimStart([char]92)
        if ([string]::IsNullOrWhiteSpace($relativeDirectory)) { $relativeDirectory = '.' }
        try { $directoryFiles = @($directory.GetFiles() | Sort-Object -Property Name) }
        catch {
            $complete = $false
            $directoryFiles = @()
            $inventoryErrors += [pscustomobject][ordered]@{
                RelativePath = $relativeDirectory
                Operation    = 'EnumerateFiles'
                Result       = 'ACCESS_FAILED'
            }
        }
        foreach ($file in $directoryFiles) {
            $relativePath = $file.FullName.Substring($rootPrefix.Length).TrimStart([char]92)
            try {
                $fileAttributes = $file.Attributes
                $fileLength = [long]$file.Length
                $fileWriteUtc = $file.LastWriteTimeUtc
            }
            catch {
                $complete = $false
                $inventoryErrors += [pscustomobject][ordered]@{
                    RelativePath = $relativePath
                    Operation    = 'ReadMetadata'
                    Result       = 'ACCESS_FAILED'
                }
                continue
            }
            if ($RuntimeConfiguration.Source.ExcludeHidden -and
                (Test-VssPipelineFileAttribute -Attributes $fileAttributes -Expected ([System.IO.FileAttributes]::Hidden))) { continue }
            if ($RuntimeConfiguration.Source.ExcludeSystem -and
                (Test-VssPipelineFileAttribute -Attributes $fileAttributes -Expected ([System.IO.FileAttributes]::System))) { continue }
            if ($RuntimeConfiguration.Source.ExcludeReparsePoints -and
                (Test-VssPipelineFileAttribute -Attributes $fileAttributes -Expected ([System.IO.FileAttributes]::ReparsePoint))) { continue }
            if ((Test-VssPipelineExcludedRelativePath -RelativePath $relativePath -Patterns ([string[]]$patterns)) -or
                (Test-VssPipelineExcludedRelativePath -RelativePath $file.Name -Patterns ([string[]]$patterns))) { continue }

            $extensionAllowed = $false
            foreach ($extension in $extensions) {
                if ([string]::Equals([string]$extension, $file.Extension, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $extensionAllowed = $true
                    break
                }
            }
            if (-not $extensionAllowed) { continue }
            try {
                if (-not $sourceRegex.IsMatch($file.Name)) { continue }
            }
            catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                throw 'Die Auswertung des Dateinamensmusters hat das konfigurierte Zeitlimit ueberschritten.'
            }

            $candidates += [pscustomobject][ordered]@{
                SourcePath       = $file.FullName
                RelativePath     = $relativePath
                Length           = $fileLength
                LastWriteTimeUtc = $fileWriteUtc.ToString('o')
            }
        }

        if ($RuntimeConfiguration.Source.Recurse) {
            try { $childDirectories = @($directory.GetDirectories() | Sort-Object -Property Name -Descending) }
            catch {
                $complete = $false
                $childDirectories = @()
                $inventoryErrors += [pscustomobject][ordered]@{
                    RelativePath = $relativeDirectory
                    Operation    = 'EnumerateDirectories'
                    Result       = 'ACCESS_FAILED'
                }
            }
            foreach ($child in $childDirectories) {
                $relativeDirectory = $child.FullName.Substring($rootPrefix.Length).TrimStart([char]92)
                try { $childAttributes = $child.Attributes }
                catch {
                    $complete = $false
                    $inventoryErrors += [pscustomobject][ordered]@{
                        RelativePath = $relativeDirectory
                        Operation    = 'ReadDirectoryMetadata'
                        Result       = 'ACCESS_FAILED'
                    }
                    continue
                }
                if ($RuntimeConfiguration.Source.ExcludeHidden -and
                    (Test-VssPipelineFileAttribute -Attributes $childAttributes -Expected ([System.IO.FileAttributes]::Hidden))) { continue }
                if ($RuntimeConfiguration.Source.ExcludeSystem -and
                    (Test-VssPipelineFileAttribute -Attributes $childAttributes -Expected ([System.IO.FileAttributes]::System))) { continue }
                if ($RuntimeConfiguration.Source.ExcludeReparsePoints -and
                    (Test-VssPipelineFileAttribute -Attributes $childAttributes -Expected ([System.IO.FileAttributes]::ReparsePoint))) { continue }
                if (Test-VssPipelineExcludedRelativePath -RelativePath $relativeDirectory -Patterns ([string[]]$patterns)) { continue }
                $pending.Push($child)
            }
        }
    }

    return [pscustomobject][ordered]@{
        Complete   = $complete
        Candidates = @($candidates | Sort-Object -Property RelativePath)
        Errors     = @($inventoryErrors)
    }
}

function Get-VssPipelineSha256 {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$LiteralPath)

    $stream = $null
    $sha256 = $null
    try {
        $stream = New-Object System.IO.FileStream($LiteralPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hash = $sha256.ComputeHash($stream)
        return (([System.BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant())
    }
    finally {
        if ($null -ne $sha256) { $sha256.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-VssPipelineSha256Text {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding($false, $true)).GetBytes($Text)
        $hash = $sha256.ComputeHash($bytes)
        return (([System.BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant())
    }
    finally { $sha256.Dispose() }
}

function Get-VssPipelineSafeLocalChildPath {
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RootPath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RelativePath
    )

    if ([System.IO.Path]::IsPathRooted($RelativePath)) { throw 'Ein lokaler Unterpfad darf nicht absolut sein.' }
    $normalizedRelative = $RelativePath.Replace([char]47, [char]92)
    foreach ($segment in @($normalizedRelative.Split([char]92))) {
        if (-not (Test-VssSafeWindowsSegment -Segment $segment -AdditionalInvalidChars ([char[]]'[]'))) {
            throw 'Ein lokaler Unterpfad enthaelt ein unsicheres Segment.'
        }
    }
    $fullRoot = [System.IO.Path]::GetFullPath($RootPath).TrimEnd([char]92)
    $combined = [System.IO.Path]::GetFullPath((Join-Path -Path $fullRoot -ChildPath $normalizedRelative))
    if (-not $combined.StartsWith($fullRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Der lokale Unterpfad verlaesst die konfigurierte Wurzel.'
    }
    return $combined
}

function Copy-VssPipelineStableFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Candidate,
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RunStagingPath
    )

    $sourceBefore = New-Object System.IO.FileInfo([string]$Candidate.SourcePath)
    if (-not $sourceBefore.Exists) {
        return [pscustomobject]@{ Status = 'DEFERRED'; Reason = 'SOURCE_MISSING'; StagedPath = $null; Sha256 = $null }
    }
    $beforeLength = [long]$sourceBefore.Length
    $beforeWriteUtc = $sourceBefore.LastWriteTimeUtc
    Start-Sleep -Seconds ([int]$Configuration.Operations.StabilityProbeSeconds)
    $sourceBefore.Refresh()
    if (-not $sourceBefore.Exists -or $beforeLength -ne [long]$sourceBefore.Length -or
        $beforeWriteUtc -ne $sourceBefore.LastWriteTimeUtc) {
        return [pscustomobject]@{ Status = 'DEFERRED'; Reason = 'SOURCE_CHANGED_DURING_PROBE'; StagedPath = $null; Sha256 = $null }
    }

    $stagedPath = Get-VssPipelineSafeLocalChildPath -RootPath $RunStagingPath -RelativePath ([string]$Candidate.RelativePath)
    $stagedDirectory = [System.IO.Path]::GetDirectoryName($stagedPath)
    [void][System.IO.Directory]::CreateDirectory($stagedDirectory)
    [System.IO.File]::Copy([string]$Candidate.SourcePath, $stagedPath, $false)
    $stagedHash = Get-VssPipelineSha256 -LiteralPath $stagedPath

    $sourceAfter = New-Object System.IO.FileInfo([string]$Candidate.SourcePath)
    if (-not $sourceAfter.Exists -or $beforeLength -ne [long]$sourceAfter.Length -or
        $beforeWriteUtc -ne $sourceAfter.LastWriteTimeUtc) {
        [System.IO.File]::Delete($stagedPath)
        return [pscustomobject]@{ Status = 'DEFERRED'; Reason = 'SOURCE_CHANGED_DURING_COPY'; StagedPath = $null; Sha256 = $null }
    }
    return [pscustomobject][ordered]@{
        Status             = 'STAGED'
        Reason             = $null
        SourcePath         = [string]$Candidate.SourcePath
        RelativePath       = [string]$Candidate.RelativePath
        StagedPath         = $stagedPath
        Sha256             = $stagedHash
        SourceLength       = $beforeLength
        SourceWriteTimeUtc = $beforeWriteUtc.ToString('o')
    }
}

function Test-VssPipelineSourceUnchanged {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$StagedSource)

    $source = New-Object System.IO.FileInfo([string]$StagedSource.SourcePath)
    if (-not $source.Exists) { return $false }
    if ([long]$source.Length -ne [long]$StagedSource.SourceLength) { return $false }
    if ($source.LastWriteTimeUtc.ToString('o') -cne [string]$StagedSource.SourceWriteTimeUtc) { return $false }
    $currentHash = Get-VssPipelineSha256 -LiteralPath $source.FullName
    return [string]::Equals($currentHash, [string]$StagedSource.Sha256, [System.StringComparison]::OrdinalIgnoreCase)
}

function Enter-VssPipelineLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$StatePath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RunId,
        [AllowNull()][string]$TargetIdentity = $null
    )

    [void]$RunId
    $identity = $TargetIdentity
    if ([string]::IsNullOrWhiteSpace($identity)) {
        $identity = [System.IO.Path]::GetFullPath($StatePath)
    }
    $lockHash = Get-VssPipelineSha256Text -Text $identity.ToLowerInvariant()
    $mutexName = 'Global\VisioSharePointSync-' + $lockHash
    $mutex = $null
    $acquired = $false
    try {
        $mutex = New-Object System.Threading.Mutex($false, $mutexName)
        try { $acquired = $mutex.WaitOne(0, $false) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw 'Ein anderer Lauf besitzt bereits den zielbezogenen Windows-Mutex.' }
    }
    catch {
        if ($null -ne $mutex -and -not $acquired) { $mutex.Dispose() }
        throw
    }
    return [pscustomobject][ordered]@{ Name = $mutexName; Mutex = $mutex; Acquired = $true }
}

function Exit-VssPipelineLock {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Lock)

    if ($null -ne $Lock.Mutex) {
        try {
            if ([bool]$Lock.Acquired) { $Lock.Mutex.ReleaseMutex() }
        }
        finally { $Lock.Mutex.Dispose() }
    }
}

function Test-VssPipelineStateObject {
    [CmdletBinding()]
    param([AllowNull()][object]$State)

    $errors = @()
    if (-not (Test-VssPipelineObjectContainer -Value $State)) {
        return @('State muss ein JSON-Objekt sein.')
    }
    $errors += @(Get-VssPipelineClosedSchemaErrors -InputObject $State -Path 'State' -AllowedKeys @('SchemaVersion', 'Items'))
    $schema = Get-VssPipelinePropertyValue -InputObject $State -Name 'SchemaVersion'
    if ($schema.Found -and (-not ($schema.Value -is [string]) -or $schema.Value -cne '1.0')) {
        $errors += 'State.SchemaVersion ist nicht unterstuetzt.'
    }
    $itemsResult = Get-VssPipelinePropertyValue -InputObject $State -Name 'Items'
    if (-not $itemsResult.Found) { return $errors }
    if (($itemsResult.Value -is [string]) -or ($itemsResult.Value -is [System.Collections.IDictionary]) -or
        -not ($itemsResult.Value -is [System.Collections.IEnumerable])) {
        $errors += 'State.Items muss ein Array sein.'
        return $errors
    }

    $allowedItemKeys = @(
        'SourceKey', 'SourceRelativePath', 'SourceSha256', 'ConversionFingerprint',
        'PdfSha256', 'TargetRelativePath', 'ApiKind', 'RemoteItemId', 'ETag',
        'LastSuccessfulCommitUtc', 'Provenance'
    )
    $seenSourceKeys = @{}
    $seenTargets = @{}
    foreach ($item in @($itemsResult.Value)) {
        $itemErrors = @(Get-VssPipelineClosedSchemaErrors -InputObject $item -Path 'State.Items[]' -AllowedKeys $allowedItemKeys)
        $errors += $itemErrors
        if ($itemErrors.Count -gt 0) { continue }
        foreach ($fieldName in $allowedItemKeys) {
            $field = Get-VssPipelinePropertyValue -InputObject $item -Name $fieldName
            $isDateObject = (($fieldName -ceq 'LastSuccessfulCommitUtc') -and
                (($field.Value -is [DateTime]) -or ($field.Value -is [DateTimeOffset])))
            if ((-not $isDateObject -and -not ($field.Value -is [string])) -or
                [string]::IsNullOrWhiteSpace([string]$field.Value)) {
                $errors += "State.Items[].$fieldName muss ein nicht leerer String sein."
            }
        }
        if (([string]$item.SourceSha256 -notmatch '^[0-9A-Fa-f]{64}$') -or
            ([string]$item.ConversionFingerprint -notmatch '^[0-9A-Fa-f]{64}$') -or
            ([string]$item.PdfSha256 -notmatch '^[0-9A-Fa-f]{64}$')) {
            $errors += 'State.Items[] enthaelt einen ungueltigen SHA-256-Wert.'
        }
        try {
            $expectedSourceKey = Get-VssPipelineSourceKey -RelativeSourcePath ([string]$item.SourceRelativePath)
            $expectedTarget = Get-VssPipelineTargetPath -RelativeSourcePath ([string]$item.SourceRelativePath)
            if (-not [string]::Equals([string]$item.SourceKey, $expectedSourceKey, [System.StringComparison]::OrdinalIgnoreCase) -or
                -not [string]::Equals([string]$item.Provenance, $expectedSourceKey, [System.StringComparison]::OrdinalIgnoreCase) -or
                -not [string]::Equals([string]$item.TargetRelativePath, $expectedTarget, [System.StringComparison]::OrdinalIgnoreCase)) {
                $errors += 'State.Items[] enthaelt inkonsistente Quell-, Ziel- oder Provenienzpfade.'
            }
        }
        catch { $errors += 'State.Items[] enthaelt einen unsicheren relativen Pfad.' }
        if (-not (Test-VssPipelineContainsOrdinal -Values @('MicrosoftGraphV1', 'SharePointRest') -Expected ([string]$item.ApiKind))) {
            $errors += 'State.Items[].ApiKind ist nicht unterstuetzt.'
        }
        $commitTime = [DateTimeOffset]::MinValue
        $commitTimeValid = $false
        if ($item.LastSuccessfulCommitUtc -is [DateTime]) {
            $dateTimeValue = [DateTime]$item.LastSuccessfulCommitUtc
            $commitTime = [DateTimeOffset]$dateTimeValue.ToUniversalTime()
            $commitTimeValid = $true
        }
        elseif ($item.LastSuccessfulCommitUtc -is [DateTimeOffset]) {
            $commitTime = ([DateTimeOffset]$item.LastSuccessfulCommitUtc).ToUniversalTime()
            $commitTimeValid = $true
        }
        else { $commitTimeValid = [DateTimeOffset]::TryParse([string]$item.LastSuccessfulCommitUtc, [ref]$commitTime) }
        if (-not $commitTimeValid -or $commitTime.Offset -ne [TimeSpan]::Zero) {
            $errors += 'State.Items[].LastSuccessfulCommitUtc muss ein gueltiger UTC-Zeitpunkt sein.'
        }
        if (([string]$item.RemoteItemId).Length -gt 1024 -or ([string]$item.ETag).Length -gt 1024) {
            $errors += 'State.Items[] enthaelt eine ueberlange Remote-ID oder einen ueberlangen eTag.'
        }
        $sourceKey = ([string]$item.SourceKey).Normalize([System.Text.NormalizationForm]::FormC).ToLowerInvariant()
        $targetKey = ([string]$item.TargetRelativePath).Replace([char]92, [char]47).Normalize([System.Text.NormalizationForm]::FormC).ToLowerInvariant()
        if ($seenSourceKeys.ContainsKey($sourceKey)) { $errors += 'State.Items[] enthaelt einen doppelten SourceKey.' }
        else { $seenSourceKeys[$sourceKey] = $true }
        if ($seenTargets.ContainsKey($targetKey)) { $errors += 'State.Items[] enthaelt einen doppelten Zielpfad.' }
        else { $seenTargets[$targetKey] = $true }
    }
    return @($errors)
}

function Read-VssPipelineState {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$LiteralPath)

    if (-not [System.IO.File]::Exists($LiteralPath)) {
        return [pscustomobject]@{ SchemaVersion = '1.0'; Items = @() }
    }
    $text = Read-VssUtf8FileStrict -LiteralPath $LiteralPath
    if (Test-VssJsonHasDuplicateObjectKeys -JsonText $text) { throw 'Die Statusdatei enthaelt doppelte JSON-Schluessel.' }
    $state = ConvertFrom-Json -InputObject $text -ErrorAction Stop
    $stateErrors = @(Test-VssPipelineStateObject -State $state)
    if ($stateErrors.Count -gt 0) { throw 'Die Statusdatei ist beschaedigt oder entspricht nicht dem State-Schema 1.0.' }
    return $state
}

function Save-VssPipelineState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$State,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$LiteralPath
    )

    $stateErrors = @(Test-VssPipelineStateObject -State $State)
    if ($stateErrors.Count -gt 0) { throw 'Der erzeugte State entspricht vor dem Schreiben nicht dem geschlossenen Schema.' }
    $directory = [System.IO.Path]::GetDirectoryName($LiteralPath)
    [void][System.IO.Directory]::CreateDirectory($directory)
    $temporaryPath = Join-Path -Path $directory -ChildPath (([System.IO.Path]::GetFileName($LiteralPath)) + '.' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = $LiteralPath + '.bak'
    $encoding = New-Object System.Text.UTF8Encoding($false, $true)
    try {
        $json = ConvertTo-Json -InputObject $State -Depth 20 -Compress
        [System.IO.File]::WriteAllText($temporaryPath, $json, $encoding)
        if ([System.IO.File]::Exists($LiteralPath)) {
            [System.IO.File]::Replace($temporaryPath, $LiteralPath, $backupPath, $true)
        }
        else { [System.IO.File]::Move($temporaryPath, $LiteralPath) }
    }
    finally {
        if ([System.IO.File]::Exists($temporaryPath)) { [System.IO.File]::Delete($temporaryPath) }
    }
}

function Get-VssPipelineSourceKey {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RelativeSourcePath)

    # Die Zielplattformen behandeln Quell- und SharePoint-Pfade praktisch
    # case-insensitive; der stabile State-Key bildet dies explizit ab.
    $targetProbe = Get-VssPipelineTargetPath -RelativeSourcePath $RelativeSourcePath
    [void]$targetProbe
    return $RelativeSourcePath.Replace([char]92, [char]47).Normalize([System.Text.NormalizationForm]::FormC).ToLowerInvariant()
}

function Get-VssPipelineStateItem {
    param(
        [Parameter(Mandatory = $true)][object]$State,
        [Parameter(Mandatory = $true)][string]$SourceKey
    )
    foreach ($item in @($State.Items)) {
        if ([string]::Equals([string]$item.SourceKey, $SourceKey, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $item
        }
    }
    return $null
}

function Set-VssPipelineStateItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$State,
        [Parameter(Mandatory = $true)][object]$Item
    )

    $items = @()
    foreach ($existing in @($State.Items)) {
        if (-not [string]::Equals([string]$existing.SourceKey, [string]$Item.SourceKey, [System.StringComparison]::OrdinalIgnoreCase)) {
            $items += $existing
        }
    }
    $items += $Item
    return [pscustomobject][ordered]@{
        SchemaVersion = '1.0'
        Items         = @($items | Sort-Object -Property SourceKey)
    }
}

function Get-VssPipelineConversionFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$SourceSha256,
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration
    )

    $profile = [ordered]@{
        OutputFormat        = [string]$Configuration.Conversion.OutputFormat
        PageRange           = [string]$Configuration.Conversion.PageRange
        Intent              = [string]$Configuration.Conversion.Intent
        DisableMacros       = [bool]$Configuration.Conversion.DisableMacros
        RefreshExternalData = [bool]$Configuration.Conversion.RefreshExternalData
        AdapterKind         = [string]$RuntimeConfiguration.Conversion.AdapterKind
        AdapterVersion      = [string]$RuntimeConfiguration.Conversion.AdapterVersion
    }
    $profileJson = ConvertTo-Json -InputObject $profile -Depth 5 -Compress
    return Get-VssPipelineSha256Text -Text ($SourceSha256.ToLowerInvariant() + '|' + $profileJson)
}

function Resolve-VssPipelineTargetCollisions {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Candidates)

    $byTarget = @{}
    $plans = @()
    foreach ($candidate in @($Candidates | Sort-Object -Property RelativePath)) {
        $targetPath = Get-VssPipelineTargetPath -RelativeSourcePath ([string]$candidate.RelativePath)
        $key = $targetPath.Normalize([System.Text.NormalizationForm]::FormC).ToLowerInvariant()
        $plan = [pscustomobject][ordered]@{
            Candidate          = $candidate
            SourceKey          = Get-VssPipelineSourceKey -RelativeSourcePath ([string]$candidate.RelativePath)
            TargetRelativePath = $targetPath
            HasCollision       = $false
        }
        if (-not $byTarget.ContainsKey($key)) { $byTarget[$key] = @() }
        $byTarget[$key] = @($byTarget[$key]) + @($plan)
        $plans += $plan
    }
    foreach ($key in @($byTarget.Keys)) {
        if (@($byTarget[$key]).Count -gt 1) {
            foreach ($plan in @($byTarget[$key])) { $plan.HasCollision = $true }
        }
    }
    return @($plans)
}

function Resolve-VssPipelineRemoteWriteDecision {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourceKey,
        [Parameter(Mandatory = $true)][string]$TargetRelativePath,
        [AllowNull()][object]$StateItem,
        [Parameter(Mandatory = $true)][object]$RemoteTarget
    )

    if (-not [bool]$RemoteTarget.Exists) {
        return [pscustomobject][ordered]@{ Action = 'Create'; Reason = 'TargetAbsent'; IfMatch = $null; RemoteItemId = $null }
    }
    if ($null -eq $StateItem) {
        return [pscustomobject][ordered]@{ Action = 'Conflict'; Reason = 'UnknownExistingTarget'; IfMatch = $null; RemoteItemId = $null }
    }
    $matches = (
        [string]::Equals([string]$StateItem.RemoteItemId, [string]$RemoteTarget.RemoteItemId, [System.StringComparison]::Ordinal) -and
        [string]::Equals([string]$StateItem.TargetRelativePath, $TargetRelativePath, [System.StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals([string]$RemoteTarget.TargetRelativePath, $TargetRelativePath, [System.StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals([string]$RemoteTarget.Provenance, $SourceKey, [System.StringComparison]::OrdinalIgnoreCase)
    )
    if (-not $matches) {
        return [pscustomobject][ordered]@{ Action = 'Conflict'; Reason = 'RemoteIdentityOrProvenanceMismatch'; IfMatch = $null; RemoteItemId = $null }
    }
    if (-not [string]::Equals([string]$StateItem.ETag, [string]$RemoteTarget.ETag, [System.StringComparison]::Ordinal)) {
        return [pscustomobject][ordered]@{ Action = 'Conflict'; Reason = 'ETagConflict'; IfMatch = $null; RemoteItemId = $null }
    }
    return [pscustomobject][ordered]@{
        Action       = 'Update'
        Reason       = 'KnownManagedTarget'
        IfMatch      = [string]$StateItem.ETag
        RemoteItemId = [string]$StateItem.RemoteItemId
    }
}

function Get-VssPipelineOrphanStateItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$State,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$CandidatePlans
    )
    $seen = @{}
    foreach ($plan in $CandidatePlans) { $seen[[string]$plan.SourceKey] = $true }
    $orphans = @()
    foreach ($item in @($State.Items)) {
        if (-not $seen.ContainsKey(([string]$item.SourceKey).ToLowerInvariant())) { $orphans += $item }
    }
    return @($orphans)
}

function ConvertTo-VssPipelineRedactedLogObject {
    param([AllowNull()][object]$InputObject)

    if ($null -eq $InputObject) { return $null }
    # Feste Positivliste statt best-effort-Redaktion. Dadurch gelangen weder
    # Upload-URLs noch Token oder absolute Quellpfade in die JSONL-Datei, auch
    # wenn ein Adapter zusaetzliche Eigenschaften an sein Event haengt.
    $copy = [ordered]@{}
    foreach ($fieldName in @('RunId', 'StageId', 'RelativePath', 'Result', 'DurationMs', 'RetryCount', 'HttpStatus', 'RequestId')) {
        $field = Get-VssPipelinePropertyValue -InputObject $InputObject -Name $fieldName
        if (-not $field.Found -or $null -eq $field.Value) { continue }
        if ($fieldName -in @('DurationMs', 'RetryCount', 'HttpStatus')) {
            $copy[$fieldName] = [long]$field.Value
            continue
        }
        $value = [string]$field.Value
        if ($fieldName -ceq 'RunId') {
            $parsedRunId = [guid]::Empty
            if (-not [guid]::TryParseExact($value, 'D', [ref]$parsedRunId)) { $value = '[REDACTED]' }
        }
        elseif ($fieldName -ceq 'RelativePath') {
            if ([System.IO.Path]::IsPathRooted($value) -or $value -match '(?i)^[a-z][a-z0-9+.-]*://' -or $value.Contains(':')) {
                $value = '[REDACTED]'
            }
            else { $value = $value.Replace([char]92, [char]47) }
            if ($value -match '(?i)(bearer\s|token|secret|password|uploadurl|sig=)') { $value = '[REDACTED]' }
        }
        elseif (($fieldName -ceq 'StageId') -or ($fieldName -ceq 'RequestId')) {
            if ($value -notmatch '^[A-Za-z0-9._-]{1,128}$') { $value = '[REDACTED]' }
        }
        elseif ($fieldName -ceq 'Result') {
            if ($value -notmatch '^[A-Z0-9_.-]{1,64}$') { $value = '[REDACTED]' }
        }
        $copy[$fieldName] = $value
    }
    return [pscustomobject]$copy
}

function Write-VssPipelineLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$LogPath,
        [Parameter(Mandatory = $true)][object]$Event
    )

    [void][System.IO.Directory]::CreateDirectory($LogPath)
    $logFile = Join-Path -Path $LogPath -ChildPath ('visio-sharepoint-sync-{0}.jsonl' -f [DateTime]::UtcNow.ToString('yyyyMMdd'))
    $redacted = ConvertTo-VssPipelineRedactedLogObject -InputObject $Event
    $line = (ConvertTo-Json -InputObject $redacted -Depth 20 -Compress) + [Environment]::NewLine
    $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($line)
    $stream = New-Object System.IO.FileStream($logFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) }
    finally { $stream.Dispose() }
    return $logFile
}

function Get-VssPipelineRetryAfterSeconds {
    param(
        [Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [scriptblock]$UtcNowProvider = { [DateTimeOffset]::UtcNow }
    )

    try {
        if ($null -ne $ErrorRecord.Exception.Data -and $null -ne $ErrorRecord.Exception.Data['RetryAfter']) {
            $dataSeconds = 0
            if ([int]::TryParse([string]$ErrorRecord.Exception.Data['RetryAfter'], [ref]$dataSeconds) -and $dataSeconds -ge 0) {
                return $dataSeconds
            }
        }
        $response = $ErrorRecord.Exception.Response
        if ($null -ne $response -and $null -ne $response.Headers) {
            $value = $null
            try { $value = $response.Headers['Retry-After'] }
            catch { $value = $null }
            if ($null -eq $value) {
                $retryAfterHeader = Get-VssPipelinePropertyValue -InputObject $response.Headers -Name 'RetryAfter'
                if ($retryAfterHeader.Found -and $null -ne $retryAfterHeader.Value) {
                    $delta = Get-VssPipelinePropertyValue -InputObject $retryAfterHeader.Value -Name 'Delta'
                    $date = Get-VssPipelinePropertyValue -InputObject $retryAfterHeader.Value -Name 'Date'
                    if ($delta.Found -and $null -ne $delta.Value) { return [int][Math]::Ceiling(([TimeSpan]$delta.Value).TotalSeconds) }
                    if ($date.Found -and $null -ne $date.Value) { $value = ([DateTimeOffset]$date.Value).ToString('r') }
                }
            }
            $seconds = 0
            if ($null -ne $value -and [int]::TryParse([string]$value, [ref]$seconds) -and $seconds -ge 0) { return $seconds }
            $retryDate = [DateTimeOffset]::MinValue
            if ($null -ne $value -and [DateTimeOffset]::TryParse([string]$value, [ref]$retryDate)) {
                $now = [DateTimeOffset](& $UtcNowProvider)
                return [int][Math]::Max(0, [Math]::Ceiling(($retryDate - $now).TotalSeconds))
            }
        }
    }
    catch { }
    return $null
}

function Get-VssPipelineHttpStatusCode {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    try {
        if ($null -ne $ErrorRecord.Exception.Data -and $null -ne $ErrorRecord.Exception.Data['StatusCode']) {
            return [int]$ErrorRecord.Exception.Data['StatusCode']
        }
        $response = $ErrorRecord.Exception.Response
        if ($null -eq $response) { return $null }
        $statusProperty = Get-VssPipelinePropertyValue -InputObject $response -Name 'StatusCode'
        if ($statusProperty.Found) { return [int]$statusProperty.Value }
    }
    catch { }
    return $null
}

function Test-VssPipelineRetryableError {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $statusCode = Get-VssPipelineHttpStatusCode -ErrorRecord $ErrorRecord
    if ($null -eq $statusCode) {
        $exception = $ErrorRecord.Exception
        if ($exception -is [System.TimeoutException] -or $exception -is [System.Threading.Tasks.TaskCanceledException]) { return $true }
        if ($exception -is [System.Net.WebException]) {
            return ($exception.Status -in @(
                [System.Net.WebExceptionStatus]::Timeout,
                [System.Net.WebExceptionStatus]::ConnectionClosed,
                [System.Net.WebExceptionStatus]::ConnectFailure,
                [System.Net.WebExceptionStatus]::NameResolutionFailure,
                [System.Net.WebExceptionStatus]::ReceiveFailure,
                [System.Net.WebExceptionStatus]::SendFailure,
                [System.Net.WebExceptionStatus]::KeepAliveFailure
            ))
        }
        return ([string]$exception.GetType().FullName -ceq 'System.Net.Http.HttpRequestException')
    }
    return (($statusCode -eq 429) -or ($statusCode -ge 500 -and $statusCode -le 599))
}

function Invoke-VssPipelineWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [Parameter(Mandatory = $true)][ValidateRange(0, 10)][int]$MaxRetryCount,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$OperationName,
        [scriptblock]$DelayAction = { param([double]$Seconds) Start-Sleep -Milliseconds ([int][Math]::Ceiling($Seconds * 1000.0)) },
        [scriptblock]$JitterProvider = { (Get-Random -Minimum 0 -Maximum 1000) / 1000.0 },
        [scriptblock]$UtcNowProvider = { [DateTimeOffset]::UtcNow }
    )

    for ($attempt = 0; $attempt -le $MaxRetryCount; $attempt++) {
        try { return & $Operation $attempt }
        catch {
            if ($attempt -ge $MaxRetryCount -or -not (Test-VssPipelineRetryableError -ErrorRecord $_)) { throw }
            $retryAfter = Get-VssPipelineRetryAfterSeconds -ErrorRecord $_ -UtcNowProvider $UtcNowProvider
            if ($null -eq $retryAfter) {
                $baseDelay = [Math]::Min([double]60.0, [Math]::Pow(2, $attempt + 1))
                $jitter = [double](& $JitterProvider)
                if ($jitter -lt 0) { $jitter = 0 }
                if ($jitter -gt 1) { $jitter = 1 }
                $retryAfter = [Math]::Min([double]60.0, [double]($baseDelay + $jitter))
            }
            $retryAfter = [Math]::Min([double]60.0, [Math]::Max([double]0.0, [double]$retryAfter))
            Write-Verbose ("{0}: Wiederholung {1}/{2} nach {3} Sekunden." -f $OperationName, ($attempt + 1), $MaxRetryCount, $retryAfter)
            & $DelayAction $retryAfter
        }
    }
}

function Remove-VssPipelineStagingRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ConfiguredStagingRoot,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RunStagingPath
    )

    $root = [System.IO.Path]::GetFullPath($ConfiguredStagingRoot).TrimEnd([char]92)
    $target = [System.IO.Path]::GetFullPath($RunStagingPath).TrimEnd([char]92)
    if ($target.Length -le $root.Length -or -not $target.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Der Staging-Laufpfad liegt nicht sicher unterhalb der konfigurierten Staging-Wurzel.'
    }
    $runDirectoryName = [System.IO.Path]::GetFileName($target)
    $runGuid = [guid]::Empty
    if (-not [guid]::TryParseExact($runDirectoryName, 'D', [ref]$runGuid)) {
        throw 'Der Staging-Laufpfad besitzt keine eindeutige Run-ID.'
    }
    if ([System.IO.Directory]::Exists($target)) {
        $pending = New-Object 'System.Collections.Generic.Stack[System.IO.DirectoryInfo]'
        $pending.Push((New-Object System.IO.DirectoryInfo($target)))
        while ($pending.Count -gt 0) {
            $directory = $pending.Pop()
            if (Test-VssPipelineFileAttribute -Attributes $directory.Attributes -Expected ([System.IO.FileAttributes]::ReparsePoint)) {
                throw 'Das Staging enthaelt einen nicht zulaessigen ReparsePoint.'
            }
            foreach ($child in @($directory.GetDirectories())) { $pending.Push($child) }
        }
        [System.IO.Directory]::Delete($target, $true)
    }
}

function Test-VssPipelineHeaderPresent {
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Headers,
        [Parameter(Mandatory = $true)][string]$Name
    )
    foreach ($key in $Headers.Keys) {
        if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Copy-VssPipelineHeaders {
    param([AllowNull()][System.Collections.IDictionary]$Headers)
    $copy = @{}
    if ($null -ne $Headers) {
        foreach ($key in $Headers.Keys) { $copy[[string]$key] = $Headers[$key] }
    }
    return $copy
}

function Send-VssGraphUploadSessionChunks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$UploadUrl,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ContentPath,
        [Parameter(Mandatory = $true)][ValidateRange(327680, 62914560)][int]$ChunkSizeBytes,
        [Parameter(Mandatory = $true)][ValidateRange(0, 10)][int]$MaxRetryCount,
        [Parameter(Mandatory = $true)][switch]$AllowExternalSideEffects,
        [AllowNull()][scriptblock]$ChunkTransport = $null
    )

    if (-not $AllowExternalSideEffects) { throw 'Graph-Upload ist ohne explizite Side-Effect-Freigabe blockiert.' }
    if (($ChunkSizeBytes % 327680) -ne 0) { throw 'Graph-Chunkgroesse muss ein Vielfaches von 327680 Bytes sein.' }
    $file = New-Object System.IO.FileInfo($ContentPath)
    if (-not $file.Exists) { throw 'Die hochzuladende PDF-Datei wurde nicht gefunden.' }
    $stream = New-Object System.IO.FileStream($ContentPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $offset = [long]0
        $lastResponse = $null
        while ($offset -lt $file.Length) {
            $remaining = $file.Length - $offset
            $requested = [int][Math]::Min([long]$ChunkSizeBytes, $remaining)
            $buffer = New-Object byte[] $requested
            $read = $stream.Read($buffer, 0, $requested)
            if ($read -le 0) { throw 'Der PDF-Datenstrom endete unerwartet.' }
            if ($read -ne $buffer.Length) {
                $payload = New-Object byte[] $read
                [System.Array]::Copy($buffer, 0, $payload, 0, $read)
            }
            else { $payload = $buffer }
            $chunkStart = $offset
            $chunkEnd = $offset + $read - 1
            $chunkHeaders = @{
                'Content-Length' = [string]$read
                'Content-Range'  = "bytes $chunkStart-$chunkEnd/$($file.Length)"
            }
            $operation = {
                param($attempt)
                if ($null -ne $ChunkTransport) {
                    return & $ChunkTransport $UploadUrl $chunkHeaders $payload $attempt
                }
                Invoke-RestMethod -Method Put -Uri $UploadUrl -Headers $chunkHeaders -Body $payload -ContentType 'application/octet-stream' -ErrorAction Stop
            }.GetNewClosure()
            $lastResponse = Invoke-VssPipelineWithRetry -Operation $operation -MaxRetryCount $MaxRetryCount -OperationName 'GraphUploadChunk'
            $nextOffset = $offset + $read
            $ranges = Get-VssPipelinePropertyValue -InputObject $lastResponse -Name 'nextExpectedRanges'
            if ($ranges.Found -and @($ranges.Value).Count -gt 0) {
                $firstRange = [string]@($ranges.Value)[0]
                if ($firstRange -notmatch '^(?<start>[0-9]+)-') { throw 'Graph lieferte einen ungueltigen Upload-Session-Fortsetzungsbereich.' }
                $nextOffset = [long]$Matches['start']
                if ($nextOffset -lt 0 -or $nextOffset -gt $file.Length) { throw 'Graph lieferte einen Upload-Offset ausserhalb der PDF-Datei.' }
            }
            elseif (($null -ne $lastResponse) -and
                (Get-VssPipelinePropertyValue -InputObject $lastResponse -Name 'id').Found) {
                $nextOffset = $file.Length
            }
            if ($nextOffset -le $offset -and $nextOffset -lt $file.Length) { throw 'Die Graph-Upload-Session meldete keinen Fortschritt.' }
            $offset = $nextOffset
            if ($stream.Position -ne $offset) { [void]$stream.Seek($offset, [System.IO.SeekOrigin]::Begin) }
        }
        return $lastResponse
    }
    finally { $stream.Dispose() }
}

function Invoke-VssGraphTransport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$RequestPlan,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ContentPath,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$AuthorizationHeaders,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration,
        [Parameter(Mandatory = $true)][ValidateRange(0, 10)][int]$MaxRetryCount,
        [Parameter(Mandatory = $true)][switch]$AllowExternalSideEffects
    )

    if (-not $AllowExternalSideEffects) { throw 'Graph-Transport ist ohne explizite Side-Effect-Freigabe blockiert.' }
    if ([string]$RequestPlan.Transport -cne 'MicrosoftGraphV1') { throw 'Der Requestplan ist nicht fuer Microsoft Graph bestimmt.' }
    if (-not (Test-VssPipelineHeaderPresent -Headers $AuthorizationHeaders -Name 'Authorization')) {
        throw 'Der Graph-Transport benoetigt einen extern bereitgestellten Authorization-Header.'
    }
    $headers = Copy-VssPipelineHeaders -Headers $AuthorizationHeaders
    foreach ($key in $RequestPlan.Headers.Keys) {
        if (-not [string]::Equals([string]$key, 'Content-Type', [System.StringComparison]::OrdinalIgnoreCase)) {
            $headers[[string]$key] = $RequestPlan.Headers[$key]
        }
    }
    if ([string]$RequestPlan.Operation -ceq 'SmallUpload') {
        $operation = {
            param($attempt)
            [void]$attempt
            Invoke-RestMethod -Method Put -Uri ([string]$RequestPlan.Uri) -Headers $headers -InFile $ContentPath -ContentType 'application/pdf' -ErrorAction Stop
        }.GetNewClosure()
        return Invoke-VssPipelineWithRetry -Operation $operation -MaxRetryCount $MaxRetryCount -OperationName 'GraphSmallUpload'
    }
    if ([string]$RequestPlan.Operation -ceq 'CreateUploadSession') {
        $conflictBehavior = if (Test-VssPipelineHeaderPresent -Headers $RequestPlan.Headers -Name 'If-None-Match') { 'fail' } else { 'replace' }
        $body = @{ item = @{ '@microsoft.graph.conflictBehavior' = $conflictBehavior } } | ConvertTo-Json -Depth 5 -Compress
        $sessionOperation = {
            param($attempt)
            [void]$attempt
            Invoke-RestMethod -Method Post -Uri ([string]$RequestPlan.Uri) -Headers $headers -Body $body -ContentType 'application/json' -ErrorAction Stop
        }.GetNewClosure()
        $session = Invoke-VssPipelineWithRetry -Operation $sessionOperation -MaxRetryCount $MaxRetryCount -OperationName 'GraphCreateUploadSession'
        if ($null -eq $session -or [string]::IsNullOrWhiteSpace([string]$session.uploadUrl)) {
            throw 'Microsoft Graph lieferte keine Upload-Session-URL.'
        }
        # uploadUrl is preauthenticated and must never be logged or returned by
        # this layer. It is consumed only by the chunk sender.
        return Send-VssGraphUploadSessionChunks -UploadUrl ([string]$session.uploadUrl) -ContentPath $ContentPath -ChunkSizeBytes ([int]$RuntimeConfiguration.Upload.ChunkSizeBytes) -MaxRetryCount $MaxRetryCount -AllowExternalSideEffects
    }
    throw 'Der Graph-Requestplan enthaelt eine nicht unterstuetzte Operation.'
}

function Get-VssRestSiteUrlFromRequestPlan {
    param([Parameter(Mandatory = $true)][object]$RequestPlan)
    $uriText = [string]$RequestPlan.Uri
    $markerIndex = $uriText.IndexOf('/_api/', [System.StringComparison]::OrdinalIgnoreCase)
    if ($markerIndex -lt 1) { throw 'Aus dem REST-Requestplan konnte keine Site-URL ermittelt werden.' }
    return $uriText.Substring(0, $markerIndex)
}

function Get-VssRestRequestDigest {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [AllowNull()][System.Collections.IDictionary]$AuthorizationHeaders,
        [bool]$UseDefaultCredentials,
        [Parameter(Mandatory = $true)][int]$MaxRetryCount
    )

    $headers = Copy-VssPipelineHeaders -Headers $AuthorizationHeaders
    $contextUri = $SiteUrl.TrimEnd('/') + '/_api/contextinfo'
    $operation = {
        param($attempt)
        [void]$attempt
        if ($UseDefaultCredentials) {
            return Invoke-RestMethod -Method Post -Uri $contextUri -Headers $headers -UseDefaultCredentials -ContentType 'application/json;odata=verbose' -ErrorAction Stop
        }
        return Invoke-RestMethod -Method Post -Uri $contextUri -Headers $headers -ContentType 'application/json;odata=verbose' -ErrorAction Stop
    }.GetNewClosure()
    $response = Invoke-VssPipelineWithRetry -Operation $operation -MaxRetryCount $MaxRetryCount -OperationName 'SharePointContextInfo'
    $digest = $null
    if ($null -ne $response -and $null -ne $response.d -and $null -ne $response.d.GetContextWebInformation) {
        $digest = $response.d.GetContextWebInformation.FormDigestValue
    }
    elseif ($null -ne $response -and $null -ne $response.FormDigestValue) { $digest = $response.FormDigestValue }
    if ([string]::IsNullOrWhiteSpace([string]$digest)) { throw 'SharePoint lieferte keinen Request-Digest.' }
    return [string]$digest
}

function Invoke-VssRestTransport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$RequestPlan,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ContentPath,
        [AllowNull()][System.Collections.IDictionary]$AuthorizationHeaders,
        [switch]$UseDefaultCredentials,
        [Parameter(Mandatory = $true)][ValidateRange(0, 10)][int]$MaxRetryCount,
        [Parameter(Mandatory = $true)][switch]$AllowExternalSideEffects
    )

    if (-not $AllowExternalSideEffects) { throw 'SharePoint-REST-Transport ist ohne explizite Side-Effect-Freigabe blockiert.' }
    if ([string]$RequestPlan.Transport -cne 'SharePointRest') { throw 'Der Requestplan ist nicht fuer SharePoint REST bestimmt.' }
    if ([string]$RequestPlan.Operation -ceq 'LargeUploadPlaceholder') {
        # >>> PLACEHOLDER [VSS-SPREST-LARGE-001] BEGIN
        throw [System.NotSupportedException]::new(
            '[VSS-SPREST-LARGE-001] Der grosse SharePoint-REST-Chunk-Upload ist fuer die konkrete Server-Version noch nicht verifiziert.'
        )
        # <<< PLACEHOLDER [VSS-SPREST-LARGE-001] END
    }
    if (-not $UseDefaultCredentials -and $null -eq $AuthorizationHeaders) {
        throw 'REST-Transport benoetigt WindowsIntegrated oder extern bereitgestellte Authorization-Header.'
    }
    $headers = Copy-VssPipelineHeaders -Headers $AuthorizationHeaders
    foreach ($key in $RequestPlan.Headers.Keys) {
        if (-not [string]::Equals([string]$key, 'X-RequestDigest', [System.StringComparison]::OrdinalIgnoreCase) -and
            -not [string]::Equals([string]$key, 'Content-Type', [System.StringComparison]::OrdinalIgnoreCase)) {
            $headers[[string]$key] = $RequestPlan.Headers[$key]
        }
    }
    $siteUrl = Get-VssRestSiteUrlFromRequestPlan -RequestPlan $RequestPlan
    $headers['X-RequestDigest'] = Get-VssRestRequestDigest -SiteUrl $siteUrl -AuthorizationHeaders $AuthorizationHeaders -UseDefaultCredentials ([bool]$UseDefaultCredentials) -MaxRetryCount $MaxRetryCount
    $headers['Accept'] = 'application/json;odata=nometadata'
    $operation = {
        param($attempt)
        [void]$attempt
        if ($UseDefaultCredentials) {
            return Invoke-RestMethod -Method Post -Uri ([string]$RequestPlan.Uri) -Headers $headers -UseDefaultCredentials -InFile $ContentPath -ContentType 'application/pdf' -ErrorAction Stop
        }
        return Invoke-RestMethod -Method Post -Uri ([string]$RequestPlan.Uri) -Headers $headers -InFile $ContentPath -ContentType 'application/pdf' -ErrorAction Stop
    }.GetNewClosure()
    return Invoke-VssPipelineWithRetry -Operation $operation -MaxRetryCount $MaxRetryCount -OperationName 'SharePointRestUpload'
}

function ConvertFrom-VssPipelineRemoteCommitResponse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('MicrosoftGraphV1', 'SharePointRest')][string]$ApiKind,
        [AllowNull()][object]$Response,
        [Parameter(Mandatory = $true)][string]$TargetRelativePath,
        [Parameter(Mandatory = $true)][string]$SourceKey,
        [scriptblock]$UtcNowProvider = { [DateTime]::UtcNow }
    )

    if ($null -eq $Response) { throw 'Der Remote-Commit lieferte keine bestaetigbare Antwort.' }
    $prebuilt = Get-VssPipelinePropertyValue -InputObject $Response -Name 'Committed'
    if ($prebuilt.Found) {
        if (-not [bool]$prebuilt.Value) { throw 'Der Remote-Adapter hat den Commit nicht bestaetigt.' }
        $remoteId = [string]$Response.RemoteItemId
        $etag = [string]$Response.ETag
        $requestId = [string]$Response.RequestId
        $committedUtc = [string]$Response.CommittedUtc
    }
    else {
        $id = Get-VssPipelinePropertyValue -InputObject $Response -Name 'id'
        if (-not $id.Found) { $id = Get-VssPipelinePropertyValue -InputObject $Response -Name 'UniqueId' }
        $etagResult = Get-VssPipelinePropertyValue -InputObject $Response -Name 'eTag'
        if (-not $etagResult.Found) { $etagResult = Get-VssPipelinePropertyValue -InputObject $Response -Name 'ETag' }
        $request = Get-VssPipelinePropertyValue -InputObject $Response -Name 'requestId'
        $remoteId = if ($id.Found) { [string]$id.Value } else { $null }
        $etag = if ($etagResult.Found) { [string]$etagResult.Value } else { $null }
        $requestId = if ($request.Found) { [string]$request.Value } else { $null }
        $committedUtc = ([DateTime](& $UtcNowProvider)).ToUniversalTime().ToString('o')
    }
    if ([string]::IsNullOrWhiteSpace($remoteId) -or [string]::IsNullOrWhiteSpace($etag)) {
        throw 'Die Remote-Antwort enthaelt keine bestaetigte Item-ID mit eTag.'
    }
    if ([string]::IsNullOrWhiteSpace($committedUtc)) {
        $committedUtc = ([DateTime](& $UtcNowProvider)).ToUniversalTime().ToString('o')
    }
    return [pscustomobject][ordered]@{
        Committed          = $true
        ApiKind            = $ApiKind
        RemoteItemId       = $remoteId
        ETag               = $etag
        RequestId          = $requestId
        CommittedUtc       = $committedUtc
        TargetRelativePath = $TargetRelativePath
        Provenance         = $SourceKey
    }
}

function Get-VssPipelineRemoteTargetMetadata {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Request)
    [void]$Request
    # >>> PLACEHOLDER [VSS-INTEGRATION-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-INTEGRATION-001] Die reale Zielabfrage samt Item-ID-, eTag- und Provenienzpruefung wartet auf eine dedizierte Testziel-Allowlist.'
    )
    # <<< PLACEHOLDER [VSS-INTEGRATION-001] END
}

function Invoke-VssVisioComConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$StagedSource,
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration
    )
    [void]$StagedSource
    [void]$Configuration
    [void]$RuntimeConfiguration
    # Der spaetere STA-isolierte Provider soll Visios Document.ExportAsFixedFormat
    # fuer PDF verwenden. Bis Makro-Policy, COM-Lebenszyklus und Pilotdateien
    # freigegeben sind, wird hier absichtlich kein COM-Objekt erzeugt.
    # >>> PLACEHOLDER [VSS-CNV-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-CNV-001] Kein produktiver, getesteter Visio-COM-Konvertierungsadapter ist ausgewaehlt.'
    )
    # <<< PLACEHOLDER [VSS-CNV-001] END
}

function Invoke-VssExternalConverterConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$StagedSource,
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration
    )
    [void]$StagedSource
    [void]$Configuration
    [void]$RuntimeConfiguration
    # >>> PLACEHOLDER [VSS-CNV-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-CNV-001] Kein produktiver, getesteter externer Konvertierungsadapter ist ausgewaehlt.'
    )
    # <<< PLACEHOLDER [VSS-CNV-001] END
}

function Invoke-VssEnsureSharePointFolders {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$RequestPlan)
    [void]$RequestPlan
    # >>> PLACEHOLDER [VSS-INTEGRATION-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-INTEGRATION-001] Reale Ziel-Allowlist, idempotente Ordneranlage und Integrationstest-Freigabe fehlen.'
    )
    # <<< PLACEHOLDER [VSS-INTEGRATION-001] END
}

function Resolve-VssSharePointRestPathContract {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Configuration)
    [void]$Configuration
    # >>> PLACEHOLDER [VSS-SPREST-PATH-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-SPREST-PATH-001] Prozent-/Rautezeichen und ResourcePath sind fuer die konkrete SharePoint-Version noch nicht verifiziert.'
    )
    # <<< PLACEHOLDER [VSS-SPREST-PATH-001] END
}

function Get-VssPipelineAuthorizationHeaders {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Configuration)
    [void]$Configuration
    # >>> PLACEHOLDER [VSS-AUTH-GRAPH-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-AUTH-GRAPH-001] Ein freigegebener, nicht protokollierender Graph-Tokenprovider fehlt.'
    )
    # <<< PLACEHOLDER [VSS-AUTH-GRAPH-001] END
}

function Invoke-VssReconcileLocalAndRemoteState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Inventory,
        [Parameter(Mandatory = $true)][object]$State,
        [Parameter(Mandatory = $true)][object]$Configuration
    )
    [void]$Inventory
    [void]$State
    [void]$Configuration
    # Dieser erste PublishOnly-Stand darf ausschliesslich melden. Es gibt bewusst
    # keinen Delete-, Move- oder Archive-Adapter.
    # >>> PLACEHOLDER [VSS-INTEGRATION-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-INTEGRATION-001] Die reale Remote-Provenienzpruefung wartet auf eine explizite Testziel-Allowlist.'
    )
    # <<< PLACEHOLDER [VSS-INTEGRATION-001] END
}

function Assert-VssPipelineGoLiveApproval {
    [CmdletBinding()]
    param()
    # >>> PLACEHOLDER [VSS-GOLIVE-001] BEGIN
    throw [System.NotSupportedException]::new(
        '[VSS-GOLIVE-001] Produktivfreigabe, Pilotnachweis und Betriebsuebergabe sind noch nicht bestaetigt.'
    )
    # <<< PLACEHOLDER [VSS-GOLIVE-001] END
}

function Assert-VssPipelineRuntimeQuarantineConfigured {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$RuntimeConfiguration)
    if ([string]$RuntimeConfiguration.Operations.QuarantinePath -ceq $script:VssPipelinePlaceholderValue) {
        # >>> PLACEHOLDER [VSS-RUNTIME-QUARANTINE-001] BEGIN
        throw [System.NotSupportedException]::new(
            '[VSS-RUNTIME-QUARANTINE-001] Der lokale Quarantaenepfad ist noch nicht festgelegt.'
        )
        # <<< PLACEHOLDER [VSS-RUNTIME-QUARANTINE-001] END
    }
}

function New-VssPipelineCapability {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet('Ready', 'Placeholder')][string]$Status,
        [AllowNull()][string]$PlaceholderId = $null
    )
    return [pscustomobject][ordered]@{ Name = $Name; Status = $Status; PlaceholderId = $PlaceholderId }
}

function Get-VssPipelineTargetIdentity {
    param([Parameter(Mandatory = $true)][object]$Configuration)

    if ([string]$Configuration.SharePoint.ApiKind -ceq 'MicrosoftGraphV1') {
        return ('MicrosoftGraphV1|{0}|{1}' -f [string]$Configuration.SharePoint.DriveId, [string]$Configuration.SharePoint.TargetFolderId)
    }
    return ('SharePointRest|{0}|{1}|{2}' -f
        ([string]$Configuration.SharePoint.SiteUrl).TrimEnd('/').ToLowerInvariant(),
        ([string]$Configuration.SharePoint.LibraryName).ToLowerInvariant(),
        ([string]$Configuration.SharePoint.TargetFolderPath).Replace([char]92, [char]47).ToLowerInvariant())
}

function Get-VssPipelineAdapterMember {
    param(
        [Parameter(Mandatory = $true)][object]$AdapterSet,
        [Parameter(Mandatory = $true)][string]$Name
    )
    $member = Get-VssPipelinePropertyValue -InputObject $AdapterSet -Name $Name
    if (-not $member.Found) { throw "Dem internen Adapterset fehlt der Handler $Name." }
    return $member.Value
}

function Invoke-VssPipelineAdapterOperation {
    param(
        [Parameter(Mandatory = $true)][object]$AdapterSet,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][object]$Request
    )
    $handler = Get-VssPipelineAdapterMember -AdapterSet $AdapterSet -Name $Name
    if (-not ($handler -is [scriptblock])) { throw "Der interne Adapterhandler $Name ist nicht aufrufbar." }
    return & $handler $Request
}

function New-VssPipelineProductionAdapterSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration
    )

    $capabilities = [ordered]@{
        Lock               = New-VssPipelineCapability -Name 'Lock' -Status Ready
        State              = New-VssPipelineCapability -Name 'State' -Status Ready
        Inventory          = New-VssPipelineCapability -Name 'Inventory' -Status Ready
        Staging            = New-VssPipelineCapability -Name 'Staging' -Status Ready
        Conversion         = New-VssPipelineCapability -Name 'Conversion' -Status Placeholder -PlaceholderId 'VSS-CNV-001'
        SourceVerification = New-VssPipelineCapability -Name 'SourceVerification' -Status Ready
        RemoteInspection   = New-VssPipelineCapability -Name 'RemoteInspection' -Status Placeholder -PlaceholderId 'VSS-INTEGRATION-001'
        FolderResolution   = New-VssPipelineCapability -Name 'FolderResolution' -Status Placeholder -PlaceholderId 'VSS-INTEGRATION-001'
        Upload             = New-VssPipelineCapability -Name 'Upload' -Status Placeholder -PlaceholderId 'VSS-INTEGRATION-001'
        Reconciliation     = New-VssPipelineCapability -Name 'Reconciliation' -Status Ready
        Logging            = New-VssPipelineCapability -Name 'Logging' -Status Ready
        Cleanup            = New-VssPipelineCapability -Name 'Cleanup' -Status Ready
        GoLive             = New-VssPipelineCapability -Name 'GoLive' -Status Placeholder -PlaceholderId 'VSS-GOLIVE-001'
    }
    if ([string]$Configuration.SharePoint.ApiKind -ceq 'MicrosoftGraphV1') {
        $capabilities['Authorization'] = New-VssPipelineCapability -Name 'Authorization' -Status Placeholder -PlaceholderId 'VSS-AUTH-GRAPH-001'
    }
    else {
        $capabilities['Authorization'] = New-VssPipelineCapability -Name 'Authorization' -Status Ready
        $capabilities['RestLargeUpload'] = New-VssPipelineCapability -Name 'RestLargeUpload' -Status Placeholder -PlaceholderId 'VSS-SPREST-LARGE-001'
        $capabilities['RestResourcePath'] = New-VssPipelineCapability -Name 'RestResourcePath' -Status Placeholder -PlaceholderId 'VSS-SPREST-PATH-001'
    }

    $acquireLock = {
        param($request)
        return Enter-VssPipelineLock -StatePath ([string]$request.Context.Configuration.Operations.StatePath) -RunId ([string]$request.Context.RunId) -TargetIdentity (Get-VssPipelineTargetIdentity -Configuration $request.Context.Configuration)
    }
    $releaseLock = { param($request) Exit-VssPipelineLock -Lock $request.Lock }
    $loadState = { param($request) Read-VssPipelineState -LiteralPath ([string]$request.Context.Configuration.Operations.StatePath) }
    $saveState = {
        param($request)
        Save-VssPipelineState -State $request.State -LiteralPath ([string]$request.Context.Configuration.Operations.StatePath)
        return $true
    }
    $inventory = { param($request) Invoke-VssPipelineInventory -Configuration $request.Context.Configuration -RuntimeConfiguration $request.Context.RuntimeConfiguration }
    $stageSource = {
        param($request)
        return Copy-VssPipelineStableFile -Candidate $request.Candidate -Configuration $request.Context.Configuration -RunStagingPath ([string]$request.Context.RunStagingPath)
    }
    $convert = {
        param($request)
        switch -CaseSensitive ([string]$request.Context.RuntimeConfiguration.Conversion.AdapterKind) {
            'VisioCom' {
                return Invoke-VssVisioComConversion -StagedSource $request.StagedSource -Configuration $request.Context.Configuration -RuntimeConfiguration $request.Context.RuntimeConfiguration
            }
            'ExternalConverter' {
                return Invoke-VssExternalConverterConversion -StagedSource $request.StagedSource -Configuration $request.Context.Configuration -RuntimeConfiguration $request.Context.RuntimeConfiguration
            }
            default {
                # >>> PLACEHOLDER [VSS-CNV-001] BEGIN
                throw [System.NotSupportedException]::new('[VSS-CNV-001] Kein produktiver Konvertierungsadapter ausgewaehlt.')
                # <<< PLACEHOLDER [VSS-CNV-001] END
            }
        }
    }
    $verifySource = { param($request) return Test-VssPipelineSourceUnchanged -StagedSource $request.StagedSource }
    $inspectTarget = { param($request) return Get-VssPipelineRemoteTargetMetadata -Request $request }
    $ensureFolders = { param($request) return Invoke-VssEnsureSharePointFolders -RequestPlan $request.RequestPlan }
    $upload = {
        param($request)
        $configuration = $request.Context.Configuration
        $runtimeConfiguration = $request.Context.RuntimeConfiguration
        $maxRetryCount = [int]$configuration.Operations.MaxRetryCount
        if ([string]$configuration.SharePoint.ApiKind -ceq 'MicrosoftGraphV1') {
            $headers = Get-VssPipelineAuthorizationHeaders -Configuration $configuration
            $response = Invoke-VssGraphTransport -RequestPlan $request.RequestPlan -ContentPath ([string]$request.Converted.PdfPath) -AuthorizationHeaders $headers -RuntimeConfiguration $runtimeConfiguration -MaxRetryCount $maxRetryCount -AllowExternalSideEffects
        }
        else {
            $useDefault = ([string]$configuration.SharePoint.AuthenticationKind -ceq 'WindowsIntegrated')
            if (-not $useDefault) {
                # >>> PLACEHOLDER [VSS-INTEGRATION-001] BEGIN
                throw [System.NotSupportedException]::new('[VSS-INTEGRATION-001] Fuer diese REST-Authentifizierungsroute fehlt der freigegebene Transportvertrag.')
                # <<< PLACEHOLDER [VSS-INTEGRATION-001] END
            }
            $response = Invoke-VssRestTransport -RequestPlan $request.RequestPlan -ContentPath ([string]$request.Converted.PdfPath) -AuthorizationHeaders $null -UseDefaultCredentials -MaxRetryCount $maxRetryCount -AllowExternalSideEffects
        }
        return ConvertFrom-VssPipelineRemoteCommitResponse -ApiKind ([string]$configuration.SharePoint.ApiKind) -Response $response -TargetRelativePath ([string]$request.TargetRelativePath) -SourceKey ([string]$request.SourceKey)
    }
    $reportOrphans = { param($request) return Get-VssPipelineOrphanStateItems -State $request.State -CandidatePlans $request.CandidatePlans }
    $writeLog = { param($request) return Write-VssPipelineLog -LogPath ([string]$request.Context.Configuration.Operations.LogPath) -Event $request.Event }
    $cleanup = {
        param($request)
        Remove-VssPipelineStagingRun -ConfiguredStagingRoot ([string]$request.Context.Configuration.Operations.StagingPath) -RunStagingPath ([string]$request.Context.RunStagingPath)
        return $true
    }
    $approve = { param($request) [void]$request; Assert-VssPipelineGoLiveApproval }

    return [pscustomobject][ordered]@{
        Name            = 'Production'
        IsSimulation    = $false
        Capabilities    = $capabilities
        AcquireLock     = $acquireLock
        ReleaseLock     = $releaseLock
        LoadState       = $loadState
        SaveState       = $saveState
        Inventory       = $inventory
        StageSource     = $stageSource
        Convert         = $convert
        VerifySource    = $verifySource
        InspectTarget   = $inspectTarget
        EnsureFolders   = $ensureFolders
        Upload          = $upload
        ReportOrphans   = $reportOrphans
        WriteLog        = $writeLog
        Cleanup         = $cleanup
        ApproveGoLive   = $approve
    }
}

function New-VssPipelineSimulationAdapterSet {
    [CmdletBinding()]
    param()

    $memory = [pscustomobject][ordered]@{
        State       = [pscustomobject][ordered]@{ SchemaVersion = '1.0'; Items = @() }
        Remote      = @{}
        Trace       = New-Object System.Collections.ArrayList
        LogEvents   = New-Object System.Collections.ArrayList
        CommitCount = 0
        SaveCount   = 0
    }
    $sourceHash = Get-VssPipelineSha256Text -Text 'SIMULATION-SOURCE-CONTENT-001'
    $pdfHash = Get-VssPipelineSha256Text -Text '%PDF-1.7 SIMULATION-PDF-CONTENT-001'
    $capabilities = [ordered]@{}
    foreach ($name in @(
        'Lock', 'State', 'Inventory', 'Staging', 'Conversion', 'SourceVerification',
        'RemoteInspection', 'FolderResolution', 'Authorization', 'Upload',
        'Reconciliation', 'Logging', 'Cleanup', 'GoLive', 'RestLargeUpload',
        'RestResourcePath'
    )) {
        $capabilities[$name] = New-VssPipelineCapability -Name $name -Status Ready
    }

    $acquireLock = {
        param($request)
        [void]$memory.Trace.Add('Lock.Acquire')
        return [pscustomobject]@{ Name = 'SIMULATION-MUTEX'; Acquired = $true }
    }.GetNewClosure()
    $releaseLock = {
        param($request)
        [void]$memory.Trace.Add('Lock.Release')
        return $true
    }.GetNewClosure()
    $loadState = {
        param($request)
        [void]$memory.Trace.Add('State.Load')
        return $memory.State
    }.GetNewClosure()
    $saveState = {
        param($request)
        $memory.State = $request.State
        $memory.SaveCount = [int]$memory.SaveCount + 1
        [void]$memory.Trace.Add(('State.Save:{0}' -f [string]$request.SourceKey))
        return $true
    }.GetNewClosure()
    $inventory = {
        param($request)
        [void]$memory.Trace.Add('Inventory')
        $candidate = [pscustomobject][ordered]@{
            SourcePath       = 'memory://simulation/Example.vsdx'
            RelativePath     = 'Simulation\Example.vsdx'
            Length           = 31
            LastWriteTimeUtc = '2030-01-02T03:04:05.0000000Z'
            SyntheticSha256  = $sourceHash
        }
        return [pscustomobject][ordered]@{ Complete = $true; Candidates = @($candidate); Errors = @() }
    }.GetNewClosure()
    $stageSource = {
        param($request)
        [void]$memory.Trace.Add(('Stage:{0}' -f [string]$request.Candidate.RelativePath))
        return [pscustomobject][ordered]@{
            Status             = 'STAGED'
            Reason             = $null
            SourcePath         = [string]$request.Candidate.SourcePath
            RelativePath       = [string]$request.Candidate.RelativePath
            StagedPath         = 'memory://staging/Example.vsdx'
            Sha256             = [string]$request.Candidate.SyntheticSha256
            SourceLength       = [long]$request.Candidate.Length
            SourceWriteTimeUtc = [string]$request.Candidate.LastWriteTimeUtc
        }
    }.GetNewClosure()
    $convert = {
        param($request)
        [void]$memory.Trace.Add(('Convert:{0}' -f [string]$request.StagedSource.RelativePath))
        return [pscustomobject][ordered]@{
            Status        = 'CONVERTED'
            PdfPath       = 'memory://pdf/Example.pdf'
            PdfSha256     = $pdfHash
            ContentLength = [long]1048576
            FakePdf       = '%PDF-1.7 SIMULATION-PDF-CONTENT-001'
        }
    }.GetNewClosure()
    $verifySource = {
        param($request)
        [void]$memory.Trace.Add(('Verify:{0}' -f [string]$request.StagedSource.RelativePath))
        return $true
    }.GetNewClosure()
    $inspectTarget = {
        param($request)
        [void]$memory.Trace.Add(('Remote.Inspect:{0}' -f [string]$request.TargetRelativePath))
        $key = ([string]$request.TargetRelativePath).ToLowerInvariant()
        if ($memory.Remote.ContainsKey($key)) { return $memory.Remote[$key] }
        return [pscustomobject][ordered]@{
            Exists             = $false
            RemoteItemId       = $null
            ETag               = $null
            TargetRelativePath = [string]$request.TargetRelativePath
            Provenance         = $null
        }
    }.GetNewClosure()
    $ensureFolders = {
        param($request)
        [void]$memory.Trace.Add(('Remote.EnsureFolders:{0}' -f [string]$request.TargetRelativePath))
        return $true
    }.GetNewClosure()
    $upload = {
        param($request)
        $memory.CommitCount = [int]$memory.CommitCount + 1
        [void]$memory.Trace.Add(('Remote.Commit:{0}' -f [string]$request.SourceKey))
        $remoteId = 'SIM-ITEM-{0:D4}' -f [int]$memory.CommitCount
        $etag = '"SIM-ETAG-{0:D4}"' -f [int]$memory.CommitCount
        $remote = [pscustomobject][ordered]@{
            Exists             = $true
            RemoteItemId       = $remoteId
            ETag               = $etag
            TargetRelativePath = [string]$request.TargetRelativePath
            Provenance         = [string]$request.SourceKey
        }
        $memory.Remote[([string]$request.TargetRelativePath).ToLowerInvariant()] = $remote
        return [pscustomobject][ordered]@{
            Committed    = $true
            ApiKind      = [string]$request.Context.Configuration.SharePoint.ApiKind
            RemoteItemId = $remoteId
            ETag         = $etag
            RequestId    = 'SIM-REQUEST-{0:D4}' -f [int]$memory.CommitCount
            CommittedUtc = '2030-01-02T03:04:06.0000000Z'
        }
    }.GetNewClosure()
    $reportOrphans = {
        param($request)
        [void]$memory.Trace.Add('State.ReportOrphans')
        return Get-VssPipelineOrphanStateItems -State $request.State -CandidatePlans $request.CandidatePlans
    }.GetNewClosure()
    $writeLog = {
        param($request)
        [void]$memory.Trace.Add(('Log:{0}' -f [string]$request.Event.StageId))
        [void]$memory.LogEvents.Add((ConvertTo-VssPipelineRedactedLogObject -InputObject $request.Event))
        return 'memory://log'
    }.GetNewClosure()
    $cleanup = {
        param($request)
        [void]$memory.Trace.Add('Staging.Cleanup')
        return $true
    }.GetNewClosure()
    $approve = {
        param($request)
        [void]$memory.Trace.Add('GoLive.FakeApproval')
        return $true
    }.GetNewClosure()

    return [pscustomobject][ordered]@{
        Name            = 'Simulation'
        IsSimulation    = $true
        Capabilities    = $capabilities
        Memory          = $memory
        AcquireLock     = $acquireLock
        ReleaseLock     = $releaseLock
        LoadState       = $loadState
        SaveState       = $saveState
        Inventory       = $inventory
        StageSource     = $stageSource
        Convert         = $convert
        VerifySource    = $verifySource
        InspectTarget   = $inspectTarget
        EnsureFolders   = $ensureFolders
        Upload          = $upload
        ReportOrphans   = $reportOrphans
        WriteLog        = $writeLog
        Cleanup         = $cleanup
        ApproveGoLive   = $approve
    }
}

function Test-VssPipelineExecutionReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration,
        [string[]]$RuntimePlaceholders = @(),
        [AllowNull()][object]$AdapterSet = $null
    )

    if ($null -eq $AdapterSet) {
        $AdapterSet = New-VssPipelineProductionAdapterSet -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration
    }
    $placeholders = @($RuntimePlaceholders)
    $reasons = @()
    $requiredCapabilities = @(
        'Lock', 'State', 'Inventory', 'Staging', 'Conversion',
        'SourceVerification', 'RemoteInspection', 'FolderResolution',
        'Authorization', 'Upload', 'Reconciliation', 'Logging', 'Cleanup', 'GoLive'
    )
    if ([string]$Configuration.SharePoint.ApiKind -ceq 'SharePointRest') {
        $requiredCapabilities += @('RestLargeUpload', 'RestResourcePath')
    }
    foreach ($capabilityName in $requiredCapabilities) {
        $capability = Get-VssPipelinePropertyValue -InputObject $AdapterSet.Capabilities -Name $capabilityName
        if (-not $capability.Found) {
            $reasons += "Dem code-definierten Adapterset fehlt die Faehigkeit $capabilityName."
            continue
        }
        if ([string]$capability.Value.Status -cne 'Ready') {
            if (-not [string]::IsNullOrWhiteSpace([string]$capability.Value.PlaceholderId)) {
                $placeholders += [string]$capability.Value.PlaceholderId
            }
            else { $reasons += "Die Adapterfaehigkeit $capabilityName ist nicht Ready." }
        }
    }
    foreach ($handlerName in @(
        'AcquireLock', 'ReleaseLock', 'LoadState', 'SaveState', 'Inventory',
        'StageSource', 'Convert', 'VerifySource', 'InspectTarget',
        'EnsureFolders', 'Upload', 'ReportOrphans', 'WriteLog', 'Cleanup',
        'ApproveGoLive'
    )) {
        $handler = Get-VssPipelinePropertyValue -InputObject $AdapterSet -Name $handlerName
        if (-not $handler.Found -or -not ($handler.Value -is [scriptblock])) {
            $reasons += "Dem code-definierten Adapterset fehlt der aufrufbare Handler $handlerName."
        }
    }
    if (-not [bool]$Configuration.Operations.SingleHostOnly) {
        $reasons += 'Operations.SingleHostOnly=false ist im ersten produktiven Stand nicht zulaessig.'
    }
    if ([string]$Configuration.Sync.Direction -cne 'PublishOnly' -or
        [string]$Configuration.Sync.DeletePolicy -cne 'Never') {
        $reasons += 'Der erste produktive Stand erlaubt ausschliesslich PublishOnly mit DeletePolicy=Never.'
    }
    if (-not [bool]$RuntimeConfiguration.Source.ExcludeHidden -or
        -not [bool]$RuntimeConfiguration.Source.ExcludeSystem -or
        -not [bool]$RuntimeConfiguration.Source.ExcludeReparsePoints) {
        $reasons += 'Execute verlangt im ersten Stand den Ausschluss von Hidden-, System- und ReparsePoint-Eintraegen.'
    }

    $uniquePlaceholders = @($placeholders | Select-Object -Unique)
    return [pscustomobject][ordered]@{
        IsReady      = [bool]($uniquePlaceholders.Count -eq 0 -and $reasons.Count -eq 0)
        Placeholders = @($placeholders | Select-Object -Unique)
        Reasons      = @($reasons)
        AdapterSet   = $AdapterSet
    }
}

function New-VssPipelineInitialStageResults {
    param([switch]$ForSimulation)
    $results = @()
    foreach ($stage in @(Get-VssPipelineStageCatalog)) {
        if ($ForSimulation) {
            $implementation = 'Simulation'
            $status = 'SIMULATION'
            $detail = 'Deterministischer In-Memory-Handler; keine externe Nebenwirkung.'
        }
        else {
            $implementation = [string]$stage.Implementation
            $status = if ($implementation -ceq 'Placeholder') { 'PLACEHOLDER' } else { 'READY' }
            $detail = if ($implementation -ceq 'Placeholder') {
                "Produktive Implementierung ist sichtbar blockiert ($($stage.PlaceholderId))."
            }
            else { 'Implementierung ist vorhanden; in diesem Modus nicht ausgefuehrt.' }
        }
        $results += New-VssPipelineStageResult -Id $stage.Id -Status $status -EffectKind $stage.EffectKind -Detail $detail -Implementation $implementation -PlaceholderId $stage.PlaceholderId -Handler $stage.Handler
    }
    return $results
}

function Set-VssPipelineStageResult {
    param(
        [Parameter(Mandatory = $true)][object[]]$Stages,
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$Detail
    )
    foreach ($stage in $Stages) {
        if ([string]$stage.Id -ceq $Id) {
            $stage.Status = $Status
            $stage.Detail = $Detail
            return
        }
    }
    throw "Unbekannte Pipeline-Stufe: $Id"
}

function Read-VssPipelinePrimaryConfigurationAssessment {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $configuration = $null
    $assessment = $null
    if (-not [System.IO.File]::Exists($LiteralPath)) {
        $assessment = New-VssAssessmentResult -Errors @('Die Konfigurationsdatei wurde nicht gefunden.') -UnresolvedDecisionIds $script:VssBlockerDecisionIds -PlannedStages @('ValidateConfiguration')
    }
    else {
        try {
            $text = Read-VssUtf8FileStrict -LiteralPath $LiteralPath
            if (Test-VssJsonHasDuplicateObjectKeys -JsonText $text) { throw 'DUPLICATE_KEYS' }
            $configuration = ConvertFrom-Json -InputObject $text -ErrorAction Stop
            $assessment = Test-VssConfigurationObject -Configuration $configuration -Mode Validate
        }
        catch {
            $message = if ($_.Exception.Message -ceq 'DUPLICATE_KEYS') {
                'Die Konfigurationsdatei enthaelt doppelte JSON-Schluessel.'
            }
            else { 'Die Konfigurationsdatei konnte nicht als gueltiges BOM-loses UTF-8-JSON gelesen werden.' }
            $assessment = New-VssAssessmentResult -Errors @($message) -UnresolvedDecisionIds $script:VssBlockerDecisionIds -PlannedStages @('ValidateConfiguration')
        }
    }
    return [pscustomobject][ordered]@{
        IsValid               = [bool]$assessment.IsValid
        Configuration        = $configuration
        Errors                = @($assessment.Errors)
        Warnings              = @($assessment.Warnings)
        UnresolvedDecisionIds = @($assessment.UnresolvedDecisionIds)
        PlannedStages         = @($assessment.PlannedStages)
    }
}

function Invoke-VssPipelineExecution {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][object]$RuntimeConfiguration,
        [Parameter(Mandatory = $true)][string]$RunId,
        [ValidateSet('Execute', 'Simulation')][string]$Mode = 'Execute',
        [AllowNull()][object]$AdapterSet = $null,
        [switch]$AllowExternalSideEffects
    )

    if ($null -eq $AdapterSet) {
        $AdapterSet = New-VssPipelineProductionAdapterSet -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration
    }
    $isSimulation = ($Mode -ceq 'Simulation')
    if (-not $isSimulation) {
        if (-not $AllowExternalSideEffects) { throw 'Die operative Pipeline wurde ohne Side-Effect-Freigabe aufgerufen.' }
        if (-not [bool]$RuntimeConfiguration.Execution.Enabled) { throw 'Die operative Pipeline wurde ohne Runtime.Execution.Enabled=true aufgerufen.' }
        $runtimeCheck = Test-VssPipelineRuntimeConfiguration -Configuration $RuntimeConfiguration
        $readiness = Test-VssPipelineExecutionReadiness -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration -RuntimePlaceholders $runtimeCheck.Placeholders -AdapterSet $AdapterSet
        if (-not $readiness.IsReady) { throw 'Die operative Pipeline wurde ohne vollstaendige Adapter-Readiness aufgerufen.' }
        [void](Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'ApproveGoLive' -Request ([pscustomobject]@{ Configuration = $Configuration }))
    }

    $stages = if ($isSimulation) { @(New-VssPipelineInitialStageResults -ForSimulation) } else { @(New-VssPipelineInitialStageResults) }
    if (-not $isSimulation) {
        # Ein erfolgreiches internes Readiness-Gate bedeutet, dass ein injiziertes
        # Code-Adapterset jeden benoetigten Handler wirklich als Ready ausweist.
        foreach ($stage in $stages) {
            $stage.Implementation = 'Ready'
            $stage.PlaceholderId = $null
            $stage.Status = 'READY'
        }
    }
    $completeStatus = if ($isSimulation) { 'SIMULATION' } else { 'COMPLETED' }
    $partialStatus = if ($isSimulation) { 'SIMULATION' } else { 'PARTIAL' }
    Set-VssPipelineStageResult -Stages $stages -Id 'ValidateConfiguration' -Status $completeStatus -Detail 'Die bereits validierte Primaerkonfiguration wird unveraendert verwendet.'
    Set-VssPipelineStageResult -Stages $stages -Id 'ValidateRuntimeConfiguration' -Status $completeStatus -Detail 'Die bereits validierte Runtime-Konfiguration wird unveraendert verwendet.'

    $runStagingPath = Get-VssPipelineSafeLocalChildPath -RootPath ([string]$Configuration.Operations.StagingPath) -RelativePath $RunId
    $context = [pscustomobject][ordered]@{
        RunId               = $RunId
        Mode                = $Mode
        Configuration       = $Configuration
        RuntimeConfiguration = $RuntimeConfiguration
        RunStagingPath      = $runStagingPath
    }
    $files = @()
    $errors = @()
    $warnings = @()
    $lock = $null
    $state = $null
    $inventory = $null
    $candidatePlans = @()
    $fatalOperationalFailure = $false
    $inventoryIncomplete = $false
    $stagingMayExist = $false
    $stagedCount = 0
    $convertedCount = 0
    $verifiedCount = 0
    $folderCount = 0
    $commitCount = 0
    $stateSaveCount = 0
    $fileFailureCount = 0
    $skippedCount = 0

    try {
        $lock = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'AcquireLock' -Request ([pscustomobject]@{ Context = $context })
        Set-VssPipelineStageResult -Stages $stages -Id 'AcquireSingleRunLock' -Status $completeStatus -Detail 'Der zielbezogene Einzellauf-Lock wurde erworben.'

        # Ein beschaedigter State wirft hier und blockiert dadurch jede spaetere
        # Remote-Pruefung oder jeden Remote-Schreibzugriff.
        $state = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'LoadState' -Request ([pscustomobject]@{ Context = $context })
        if (@(Test-VssPipelineStateObject -State $state).Count -gt 0) {
            throw 'Der geladene State entspricht nicht dem geschlossenen Schema.'
        }

        $inventory = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'Inventory' -Request ([pscustomobject]@{ Context = $context })
        $candidatePlans = @(Resolve-VssPipelineTargetCollisions -Candidates @($inventory.Candidates))
        if ([bool]$inventory.Complete) {
            Set-VssPipelineStageResult -Stages $stages -Id 'InventorySource' -Status $completeStatus -Detail ("Vollstaendige Inventur mit {0} Kandidat(en)." -f $candidatePlans.Count)
        }
        else {
            $inventoryIncomplete = $true
            Set-VssPipelineStageResult -Stages $stages -Id 'InventorySource' -Status $partialStatus -Detail ("Unvollstaendige Inventur; {0} erreichbare Kandidat(en) bleiben verarbeitbar." -f $candidatePlans.Count)
            $warnings += 'Die Quellinventur war unvollstaendig; verwaiste State-Eintraege werden in diesem Lauf nicht ausgewertet.'
            $errors += 'Mindestens ein freigegebener relativer Quellbereich konnte nicht vollstaendig inventarisiert werden.'
        }

        foreach ($plan in $candidatePlans) {
            $relativePath = [string]$plan.Candidate.RelativePath
            $targetRelativePath = [string]$plan.TargetRelativePath
            $sourceKey = [string]$plan.SourceKey
            $fileResult = $null
            $fileStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $remoteCommit = $null
            try {
                if ([bool]$plan.HasCollision) {
                    $fileFailureCount++
                    $errors += "Zielpfadkonflikt fuer freigegebenen relativen Quellpfad: $relativePath"
                    $fileResult = [pscustomobject][ordered]@{
                        SourceRelativePath   = $relativePath
                        TargetRelativePath   = $targetRelativePath
                        SourceSha256         = $null
                        ConversionFingerprint = $null
                        PdfSha256            = $null
                        Status               = 'CONFLICT_TARGET_PATH'
                        RemoteItemId         = $null
                        ETag                 = $null
                        RequestPlan          = $null
                    }
                    continue
                }

                $stagingMayExist = $true
                $staged = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'StageSource' -Request ([pscustomobject]@{ Context = $context; Candidate = $plan.Candidate })
                if ([string]$staged.Status -cne 'STAGED') {
                    $skippedCount++
                    $fileResult = [pscustomobject][ordered]@{
                        SourceRelativePath   = $relativePath
                        TargetRelativePath   = $targetRelativePath
                        SourceSha256         = $null
                        ConversionFingerprint = $null
                        PdfSha256            = $null
                        Status               = 'DEFERRED_SOURCE_UNSTABLE'
                        RemoteItemId         = $null
                        ETag                 = $null
                        RequestPlan          = $null
                    }
                    continue
                }
                $stagedCount++
                $sourceHash = [string]$staged.Sha256
                $fingerprint = Get-VssPipelineConversionFingerprint -SourceSha256 $sourceHash -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration
                $stateItem = Get-VssPipelineStateItem -State $state -SourceKey $sourceKey
                if ($null -ne $stateItem -and
                    [string]::Equals([string]$stateItem.SourceSha256, $sourceHash, [System.StringComparison]::OrdinalIgnoreCase) -and
                    [string]::Equals([string]$stateItem.ConversionFingerprint, $fingerprint, [System.StringComparison]::OrdinalIgnoreCase) -and
                    [string]::Equals([string]$stateItem.TargetRelativePath, $targetRelativePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $skippedCount++
                    $fileResult = [pscustomobject][ordered]@{
                        SourceRelativePath   = $relativePath
                        TargetRelativePath   = $targetRelativePath
                        SourceSha256         = $sourceHash
                        ConversionFingerprint = $fingerprint
                        PdfSha256            = [string]$stateItem.PdfSha256
                        Status               = 'SKIPPED_UNCHANGED'
                        RemoteItemId         = [string]$stateItem.RemoteItemId
                        ETag                 = [string]$stateItem.ETag
                        RequestPlan          = $null
                    }
                    continue
                }

                $converted = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'Convert' -Request ([pscustomobject]@{ Context = $context; StagedSource = $staged })
                if ([string]$converted.Status -cne 'CONVERTED' -or
                    [string]$converted.PdfSha256 -notmatch '^[0-9A-Fa-f]{64}$' -or
                    [long]$converted.ContentLength -le 0) {
                    throw 'Der Konvertierungsadapter lieferte keinen gueltigen PDF-Vertrag.'
                }
                $convertedCount++

                $remoteTarget = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'InspectTarget' -Request ([pscustomobject]@{
                    Context = $context; SourceKey = $sourceKey; TargetRelativePath = $targetRelativePath; StateItem = $stateItem
                })
                $directive = Resolve-VssPipelineRemoteWriteDecision -SourceKey $sourceKey -TargetRelativePath $targetRelativePath -StateItem $stateItem -RemoteTarget $remoteTarget
                if ([string]$directive.Action -ceq 'Conflict') {
                    $fileFailureCount++
                    $errors += "Remote-Konflikt fuer freigegebenen relativen Quellpfad: $relativePath"
                    $fileResult = [pscustomobject][ordered]@{
                        SourceRelativePath   = $relativePath
                        TargetRelativePath   = $targetRelativePath
                        SourceSha256         = $sourceHash
                        ConversionFingerprint = $fingerprint
                        PdfSha256            = [string]$converted.PdfSha256
                        Status               = 'CONFLICT_REMOTE'
                        RemoteItemId         = $null
                        ETag                 = $null
                        RequestPlan          = $null
                    }
                    continue
                }

                $requestPlan = New-VssPipelineRequestPlan -Configuration $Configuration -RuntimeConfiguration $RuntimeConfiguration -TargetRelativePath $targetRelativePath -ContentLength ([long]$converted.ContentLength) -CommitDirective $directive
                $unchanged = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'VerifySource' -Request ([pscustomobject]@{ Context = $context; StagedSource = $staged })
                if (-not [bool]$unchanged) {
                    $skippedCount++
                    $fileResult = [pscustomobject][ordered]@{
                        SourceRelativePath   = $relativePath
                        TargetRelativePath   = $targetRelativePath
                        SourceSha256         = $sourceHash
                        ConversionFingerprint = $fingerprint
                        PdfSha256            = [string]$converted.PdfSha256
                        Status               = 'DEFERRED_SOURCE_CHANGED'
                        RemoteItemId         = $null
                        ETag                 = $null
                        RequestPlan          = $requestPlan
                    }
                    continue
                }
                $verifiedCount++

                [void](Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'EnsureFolders' -Request ([pscustomobject]@{
                    Context = $context; RequestPlan = $requestPlan; TargetRelativePath = $targetRelativePath
                }))
                $folderCount++
                $remoteCommit = Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'Upload' -Request ([pscustomobject]@{
                    Context = $context; RequestPlan = $requestPlan; Converted = $converted; Directive = $directive
                    SourceKey = $sourceKey; TargetRelativePath = $targetRelativePath
                })
                if (-not [bool]$remoteCommit.Committed -or
                    [string]::IsNullOrWhiteSpace([string]$remoteCommit.RemoteItemId) -or
                    [string]::IsNullOrWhiteSpace([string]$remoteCommit.ETag)) {
                    throw 'Der Remote-Adapter hat den Commit nicht eindeutig bestaetigt.'
                }
                $commitApi = Get-VssPipelinePropertyValue -InputObject $remoteCommit -Name 'ApiKind'
                if ($commitApi.Found -and -not [string]::Equals([string]$commitApi.Value, [string]$Configuration.SharePoint.ApiKind, [System.StringComparison]::Ordinal)) {
                    throw 'Der Remote-Adapter bestaetigte einen unerwarteten API-Vertrag.'
                }
                $commitTarget = Get-VssPipelinePropertyValue -InputObject $remoteCommit -Name 'TargetRelativePath'
                if ($commitTarget.Found -and -not [string]::Equals([string]$commitTarget.Value, $targetRelativePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw 'Der Remote-Adapter bestaetigte einen unerwarteten Zielpfad.'
                }
                $commitCount++

                # Die einzige State-Mutation steht absichtlich unmittelbar nach
                # dem bestaetigten Commit. Vor diesem Punkt gibt es keinen Save.
                $newStateItem = [pscustomobject][ordered]@{
                    SourceKey                = $sourceKey
                    SourceRelativePath       = $relativePath
                    SourceSha256             = $sourceHash.ToLowerInvariant()
                    ConversionFingerprint   = $fingerprint.ToLowerInvariant()
                    PdfSha256                = ([string]$converted.PdfSha256).ToLowerInvariant()
                    TargetRelativePath       = $targetRelativePath
                    ApiKind                  = [string]$Configuration.SharePoint.ApiKind
                    RemoteItemId             = [string]$remoteCommit.RemoteItemId
                    ETag                     = [string]$remoteCommit.ETag
                    LastSuccessfulCommitUtc  = [string]$remoteCommit.CommittedUtc
                    Provenance               = $sourceKey
                }
                $state = Set-VssPipelineStateItem -State $state -Item $newStateItem
                [void](Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'SaveState' -Request ([pscustomobject]@{
                    Context = $context; State = $state; SourceKey = $sourceKey
                }))
                $stateSaveCount++
                $fileResult = [pscustomobject][ordered]@{
                    SourceRelativePath   = $relativePath
                    TargetRelativePath   = $targetRelativePath
                    SourceSha256         = $sourceHash
                    ConversionFingerprint = $fingerprint
                    PdfSha256            = [string]$converted.PdfSha256
                    Status               = if ([string]$directive.Action -ceq 'Update') { 'UPDATED' } else { 'UPLOADED' }
                    RemoteItemId         = [string]$remoteCommit.RemoteItemId
                    ETag                 = [string]$remoteCommit.ETag
                    RequestPlan          = $requestPlan
                }
            }
            catch {
                $fileFailureCount++
                $errors += "Dateiverarbeitung fehlgeschlagen fuer freigegebenen relativen Quellpfad: $relativePath"
                $fileResult = [pscustomobject][ordered]@{
                    SourceRelativePath   = $relativePath
                    TargetRelativePath   = $targetRelativePath
                    SourceSha256         = $null
                    ConversionFingerprint = $null
                    PdfSha256            = $null
                    Status               = if ($null -ne $remoteCommit -and [bool]$remoteCommit.Committed) { 'REMOTE_COMMITTED_STATE_FAILED' } else { 'FAILED' }
                    RemoteItemId         = if ($null -ne $remoteCommit) { [string]$remoteCommit.RemoteItemId } else { $null }
                    ETag                 = if ($null -ne $remoteCommit) { [string]$remoteCommit.ETag } else { $null }
                    RequestPlan          = $null
                }
            }
            finally {
                $fileStopwatch.Stop()
                if ($null -eq $fileResult) {
                    $fileResult = [pscustomobject][ordered]@{
                        SourceRelativePath = $relativePath; TargetRelativePath = $targetRelativePath
                        SourceSha256 = $null; ConversionFingerprint = $null; PdfSha256 = $null
                        Status = 'FAILED'; RemoteItemId = $null; ETag = $null; RequestPlan = $null
                    }
                }
                $files += $fileResult
                $logEvent = [pscustomobject][ordered]@{
                    RunId = $RunId; StageId = 'UploadOrUpdatePdf'; RelativePath = $relativePath
                    Result = [string]$fileResult.Status
                    DurationMs = if ($isSimulation) { 0 } else { [long]$fileStopwatch.ElapsedMilliseconds }
                    RetryCount = 0; HttpStatus = $null
                    RequestId = if ($null -ne $remoteCommit) { [string]$remoteCommit.RequestId } else { $null }
                }
                try {
                    [void](Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'WriteLog' -Request ([pscustomobject]@{ Context = $context; Event = $logEvent }))
                }
                catch { $warnings += "Das strukturierte Dateilog konnte fuer den relativen Pfad nicht geschrieben werden: $relativePath" }
            }
        }

        Set-VssPipelineStageResult -Stages $stages -Id 'StageStableSource' -Status $(if ($fileFailureCount -gt 0) { $partialStatus } else { $completeStatus }) -Detail ("{0} stabile Snapshot(s), {1} Zurueckstellung(en)." -f $stagedCount, $skippedCount)
        Set-VssPipelineStageResult -Stages $stages -Id 'ConvertVisioToPdf' -Status $(if ($fileFailureCount -gt 0) { $partialStatus } else { $completeStatus }) -Detail ("{0} PDF-Konvertierung(en)." -f $convertedCount)
        Set-VssPipelineStageResult -Stages $stages -Id 'VerifySourceUnchanged' -Status $completeStatus -Detail ("{0} unmittelbar vor Remote-Schreiben bestaetigte Quelle(n)." -f $verifiedCount)
        Set-VssPipelineStageResult -Stages $stages -Id 'EnsureSharePointFolders' -Status $(if ($fileFailureCount -gt 0) { $partialStatus } else { $completeStatus }) -Detail ("{0} Zielordnerauflosung(en)." -f $folderCount)
        Set-VssPipelineStageResult -Stages $stages -Id 'UploadOrUpdatePdf' -Status $(if ($fileFailureCount -gt 0) { $partialStatus } else { $completeStatus }) -Detail ("{0} bestaetigte Remote-Commit(s)." -f $commitCount)
        Set-VssPipelineStageResult -Stages $stages -Id 'PersistSyncState' -Status $(if ($commitCount -ne $stateSaveCount) { $partialStatus } else { $completeStatus }) -Detail ("{0} atomare State-Speicherung(en) nach bestaetigtem Commit." -f $stateSaveCount)

        if ([bool]$inventory.Complete) {
            $orphans = @(Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'ReportOrphans' -Request ([pscustomobject]@{ Context = $context; State = $state; CandidatePlans = $candidatePlans }))
            if ($orphans.Count -gt 0) { $warnings += ("{0} verwaiste State-Eintraege wurden nur gemeldet; es erfolgte keine Move-, Archive- oder Delete-Aktion." -f $orphans.Count) }
            Set-VssPipelineStageResult -Stages $stages -Id 'ReconcileLocalAndRemoteState' -Status $completeStatus -Detail ("{0} verwaiste State-Eintraege nur gemeldet; keine Remote-Aktion." -f $orphans.Count)
        }
        else {
            Set-VssPipelineStageResult -Stages $stages -Id 'ReconcileLocalAndRemoteState' -Status 'SKIPPED' -Detail 'Wegen unvollstaendiger Inventur unterdrueckt; keinerlei Orphan-Aktion.'
        }

        $summaryEvent = [pscustomobject][ordered]@{
            RunId = $RunId; StageId = 'SummarizeRun'; RelativePath = $null
            Result = if ($fileFailureCount -gt 0) { 'PARTIAL' } else { 'SUCCEEDED' }
            DurationMs = 0; RetryCount = 0; HttpStatus = $null; RequestId = $null
        }
        try {
            [void](Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'WriteLog' -Request ([pscustomobject]@{ Context = $context; Event = $summaryEvent }))
            Set-VssPipelineStageResult -Stages $stages -Id 'WriteStructuredLog' -Status $completeStatus -Detail 'Allowlist-basiertes Laufprotokoll geschrieben.'
        }
        catch {
            $warnings += 'Das strukturierte Laufprotokoll konnte nicht geschrieben werden.'
            Set-VssPipelineStageResult -Stages $stages -Id 'WriteStructuredLog' -Status $partialStatus -Detail 'Protokollierung fehlgeschlagen; keine geheimen Fehlerdetails ausgegeben.'
        }
    }
    catch {
        $fatalOperationalFailure = $true
        $errors += 'Der operative Ablauf wurde vor Remote-Schreiben oder zwischen kontrollierten Stufen abgebrochen.'
    }
    finally {
        if ($null -ne $lock) {
            try {
                [void](Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'Cleanup' -Request ([pscustomobject]@{ Context = $context; MayExist = $stagingMayExist }))
                Set-VssPipelineStageResult -Stages $stages -Id 'CleanupStaging' -Status $completeStatus -Detail 'Laufbezogenes Staging wurde sicher bereinigt.'
            }
            catch {
                $fatalOperationalFailure = $true
                $errors += 'Das laufbezogene Staging konnte nicht vollstaendig bereinigt werden.'
                Set-VssPipelineStageResult -Stages $stages -Id 'CleanupStaging' -Status $partialStatus -Detail 'Bereinigung meldete einen operativen Fehler.'
            }
            try {
                [void](Invoke-VssPipelineAdapterOperation -AdapterSet $AdapterSet -Name 'ReleaseLock' -Request ([pscustomobject]@{ Context = $context; Lock = $lock }))
                Set-VssPipelineStageResult -Stages $stages -Id 'ReleaseSingleRunLock' -Status $completeStatus -Detail 'Der Einzellauf-Lock wurde im finally freigegeben.'
            }
            catch {
                $fatalOperationalFailure = $true
                $errors += 'Der Einzellauf-Lock konnte nicht kontrolliert freigegeben werden.'
                Set-VssPipelineStageResult -Stages $stages -Id 'ReleaseSingleRunLock' -Status $partialStatus -Detail 'Lock-Freigabe meldete einen operativen Fehler.'
            }
        }
    }

    $isPartial = ($fatalOperationalFailure -or $inventoryIncomplete -or $fileFailureCount -gt 0 -or $commitCount -ne $stateSaveCount)
    Set-VssPipelineStageResult -Stages $stages -Id 'SummarizeRun' -Status $(if ($isPartial) { $partialStatus } else { $completeStatus }) -Detail ("{0} Dateiresultat(e), {1} Commit(s), {2} Dateifehler." -f $files.Count, $commitCount, $fileFailureCount)
    return [pscustomobject][ordered]@{
        Status   = if ($isPartial) { 'PARTIAL' } elseif ($isSimulation) { 'SIMULATED' } else { 'SUCCEEDED' }
        ExitCode = if ($isPartial) { 4 } else { 0 }
        Stages   = @($stages)
        Files    = @($files)
        Errors   = @($errors)
        Warnings = @($warnings)
        Trace    = if ($null -ne (Get-VssPipelinePropertyValue -InputObject $AdapterSet -Name 'Memory').Value) { @($AdapterSet.Memory.Trace) } else { @() }
    }
}

function Invoke-VisioSharePointSyncPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ConfigurationPath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RuntimeConfigurationPath,
        [ValidateSet('Validate', 'Simulate', 'Execute')][string]$Mode = 'Validate',
        [switch]$AllowExternalSideEffects
    )

    $runId = [guid]::NewGuid().ToString('D')
    $startedUtc = [DateTime]::UtcNow.ToString('o')
    $stages = @(New-VssPipelineInitialStageResults)
    $files = @()
    $errors = @()
    $warnings = @()
    $placeholders = @()

    try {
        # Die Datei wird genau einmal strikt gelesen; dasselbe Objekt durchlaeuft
        # den bestehenden Core-Validator und wird danach weiterverwendet. So gibt
        # es kein zweites, zeitlich abweichendes JSON-Read (TOCTOU).
        $configurationAssessment = Read-VssPipelinePrimaryConfigurationAssessment -LiteralPath $ConfigurationPath
        $errors += @($configurationAssessment.Errors)
        $warnings += @($configurationAssessment.Warnings)
        foreach ($decisionId in @($configurationAssessment.UnresolvedDecisionIds)) {
            $errors += "Offene Entscheidungs-ID: $decisionId"
        }
        if ($configurationAssessment.IsValid) {
            Set-VssPipelineStageResult -Stages $stages -Id 'ValidateConfiguration' -Status 'READY' -Detail 'Konfiguration ist vollstaendig und gueltig.'
        }
        else {
            Set-VssPipelineStageResult -Stages $stages -Id 'ValidateConfiguration' -Status 'FAILED' -Detail 'Konfiguration ist ungueltig oder unvollstaendig.'
        }

        $runtimeAssessment = Read-VssPipelineRuntimeConfiguration -LiteralPath $RuntimeConfigurationPath
        $errors += @($runtimeAssessment.Errors)
        $warnings += @($runtimeAssessment.Warnings)
        $placeholders += @($runtimeAssessment.Placeholders)
        if ($runtimeAssessment.IsValid) {
            Set-VssPipelineStageResult -Stages $stages -Id 'ValidateRuntimeConfiguration' -Status 'READY' -Detail 'Laufzeitkonfiguration entspricht dem geschlossenen Schema.'
        }
        else {
            Set-VssPipelineStageResult -Stages $stages -Id 'ValidateRuntimeConfiguration' -Status 'FAILED' -Detail 'Laufzeitkonfiguration ist ungueltig.'
        }

        if (-not $configurationAssessment.IsValid -or -not $runtimeAssessment.IsValid) {
            return New-VssPipelineResult -RunId $runId -Mode $Mode -Status 'INVALID' -ExitCode 2 -StartedUtc $startedUtc -FinishedUtc ([DateTime]::UtcNow.ToString('o')) -Stages $stages -Files $files -Errors $errors -Warnings $warnings -Placeholders $placeholders
        }

        $configuration = $configurationAssessment.Configuration
        $runtimeConfiguration = $runtimeAssessment.Configuration
        $productionAdapterSet = New-VssPipelineProductionAdapterSet -Configuration $configuration -RuntimeConfiguration $runtimeConfiguration
        $readiness = Test-VssPipelineExecutionReadiness -Configuration $configuration -RuntimeConfiguration $runtimeConfiguration -RuntimePlaceholders $placeholders -AdapterSet $productionAdapterSet
        $placeholders = @($readiness.Placeholders)

        if ($Mode -eq 'Validate') {
            if ($AllowExternalSideEffects) { $warnings += 'AllowExternalSideEffects wird im Modus Validate ignoriert.' }
            if (-not [bool]$runtimeConfiguration.Execution.Enabled) {
                $warnings += 'Runtime.Execution.Enabled ist false; Execute bleibt zusaetzlich administrativ gesperrt.'
            }
            return New-VssPipelineResult -RunId $runId -Mode $Mode -Status 'VALID' -ExitCode 0 -StartedUtc $startedUtc -FinishedUtc ([DateTime]::UtcNow.ToString('o')) -Stages $stages -Files $files -Errors $errors -Warnings $warnings -Placeholders $placeholders
        }

        if ($Mode -eq 'Simulate') {
            if ($AllowExternalSideEffects) { $warnings += 'AllowExternalSideEffects wird im Modus Simulate ignoriert; die Simulation bleibt nebenwirkungsfrei.' }
            $simulation = Invoke-VssPipelineSimulation -Configuration $configuration -RuntimeConfiguration $runtimeConfiguration -RunId $runId
            $errors += @($simulation.Errors)
            $warnings += @($simulation.Warnings)
            return New-VssPipelineResult -RunId $runId -Mode $Mode -Status ([string]$simulation.Status) -ExitCode ([int]$simulation.ExitCode) -StartedUtc $startedUtc -FinishedUtc ([DateTime]::UtcNow.ToString('o')) -Stages @($simulation.Stages) -Files @($simulation.Files) -Errors $errors -Warnings $warnings -Placeholders $placeholders
        }

        # Execute double gate. No source probe, path creation, lock, COM object or
        # HTTP request is allowed above or inside these readiness checks.
        $gateErrors = @()
        if (-not [bool]$runtimeConfiguration.Execution.Enabled) {
            $gateErrors += 'Execute ist blockiert: Runtime.Execution.Enabled ist false.'
        }
        if (-not $AllowExternalSideEffects) {
            $gateErrors += 'Execute ist blockiert: -AllowExternalSideEffects wurde nicht angegeben.'
        }
        if (-not $readiness.IsReady -or @($placeholders).Count -gt 0) {
            $gateErrors += 'Execute ist blockiert: mindestens ein stabiler Implementierungs- oder Freigabeplatzhalter ist offen.'
        }
        $gateErrors += @($readiness.Reasons | ForEach-Object { "Execute ist blockiert: $_" })
        if ($gateErrors.Count -gt 0) {
            $errors += $gateErrors
            foreach ($stage in $stages) {
                if (($stage.Id -cne 'ValidateConfiguration') -and ($stage.Id -cne 'ValidateRuntimeConfiguration')) {
                    $stage.Status = 'BLOCKED'
                    $stage.Detail = 'Vor jeder externen Nebenwirkung durch Execute-Gate oder Readiness blockiert.'
                }
            }
            return New-VssPipelineResult -RunId $runId -Mode $Mode -Status 'BLOCKED' -ExitCode 3 -StartedUtc $startedUtc -FinishedUtc ([DateTime]::UtcNow.ToString('o')) -Stages $stages -Files $files -Errors $errors -Warnings $warnings -Placeholders $placeholders
        }

        $execution = Invoke-VssPipelineExecution -Configuration $configuration -RuntimeConfiguration $runtimeConfiguration -RunId $runId -Mode Execute -AdapterSet $productionAdapterSet -AllowExternalSideEffects
        return New-VssPipelineResult -RunId $runId -Mode $Mode -Status ([string]$execution.Status) -ExitCode ([int]$execution.ExitCode) -StartedUtc $startedUtc -FinishedUtc ([DateTime]::UtcNow.ToString('o')) -Stages @($execution.Stages) -Files @($execution.Files) -Errors @($execution.Errors) -Warnings @($execution.Warnings) -Placeholders @()
    }
    catch {
        # Never expose exception text: configuration fragments, paths, request
        # headers or preauthenticated URLs could otherwise reach console output.
        return New-VssPipelineResult -RunId $runId -Mode $Mode -Status 'INTERNAL_ERROR' -ExitCode 1 -StartedUtc $startedUtc -FinishedUtc ([DateTime]::UtcNow.ToString('o')) -Stages $stages -Files $files -Errors @('Interner Fehler im Pipeline-Geruest.') -Warnings $warnings -Placeholders $placeholders
    }
}

function Format-VssPipelineResult {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Result)

    $lines = @(
        'Visio-SharePoint-Pipeline',
        "RunId: $($Result.RunId)",
        "Modus: $($Result.Mode)",
        "Status: $($Result.Status)",
        "Exitcode: $($Result.ExitCode)",
        "StartedUtc: $($Result.StartedUtc)",
        "FinishedUtc: $($Result.FinishedUtc)",
        'Stufen:'
    )
    foreach ($stage in @($Result.Stages)) {
        $implementationLabel = ([string]$stage.Implementation).ToUpperInvariant()
        $placeholderLabel = if ([string]::IsNullOrWhiteSpace([string]$stage.PlaceholderId)) { '' } else { " [$($stage.PlaceholderId)]" }
        $lines += "  - $($stage.Id) [$($stage.Status)] <$implementationLabel>$placeholderLabel ($($stage.EffectKind)): $($stage.Detail)"
    }
    $lines += 'Dateien:'
    if (@($Result.Files).Count -eq 0) { $lines += '  - keine' }
    foreach ($file in @($Result.Files)) {
        $source = Get-VssPipelinePropertyValue -InputObject $file -Name 'SourceRelativePath'
        $target = Get-VssPipelinePropertyValue -InputObject $file -Name 'TargetRelativePath'
        $status = Get-VssPipelinePropertyValue -InputObject $file -Name 'Status'
        $lines += "  - $($source.Value) -> $($target.Value) [$($status.Value)]"
        $request = Get-VssPipelinePropertyValue -InputObject $file -Name 'RequestPlan'
        if ($request.Found -and $null -ne $request.Value) {
            # Keine vollstaendige Site-, Drive-, Folder- oder Upload-URL auf der
            # Konsole; die abstrakte Route reicht fuer Validate/Simulate aus.
            $lines += "    RequestPlan: $($request.Value.Transport) $($request.Value.Method) ($($request.Value.Operation); $($request.Value.BodyKind); ExternalSideEffect=$($request.Value.ExternalSideEffect))"
        }
        $remoteId = Get-VssPipelinePropertyValue -InputObject $file -Name 'RemoteItemId'
        $etag = Get-VssPipelinePropertyValue -InputObject $file -Name 'ETag'
        if ($remoteId.Found -and -not [string]::IsNullOrWhiteSpace([string]$remoteId.Value)) {
            $lines += "    Remote: ItemId=$($remoteId.Value); eTag=$($etag.Value)"
        }
    }
    if (@($Result.Errors).Count -gt 0) {
        $lines += 'Fehler:'
        foreach ($item in @($Result.Errors)) { $lines += "  - $item" }
    }
    if (@($Result.Warnings).Count -gt 0) {
        $lines += 'Warnungen:'
        foreach ($item in @($Result.Warnings)) { $lines += "  - $item" }
    }
    $lines += 'Platzhalter:'
    if (@($Result.Placeholders).Count -eq 0) { $lines += '  - keine' }
    foreach ($item in @($Result.Placeholders)) { $lines += "  - $item" }
    return $lines
}

$corePath = Join-Path -Path $PSScriptRoot -ChildPath 'VisioSharePointSync.Core.ps1'
$bootstrapRunId = [guid]::NewGuid().ToString('D')
$bootstrapStartedUtc = [DateTime]::UtcNow.ToString('o')
try {
    . $corePath
    $pipelineResult = Invoke-VisioSharePointSyncPipeline -ConfigurationPath $ConfigurationPath -RuntimeConfigurationPath $RuntimeConfigurationPath -Mode $Mode -AllowExternalSideEffects:$AllowExternalSideEffects
    Format-VssPipelineResult -Result $pipelineResult | Write-Output
    exit ([int]$pipelineResult.ExitCode)
}
catch {
    $pipelineResult = New-VssPipelineResult -RunId $bootstrapRunId -Mode $Mode -Status 'INTERNAL_ERROR' -ExitCode 1 -StartedUtc $bootstrapStartedUtc -FinishedUtc ([DateTime]::UtcNow.ToString('o')) -Stages @() -Files @() -Errors @('Interner Fehler beim Laden des Pipeline-Geruests.') -Warnings @() -Placeholders @()
    Format-VssPipelineResult -Result $pipelineResult | Write-Output
    exit 1
}
