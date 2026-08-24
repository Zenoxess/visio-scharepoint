Set-StrictMode -Version 2.0

# Complete set of IDs marked "Blocker" in docs/OFFENE-FRAGEN.md. Exit code 0 is
# impossible until every ID is explicitly acknowledged in ResolvedDecisionIds.
$script:VssBlockerDecisionIds = @(
    'ENV-001', 'ENV-002', 'ENV-003',
    'SRC-001', 'SRC-002',
    'CNV-001', 'CNV-002',
    'SP-001', 'SP-002',
    'SEC-001', 'SEC-002',
    'SYN-001', 'SYN-002', 'SYN-003', 'SYN-005',
    'OPS-001',
    'ACC-001', 'ACC-003'
)

$script:VssTopLevelKeys = @(
    'SchemaVersion', 'ResolvedDecisionIds', 'Source', 'Conversion',
    'SharePoint', 'Sync', 'Operations'
)

$script:VssSectionKeys = @{
    Source = @('RootPath', 'FileNamePattern', 'Extensions', 'RegexTimeoutMilliseconds')
    Conversion = @('Provider', 'OutputFormat', 'PageRange', 'Intent', 'DisableMacros', 'RefreshExternalData')
    SharePoint = @(
        'Platform', 'ApiKind', 'AuthenticationKind', 'CertificateStoreLocation',
        'WorkloadIdentityFilePath',
        'TenantId', 'ClientId', 'CertificateThumbprint',
        'SiteId', 'DriveId', 'TargetFolderId',
        'SiteUrl', 'LibraryName', 'TargetFolderPath'
    )
    Sync = @('Direction', 'ConflictPolicy', 'RenamePolicy', 'DeletePolicy', 'FingerprintMode')
    Operations = @(
        'StatePath', 'StagingPath', 'LogPath', 'IntervalMinutes',
        'StabilityProbeSeconds', 'ConversionTimeoutSeconds', 'MaxRetryCount',
        'SingleHostOnly'
    )
}

# SharePoint target and credential-reference fields are conditionally required by
# platform/API/authentication rules later in validation.
$script:VssConfigurationFields = @(
    [pscustomobject]@{ Path = 'SchemaVersion';                       QuestionId = 'ENV-004'; Type = 'String';      Required = $true;  Allowed = @('1.0') },
    [pscustomobject]@{ Path = 'Source.RootPath';                     QuestionId = 'SRC-001'; Type = 'String';      Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Source.FileNamePattern';              QuestionId = 'SRC-002'; Type = 'String';      Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Source.Extensions';                   QuestionId = 'SRC-003'; Type = 'StringArray'; Required = $true;  Allowed = @('.vsd', '.vsdx', '.vsdm') },
    [pscustomobject]@{ Path = 'Source.RegexTimeoutMilliseconds';     QuestionId = 'SRC-002'; Type = 'Integer';     Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Conversion.Provider';                 QuestionId = 'CNV-001'; Type = 'String';      Required = $true;  Allowed = @('VisioCom', 'ExternalConverter') },
    [pscustomobject]@{ Path = 'Conversion.OutputFormat';             QuestionId = 'CNV-003'; Type = 'String';      Required = $true;  Allowed = @('PDF') },
    [pscustomobject]@{ Path = 'Conversion.PageRange';                QuestionId = 'CNV-003'; Type = 'String';      Required = $true;  Allowed = @('AllForegroundPages') },
    [pscustomobject]@{ Path = 'Conversion.Intent';                   QuestionId = 'CNV-003'; Type = 'String';      Required = $true;  Allowed = @('Print', 'Screen') },
    [pscustomobject]@{ Path = 'Conversion.DisableMacros';            QuestionId = 'CNV-004'; Type = 'Boolean';     Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Conversion.RefreshExternalData';      QuestionId = 'CNV-004'; Type = 'Boolean';     Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.Platform';                 QuestionId = 'SP-001';  Type = 'String';      Required = $true;  Allowed = @('SharePointOnline', 'SharePointServer') },
    [pscustomobject]@{ Path = 'SharePoint.ApiKind';                  QuestionId = 'SP-001';  Type = 'String';      Required = $true;  Allowed = @('MicrosoftGraphV1', 'SharePointRest') },
    [pscustomobject]@{ Path = 'SharePoint.AuthenticationKind';       QuestionId = 'SEC-001'; Type = 'String';      Required = $true;  Allowed = @('CertificateAppOnly', 'ManagedIdentity', 'WorkloadIdentity', 'WindowsIntegrated', 'Delegated') },
    [pscustomobject]@{ Path = 'SharePoint.CertificateStoreLocation'; QuestionId = 'SEC-003'; Type = 'String';      Required = $false; Allowed = @('LocalMachine', 'CurrentUser') },
    [pscustomobject]@{ Path = 'SharePoint.WorkloadIdentityFilePath';  QuestionId = 'SEC-001'; Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.TenantId';                 QuestionId = 'SP-002';  Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.ClientId';                 QuestionId = 'SEC-001'; Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.CertificateThumbprint';    QuestionId = 'SEC-003'; Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.SiteId';                   QuestionId = 'SP-002';  Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.DriveId';                  QuestionId = 'SP-002';  Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.TargetFolderId';           QuestionId = 'SP-002';  Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.SiteUrl';                  QuestionId = 'SP-002';  Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.LibraryName';              QuestionId = 'SP-002';  Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'SharePoint.TargetFolderPath';         QuestionId = 'SP-002';  Type = 'String';      Required = $false; Allowed = @() },
    [pscustomobject]@{ Path = 'Sync.Direction';                      QuestionId = 'SYN-001'; Type = 'String';      Required = $true;  Allowed = @('PublishOnly', 'Mirror') },
    [pscustomobject]@{ Path = 'Sync.ConflictPolicy';                 QuestionId = 'SYN-003'; Type = 'String';      Required = $true;  Allowed = @('SourceWins', 'Fail') },
    [pscustomobject]@{ Path = 'Sync.RenamePolicy';                   QuestionId = 'SYN-005'; Type = 'String';      Required = $true;  Allowed = @('CreateNew', 'MoveWhenStableId') },
    [pscustomobject]@{ Path = 'Sync.DeletePolicy';                   QuestionId = 'SYN-005'; Type = 'String';      Required = $true;  Allowed = @('Never', 'ArchiveAfterGrace', 'DeleteAfterGrace') },
    [pscustomobject]@{ Path = 'Sync.FingerprintMode';                QuestionId = 'SYN-002'; Type = 'String';      Required = $true;  Allowed = @('MetadataThenSha256', 'Sha256EveryRun') },
    [pscustomobject]@{ Path = 'Operations.StatePath';                QuestionId = 'OPS-003'; Type = 'String';      Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Operations.StagingPath';              QuestionId = 'OPS-005'; Type = 'String';      Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Operations.LogPath';                  QuestionId = 'OPS-004'; Type = 'String';      Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Operations.IntervalMinutes';          QuestionId = 'OPS-001'; Type = 'Integer';     Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Operations.StabilityProbeSeconds';    QuestionId = 'SRC-004'; Type = 'Integer';     Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Operations.ConversionTimeoutSeconds'; QuestionId = 'CNV-005'; Type = 'Integer';     Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Operations.MaxRetryCount';            QuestionId = 'OPS-002'; Type = 'Integer';     Required = $true;  Allowed = @() },
    [pscustomobject]@{ Path = 'Operations.SingleHostOnly';           QuestionId = 'OPS-002'; Type = 'Boolean';     Required = $true;  Allowed = @() }
)

function Get-VssBlockerDecisionIds {
    return @($script:VssBlockerDecisionIds)
}

function Test-VssObjectContainer {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [System.Collections.IDictionary]) { return $true }
    return ($Value -is [System.Management.Automation.PSCustomObject])
}

function Get-VssDirectPropertyNames {
    param([Parameter(Mandatory = $true)][object]$InputObject)
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) { Write-Output ([string]$key) }
        return
    }
    foreach ($property in $InputObject.PSObject.Properties) { Write-Output $property.Name }
}

function Get-VssPropertyValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $current = $InputObject
    foreach ($segment in $Path.Split('.')) {
        if ($null -eq $current) { return [pscustomobject]@{ Found = $false; Value = $null } }
        if ($current -is [System.Collections.IDictionary]) {
            $matchingKey = $null
            foreach ($key in $current.Keys) {
                if ([string]::Equals([string]$key, $segment, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $matchingKey = $key
                    break
                }
            }
            if ($null -eq $matchingKey) { return [pscustomobject]@{ Found = $false; Value = $null } }
            $current = $current[$matchingKey]
            continue
        }
        $matchingProperty = $null
        foreach ($property in $current.PSObject.Properties) {
            if ([string]::Equals($property.Name, $segment, [System.StringComparison]::OrdinalIgnoreCase)) {
                $matchingProperty = $property
                break
            }
        }
        if ($null -eq $matchingProperty) { return [pscustomobject]@{ Found = $false; Value = $null } }
        $current = $matchingProperty.Value
    }
    return [pscustomobject]@{ Found = $true; Value = $current }
}

function Test-VssUndecidedValue {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) {
        return ([string]::IsNullOrWhiteSpace($Value) -or $Value.Trim().Equals('Undecided', [System.StringComparison]::OrdinalIgnoreCase))
    }
    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [System.Collections.IDictionary])) {
        $items = @($Value)
        if ($items.Count -eq 0) { return $true }
        foreach ($item in $items) {
            if ($null -eq $item) { return $true }
            if (($item -is [string]) -and ([string]::IsNullOrWhiteSpace($item) -or $item.Trim().Equals('Undecided', [System.StringComparison]::OrdinalIgnoreCase))) { return $true }
        }
    }
    return $false
}

function Test-VssIntegerValue {
    param([AllowNull()][object]$Value)
    return (
        ($Value -is [byte]) -or ($Value -is [sbyte]) -or
        ($Value -is [int16]) -or ($Value -is [uint16]) -or
        ($Value -is [int32]) -or ($Value -is [uint32]) -or
        ($Value -is [int64]) -or ($Value -is [uint64])
    )
}

function Test-VssStringArrayValue {
    param([AllowNull()][object]$Value)
    if (($Value -is [string]) -or ($Value -is [System.Collections.IDictionary]) -or -not ($Value -is [System.Collections.IEnumerable])) { return $false }
    foreach ($item in @($Value)) { if (-not ($item -is [string])) { return $false } }
    return $true
}

function Test-VssStringArrayContainsOrdinal {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Values,
        [Parameter(Mandatory = $true)][string]$Expected
    )
    foreach ($value in $Values) {
        if ([string]::Equals($value, $Expected, [System.StringComparison]::Ordinal)) { return $true }
    }
    return $false
}

function Test-VssConfigurationFieldHasValue {
    param(
        [Parameter(Mandatory = $true)][object]$Configuration,
        [Parameter(Mandatory = $true)][string]$Path
    )
    $result = Get-VssPropertyValue -InputObject $Configuration -Path $Path
    return ($result.Found -and -not (Test-VssUndecidedValue -Value $result.Value))
}

function Test-VssSecretLikeKeyName {
    param([Parameter(Mandatory = $true)][string]$Name)
    $normalized = [System.Text.RegularExpressions.Regex]::Replace($Name, '[^A-Za-z0-9]', '').ToLowerInvariant()
    return ($normalized -match '(authorization|uploadurl|uploadsessionurl|preauthenticatedurl|clientassertion|pfxpassphrase|clientsecret|password|passwd|pwd|privatekey|accesstoken|refreshtoken|bearertoken|sasurl|apikey|connectionstring|credential|secret|token)')
}

function Test-VssContainsProhibitedKey {
    param([AllowNull()][object]$InputObject)
    if ($null -eq $InputObject) { return $false }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            if (Test-VssSecretLikeKeyName -Name ([string]$key)) { return $true }
            if (Test-VssContainsProhibitedKey -InputObject $InputObject[$key]) { return $true }
        }
        return $false
    }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in $InputObject.PSObject.Properties) {
            if (Test-VssSecretLikeKeyName -Name $property.Name) { return $true }
            if (Test-VssContainsProhibitedKey -InputObject $property.Value) { return $true }
        }
        return $false
    }
    if (($InputObject -is [System.Collections.IEnumerable]) -and -not ($InputObject -is [string])) {
        foreach ($item in $InputObject) { if (Test-VssContainsProhibitedKey -InputObject $item) { return $true } }
    }
    return $false
}

function Test-VssContainsUnknownConfigurationKey {
    param([Parameter(Mandatory = $true)][object]$Configuration)
    foreach ($name in @(Get-VssDirectPropertyNames -InputObject $Configuration)) {
        if (-not (Test-VssStringArrayContainsOrdinal -Values $script:VssTopLevelKeys -Expected $name)) { return $true }
    }
    foreach ($sectionName in @('Source', 'Conversion', 'SharePoint', 'Sync', 'Operations')) {
        $section = Get-VssPropertyValue -InputObject $Configuration -Path $sectionName
        if ($section.Found -and (Test-VssObjectContainer -Value $section.Value)) {
            foreach ($name in @(Get-VssDirectPropertyNames -InputObject $section.Value)) {
                if (-not (Test-VssStringArrayContainsOrdinal -Values @($script:VssSectionKeys[$sectionName]) -Expected $name)) { return $true }
            }
        }
    }
    return $false
}

function Read-VssUtf8FileStrict {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$LiteralPath)
    $bytes = [System.IO.File]::ReadAllBytes($LiteralPath)
    if (($bytes.Length -ge 3) -and ($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB) -and ($bytes[2] -eq 0xBF)) {
        throw 'UTF-8 BOM is not permitted.'
    }
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
    return $utf8.GetString($bytes)
}

function ConvertFrom-VssJsonStringToken {
    param([Parameter(Mandatory = $true)][string]$Token)
    if (($Token.Length -lt 2) -or ($Token[0] -ne '"') -or ($Token[$Token.Length - 1] -ne '"')) {
        throw 'Invalid JSON string token.'
    }
    $builder = New-Object System.Text.StringBuilder
    $endIndex = $Token.Length - 1
    for ($index = 1; $index -lt $endIndex; $index++) {
        $character = $Token[$index]
        if ($character -ne '\') {
            if ([int]$character -lt 0x20) { throw 'Invalid JSON control character.' }
            [void]$builder.Append($character)
            continue
        }
        $index++
        if ($index -ge $endIndex) { throw 'Invalid JSON escape.' }
        $escapeCharacter = $Token[$index]
        switch ($escapeCharacter) {
            '"' { [void]$builder.Append('"') }
            '\' { [void]$builder.Append('\') }
            '/' { [void]$builder.Append('/') }
            'b' { [void]$builder.Append([char]8) }
            'f' { [void]$builder.Append([char]12) }
            'n' { [void]$builder.Append([char]10) }
            'r' { [void]$builder.Append([char]13) }
            't' { [void]$builder.Append([char]9) }
            'u' {
                if (($index + 4) -ge $endIndex) { throw 'Invalid JSON unicode escape.' }
                $hexText = $Token.Substring($index + 1, 4)
                if ($hexText -cnotmatch '^[0-9A-Fa-f]{4}$') { throw 'Invalid JSON unicode escape.' }
                [void]$builder.Append([char][Convert]::ToInt32($hexText, 16))
                $index += 4
            }
            default { throw 'Invalid JSON escape.' }
        }
    }
    return $builder.ToString()
}

function Test-VssJsonHasDuplicateObjectKeys {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$JsonText)

    # ConvertFrom-Json silently applies last-wins semantics to exactly repeated
    # property names on Windows PowerShell 5.1. This lexer runs first and keeps a
    # case-insensitive key set per object scope. It never returns a key or value.
    $contexts = New-Object System.Collections.Stack
    $index = 0
    while ($index -lt $JsonText.Length) {
        $character = $JsonText[$index]
        if ([char]::IsWhiteSpace($character)) { $index++; continue }

        if ($character -eq '{') {
            $contexts.Push([pscustomobject]@{ Kind = 'Object'; Keys = @() })
            $index++
            continue
        }
        if ($character -eq '[') {
            $contexts.Push([pscustomobject]@{ Kind = 'Array'; Keys = @() })
            $index++
            continue
        }
        if (($character -eq '}') -or ($character -eq ']')) {
            if ($contexts.Count -eq 0) { throw 'Invalid JSON container nesting.' }
            $context = $contexts.Pop()
            if ((($character -eq '}') -and ($context.Kind -ne 'Object')) -or
                (($character -eq ']') -and ($context.Kind -ne 'Array'))) {
                throw 'Invalid JSON container nesting.'
            }
            $index++
            continue
        }
        if ($character -ne '"') { $index++; continue }

        $stringStart = $index
        $index++
        $escaped = $false
        $closed = $false
        while ($index -lt $JsonText.Length) {
            $stringCharacter = $JsonText[$index]
            if ($escaped) { $escaped = $false; $index++; continue }
            if ($stringCharacter -eq '\') { $escaped = $true; $index++; continue }
            if ($stringCharacter -eq '"') { $closed = $true; $index++; break }
            $index++
        }
        if (-not $closed) { throw 'Invalid JSON string.' }

        $lookAhead = $index
        while (($lookAhead -lt $JsonText.Length) -and [char]::IsWhiteSpace($JsonText[$lookAhead])) { $lookAhead++ }
        if (($lookAhead -ge $JsonText.Length) -or ($JsonText[$lookAhead] -ne ':')) { continue }
        if (($contexts.Count -eq 0) -or ($contexts.Peek().Kind -ne 'Object')) { throw 'Invalid JSON property position.' }

        $keyToken = $JsonText.Substring($stringStart, $index - $stringStart)
        $decodedKey = ConvertFrom-VssJsonStringToken -Token $keyToken
        $objectContext = $contexts.Peek()
        foreach ($existingKey in @($objectContext.Keys)) {
            if ([string]::Equals($existingKey, $decodedKey, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
        $objectContext.Keys += $decodedKey
    }
    if ($contexts.Count -ne 0) { throw 'Invalid JSON container nesting.' }
    return $false
}

function Test-VssReservedWindowsSegment {
    param([Parameter(Mandatory = $true)][string]$Segment)
    $baseName = $Segment.Split('.')[0].ToUpperInvariant()
    $baseName = $baseName.Replace([char]0x00B9, '1').Replace([char]0x00B2, '2').Replace([char]0x00B3, '3')
    return ($baseName -match '^(CON|PRN|AUX|NUL|CLOCK\$|CONIN\$|CONOUT\$|COM[1-9]|LPT[1-9])$')
}

function Test-VssSafeWindowsSegment {
    param(
        [Parameter(Mandatory = $true)][string]$Segment,
        [bool]$RejectWhitespace = $false,
        [char[]]$AdditionalInvalidChars = @()
    )
    if ([string]::IsNullOrWhiteSpace($Segment)) { return $false }
    if (($Segment -eq '.') -or ($Segment -eq '..')) { return $false }
    if ($RejectWhitespace -and ($Segment -match '\s')) { return $false }
    $invalidChars = [char[]]('<>:"|?*' + (-join $AdditionalInvalidChars))
    if ($Segment.IndexOfAny($invalidChars) -ge 0) { return $false }
    foreach ($character in $Segment.ToCharArray()) { if ([char]::IsControl($character)) { return $false } }
    if ($Segment.EndsWith('.') -or $Segment.EndsWith(' ')) { return $false }
    if (Test-VssReservedWindowsSegment -Segment $Segment) { return $false }
    return $true
}

function Test-VssUncRootPathSyntax {
    param([AllowNull()][object]$Value)
    if (-not ($Value -is [string])) { return $false }
    if (($Value.Length -eq 0) -or ($Value -ne $Value.Trim())) { return $false }
    if (-not $Value.StartsWith('\\')) { return $false }
    if ($Value.StartsWith('\\?\', [System.StringComparison]::OrdinalIgnoreCase) -or $Value.StartsWith('\\.\', [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    if ($Value.Contains('/')) { return $false }
    $trimmed = $Value.TrimEnd([char]92)
    if (($Value.Length - $trimmed.Length) -gt 1) { return $false }
    if ($trimmed.Length -le 2) { return $false }
    $segments = @($trimmed.Substring(2).Split([char[]]@([char]92)))
    if ($segments.Count -lt 2) { return $false }
    for ($index = 0; $index -lt $segments.Count; $index++) {
        if (-not (Test-VssSafeWindowsSegment -Segment $segments[$index] -RejectWhitespace ($index -eq 0))) { return $false }
    }
    return $true
}

function Test-VssAbsoluteLocalPathSyntax {
    param([AllowNull()][object]$Value)
    if (-not ($Value -is [string])) { return $false }
    if (($Value.Length -eq 0) -or ($Value -ne $Value.Trim())) { return $false }
    if ($Value.StartsWith('\\') -or $Value.Contains('/')) { return $false }
    if ($Value -notmatch '^[A-Za-z]:\\') { return $false }
    $trimmed = $Value.TrimEnd([char]92)
    if (($Value.Length - $trimmed.Length) -gt 1) { return $false }
    if ($trimmed.Length -le 3) { return $false }
    $segments = @($trimmed.Substring(3).Split([char[]]@([char]92)))
    foreach ($segment in $segments) {
        if (-not (Test-VssSafeWindowsSegment -Segment $segment -AdditionalInvalidChars ([char[]]'[]'))) { return $false }
    }
    return $true
}

function Test-VssLocalPathOverlap {
    param([Parameter(Mandatory = $true)][string[]]$Paths)
    $normalized = @()
    foreach ($path in $Paths) { $normalized += $path.TrimEnd([char]92).ToUpperInvariant() }
    for ($left = 0; $left -lt $normalized.Count; $left++) {
        for ($right = $left + 1; $right -lt $normalized.Count; $right++) {
            if ($normalized[$left] -eq $normalized[$right]) { return $true }
            if ($normalized[$left].StartsWith($normalized[$right] + '\')) { return $true }
            if ($normalized[$right].StartsWith($normalized[$left] + '\')) { return $true }
        }
    }
    return $false
}

function Test-VssRelativeSharePointPathSyntax {
    param([AllowNull()][object]$Value)
    if (-not ($Value -is [string])) { return $false }
    if (($Value.Length -eq 0) -or ($Value -ne $Value.Trim())) { return $false }
    if ($Value.StartsWith('/') -or $Value.StartsWith('\') -or $Value.Contains('\')) { return $false }
    if ($Value -match '^[A-Za-z][A-Za-z0-9+.-]*:') { return $false }
    $segments = @($Value.Split('/'))
    foreach ($segment in $segments) {
        if (-not (Test-VssSafeWindowsSegment -Segment $segment -AdditionalInvalidChars ([char[]]'[]'))) { return $false }
    }
    return $true
}

function Test-VssSharePointLibraryNameSyntax {
    param([AllowNull()][object]$Value)
    if (-not ($Value -is [string])) { return $false }
    if (($Value.Length -eq 0) -or ($Value -ne $Value.Trim())) { return $false }
    if ($Value.IndexOfAny([char[]]'\/[]') -ge 0) { return $false }
    return (Test-VssSafeWindowsSegment -Segment $Value)
}

function Test-VssAbsoluteHttpUrlSyntax {
    param([AllowNull()][object]$Value)
    if (-not ($Value -is [string])) { return $false }
    if (($Value.Length -eq 0) -or ($Value -ne $Value.Trim())) { return $false }
    if ($Value.Contains('\')) { return $false }
    if ($Value -match '(?i)%2f|%5c') { return $false }
    if ($Value -match '(?i)(?:^|/)(?:(?:\.|%2e){1,2})(?:/|$)') { return $false }
    $parsed = $null
    if (-not [uri]::TryCreate($Value, [System.UriKind]::Absolute, [ref]$parsed)) { return $false }
    if ($parsed.Scheme -ne 'https') { return $false }
    if ([string]::IsNullOrWhiteSpace($parsed.Host) -or -not [string]::IsNullOrEmpty($parsed.UserInfo)) { return $false }
    if (-not [string]::IsNullOrEmpty($parsed.Query) -or -not [string]::IsNullOrEmpty($parsed.Fragment)) { return $false }
    return $true
}

function Test-VssGraphSiteIdSyntax {
    param([AllowNull()][object]$Value)
    if (-not ($Value -is [string])) { return $false }
    if (($Value.Length -eq 0) -or ($Value -ne $Value.Trim()) -or ($Value -match '\s')) { return $false }
    $parts = @($Value.Split(','))
    if ($parts.Count -ne 3) { return $false }
    if (($parts[0].Length -gt 253) -or ($parts[0].IndexOf('.') -lt 1) -or
        ([System.Uri]::CheckHostName($parts[0]) -ne [System.UriHostNameType]::Dns)) { return $false }
    foreach ($guidText in @($parts[1], $parts[2])) {
        $parsedGuid = [guid]::Empty
        if (-not [guid]::TryParseExact($guidText, 'D', [ref]$parsedGuid) -or ($parsedGuid -eq [guid]::Empty)) { return $false }
    }
    return $true
}

function Test-VssGraphOpaqueIdSyntax {
    param([AllowNull()][object]$Value)
    if (-not ($Value -is [string])) { return $false }
    if (($Value.Length -lt 1) -or ($Value.Length -gt 512) -or ($Value -ne $Value.Trim())) { return $false }
    if (($Value -eq '.') -or ($Value -eq '..')) { return $false }
    return ($Value -cmatch '^[A-Za-z0-9][A-Za-z0-9._!~-]{0,511}$')
}

function New-VssSourceRegex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Pattern,
        [Parameter(Mandatory = $true)][ValidateRange(50, 5000)][int]$TimeoutMilliseconds
    )
    $options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    return [System.Text.RegularExpressions.Regex]::new($Pattern, $options, [TimeSpan]::FromMilliseconds($TimeoutMilliseconds))
}

function Get-VssPlannedStages {
    param([ValidateSet('Validate', 'DryRun')][string]$Mode)
    if ($Mode -eq 'Validate') { return @('ValidateConfiguration') }
    return @(
        'ValidateConfiguration', 'AcquireSingleRunLock', 'InventorySource',
        'StageStableSource', 'ConvertVisioToPdf', 'EnsureSharePointFolders',
        'UploadOrUpdatePdf', 'ReconcileLocalAndRemoteState', 'PersistSyncState',
        'SummarizeRun'
    )
}

function New-VssAssessmentResult {
    param(
        [string[]]$Errors = @(),
        [string[]]$Warnings = @(),
        [string[]]$UnresolvedDecisionIds = @(),
        [string[]]$PlannedStages = @()
    )
    $errorItems = @($Errors)
    $warningItems = @($Warnings)
    $unresolvedItems = @($UnresolvedDecisionIds)
    return [pscustomobject]@{
        IsValid               = [bool](($errorItems.Count -eq 0) -and ($unresolvedItems.Count -eq 0))
        Errors                = $errorItems
        Warnings              = $warningItems
        UnresolvedDecisionIds = $unresolvedItems
        PlannedStages         = @($PlannedStages)
    }
}

function Test-VssConfigurationObject {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Configuration,
        [ValidateSet('Validate', 'DryRun')][string]$Mode = 'Validate'
    )

    $errors = @()
    $warnings = @()
    $unresolved = @()
    $plannedStages = @(Get-VssPlannedStages -Mode $Mode)
    $resolvedDecisionIds = @()

    if (Test-VssObjectContainer -Value $Configuration) {
        $resolvedResult = Get-VssPropertyValue -InputObject $Configuration -Path 'ResolvedDecisionIds'
        if (-not $resolvedResult.Found -or -not (Test-VssStringArrayValue -Value $resolvedResult.Value)) {
            $errors += 'ResolvedDecisionIds muss als String-Array angegeben werden.'
        }
        else {
            $resolvedDecisionIds = @($resolvedResult.Value)
            $validResolvedList = $true
            foreach ($decisionId in $resolvedDecisionIds) {
                if ([string]::IsNullOrWhiteSpace($decisionId) -or
                    -not (Test-VssStringArrayContainsOrdinal -Values $script:VssBlockerDecisionIds -Expected $decisionId)) {
                    $validResolvedList = $false
                    break
                }
            }
            for ($left = 0; $left -lt $resolvedDecisionIds.Count; $left++) {
                for ($right = $left + 1; $right -lt $resolvedDecisionIds.Count; $right++) {
                    if ([string]::Equals($resolvedDecisionIds[$left], $resolvedDecisionIds[$right], [System.StringComparison]::Ordinal)) {
                        $validResolvedList = $false
                    }
                }
            }
            if (-not $validResolvedList) { $errors += 'ResolvedDecisionIds enthaelt ungueltige oder doppelte Eintraege.' }
        }
    }

    foreach ($blockerId in $script:VssBlockerDecisionIds) {
        if (-not (Test-VssStringArrayContainsOrdinal -Values ([string[]]$resolvedDecisionIds) -Expected $blockerId)) { $unresolved += $blockerId }
    }

    if (-not (Test-VssObjectContainer -Value $Configuration)) {
        $errors += 'Die Wurzel der Konfiguration muss ein JSON-Objekt sein.'
        return New-VssAssessmentResult -Errors $errors -Warnings $warnings -UnresolvedDecisionIds $unresolved -PlannedStages $plannedStages
    }

    if (Test-VssContainsProhibitedKey -InputObject $Configuration) {
        $errors += 'Die Konfiguration enthaelt mindestens einen verbotenen geheimnisartigen Schluessel.'
    }
    if (Test-VssContainsUnknownConfigurationKey -Configuration $Configuration) {
        $errors += 'Die Konfiguration enthaelt mindestens einen unbekannten Schluessel.'
    }

    foreach ($sectionName in @('Source', 'Conversion', 'SharePoint', 'Sync', 'Operations')) {
        $section = Get-VssPropertyValue -InputObject $Configuration -Path $sectionName
        if ($section.Found -and -not (Test-VssUndecidedValue -Value $section.Value) -and -not (Test-VssObjectContainer -Value $section.Value)) {
            $errors += "$sectionName muss ein JSON-Objekt sein."
        }
    }

    foreach ($field in $script:VssConfigurationFields) {
        $resolved = Get-VssPropertyValue -InputObject $Configuration -Path $field.Path
        if (-not $resolved.Found -or (Test-VssUndecidedValue -Value $resolved.Value)) {
            if ($field.Required -and ($unresolved -notcontains $field.QuestionId)) { $unresolved += $field.QuestionId }
            continue
        }
        $value = $resolved.Value
        $typeIsValid = switch ($field.Type) {
            'String'      { $value -is [string] }
            'Boolean'     { $value -is [bool] }
            'Integer'     { Test-VssIntegerValue -Value $value }
            'StringArray' { Test-VssStringArrayValue -Value $value }
            default       { $false }
        }
        if (-not $typeIsValid) { $errors += "[$($field.QuestionId)] $($field.Path) hat einen ungueltigen Datentyp."; continue }
        if (@($field.Allowed).Count -gt 0) {
            if ($field.Type -eq 'StringArray') {
                foreach ($item in @($value)) {
                    if (-not (Test-VssStringArrayContainsOrdinal -Values ([string[]]@($field.Allowed)) -Expected $item)) {
                        $errors += "[$($field.QuestionId)] $($field.Path) enthaelt einen nicht unterstuetzten Wert."
                        break
                    }
                }
            }
            elseif (-not (Test-VssStringArrayContainsOrdinal -Values ([string[]]@($field.Allowed)) -Expected $value)) {
                $errors += "[$($field.QuestionId)] $($field.Path) enthaelt einen nicht unterstuetzten Wert."
            }
        }
    }

    $platform = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.Platform'
    $apiKind = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.ApiKind'
    $authenticationKind = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.AuthenticationKind'
    if ($platform.Found -and $apiKind.Found -and ($platform.Value -ceq 'SharePointServer') -and ($apiKind.Value -ceq 'MicrosoftGraphV1')) {
        $errors += '[SP-001] SharePointServer kann nicht mit MicrosoftGraphV1 kombiniert werden.'
    }
    if ($platform.Found -and $authenticationKind.Found -and ($platform.Value -ceq 'SharePointServer') -and
        -not (Test-VssUndecidedValue -Value $authenticationKind.Value) -and ($authenticationKind.Value -cne 'WindowsIntegrated')) {
        $errors += '[SP-001, SEC-001] SharePointServer unterstuetzt in diesem Geruest nur WindowsIntegrated.'
    }
    if ($platform.Found -and $authenticationKind.Found -and ($platform.Value -ceq 'SharePointOnline') -and
        ($authenticationKind.Value -ceq 'WindowsIntegrated')) {
        $errors += '[SP-001, SEC-001] SharePointOnline kann nicht mit WindowsIntegrated kombiniert werden.'
    }
    if ($authenticationKind.Found -and ($authenticationKind.Value -ceq 'Delegated')) {
        $errors += '[SEC-001] Delegated ist fuer dieses unbeaufsichtigte Geruest nicht zulaessig.'
    }

    $conditionalRequirements = @()
    if ($platform.Found -and $apiKind.Found -and ($platform.Value -ceq 'SharePointOnline') -and ($apiKind.Value -ceq 'MicrosoftGraphV1')) {
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.SiteId'; QuestionId = 'SP-002' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.DriveId'; QuestionId = 'SP-002' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.TargetFolderId'; QuestionId = 'SP-002' }
    }
    if ($apiKind.Found -and ($apiKind.Value -ceq 'MicrosoftGraphV1')) {
        foreach ($restTargetPath in @('SharePoint.SiteUrl', 'SharePoint.LibraryName', 'SharePoint.TargetFolderPath')) {
            if (Test-VssConfigurationFieldHasValue -Configuration $Configuration -Path $restTargetPath) {
                $errors += '[SP-002] MicrosoftGraphV1 darf keine REST-Zielpfadfelder enthalten.'
                break
            }
        }
    }
    if ($apiKind.Found -and ($apiKind.Value -ceq 'SharePointRest')) {
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.SiteUrl'; QuestionId = 'SP-002' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.LibraryName'; QuestionId = 'SP-002' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.TargetFolderPath'; QuestionId = 'SP-002' }
        foreach ($graphTargetPath in @('SharePoint.SiteId', 'SharePoint.DriveId', 'SharePoint.TargetFolderId')) {
            if (Test-VssConfigurationFieldHasValue -Configuration $Configuration -Path $graphTargetPath) {
                $errors += '[SP-002] SharePointRest darf keine Graph-Ziel-ID-Felder enthalten.'
                break
            }
        }
    }
    if ($authenticationKind.Found -and ($authenticationKind.Value -ceq 'CertificateAppOnly')) {
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.TenantId'; QuestionId = 'SP-002' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.ClientId'; QuestionId = 'SEC-001' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.CertificateThumbprint'; QuestionId = 'SEC-003' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.CertificateStoreLocation'; QuestionId = 'SEC-003' }
    }
    if ($authenticationKind.Found -and (($authenticationKind.Value -ceq 'WorkloadIdentity') -or ($authenticationKind.Value -ceq 'Delegated'))) {
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.TenantId'; QuestionId = 'SP-002' }
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.ClientId'; QuestionId = 'SEC-001' }
    }
    if ($authenticationKind.Found -and ($authenticationKind.Value -ceq 'WorkloadIdentity')) {
        $conditionalRequirements += [pscustomobject]@{ Path = 'SharePoint.WorkloadIdentityFilePath'; QuestionId = 'SEC-001' }
    }
    foreach ($requirement in $conditionalRequirements) {
        $requiredResult = Get-VssPropertyValue -InputObject $Configuration -Path $requirement.Path
        if (-not $requiredResult.Found -or (Test-VssUndecidedValue -Value $requiredResult.Value)) {
            if ($unresolved -notcontains $requirement.QuestionId) { $unresolved += $requirement.QuestionId }
        }
    }

    $forbiddenAuthenticationFields = @()
    if ($authenticationKind.Found) {
        switch -CaseSensitive ([string]$authenticationKind.Value) {
            'CertificateAppOnly' { $forbiddenAuthenticationFields = @('SharePoint.WorkloadIdentityFilePath') }
            'ManagedIdentity' {
                $forbiddenAuthenticationFields = @(
                    'SharePoint.TenantId', 'SharePoint.CertificateThumbprint',
                    'SharePoint.CertificateStoreLocation', 'SharePoint.WorkloadIdentityFilePath'
                )
            }
            'WorkloadIdentity' { $forbiddenAuthenticationFields = @('SharePoint.CertificateThumbprint', 'SharePoint.CertificateStoreLocation') }
            'WindowsIntegrated' {
                $forbiddenAuthenticationFields = @(
                    'SharePoint.TenantId', 'SharePoint.ClientId', 'SharePoint.CertificateThumbprint',
                    'SharePoint.CertificateStoreLocation', 'SharePoint.WorkloadIdentityFilePath'
                )
            }
            'Delegated' { $forbiddenAuthenticationFields = @('SharePoint.CertificateThumbprint', 'SharePoint.CertificateStoreLocation', 'SharePoint.WorkloadIdentityFilePath') }
        }
    }
    foreach ($forbiddenAuthenticationField in $forbiddenAuthenticationFields) {
        if (Test-VssConfigurationFieldHasValue -Configuration $Configuration -Path $forbiddenAuthenticationField) {
            $errors += '[SEC-001] Die gewaehlte Authentifizierungsart enthaelt nicht zulaessige Zusatzfelder.'
            break
        }
    }

    foreach ($guidPath in @('SharePoint.TenantId', 'SharePoint.ClientId')) {
        $guidResult = Get-VssPropertyValue -InputObject $Configuration -Path $guidPath
        if ($guidResult.Found -and ($guidResult.Value -is [string]) -and -not (Test-VssUndecidedValue -Value $guidResult.Value)) {
            $parsedGuid = [guid]::Empty
            if (-not [guid]::TryParse($guidResult.Value, [ref]$parsedGuid)) {
                $guidQuestionId = if ($guidPath -eq 'SharePoint.ClientId') { 'SEC-001' } else { 'SP-002' }
                $errors += "[$guidQuestionId] $guidPath muss eine GUID sein."
            }
        }
    }

    $thumbprint = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.CertificateThumbprint'
    if ($thumbprint.Found -and ($thumbprint.Value -is [string]) -and -not (Test-VssUndecidedValue -Value $thumbprint.Value) -and ($thumbprint.Value -notmatch '^[A-Fa-f0-9]{40}$')) {
        $errors += '[SEC-003] SharePoint.CertificateThumbprint muss aus genau 40 Hex-Zeichen bestehen.'
    }

    $siteId = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.SiteId'
    if ($siteId.Found -and -not (Test-VssUndecidedValue -Value $siteId.Value) -and -not (Test-VssGraphSiteIdSyntax -Value $siteId.Value)) {
        $errors += '[SP-002] SharePoint.SiteId hat keine sichere Graph-Site-ID-Syntax.'
    }
    foreach ($opaqueIdPath in @('SharePoint.DriveId', 'SharePoint.TargetFolderId')) {
        $opaqueId = Get-VssPropertyValue -InputObject $Configuration -Path $opaqueIdPath
        if ($opaqueId.Found -and -not (Test-VssUndecidedValue -Value $opaqueId.Value) -and -not (Test-VssGraphOpaqueIdSyntax -Value $opaqueId.Value)) {
            $errors += "[SP-002] $opaqueIdPath hat keine sichere opaque Graph-ID-Syntax."
        }
    }

    $siteUrl = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.SiteUrl'
    if ($siteUrl.Found -and -not (Test-VssUndecidedValue -Value $siteUrl.Value) -and -not (Test-VssAbsoluteHttpUrlSyntax -Value $siteUrl.Value)) {
        $errors += '[SP-002] SharePoint.SiteUrl muss eine absolute HTTPS-URL ohne Benutzerinformation, Query oder Fragment sein.'
    }
    $libraryName = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.LibraryName'
    if ($libraryName.Found -and -not (Test-VssUndecidedValue -Value $libraryName.Value) -and -not (Test-VssSharePointLibraryNameSyntax -Value $libraryName.Value)) {
        $errors += '[SP-002] SharePoint.LibraryName hat eine ungueltige Pfadsyntax.'
    }
    $targetFolderPath = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.TargetFolderPath'
    if ($targetFolderPath.Found -and -not (Test-VssUndecidedValue -Value $targetFolderPath.Value) -and -not (Test-VssRelativeSharePointPathSyntax -Value $targetFolderPath.Value)) {
        $errors += '[SP-002] SharePoint.TargetFolderPath muss ein sicherer relativer Pfad sein.'
    }
    $workloadIdentityFilePath = Get-VssPropertyValue -InputObject $Configuration -Path 'SharePoint.WorkloadIdentityFilePath'
    if ($workloadIdentityFilePath.Found -and -not (Test-VssUndecidedValue -Value $workloadIdentityFilePath.Value) -and
        -not (Test-VssAbsoluteLocalPathSyntax -Value $workloadIdentityFilePath.Value)) {
        $errors += '[SEC-001] SharePoint.WorkloadIdentityFilePath muss ein sicherer absoluter lokaler Nicht-Root-Pfad sein.'
    }

    $rootPath = Get-VssPropertyValue -InputObject $Configuration -Path 'Source.RootPath'
    if ($rootPath.Found -and -not (Test-VssUndecidedValue -Value $rootPath.Value) -and -not (Test-VssUncRootPathSyntax -Value $rootPath.Value)) {
        $errors += '[SRC-001] Source.RootPath muss ein syntaktisch sicherer UNC-Pfad mit Server und Freigabe sein.'
    }

    $integerBounds = @(
        [pscustomobject]@{ Path = 'Source.RegexTimeoutMilliseconds'; QuestionId = 'SRC-002'; Minimum = 50; Maximum = 5000 },
        [pscustomobject]@{ Path = 'Operations.IntervalMinutes'; QuestionId = 'OPS-001'; Minimum = 1; Maximum = 1440 },
        [pscustomobject]@{ Path = 'Operations.StabilityProbeSeconds'; QuestionId = 'SRC-004'; Minimum = 1; Maximum = 300 },
        [pscustomobject]@{ Path = 'Operations.ConversionTimeoutSeconds'; QuestionId = 'CNV-005'; Minimum = 30; Maximum = 7200 },
        [pscustomobject]@{ Path = 'Operations.MaxRetryCount'; QuestionId = 'OPS-002'; Minimum = 0; Maximum = 10 }
    )
    foreach ($bound in $integerBounds) {
        $boundResult = Get-VssPropertyValue -InputObject $Configuration -Path $bound.Path
        if ($boundResult.Found -and (Test-VssIntegerValue -Value $boundResult.Value) -and (($boundResult.Value -lt $bound.Minimum) -or ($boundResult.Value -gt $bound.Maximum))) {
            $errors += "[$($bound.QuestionId)] $($bound.Path) liegt ausserhalb des erlaubten Bereichs."
        }
    }

    $pattern = Get-VssPropertyValue -InputObject $Configuration -Path 'Source.FileNamePattern'
    $regexTimeout = Get-VssPropertyValue -InputObject $Configuration -Path 'Source.RegexTimeoutMilliseconds'
    if ($pattern.Found -and ($pattern.Value -is [string]) -and -not (Test-VssUndecidedValue -Value $pattern.Value) -and
        $regexTimeout.Found -and (Test-VssIntegerValue -Value $regexTimeout.Value) -and
        ($regexTimeout.Value -ge 50) -and ($regexTimeout.Value -le 5000)) {
        try { [void](New-VssSourceRegex -Pattern $pattern.Value -TimeoutMilliseconds ([int]$regexTimeout.Value)) }
        catch { $errors += '[SRC-002] Source.FileNamePattern ist kein gueltiger regulaerer Ausdruck.' }
    }

    $localPathValues = @()
    $hasWorkloadIdentityFilePath = $false
    if ($workloadIdentityFilePath.Found -and -not (Test-VssUndecidedValue -Value $workloadIdentityFilePath.Value) -and
        (Test-VssAbsoluteLocalPathSyntax -Value $workloadIdentityFilePath.Value)) {
        $localPathValues += $workloadIdentityFilePath.Value
        $hasWorkloadIdentityFilePath = $true
    }
    foreach ($operationPath in @('Operations.StatePath', 'Operations.StagingPath', 'Operations.LogPath')) {
        $pathResult = Get-VssPropertyValue -InputObject $Configuration -Path $operationPath
        if ($pathResult.Found -and -not (Test-VssUndecidedValue -Value $pathResult.Value)) {
            if (-not (Test-VssAbsoluteLocalPathSyntax -Value $pathResult.Value)) {
                $operationQuestionId = switch ($operationPath) {
                    'Operations.StatePath' { 'OPS-003' }
                    'Operations.StagingPath' { 'OPS-005' }
                    'Operations.LogPath' { 'OPS-004' }
                }
                $errors += "[$operationQuestionId] $operationPath muss ein sicherer absoluter lokaler Nicht-Root-Pfad sein."
            }
            else { $localPathValues += $pathResult.Value }
        }
    }
    if (($localPathValues.Count -ge 2) -and (Test-VssLocalPathOverlap -Paths $localPathValues)) {
        if ($hasWorkloadIdentityFilePath) {
            $errors += '[SEC-001, OPS-003, OPS-004, OPS-005] Lokale Konfigurationspfade duerfen weder identisch sein noch ineinander liegen.'
        }
        else {
            $errors += '[OPS-003, OPS-004, OPS-005] Operations-Pfade duerfen weder identisch sein noch ineinander liegen.'
        }
    }

    $direction = Get-VssPropertyValue -InputObject $Configuration -Path 'Sync.Direction'
    $deletePolicy = Get-VssPropertyValue -InputObject $Configuration -Path 'Sync.DeletePolicy'
    if ($direction.Found -and $deletePolicy.Found -and ($direction.Value -eq 'PublishOnly') -and ($deletePolicy.Value -ne 'Never') -and -not (Test-VssUndecidedValue -Value $deletePolicy.Value)) {
        $errors += '[SYN-001, SYN-005] Sync.DeletePolicy muss bei PublishOnly auf Never stehen.'
    }

    $disableMacros = Get-VssPropertyValue -InputObject $Configuration -Path 'Conversion.DisableMacros'
    if ($disableMacros.Found -and ($disableMacros.Value -is [bool]) -and -not $disableMacros.Value) {
        $warnings += 'Conversion.DisableMacros ist false; das spaetere Ausfuehren von Dokumentmakros waere ein Sicherheitsrisiko.'
    }

    return New-VssAssessmentResult -Errors $errors -Warnings $warnings -UnresolvedDecisionIds $unresolved -PlannedStages $plannedStages
}

function Invoke-VisioSharePointSyncAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ConfigurationPath,
        [ValidateSet('Validate', 'DryRun')][string]$Mode = 'Validate'
    )
    $plannedStages = @(Get-VssPlannedStages -Mode $Mode)
    if (-not [System.IO.File]::Exists($ConfigurationPath)) {
        return New-VssAssessmentResult -Errors @('Die Konfigurationsdatei wurde nicht gefunden.') -UnresolvedDecisionIds $script:VssBlockerDecisionIds -PlannedStages $plannedStages
    }
    try { $jsonText = Read-VssUtf8FileStrict -LiteralPath $ConfigurationPath }
    catch {
        return New-VssAssessmentResult -Errors @('Die Konfigurationsdatei konnte nicht als BOM-loses gueltiges UTF-8 gelesen werden.') -UnresolvedDecisionIds $script:VssBlockerDecisionIds -PlannedStages $plannedStages
    }
    try {
        if (Test-VssJsonHasDuplicateObjectKeys -JsonText $jsonText) {
            return New-VssAssessmentResult -Errors @('Die Konfigurationsdatei enthaelt doppelte JSON-Schluessel.') -UnresolvedDecisionIds $script:VssBlockerDecisionIds -PlannedStages $plannedStages
        }
        $configuration = ConvertFrom-Json -InputObject $jsonText -ErrorAction Stop
    }
    catch {
        return New-VssAssessmentResult -Errors @('Die Konfigurationsdatei enthaelt kein gueltiges JSON.') -UnresolvedDecisionIds $script:VssBlockerDecisionIds -PlannedStages $plannedStages
    }
    return Test-VssConfigurationObject -Configuration $configuration -Mode $Mode
}

function Get-VisioSharePointSyncExitCode {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Result)
    if ((@($Result.Errors).Count -gt 0) -or (@($Result.UnresolvedDecisionIds).Count -gt 0)) { return 2 }
    return 0
}

function Format-VisioSharePointSyncAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [ValidateSet('Validate', 'DryRun')][string]$Mode = 'Validate'
    )
    $exitCode = Get-VisioSharePointSyncExitCode -Result $Result
    if (@($Result.Errors).Count -gt 0) { $status = 'UNGUELTIG' }
    elseif (@($Result.UnresolvedDecisionIds).Count -gt 0) { $status = 'UNVOLLSTAENDIG' }
    else { $status = 'GUELTIG' }
    $lines = @('Visio-SharePoint-Sync (Geruest)', "Modus: $Mode", "Status: $status", "Exitcode: $exitCode")
    if (@($Result.Errors).Count -gt 0) {
        $lines += 'Fehler:'
        foreach ($item in @($Result.Errors)) { $lines += "  - $item" }
    }
    if (@($Result.Warnings).Count -gt 0) {
        $lines += 'Warnungen:'
        foreach ($item in @($Result.Warnings)) { $lines += "  - $item" }
    }
    if (@($Result.UnresolvedDecisionIds).Count -gt 0) {
        $lines += 'Offene Fragen-IDs:'
        foreach ($item in @($Result.UnresolvedDecisionIds)) { $lines += "  - $item" }
    }
    $lines += if ($Mode -eq 'DryRun') { 'Geplante Stufen (nur Anzeige, nicht ausgefuehrt):' } else { 'Ausgefuehrte Stufen:' }
    foreach ($item in @($Result.PlannedStages)) { $lines += "  - $item" }
    return $lines
}
