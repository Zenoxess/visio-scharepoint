[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Visio automation is most reliable in an STA process. PowerShell 7 normally
# starts in MTA, so transparently restart the same mirror run in STA.
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'Das Spiegel-Skript muss aus einer Datei gestartet werden.'
    }
    $hostExecutable = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    & $hostExecutable -NoLogo -NoProfile -STA -File $PSCommandPath -ConfigurationPath $ConfigurationPath
    exit $LASTEXITCODE
}

$script:MirrorLogPath = $null
$script:MirrorUtf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:AllowedExtensions = @('.vsd', '.vsdx', '.vsdm')
$script:SharePointTimeoutSeconds = 300
$script:SharePointHeaders = @{
    Accept                         = 'application/json;odata=verbose'
    'X-FORMS_BASED_AUTH_ACCEPTED' = 'f'
}

function Write-MirrorLog {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message
    )

    $line = '{0} [{1}] {2}' -f [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    [Console]::WriteLine($line)
    if (-not [string]::IsNullOrWhiteSpace($script:MirrorLogPath)) {
        [IO.File]::AppendAllText($script:MirrorLogPath, $line + [Environment]::NewLine, $script:MirrorUtf8NoBom)
    }
}

function Test-ObjectProperty {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) { return $false }
    return ($null -ne $InputObject.PSObject.Properties[$Name])
}

function Test-SafeSharePointSegment {
    param([Parameter(Mandatory = $true)][string]$Segment)

    if ([string]::IsNullOrWhiteSpace($Segment) -or $Segment -ne $Segment.Trim()) { return $false }
    if ($Segment -eq '.' -or $Segment -eq '..' -or $Segment.EndsWith('.')) { return $false }
    if ($Segment.Length -gt 128) { return $false }
    if ($Segment.IndexOfAny([char[]]'~"#%&*:<>?/\{|}[]') -ge 0) { return $false }
    foreach ($character in $Segment.ToCharArray()) {
        if ([char]::IsControl($character)) { return $false }
    }
    if ($Segment -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$') { return $false }
    if ($Segment -match '^(?i:_vti_)') { return $false }
    return $true
}

function Get-NormalizedRelativePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    return $Path.Replace([char]92, [char]47).Trim([char]47)
}

function Test-AbsoluteWindowsPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    return $Path -match '^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+(?:[\\/]|$))'
}

function Get-PathKey {
    param([Parameter(Mandatory = $true)][string]$Path)

    return (Get-NormalizedRelativePath -Path $Path).Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
}

function ConvertTo-ODataLiteral {
    param([Parameter(Mandatory = $true)][string]$Value)

    return $Value.Replace("'", "''")
}

function Join-ServerRelativeUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Base,
        [Parameter(Mandatory = $true)][string]$Child
    )

    $left = $Base.TrimEnd([char]47)
    $right = (Get-NormalizedRelativePath -Path $Child)
    if ([string]::IsNullOrWhiteSpace($right)) { return $left }
    return $left + '/' + $right
}

function New-AbsoluteUriText {
    param([Parameter(Mandatory = $true)][string]$Text)

    return ([uri]$Text).AbsoluteUri
}

function Read-MirrorConfiguration {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $fullConfigurationPath = [IO.Path]::GetFullPath($LiteralPath)
    if (-not [IO.File]::Exists($fullConfigurationPath)) {
        throw 'Die Konfigurationsdatei wurde nicht gefunden.'
    }

    try {
        $configuration = [IO.File]::ReadAllText($fullConfigurationPath, (New-Object Text.UTF8Encoding($false, $true))) | ConvertFrom-Json
    }
    catch {
        throw 'Die Konfigurationsdatei ist kein gueltiges UTF-8-JSON.'
    }

    $required = @('SourcePath', 'SharePointSiteUrl', 'LibraryName', 'TargetFolderPath', 'TargetFolderUniqueId', 'LogPath')
    if ($null -eq $configuration -or $configuration -is [array]) {
        throw 'Die Konfiguration muss genau ein JSON-Objekt enthalten.'
    }
    foreach ($name in @($configuration.PSObject.Properties.Name)) {
        if ($required -cnotcontains $name) { throw 'Die Konfiguration enthaelt einen unbekannten Schluessel.' }
    }
    foreach ($name in $required) {
        if (-not (Test-ObjectProperty -InputObject $configuration -Name $name) -or
            -not ($configuration.$name -is [string]) -or
            [string]::IsNullOrWhiteSpace([string]$configuration.$name)) {
            throw "Der Konfigurationswert $name fehlt oder ist leer."
        }
    }

    $rawSourcePath = [string]$configuration.SourcePath
    if (-not (Test-AbsoluteWindowsPath -Path $rawSourcePath)) {
        throw 'SourcePath muss ein absoluter Windows-Ordner sein.'
    }
    $sourcePath = [IO.Path]::GetFullPath($rawSourcePath).TrimEnd([char]92)
    $sourceRoot = [IO.Path]::GetPathRoot($sourcePath)
    if (-not [IO.Path]::IsPathRooted($sourcePath) -or [string]::IsNullOrWhiteSpace($sourceRoot) -or
        [string]::Equals($sourcePath, $sourceRoot.TrimEnd([char]92), [StringComparison]::OrdinalIgnoreCase)) {
        throw 'SourcePath muss ein absoluter Windows-Ordner unterhalb einer Laufwerks- oder Freigabewurzel sein.'
    }

    $siteUri = $null
    try { $siteUri = [uri]([string]$configuration.SharePointSiteUrl) }
    catch { throw 'SharePointSiteUrl ist keine gueltige absolute URL.' }
    if (-not $siteUri.IsAbsoluteUri -or $siteUri.Scheme -cne 'https' -or
        -not [string]::IsNullOrWhiteSpace($siteUri.UserInfo) -or
        -not [string]::IsNullOrWhiteSpace($siteUri.Query) -or
        -not [string]::IsNullOrWhiteSpace($siteUri.Fragment)) {
        throw 'SharePointSiteUrl muss eine HTTPS-URL ohne Zugangsdaten, Query oder Fragment sein.'
    }

    $libraryName = ([string]$configuration.LibraryName).Trim()
    if (-not (Test-SafeSharePointSegment -Segment $libraryName)) {
        throw 'LibraryName enthaelt einen nicht unterstuetzten Bibliothekstitel.'
    }

    $rawTargetFolderPath = [string]$configuration.TargetFolderPath
    if ($rawTargetFolderPath.StartsWith('/') -or $rawTargetFolderPath.StartsWith('\') -or $rawTargetFolderPath.Contains(':')) {
        throw 'TargetFolderPath muss relativ zur Bibliothekswurzel sein.'
    }
    $targetFolderPath = Get-NormalizedRelativePath -Path $rawTargetFolderPath
    if ([string]::IsNullOrWhiteSpace($targetFolderPath)) {
        throw 'TargetFolderPath muss einen dedizierten Unterordner innerhalb der Bibliothek bezeichnen.'
    }
    foreach ($segment in @($targetFolderPath.Split([char]47))) {
        if (-not (Test-SafeSharePointSegment -Segment $segment)) {
            throw 'TargetFolderPath enthaelt ein nicht unterstuetztes SharePoint-Pfadsegment.'
        }
    }

    $targetFolderId = [guid]::Empty
    if (-not [guid]::TryParse([string]$configuration.TargetFolderUniqueId, [ref]$targetFolderId) -or
        $targetFolderId -eq [guid]::Empty) {
        throw 'TargetFolderUniqueId muss eine gueltige, nicht leere GUID sein.'
    }

    $rawLogPath = [string]$configuration.LogPath
    if (-not (Test-AbsoluteWindowsPath -Path $rawLogPath)) { throw 'LogPath muss ein absoluter Dateipfad sein.' }
    $logPath = [IO.Path]::GetFullPath($rawLogPath)

    return [pscustomobject][ordered]@{
        SourcePath           = $sourcePath
        SharePointSiteUrl    = $siteUri.AbsoluteUri.TrimEnd([char]47)
        LibraryName          = $libraryName
        TargetFolderPath     = $targetFolderPath
        TargetFolderUniqueId = $targetFolderId
        LogPath              = $logPath
    }
}

function Initialize-MirrorLog {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $parent = [IO.Path]::GetDirectoryName($LiteralPath)
    if ([string]::IsNullOrWhiteSpace($parent)) { throw 'LogPath besitzt keinen gueltigen Elternordner.' }
    if (-not [IO.Directory]::Exists($parent)) { [void][IO.Directory]::CreateDirectory($parent) }
    $script:MirrorLogPath = $LiteralPath
}

function Get-SourceInventory {
    param([Parameter(Mandatory = $true)][string]$RootPath)

    $root = New-Object IO.DirectoryInfo($RootPath)
    if (-not $root.Exists) { throw 'Der konfigurierte Quellordner ist nicht erreichbar.' }
    if (($root.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw 'Der konfigurierte Quellordner darf kein Reparse Point sein.'
    }
    if (-not (Test-SafeSharePointSegment -Segment $root.Name)) {
        throw 'Der Name des Quellordners kann nicht sicher nach SharePoint gespiegelt werden.'
    }

    $rootPrefix = $root.FullName.TrimEnd([char]92)
    $pending = New-Object 'Collections.Generic.Stack[IO.DirectoryInfo]'
    $pending.Push($root)
    $items = @()
    $targetOwners = @{}

    while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        try {
            $files = @($directory.GetFiles() | Sort-Object Name)
            $directories = @($directory.GetDirectories() | Sort-Object Name -Descending)
        }
        catch {
            throw 'Mindestens ein Quellordner konnte nicht vollstaendig gelesen werden.'
        }

        foreach ($childDirectory in $directories) {
            try { $attributes = $childDirectory.Attributes }
            catch { throw 'Die Attribute eines Quellordners konnten nicht gelesen werden.' }
            if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'Ein Unterordner ist ein Reparse Point; der Spiegel-Lauf wurde sicher abgebrochen.'
            }
            $pending.Push($childDirectory)
        }

        foreach ($file in $files) {
            if ($file.Name.StartsWith('~$', [StringComparison]::OrdinalIgnoreCase)) { continue }
            $isVisio = $false
            foreach ($extension in $script:AllowedExtensions) {
                if ([string]::Equals($file.Extension, $extension, [StringComparison]::OrdinalIgnoreCase)) {
                    $isVisio = $true
                    break
                }
            }
            if (-not $isVisio) { continue }

            try { $attributes = $file.Attributes }
            catch { throw 'Die Attribute einer Visio-Datei konnten nicht gelesen werden.' }
            if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'Eine Visio-Datei ist ein Reparse Point; der Spiegel-Lauf wurde sicher abgebrochen.'
            }

            $relativeSourcePath = $file.FullName.Substring($rootPrefix.Length).TrimStart([char]92)
            $sourceSegments = @($relativeSourcePath.Replace([char]92, [char]47).Split([char]47))
            foreach ($segment in $sourceSegments) {
                if (-not (Test-SafeSharePointSegment -Segment $segment)) {
                    throw 'Mindestens ein Visio-Pfad kann nicht sicher nach SharePoint gespiegelt werden.'
                }
            }

            $baseName = [IO.Path]::GetFileNameWithoutExtension($file.Name)
            if ([string]::IsNullOrWhiteSpace($baseName)) {
                throw 'Eine Visio-Datei besitzt keinen gueltigen PDF-Basisnamen.'
            }
            $targetFileName = $baseName + '.pdf'
            if (-not (Test-SafeSharePointSegment -Segment $targetFileName)) {
                throw 'Mindestens ein erzeugter PDF-Dateiname ist in SharePoint nicht zulaessig.'
            }
            $targetSegments = @($root.Name) + $sourceSegments
            $targetSegments[$targetSegments.Count - 1] = $targetFileName
            $targetRelativePath = $targetSegments -join '/'
            $targetKey = Get-PathKey -Path $targetRelativePath
            if ($targetOwners.ContainsKey($targetKey)) {
                throw 'Mehrere Visio-Dateien wuerden denselben PDF-Zielpfad erzeugen.'
            }
            $targetOwners[$targetKey] = $relativeSourcePath

            try {
                $items += [pscustomobject][ordered]@{
                    SourcePath             = $file.FullName
                    SourceRelativePath     = $relativeSourcePath
                    SourceLength           = [long]$file.Length
                    SourceLastWriteUtcTicks = [long]$file.LastWriteTimeUtc.Ticks
                    TargetRelativePath     = $targetRelativePath
                }
            }
            catch {
                throw 'Die Metadaten einer Visio-Datei konnten nicht gelesen werden.'
            }
        }
    }

    return @($items | Sort-Object SourceRelativePath)
}

function Get-ExpectedFolders {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Inventory)

    $folders = @{}
    foreach ($item in $Inventory) {
        $segments = @(([string]$item.TargetRelativePath).Split([char]47))
        for ($count = 1; $count -lt $segments.Count; $count++) {
            $relativeFolder = $segments[0..($count - 1)] -join '/'
            $folderKey = Get-PathKey -Path $relativeFolder
            if ($folders.ContainsKey($folderKey) -and
                -not [string]::Equals([string]$folders[$folderKey], $relativeFolder, [StringComparison]::Ordinal)) {
                throw 'Mehrere Quellordner wuerden auf denselben SharePoint-Zielordner abgebildet.'
            }
            $folders[$folderKey] = $relativeFolder
        }
    }
    return @($folders.Values | Sort-Object @{ Expression = { @($_.Split([char]47)).Count } }, @{ Expression = { $_ } })
}

function Assert-NoFileFolderCollisions {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Inventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ExpectedFolders
    )

    $folderKeys = @{}
    foreach ($folder in $ExpectedFolders) { $folderKeys[(Get-PathKey -Path $folder)] = $true }
    foreach ($item in $Inventory) {
        if ($folderKeys.ContainsKey((Get-PathKey -Path ([string]$item.TargetRelativePath)))) {
            throw 'Ein PDF-Zielpfad kollidiert mit einem benoetigten Zielordner.'
        }
    }
}

function Test-SameSourceInventory {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Expected,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Actual
    )

    if ($Expected.Count -ne $Actual.Count) { return $false }
    for ($index = 0; $index -lt $Expected.Count; $index++) {
        if (-not [string]::Equals([string]$Expected[$index].SourceRelativePath, [string]$Actual[$index].SourceRelativePath, [StringComparison]::OrdinalIgnoreCase) -or
            [long]$Expected[$index].SourceLength -ne [long]$Actual[$index].SourceLength -or
            [long]$Expected[$index].SourceLastWriteUtcTicks -ne [long]$Actual[$index].SourceLastWriteUtcTicks) {
            return $false
        }
    }
    return $true
}

function Release-ComObject {
    param([AllowNull()][object]$InputObject)

    if ($null -ne $InputObject -and [Runtime.InteropServices.Marshal]::IsComObject($InputObject)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($InputObject) }
        catch {}
    }
}

function Assert-ValidPdf {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    if (-not [IO.File]::Exists($LiteralPath)) { throw 'Visio hat keine PDF-Datei erzeugt.' }
    $stream = $null
    try {
        $stream = [IO.File]::OpenRead($LiteralPath)
        if ($stream.Length -lt 5) { throw 'Die erzeugte PDF-Datei ist leer oder unvollstaendig.' }
        $header = New-Object byte[] 5
        if ($stream.Read($header, 0, 5) -ne 5 -or [Text.Encoding]::ASCII.GetString($header) -cne '%PDF-') {
            throw 'Die erzeugte Datei besitzt keinen gueltigen PDF-Header.'
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Convert-InventoryToPdf {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Inventory,
        [Parameter(Mandatory = $true)][string]$RunPath
    )

    $visio = $documents = $null
    $artifacts = @()
    $quitFailure = $false
    try {
        $visioType = [Type]::GetTypeFromProgID('Visio.InvisibleApp', $true)
        $visio = [Activator]::CreateInstance($visioType)
        $visio.AlertResponse = 7
        $visio.EventsEnabled = $false
        $visio.ShowChanges = $false
        $visio.ShowProgress = $false
        $documents = $visio.Documents

        # Read-only, no MRU, hidden, macros disabled, no workspace and no refresh prompt.
        $openFlags = [int16](2 -bor 8 -bor 64 -bor 128 -bor 256 -bor 1024)
        for ($index = 0; $index -lt $Inventory.Count; $index++) {
            $item = $Inventory[$index]
            $document = $recordsets = $null
            try {
                $stagedSource = Join-Path $RunPath (('{0:D6}{1}' -f $index, [IO.Path]::GetExtension([string]$item.SourcePath)))
                $pdfPath = Join-Path $RunPath (('{0:D6}.pdf' -f $index))
                [IO.File]::Copy([string]$item.SourcePath, $stagedSource, $false)
                $document = $documents.OpenEx([IO.Path]::GetFullPath($stagedSource), $openFlags)
                if ([bool]$visio.DataFeaturesEnabled) {
                    $recordsets = $document.DataRecordsets
                    for ($recordsetIndex = 1; $recordsetIndex -le [int]$recordsets.Count; $recordsetIndex++) {
                        $recordset = $null
                        try {
                            $recordset = $recordsets.Item($recordsetIndex)
                            $recordset.RefreshInterval = 0
                        }
                        finally {
                            Release-ComObject $recordset
                            $recordset = $null
                        }
                    }
                }
                [void]$document.ExportAsFixedFormat(
                    1, [IO.Path]::GetFullPath($pdfPath), 1, 0,
                    1, -1, $false, $false, $true, $true, $false
                )
                Assert-ValidPdf -LiteralPath $pdfPath
                $artifacts += [pscustomobject][ordered]@{
                    TargetRelativePath = [string]$item.TargetRelativePath
                    PdfPath            = $pdfPath
                }
            }
            catch {
                throw ("Die Visio-Konvertierung ist fehlgeschlagen: {0}. Ursache: {1}" -f $item.SourceRelativePath, $_.Exception.Message)
            }
            finally {
                Release-ComObject $recordsets
                $recordsets = $null
                if ($null -ne $document) {
                    try {
                        $document.Saved = $true
                        [void]$document.Close()
                    }
                    finally {
                        Release-ComObject $document
                        $document = $null
                    }
                }
            }
        }
    }
    finally {
        Release-ComObject $documents
        $documents = $null
        if ($null -ne $visio) {
            try { [void]$visio.Quit() }
            catch {
                $quitFailure = $true
                try { Write-MirrorLog -Level ERROR -Message 'Die Visio-Instanz konnte nicht sauber beendet werden.' }
                catch {}
            }
            finally {
                Release-ComObject $visio
                $visio = $null
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    if ($quitFailure) { throw 'Die Visio-Instanz konnte nicht sauber beendet werden.' }
    return @($artifacts)
}

function Get-HttpStatusCode {
    param([Parameter(Mandatory = $true)][object]$ErrorRecord)

    try {
        if ($null -ne $ErrorRecord.Exception.Response) {
            return [int]$ErrorRecord.Exception.Response.StatusCode
        }
    }
    catch {}
    return $null
}

function Invoke-SpGet {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [switch]$AllowNotFound
    )

    try {
        return Invoke-RestMethod -Method Get -Uri $Uri -Headers $script:SharePointHeaders -UseDefaultCredentials -TimeoutSec $script:SharePointTimeoutSeconds -ErrorAction Stop
    }
    catch {
        $status = Get-HttpStatusCode -ErrorRecord $_
        if ($AllowNotFound -and $status -eq 404) { return $null }
        $label = if ($null -eq $status) { 'unbekannt' } else { [string]$status }
        throw "SharePoint-REST-Lesezugriff fehlgeschlagen (HTTP $label)."
    }
}

function Get-SpDigest {
    param([Parameter(Mandatory = $true)][string]$SiteUrl)

    try {
        $response = Invoke-RestMethod -Method Post -Uri ($SiteUrl + '/_api/contextinfo') -Headers $script:SharePointHeaders -UseDefaultCredentials -TimeoutSec $script:SharePointTimeoutSeconds -ContentType 'application/json;odata=verbose;charset=utf-8' -ErrorAction Stop
        $digest = [string]$response.d.GetContextWebInformation.FormDigestValue
        if ([string]::IsNullOrWhiteSpace($digest)) { throw 'empty' }
        return $digest
    }
    catch {
        $status = Get-HttpStatusCode -ErrorRecord $_
        $label = if ($null -eq $status) { 'unbekannt' } else { [string]$status }
        throw "SharePoint lieferte keinen Request-Digest (HTTP $label)."
    }
}

function New-SpMutationHeaders {
    param([Parameter(Mandatory = $true)][string]$SiteUrl)

    return @{
        Accept                         = 'application/json;odata=verbose'
        'X-FORMS_BASED_AUTH_ACCEPTED' = 'f'
        'X-RequestDigest'              = (Get-SpDigest -SiteUrl $SiteUrl)
    }
}

function Invoke-SpJsonPost {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][object]$Body
    )

    try {
        $json = $Body | ConvertTo-Json -Depth 6 -Compress
        $jsonBytes = $script:MirrorUtf8NoBom.GetBytes($json)
        return Invoke-RestMethod -Method Post -Uri $Uri -Headers (New-SpMutationHeaders -SiteUrl $SiteUrl) -UseDefaultCredentials -TimeoutSec $script:SharePointTimeoutSeconds -ContentType 'application/json;odata=verbose;charset=utf-8' -Body $jsonBytes -ErrorAction Stop
    }
    catch {
        $status = Get-HttpStatusCode -ErrorRecord $_
        $label = if ($null -eq $status) { 'unbekannt' } else { [string]$status }
        throw "SharePoint-REST-Schreibzugriff fehlgeschlagen (HTTP $label)."
    }
}

function Invoke-SpFilePost {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$LiteralPath
    )

    try {
        return Invoke-RestMethod -Method Post -Uri $Uri -Headers (New-SpMutationHeaders -SiteUrl $SiteUrl) -UseDefaultCredentials -TimeoutSec $script:SharePointTimeoutSeconds -ContentType 'application/pdf' -InFile $LiteralPath -ErrorAction Stop
    }
    catch {
        $status = Get-HttpStatusCode -ErrorRecord $_
        $label = if ($null -eq $status) { 'unbekannt' } else { [string]$status }
        throw "Der PDF-Upload nach SharePoint ist fehlgeschlagen (HTTP $label)."
    }
}

function Invoke-SpActionPost {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$Uri
    )

    try {
        return Invoke-RestMethod -Method Post -Uri $Uri -Headers (New-SpMutationHeaders -SiteUrl $SiteUrl) -UseDefaultCredentials -TimeoutSec $script:SharePointTimeoutSeconds -ContentType 'application/json;odata=verbose;charset=utf-8' -ErrorAction Stop
    }
    catch {
        $status = Get-HttpStatusCode -ErrorRecord $_
        $label = if ($null -eq $status) { 'unbekannt' } else { [string]$status }
        throw "Die SharePoint-Papierkorbaktion ist fehlgeschlagen (HTTP $label)."
    }
}

function Get-SpFolderMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$ServerRelativeUrl,
        [switch]$AllowNotFound
    )

    $literal = ConvertTo-ODataLiteral -Value $ServerRelativeUrl
    $uri = New-AbsoluteUriText -Text ($SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$literal')?`$select=UniqueId,ServerRelativeUrl")
    $response = Invoke-SpGet -Uri $uri -AllowNotFound:$AllowNotFound
    if ($null -eq $response) { return $null }
    return $response.d
}

function Get-SpContext {
    param([Parameter(Mandatory = $true)][object]$Configuration)

    $libraryLiteral = ConvertTo-ODataLiteral -Value ([string]$Configuration.LibraryName)
    $libraryUri = New-AbsoluteUriText -Text ($Configuration.SharePointSiteUrl + "/_api/web/lists/getbytitle('$libraryLiteral')?`$select=ForceCheckout,RootFolder/ServerRelativeUrl,RootFolder/UniqueId&`$expand=RootFolder")
    $libraryResponse = Invoke-SpGet -Uri $libraryUri
    if ($libraryResponse.d.ForceCheckout -eq $true) {
        throw 'Die Zielbibliothek verlangt Auschecken; der Minimalbetrieb setzt ForceCheckout=false voraus.'
    }
    $libraryRootUrl = [string]$libraryResponse.d.RootFolder.ServerRelativeUrl
    if ([string]::IsNullOrWhiteSpace($libraryRootUrl)) { throw 'SharePoint lieferte keinen Bibliothekswurzelpfad.' }

    $targetRootUrl = Join-ServerRelativeUrl -Base $libraryRootUrl -Child ([string]$Configuration.TargetFolderPath)
    $target = Get-SpFolderMetadata -SiteUrl $Configuration.SharePointSiteUrl -ServerRelativeUrl $targetRootUrl
    $actualId = [guid]::Empty
    if ($null -eq $target -or -not [guid]::TryParse([string]$target.UniqueId, [ref]$actualId) -or
        $actualId -ne [guid]$Configuration.TargetFolderUniqueId -or
        -not [string]::Equals([string]$target.ServerRelativeUrl, $targetRootUrl, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Der SharePoint-Zielordner stimmt nicht mit Pfad und TargetFolderUniqueId ueberein.'
    }
    return [pscustomobject][ordered]@{
        SiteUrl       = [string]$Configuration.SharePointSiteUrl
        TargetRootUrl = $targetRootUrl
        TargetRootId  = $actualId
    }
}

function Assert-SharePointPathLengths {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Inventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ExpectedFolders
    )

    if (([string]$Context.TargetRootUrl).Length -gt 260) {
        throw 'Der SharePoint-Zielordnerpfad ueberschreitet 260 Zeichen.'
    }
    foreach ($folder in $ExpectedFolders) {
        if ((Join-ServerRelativeUrl -Base $Context.TargetRootUrl -Child $folder).Length -gt 260) {
            throw 'Mindestens ein SharePoint-Zielordnerpfad ueberschreitet 260 Zeichen.'
        }
    }
    foreach ($item in $Inventory) {
        if ((Join-ServerRelativeUrl -Base $Context.TargetRootUrl -Child ([string]$item.TargetRelativePath)).Length -gt 260) {
            throw 'Mindestens ein SharePoint-PDF-Zielpfad ueberschreitet 260 Zeichen.'
        }
    }
}

function Assert-SpTargetIdentity {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][guid]$ExpectedId
    )

    $folder = Get-SpFolderMetadata -SiteUrl $Context.SiteUrl -ServerRelativeUrl $Context.TargetRootUrl
    $actualId = [guid]::Empty
    if ($null -eq $folder -or
        -not [guid]::TryParse([string]$folder.UniqueId, [ref]$actualId) -or
        $actualId -ne $ExpectedId -or
        -not [string]::Equals([string]$folder.ServerRelativeUrl, [string]$Context.TargetRootUrl, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Die SharePoint-Zielidentitaet hat sich waehrend des Laufs geaendert.'
    }
}

function Ensure-SpFolders {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$RelativeFolders
    )

    foreach ($relativeFolder in $RelativeFolders) {
        $serverRelativeUrl = Join-ServerRelativeUrl -Base $Context.TargetRootUrl -Child $relativeFolder
        if ($null -ne (Get-SpFolderMetadata -SiteUrl $Context.SiteUrl -ServerRelativeUrl $serverRelativeUrl -AllowNotFound)) { continue }
        Assert-SpTargetIdentity -Context $Context -ExpectedId $Context.TargetRootId
        $body = [ordered]@{
            '__metadata'     = [ordered]@{ type = 'SP.Folder' }
            ServerRelativeUrl = $serverRelativeUrl
        }
        [void](Invoke-SpJsonPost -SiteUrl $Context.SiteUrl -Uri ($Context.SiteUrl + '/_api/web/folders') -Body $body)
        if ($null -eq (Get-SpFolderMetadata -SiteUrl $Context.SiteUrl -ServerRelativeUrl $serverRelativeUrl -AllowNotFound)) {
            throw 'Ein SharePoint-Zielordner wurde nach der Anlage nicht gefunden.'
        }
    }
}

function Send-SpPdfs {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Artifacts
    )

    foreach ($artifact in $Artifacts) {
        $segments = @(([string]$artifact.TargetRelativePath).Split([char]47))
        $fileName = $segments[$segments.Count - 1]
        $relativeParent = if ($segments.Count -gt 1) { $segments[0..($segments.Count - 2)] -join '/' } else { '' }
        $parentUrl = Join-ServerRelativeUrl -Base $Context.TargetRootUrl -Child $relativeParent
        $parentLiteral = ConvertTo-ODataLiteral -Value $parentUrl
        $fileLiteral = ConvertTo-ODataLiteral -Value $fileName
        $uri = New-AbsoluteUriText -Text ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$parentLiteral')/Files/add(url='$fileLiteral',overwrite=true)")
        Assert-SpTargetIdentity -Context $Context -ExpectedId $Context.TargetRootId
        [void](Invoke-SpFilePost -SiteUrl $Context.SiteUrl -Uri $uri -LiteralPath ([string]$artifact.PdfPath))
    }
}

function Assert-SpApiUri {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$Uri
    )

    $site = [uri]$SiteUrl
    $candidate = [uri]$Uri
    if ($candidate.Scheme -cne 'https' -or
        -not [string]::Equals($candidate.Host, $site.Host, [StringComparison]::OrdinalIgnoreCase) -or
        $candidate.Port -ne $site.Port -or
        $candidate.AbsolutePath.IndexOf('/_api/', [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw 'SharePoint lieferte einen unsicheren Paging-Link.'
    }
}

function Get-SpPagedItems {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$InitialUri
    )

    $items = @()
    $next = $InitialUri
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        Assert-SpApiUri -SiteUrl $SiteUrl -Uri $next
        $response = Invoke-SpGet -Uri $next
        $page = @($response.d.results)
        $items += $page
        $next = if (Test-ObjectProperty -InputObject $response.d -Name '__next') { [string]$response.d.__next } else { $null }
        if ([string]::IsNullOrWhiteSpace($next) -and $page.Count -ge 5000) {
            throw 'SharePoint lieferte eine moeglicherweise abgeschnittene 5000er-Inventurseite ohne Fortsetzungslink.'
        }
    }
    return @($items)
}

function Get-SpTargetInventory {
    param([Parameter(Mandatory = $true)][object]$Context)

    $files = @()
    $folders = @()
    $pending = New-Object 'Collections.Generic.Stack[object]'
    $pending.Push([pscustomobject]@{ ServerRelativeUrl = $Context.TargetRootUrl; RelativePath = '' })

    while ($pending.Count -gt 0) {
        $current = $pending.Pop()
        $folderLiteral = ConvertTo-ODataLiteral -Value ([string]$current.ServerRelativeUrl)
        $filesUri = New-AbsoluteUriText -Text ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$folderLiteral')/Files?`$select=Name,ServerRelativeUrl,UniqueId&`$top=5000")
        foreach ($file in @(Get-SpPagedItems -SiteUrl $Context.SiteUrl -InitialUri $filesUri)) {
            $relativePath = if ([string]::IsNullOrWhiteSpace([string]$current.RelativePath)) { [string]$file.Name } else { [string]$current.RelativePath + '/' + [string]$file.Name }
            $files += [pscustomobject]@{ RelativePath = $relativePath; ServerRelativeUrl = [string]$file.ServerRelativeUrl }
        }

        $foldersUri = New-AbsoluteUriText -Text ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$folderLiteral')/Folders?`$select=Name,ServerRelativeUrl,UniqueId&`$top=5000")
        foreach ($folder in @(Get-SpPagedItems -SiteUrl $Context.SiteUrl -InitialUri $foldersUri)) {
            $relativePath = if ([string]::IsNullOrWhiteSpace([string]$current.RelativePath)) { [string]$folder.Name } else { [string]$current.RelativePath + '/' + [string]$folder.Name }
            $folders += [pscustomobject]@{ RelativePath = $relativePath; ServerRelativeUrl = [string]$folder.ServerRelativeUrl }
            $pending.Push([pscustomobject]@{ ServerRelativeUrl = [string]$folder.ServerRelativeUrl; RelativePath = $relativePath })
        }
    }

    return [pscustomobject][ordered]@{ Files = @($files); Folders = @($folders) }
}

function Assert-SpRemoteInventory {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][object]$RemoteInventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SourceInventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ExpectedFolders
    )

    $rootPrefix = ([string]$Context.TargetRootUrl).TrimEnd([char]47) + '/'
    foreach ($remoteItem in @($RemoteInventory.Files) + @($RemoteInventory.Folders)) {
        $serverPath = [string]$remoteItem.ServerRelativeUrl
        if (-not $serverPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $serverPath.Length -gt 260) {
            throw 'Die SharePoint-Inventur enthielt einen Pfad ausserhalb des freigegebenen Zielordners.'
        }
        foreach ($segment in @(([string]$remoteItem.RelativePath).Split([char]47))) {
            if (-not (Test-SafeSharePointSegment -Segment $segment)) {
                throw 'Die SharePoint-Inventur enthaelt einen nicht sicher adressierbaren Namen.'
            }
        }
    }

    $remoteFileKeys = @{}
    foreach ($file in @($RemoteInventory.Files)) { $remoteFileKeys[(Get-PathKey -Path ([string]$file.RelativePath))] = $true }
    foreach ($item in $SourceInventory) {
        if (-not $remoteFileKeys.ContainsKey((Get-PathKey -Path ([string]$item.TargetRelativePath)))) {
            throw 'Nach dem Upload fehlt mindestens eine erwartete PDF-Datei in SharePoint.'
        }
    }

    $remoteFolderKeys = @{}
    foreach ($folder in @($RemoteInventory.Folders)) { $remoteFolderKeys[(Get-PathKey -Path ([string]$folder.RelativePath))] = $true }
    foreach ($expectedFolder in $ExpectedFolders) {
        if (-not $remoteFolderKeys.ContainsKey((Get-PathKey -Path $expectedFolder))) {
            throw 'Nach dem Upload fehlt mindestens ein erwarteter Zielordner in SharePoint.'
        }
    }
}

function Invoke-SpRecycleExtras {
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][object]$RemoteInventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SourceInventory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ExpectedFolders
    )

    $expectedFiles = @{}
    foreach ($item in $SourceInventory) { $expectedFiles[(Get-PathKey -Path ([string]$item.TargetRelativePath))] = $true }
    $expectedFolderKeys = @{}
    foreach ($folder in $ExpectedFolders) { $expectedFolderKeys[(Get-PathKey -Path $folder)] = $true }

    $extraFiles = @($RemoteInventory.Files | Where-Object { -not $expectedFiles.ContainsKey((Get-PathKey -Path ([string]$_.RelativePath))) } |
        Sort-Object @{ Expression = { @(([string]$_.RelativePath).Split([char]47)).Count }; Descending = $true }, RelativePath)
    foreach ($file in $extraFiles) {
        $literal = ConvertTo-ODataLiteral -Value ([string]$file.ServerRelativeUrl)
        $uri = New-AbsoluteUriText -Text ($Context.SiteUrl + "/_api/web/GetFileByServerRelativeUrl('$literal')/recycle()")
        Assert-SpTargetIdentity -Context $Context -ExpectedId $Context.TargetRootId
        [void](Invoke-SpActionPost -SiteUrl $Context.SiteUrl -Uri $uri)
    }

    $extraFolders = @($RemoteInventory.Folders | Where-Object { -not $expectedFolderKeys.ContainsKey((Get-PathKey -Path ([string]$_.RelativePath))) } |
        Sort-Object @{ Expression = { @(([string]$_.RelativePath).Split([char]47)).Count }; Descending = $true }, RelativePath)
    foreach ($folder in $extraFolders) {
        $literal = ConvertTo-ODataLiteral -Value ([string]$folder.ServerRelativeUrl)
        $uri = New-AbsoluteUriText -Text ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$literal')/recycle()")
        Assert-SpTargetIdentity -Context $Context -ExpectedId $Context.TargetRootId
        [void](Invoke-SpActionPost -SiteUrl $Context.SiteUrl -Uri $uri)
    }

    return [pscustomobject]@{ RecycledFiles = $extraFiles.Count; RecycledFolders = $extraFolders.Count }
}

function New-MirrorMutexName {
    param([Parameter(Mandatory = $true)][object]$Configuration)

    $identity = ('{0}|{1}|{2}|{3}' -f $Configuration.SharePointSiteUrl, $Configuration.LibraryName, $Configuration.TargetFolderPath, $Configuration.TargetFolderUniqueId).ToLowerInvariant()
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($identity)) }
    finally { $sha.Dispose() }
    $suffix = ([BitConverter]::ToString($hash)).Replace('-', '').Substring(0, 24)
    return 'Global\PPSI_VisioSharePointMirror_' + $suffix
}

$configuration = $null
$mutex = $null
$mutexAcquired = $false
$runPath = $null
$exitCode = 1

try {
    $configuration = Read-MirrorConfiguration -LiteralPath $ConfigurationPath
    Initialize-MirrorLog -LiteralPath $configuration.LogPath
    Write-MirrorLog -Level INFO -Message 'Spiegel-Lauf gestartet.'

    $mutex = New-Object Threading.Mutex($false, (New-MirrorMutexName -Configuration $configuration))
    try { $mutexAcquired = $mutex.WaitOne(0, $false) }
    catch [Threading.AbandonedMutexException] { $mutexAcquired = $true }
    if (-not $mutexAcquired) { throw 'Fuer dieses SharePoint-Ziel laeuft bereits ein Spiegel-Lauf.' }

    $sourceInventory = @(Get-SourceInventory -RootPath $configuration.SourcePath)
    $expectedFolders = @(Get-ExpectedFolders -Inventory $sourceInventory)
    Assert-NoFileFolderCollisions -Inventory $sourceInventory -ExpectedFolders $expectedFolders
    Write-MirrorLog -Level INFO -Message ("{0} Visio-Datei(en) gefunden." -f $sourceInventory.Count)

    $spContext = Get-SpContext -Configuration $configuration
    Assert-SharePointPathLengths -Context $spContext -Inventory $sourceInventory -ExpectedFolders $expectedFolders
    Write-MirrorLog -Level INFO -Message 'SharePoint-Ziel und Windows-Anmeldung wurden bestaetigt.'

    $runPath = Join-Path ([IO.Path]::GetTempPath()) ('PPSI-VisioSharePointMirror-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($runPath)
    $artifacts = @(Convert-InventoryToPdf -Inventory $sourceInventory -RunPath $runPath)
    Write-MirrorLog -Level INFO -Message ("{0} PDF-Datei(en) erzeugt." -f $artifacts.Count)

    $inventoryBeforeWrite = @(Get-SourceInventory -RootPath $configuration.SourcePath)
    if (-not (Test-SameSourceInventory -Expected $sourceInventory -Actual $inventoryBeforeWrite)) {
        throw 'Die Quelle hat sich waehrend der Konvertierung geaendert; SharePoint blieb unveraendert.'
    }

    Assert-SpTargetIdentity -Context $spContext -ExpectedId $configuration.TargetFolderUniqueId
    Ensure-SpFolders -Context $spContext -RelativeFolders $expectedFolders
    Send-SpPdfs -Context $spContext -Artifacts $artifacts
    Write-MirrorLog -Level INFO -Message ("{0} PDF-Datei(en) nach SharePoint hochgeladen." -f $artifacts.Count)

    $inventoryBeforeRecycle = @(Get-SourceInventory -RootPath $configuration.SourcePath)
    if (-not (Test-SameSourceInventory -Expected $sourceInventory -Actual $inventoryBeforeRecycle)) {
        throw 'Die Quelle hat sich waehrend des Uploads geaendert; Papierkorbaktionen wurden unterdrueckt.'
    }
    Assert-SpTargetIdentity -Context $spContext -ExpectedId $configuration.TargetFolderUniqueId
    $remoteInventory = Get-SpTargetInventory -Context $spContext
    Assert-SpTargetIdentity -Context $spContext -ExpectedId $configuration.TargetFolderUniqueId
    Assert-SpRemoteInventory -Context $spContext -RemoteInventory $remoteInventory -SourceInventory $sourceInventory -ExpectedFolders $expectedFolders
    $inventoryAfterRemoteRead = @(Get-SourceInventory -RootPath $configuration.SourcePath)
    if (-not (Test-SameSourceInventory -Expected $sourceInventory -Actual $inventoryAfterRemoteRead)) {
        throw 'Die Quelle hat sich waehrend der SharePoint-Inventur geaendert; Papierkorbaktionen wurden unterdrueckt.'
    }
    Assert-SpTargetIdentity -Context $spContext -ExpectedId $configuration.TargetFolderUniqueId
    $recycleResult = Invoke-SpRecycleExtras -Context $spContext -RemoteInventory $remoteInventory -SourceInventory $sourceInventory -ExpectedFolders $expectedFolders

    Write-MirrorLog -Level INFO -Message ("Spiegelabgleich abgeschlossen: {0} Datei(en) und {1} Ordner recycelt." -f $recycleResult.RecycledFiles, $recycleResult.RecycledFolders)
    $exitCode = 0
}
catch {
    $message = if ([string]::IsNullOrWhiteSpace($_.Exception.Message)) { 'Unbekannter Fehler im Spiegel-Lauf.' } else { $_.Exception.Message }
    try { Write-MirrorLog -Level ERROR -Message $message }
    catch { [Console]::Error.WriteLine($message) }
    $exitCode = 1
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($runPath) -and [IO.Directory]::Exists($runPath)) {
        try { [IO.Directory]::Delete($runPath, $true) }
        catch {
            $exitCode = 1
            try { Write-MirrorLog -Level ERROR -Message 'Das temporaere Laufverzeichnis konnte nicht vollstaendig entfernt werden.' }
            catch {}
        }
    }
    if ($mutexAcquired -and $null -ne $mutex) {
        try { $mutex.ReleaseMutex() }
        catch {}
    }
    if ($null -ne $mutex) { $mutex.Dispose() }
}

exit $exitCode
