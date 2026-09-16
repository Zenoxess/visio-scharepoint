<#
.SYNOPSIS
Spiegelt Visio-Dateien als PDFs nach SharePoint.

.DESCRIPTION
Einziger Betriebsmodus: vollstaendiger Spiegelabgleich. Das Skript sucht
rekursiv nach .vsd, .vsdx und .vsdm, konvertiert mit lokalem Visio und bildet
den Pfad inklusive Quellordnername ab:

P:\Quelle\A\B\Datei.vsdx -> <TargetFolderPath>\Quelle\A\B\Datei.pdf

Die Konfiguration enthaelt nur SourcePath, SharePointSiteUrl, LibraryName,
TargetFolderPath, TargetFolderUniqueId und LogPath. Werte mit
__PLATZHALTER_...__ sind Vorlagen und stoppen den Lauf.

Der SharePoint-Zielordner ist dediziert: Inhalte darunter duerfen nach einem
vollstaendig erfolgreichen Lauf in den Papierkorb verschoben werden. Der
Zielordner selbst wird nie geloescht. Es gibt keine Preview-, Test-, DryRun-,
State- oder Scheduler-Logik.
#>

#requires -Version 5.1
#requires -PSEdition Desktop

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Visio-COM braucht STA. Falls ein fremder Host MTA nutzt, startet sich das Skript
# einmalig selbst in STA neu.
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) { throw 'Das Skript muss aus einer Datei gestartet werden.' }
    & ([Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -NoLogo -NoProfile -STA -File $PSCommandPath -ConfigurationPath $ConfigurationPath
    exit $LASTEXITCODE
}

$script:LogPath = $null
$script:Utf8NoBom = New-Object Text.UTF8Encoding($false)
$script:AllowedExtensions = @('.vsd', '.vsdx', '.vsdm')
$script:SpTimeoutSeconds = 300
$script:SpHeaders = @{
    Accept                         = 'application/json;odata=verbose'
    'X-FORMS_BASED_AUTH_ACCEPTED' = 'f'
}

function Write-Log {
    param([ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level, [string]$Message)
    $line = '{0} [{1}] {2}' -f [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    [Console]::WriteLine($line)
    if ($script:LogPath) { [IO.File]::AppendAllText($script:LogPath, $line + [Environment]::NewLine, $script:Utf8NoBom) }
}

function Has-Property {
    param([AllowNull()][object]$InputObject, [string]$Name)
    return ($null -ne $InputObject -and $null -ne $InputObject.PSObject.Properties[$Name])
}

function ConvertTo-RelativeUrl {
    param([string]$Path)
    return $Path.Replace([char]92, [char]47).Trim([char]47)
}

function Get-PathKey {
    param([string]$Path)
    return (ConvertTo-RelativeUrl $Path).Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
}

function ConvertTo-ODataLiteral {
    param([string]$Value)
    return $Value.Replace("'", "''")
}

function Join-SpUrl {
    param([string]$Base, [string]$Child)
    $childUrl = ConvertTo-RelativeUrl $Child
    if ([string]::IsNullOrWhiteSpace($childUrl)) { return $Base.TrimEnd([char]47) }
    return $Base.TrimEnd([char]47) + '/' + $childUrl
}

function New-AbsoluteUri {
    param([string]$Text)
    return ([uri]$Text).AbsoluteUri
}

function Test-AbsoluteWindowsPath {
    param([string]$Path)
    return $Path -match '^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+(?:[\\/]|$))'
}

function Test-SafeSharePointSegment {
    param([string]$Segment)
    if ([string]::IsNullOrWhiteSpace($Segment) -or $Segment -ne $Segment.Trim()) { return $false }
    if ($Segment -eq '.' -or $Segment -eq '..' -or $Segment.EndsWith('.') -or $Segment.Length -gt 128) { return $false }
    if ($Segment.IndexOfAny([char[]]'~"#%&*:<>?/\{|}[]') -ge 0) { return $false }
    foreach ($char in $Segment.ToCharArray()) { if ([char]::IsControl($char)) { return $false } }
    if ($Segment -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$') { return $false }
    if ($Segment -match '^(?i:_vti_)') { return $false }
    return $true
}

function Assert-SafeRelativePath {
    param([string]$Path, [string]$ErrorMessage)
    foreach ($segment in @((ConvertTo-RelativeUrl $Path).Split([char]47))) {
        if (-not (Test-SafeSharePointSegment $segment)) { throw $ErrorMessage }
    }
}

function Read-Configuration {
    param([string]$Path)
    $configPath = [IO.Path]::GetFullPath($Path)
    if (-not [IO.File]::Exists($configPath)) { throw 'Die Konfigurationsdatei wurde nicht gefunden.' }

    try {
        $config = [IO.File]::ReadAllText($configPath, (New-Object Text.UTF8Encoding($false, $true))) | ConvertFrom-Json
    }
    catch {
        throw 'Die Konfigurationsdatei ist kein gueltiges UTF-8-JSON.'
    }

    $required = @('SourcePath', 'SharePointSiteUrl', 'LibraryName', 'TargetFolderPath', 'TargetFolderUniqueId', 'LogPath')
    if ($null -eq $config -or $config -is [array]) { throw 'Die Konfiguration muss genau ein JSON-Objekt enthalten.' }
    foreach ($name in @($config.PSObject.Properties.Name)) {
        if ($required -cnotcontains $name) { throw "Unbekannter Konfigurationswert: $name." }
    }
    foreach ($name in $required) {
        if (-not (Has-Property $config $name) -or -not ($config.$name -is [string]) -or [string]::IsNullOrWhiteSpace([string]$config.$name)) {
            throw "Der Konfigurationswert $name fehlt oder ist leer."
        }
        # MARKIERTER PLATZHALTER: Die ausgelieferte mirror.json muss vor Betrieb befuellt werden.
        if ([string]$config.$name -like '__PLATZHALTER_*__') {
            throw "Der Konfigurationswert $name ist noch ein __PLATZHALTER_...__ und muss ersetzt werden."
        }
    }

    $sourcePath = [IO.Path]::GetFullPath([string]$config.SourcePath).TrimEnd([char]92)
    $sourceRoot = [IO.Path]::GetPathRoot($sourcePath)
    if (-not (Test-AbsoluteWindowsPath $sourcePath) -or [string]::IsNullOrWhiteSpace($sourceRoot) -or
        [string]::Equals($sourcePath, $sourceRoot.TrimEnd([char]92), [StringComparison]::OrdinalIgnoreCase)) {
        throw 'SourcePath muss ein absoluter Windows-Ordner unterhalb einer Laufwerks- oder Freigabewurzel sein.'
    }

    try { $siteUri = [uri]([string]$config.SharePointSiteUrl) }
    catch { throw 'SharePointSiteUrl ist keine gueltige absolute URL.' }
    if (-not $siteUri.IsAbsoluteUri -or $siteUri.Scheme -cne 'https' -or $siteUri.UserInfo -or $siteUri.Query -or $siteUri.Fragment) {
        throw 'SharePointSiteUrl muss eine HTTPS-URL ohne Zugangsdaten, Query oder Fragment sein.'
    }

    $libraryName = ([string]$config.LibraryName).Trim()
    if (-not (Test-SafeSharePointSegment $libraryName)) { throw 'LibraryName enthaelt einen nicht unterstuetzten Namen.' }

    $targetFolderPath = ConvertTo-RelativeUrl ([string]$config.TargetFolderPath)
    if ([string]$config.TargetFolderPath -match '^[\\/]|:' -or [string]::IsNullOrWhiteSpace($targetFolderPath)) {
        throw 'TargetFolderPath muss ein relativer, dedizierter Unterordner der Bibliothek sein.'
    }
    Assert-SafeRelativePath $targetFolderPath 'TargetFolderPath enthaelt ein nicht unterstuetztes SharePoint-Pfadsegment.'

    $targetId = [guid]::Empty
    if (-not [guid]::TryParse([string]$config.TargetFolderUniqueId, [ref]$targetId) -or $targetId -eq [guid]::Empty) {
        throw 'TargetFolderUniqueId muss eine gueltige, nicht leere GUID sein.'
    }

    $logPath = [IO.Path]::GetFullPath([string]$config.LogPath)
    if (-not (Test-AbsoluteWindowsPath $logPath)) { throw 'LogPath muss ein absoluter Dateipfad sein.' }

    [pscustomobject][ordered]@{
        SourcePath           = $sourcePath
        SharePointSiteUrl    = $siteUri.AbsoluteUri.TrimEnd([char]47)
        LibraryName          = $libraryName
        TargetFolderPath     = $targetFolderPath
        TargetFolderUniqueId = $targetId
        LogPath              = $logPath
    }
}

function Initialize-Log {
    param([string]$Path)
    $parent = [IO.Path]::GetDirectoryName($Path)
    if ([string]::IsNullOrWhiteSpace($parent)) { throw 'LogPath besitzt keinen gueltigen Elternordner.' }
    if (-not [IO.Directory]::Exists($parent)) { [void][IO.Directory]::CreateDirectory($parent) }
    $script:LogPath = $Path
}

function Get-SourceInventory {
    param([string]$RootPath)
    $root = New-Object IO.DirectoryInfo($RootPath)
    if (-not $root.Exists) { throw 'Der konfigurierte Quellordner ist nicht erreichbar.' }
    if (($root.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Der Quellordner darf kein Reparse Point sein.' }
    if (-not (Test-SafeSharePointSegment $root.Name)) { throw 'Der Quellordnername kann nicht sicher nach SharePoint gespiegelt werden.' }

    $rootPrefix = $root.FullName.TrimEnd([char]92)
    $stack = New-Object 'Collections.Generic.Stack[IO.DirectoryInfo]'
    $stack.Push($root)
    $items = @()
    $targetOwners = @{}

    # Der Stack ersetzt Rekursion ohne Tiefenlimit. Jeder Reparse Point bricht ab,
    # damit der Lauf den freigegebenen Quellbaum nicht verlaesst.
    while ($stack.Count -gt 0) {
        $directory = $stack.Pop()
        try {
            $files = @($directory.GetFiles() | Sort-Object Name)
            $dirs = @($directory.GetDirectories() | Sort-Object Name -Descending)
        }
        catch {
            throw 'Mindestens ein Quellordner konnte nicht vollstaendig gelesen werden.'
        }

        foreach ($dir in $dirs) {
            if (($dir.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Ein Unterordner ist ein Reparse Point.' }
            $stack.Push($dir)
        }

        foreach ($file in $files) {
            if ($file.Name.StartsWith('~$', [StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($script:AllowedExtensions -notcontains $file.Extension.ToLowerInvariant()) { continue }
            if (($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Eine Visio-Datei ist ein Reparse Point.' }

            $sourceRel = $file.FullName.Substring($rootPrefix.Length).TrimStart([char]92)
            $sourceSegments = @($sourceRel.Replace([char]92, [char]47).Split([char]47))
            foreach ($segment in $sourceSegments) {
                if (-not (Test-SafeSharePointSegment $segment)) { throw 'Mindestens ein Visio-Pfad ist fuer SharePoint nicht zulaessig.' }
            }

            $pdfName = [IO.Path]::GetFileNameWithoutExtension($file.Name) + '.pdf'
            if (-not (Test-SafeSharePointSegment $pdfName)) { throw 'Mindestens ein erzeugter PDF-Dateiname ist nicht zulaessig.' }
            $targetSegments = @($root.Name) + $sourceSegments
            $targetSegments[$targetSegments.Count - 1] = $pdfName
            $targetRel = $targetSegments -join '/'
            $key = Get-PathKey $targetRel
            if ($targetOwners.ContainsKey($key)) { throw 'Mehrere Visio-Dateien wuerden denselben PDF-Zielpfad erzeugen.' }
            $targetOwners[$key] = $sourceRel

            $items += [pscustomobject][ordered]@{
                SourcePath              = $file.FullName
                SourceRelativePath      = $sourceRel
                SourceLength            = [long]$file.Length
                SourceLastWriteUtcTicks = [long]$file.LastWriteTimeUtc.Ticks
                TargetRelativePath      = $targetRel
            }
        }
    }
    return @($items | Sort-Object SourceRelativePath)
}

function New-PathSet {
    param([AllowEmptyCollection()][string[]]$Paths)
    $set = @{}
    foreach ($path in $Paths) { $set[(Get-PathKey $path)] = $true }
    return $set
}

function New-MirrorPlan {
    param([AllowEmptyCollection()][object[]]$Inventory)
    $files = New-PathSet -Paths @($Inventory | ForEach-Object { $_.TargetRelativePath })
    $folders = @{}
    foreach ($item in $Inventory) {
        $segments = @(([string]$item.TargetRelativePath).Split([char]47))
        for ($i = 1; $i -lt $segments.Count; $i++) {
            $folder = $segments[0..($i - 1)] -join '/'
            $key = Get-PathKey $folder
            if ($folders.ContainsKey($key) -and $folders[$key] -cne $folder) { throw 'Mehrere Quellordner wuerden denselben Zielordner erzeugen.' }
            $folders[$key] = $folder
        }
    }
    foreach ($key in $files.psbase.Keys) {
        if ($folders.ContainsKey($key)) { throw 'Ein PDF-Zielpfad kollidiert mit einem Zielordner.' }
    }
    return [pscustomobject]@{
        FileKeys   = $files
        FolderKeys = $folders
        Folders    = @($folders.psbase.Values | Sort-Object @{ Expression = { @($_.Split([char]47)).Count } }, @{ Expression = { $_ } })
    }
}

function Test-SameInventory {
    param([AllowEmptyCollection()][object[]]$Expected, [AllowEmptyCollection()][object[]]$Actual)
    if ($Expected.Count -ne $Actual.Count) { return $false }
    for ($i = 0; $i -lt $Expected.Count; $i++) {
        if (-not [string]::Equals([string]$Expected[$i].SourceRelativePath, [string]$Actual[$i].SourceRelativePath, [StringComparison]::OrdinalIgnoreCase)) { return $false }
        if ([long]$Expected[$i].SourceLength -ne [long]$Actual[$i].SourceLength) { return $false }
        if ([long]$Expected[$i].SourceLastWriteUtcTicks -ne [long]$Actual[$i].SourceLastWriteUtcTicks) { return $false }
    }
    return $true
}

function Release-Com {
    param([AllowNull()][object]$Object)
    if ($null -ne $Object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) }
}

function Assert-ValidPdf {
    param([string]$Path)
    if (-not [IO.File]::Exists($Path) -or (Get-Item -LiteralPath $Path).Length -lt 5) { throw 'Eine PDF-Datei wurde nicht korrekt erzeugt.' }
    $stream = [IO.File]::OpenRead($Path)
    try {
        $bytes = New-Object byte[] 5
        if ($stream.Read($bytes, 0, 5) -ne 5 -or [Text.Encoding]::ASCII.GetString($bytes) -ne '%PDF-') {
            throw 'Eine erzeugte Datei ist keine gueltige PDF-Datei.'
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Convert-ToPdf {
    param([AllowEmptyCollection()][object[]]$Inventory, [string]$RunPath)
    if ($Inventory.Count -eq 0) { return @() }
    $workPath = Join-Path $RunPath 'work'
    $pdfPath = Join-Path $RunPath 'pdf'
    [void][IO.Directory]::CreateDirectory($workPath)
    [void][IO.Directory]::CreateDirectory($pdfPath)

    $visio = $null
    $artifacts = @()
    try {
        try { $visio = New-Object -ComObject Visio.Application }
        catch { throw 'Microsoft Visio konnte nicht als COM-Anwendung gestartet werden.' }
        $visio.Visible = $false
        $visio.AlertResponse = 7

        foreach ($item in $Inventory) {
            $copyPath = Join-Path $workPath ([guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension([string]$item.SourcePath))
            $outPath = Join-Path $pdfPath ([guid]::NewGuid().ToString('N') + '.pdf')
            [IO.File]::Copy([string]$item.SourcePath, $copyPath, $false)

            $document = $null
            try {
                # Read-only, keine Dateiliste, verborgen, Makros aus, kein Workspace.
                $document = $visio.Documents.OpenEx($copyPath, (2 -bor 8 -bor 64 -bor 128 -bor 256))
                # PDF, Druckqualitaet, alle Vordergrundseiten; Visio-Defaults inkl. Hintergrund.
                $document.ExportAsFixedFormat(1, $outPath, 1, 0)
            }
            catch {
                throw ("Visio-Datei konnte nicht konvertiert werden: {0}" -f [string]$item.SourceRelativePath)
            }
            finally {
                if ($null -ne $document) {
                    try { $document.Close() } catch {}
                    Release-Com $document
                }
            }
            Assert-ValidPdf $outPath
            $artifacts += [pscustomobject][ordered]@{ TargetRelativePath = [string]$item.TargetRelativePath; PdfPath = $outPath }
        }
    }
    finally {
        if ($null -ne $visio) {
            try { $visio.Quit() } catch {}
            Release-Com $visio
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    return @($artifacts)
}

function Get-HttpStatus {
    param([Management.Automation.ErrorRecord]$ErrorRecord)
    if ((Has-Property $ErrorRecord.Exception 'Response') -and
        (Has-Property $ErrorRecord.Exception.Response 'StatusCode')) {
        return [int]$ErrorRecord.Exception.Response.StatusCode
    }
    return $null
}

function Invoke-SpRequest {
    param(
        [string]$Uri, [ValidateSet('Get', 'Post')][string]$Method = 'Get',
        [string]$SiteUrl, [AllowNull()][object]$Body,
        [string]$ContentType = 'application/json;odata=verbose', [switch]$AllowNotFound
    )
    $request = @{
        Method = $Method; Uri = $Uri; Headers = $script:SpHeaders.Clone()
        UseDefaultCredentials = $true; TimeoutSec = $script:SpTimeoutSeconds
    }
    try {
        # contextinfo selbst braucht keinen Digest. Schreibaufrufe geben SiteUrl an.
        if ($Method -eq 'Post' -and $SiteUrl) { $request.Headers['X-RequestDigest'] = Get-SpDigest $SiteUrl }
        if ($null -ne $Body) { $request.Body = $Body; $request.ContentType = $ContentType }
        return Invoke-RestMethod @request
    }
    catch {
        if ($Method -eq 'Get' -and $AllowNotFound -and (Get-HttpStatus $_) -eq 404) { return $null }
        # Urspruengliche HTTP-/Verbindungsursache fuer die Diagnose erhalten.
        throw
    }
}

function Get-SpDigest {
    param([string]$SiteUrl)
    $response = Invoke-SpRequest -Method Post -Uri (New-AbsoluteUri ($SiteUrl + '/_api/contextinfo'))
    return [string]$response.d.GetContextWebInformation.FormDigestValue
}

function Get-SpFolder {
    param([string]$SiteUrl, [string]$ServerRelativeUrl, [switch]$AllowNotFound)
    $literal = ConvertTo-ODataLiteral $ServerRelativeUrl
    $uri = New-AbsoluteUri ($SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$literal')?`$select=Name,ServerRelativeUrl,UniqueId")
    return Invoke-SpRequest -Uri $uri -AllowNotFound:$AllowNotFound
}

function Get-SpContext {
    param([object]$Config)
    $libraryLiteral = ConvertTo-ODataLiteral ([string]$Config.LibraryName)
    $uri = New-AbsoluteUri ($Config.SharePointSiteUrl + "/_api/web/lists/getbytitle('$libraryLiteral')?`$select=ForceCheckout,RootFolder/ServerRelativeUrl&`$expand=RootFolder")
    $library = Invoke-SpRequest $uri
    if ($library.d.ForceCheckout -eq $true) { throw 'Die Zielbibliothek verlangt Auschecken; das Skript setzt ForceCheckout=false voraus.' }

    $libraryRoot = [string]$library.d.RootFolder.ServerRelativeUrl
    if ([string]::IsNullOrWhiteSpace($libraryRoot)) { throw 'SharePoint lieferte keinen Bibliothekswurzelpfad.' }
    $context = [pscustomobject][ordered]@{
        SiteUrl       = [string]$Config.SharePointSiteUrl
        TargetRootUrl = Join-SpUrl $libraryRoot ([string]$Config.TargetFolderPath)
        TargetRootId  = [guid]$Config.TargetFolderUniqueId
    }
    Assert-SpTarget $context
    return $context
}

function Assert-SpTarget {
    param([object]$Context)
    # Vor kritischen Schreib-/Recycle-Aktionen erneut pruefen: Pfad und GUID muessen
    # noch auf denselben dedizierten Zielordner zeigen.
    $folder = Get-SpFolder -SiteUrl $Context.SiteUrl -ServerRelativeUrl $Context.TargetRootUrl
    $actualId = [guid]::Empty
    if ($null -eq $folder -or -not [guid]::TryParse([string]$folder.d.UniqueId, [ref]$actualId) -or
        $actualId -ne [guid]$Context.TargetRootId -or
        -not [string]::Equals([string]$folder.d.ServerRelativeUrl, [string]$Context.TargetRootUrl, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Der SharePoint-Zielordner stimmt nicht mit Pfad und TargetFolderUniqueId ueberein.'
    }
}

function Assert-SpPathLengths {
    param([object]$Context, [AllowEmptyCollection()][object[]]$Inventory, [AllowEmptyCollection()][string[]]$Folders)
    if (([string]$Context.TargetRootUrl).Length -gt 260) { throw 'Der SharePoint-Zielordnerpfad ueberschreitet 260 Zeichen.' }
    foreach ($folder in $Folders) {
        if ((Join-SpUrl $Context.TargetRootUrl $folder).Length -gt 260) { throw 'Mindestens ein SharePoint-Zielordnerpfad ueberschreitet 260 Zeichen.' }
    }
    foreach ($item in $Inventory) {
        if ((Join-SpUrl $Context.TargetRootUrl ([string]$item.TargetRelativePath)).Length -gt 260) { throw 'Mindestens ein SharePoint-PDF-Zielpfad ueberschreitet 260 Zeichen.' }
    }
}

function Ensure-SpFolders {
    param([object]$Context, [AllowEmptyCollection()][string[]]$Folders)
    foreach ($folder in $Folders) {
        $serverUrl = Join-SpUrl $Context.TargetRootUrl $folder
        if ($null -ne (Get-SpFolder -SiteUrl $Context.SiteUrl -ServerRelativeUrl $serverUrl -AllowNotFound)) { continue }
        Assert-SpTarget $Context
        $body = ([ordered]@{ '__metadata' = [ordered]@{ type = 'SP.Folder' }; ServerRelativeUrl = $serverUrl } | ConvertTo-Json -Depth 4)
        [void](Invoke-SpRequest -Method Post -SiteUrl $Context.SiteUrl -Uri ($Context.SiteUrl + '/_api/web/folders') -Body $body)
        if ($null -eq (Get-SpFolder -SiteUrl $Context.SiteUrl -ServerRelativeUrl $serverUrl -AllowNotFound)) { throw 'Ein SharePoint-Zielordner wurde nach der Anlage nicht gefunden.' }
    }
}

function Send-SpPdfs {
    param([object]$Context, [AllowEmptyCollection()][object[]]$Artifacts)
    foreach ($artifact in $Artifacts) {
        $segments = @(([string]$artifact.TargetRelativePath).Split([char]47))
        $fileName = $segments[$segments.Count - 1]
        $parentRel = if ($segments.Count -gt 1) { $segments[0..($segments.Count - 2)] -join '/' } else { '' }
        $parentUrl = Join-SpUrl $Context.TargetRootUrl $parentRel
        $uri = New-AbsoluteUri ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$(ConvertTo-ODataLiteral $parentUrl)')/Files/add(url='$(ConvertTo-ODataLiteral $fileName)',overwrite=true)")
        Assert-SpTarget $Context
        [void](Invoke-SpRequest -Method Post -SiteUrl $Context.SiteUrl -Uri $uri -Body ([IO.File]::ReadAllBytes([string]$artifact.PdfPath)) -ContentType 'application/octet-stream')
    }
}

function Assert-SpApiUri {
    param([string]$SiteUrl, [string]$Uri)
    $site = [uri]$SiteUrl
    $candidate = [uri]$Uri
    if ($candidate.Scheme -cne 'https' -or
        -not [string]::Equals($candidate.Host, $site.Host, [StringComparison]::OrdinalIgnoreCase) -or
        $candidate.Port -ne $site.Port -or
        $candidate.AbsolutePath.IndexOf('/_api/', [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw 'SharePoint lieferte einen unsicheren Paging-Link.'
    }
}

function Get-SpPages {
    param([string]$SiteUrl, [string]$InitialUri)
    $items = @()
    $next = $InitialUri
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        Assert-SpApiUri -SiteUrl $SiteUrl -Uri $next
        $response = Invoke-SpRequest $next
        $page = @($response.d.results)
        $items += $page
        $next = if (Has-Property $response.d '__next') { [string]$response.d.__next } else { $null }
        if ([string]::IsNullOrWhiteSpace($next) -and $page.Count -ge 5000) { throw 'SharePoint lieferte eine abgeschnittene Inventurseite ohne Fortsetzungslink.' }
    }
    return @($items)
}

function Get-SpTargetInventory {
    param([object]$Context)
    $files = @()
    $folders = @()
    $stack = New-Object 'Collections.Generic.Stack[object]'
    $stack.Push([pscustomobject]@{ ServerRelativeUrl = $Context.TargetRootUrl; RelativePath = '' })

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        $literal = ConvertTo-ODataLiteral ([string]$current.ServerRelativeUrl)
        $fileUri = New-AbsoluteUri ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$literal')/Files?`$select=Name,ServerRelativeUrl&`$top=5000")
        foreach ($file in @(Get-SpPages -SiteUrl $Context.SiteUrl -InitialUri $fileUri)) {
            $rel = if ([string]::IsNullOrWhiteSpace([string]$current.RelativePath)) { [string]$file.Name } else { [string]$current.RelativePath + '/' + [string]$file.Name }
            $files += [pscustomobject]@{ RelativePath = $rel; ServerRelativeUrl = [string]$file.ServerRelativeUrl }
        }

        $folderUri = New-AbsoluteUri ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$literal')/Folders?`$select=Name,ServerRelativeUrl&`$top=5000")
        foreach ($folder in @(Get-SpPages -SiteUrl $Context.SiteUrl -InitialUri $folderUri)) {
            $rel = if ([string]::IsNullOrWhiteSpace([string]$current.RelativePath)) { [string]$folder.Name } else { [string]$current.RelativePath + '/' + [string]$folder.Name }
            $folders += [pscustomobject]@{ RelativePath = $rel; ServerRelativeUrl = [string]$folder.ServerRelativeUrl }
            $stack.Push([pscustomobject]@{ ServerRelativeUrl = [string]$folder.ServerRelativeUrl; RelativePath = $rel })
        }
    }
    return [pscustomobject][ordered]@{ Files = @($files); Folders = @($folders) }
}

function Assert-SpRemoteInventory {
    param([object]$Context, [object]$Remote, [object]$Plan)
    $rootPrefix = ([string]$Context.TargetRootUrl).TrimEnd([char]47) + '/'
    foreach ($remoteItem in @($Remote.Files) + @($Remote.Folders)) {
        $serverPath = [string]$remoteItem.ServerRelativeUrl
        if (-not $serverPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $serverPath.Length -gt 260) {
            throw 'Die SharePoint-Inventur enthielt einen Pfad ausserhalb des Zielordners.'
        }
        Assert-SafeRelativePath ([string]$remoteItem.RelativePath) 'Die SharePoint-Inventur enthaelt einen nicht sicher adressierbaren Namen.'
    }

    $remoteFiles = New-PathSet -Paths @($Remote.Files | ForEach-Object { $_.RelativePath })
    foreach ($key in $Plan.FileKeys.psbase.Keys) {
        if (-not $remoteFiles.ContainsKey($key)) { throw 'Nach dem Upload fehlt mindestens eine erwartete PDF-Datei.' }
    }

    $remoteFolders = New-PathSet -Paths @($Remote.Folders | ForEach-Object { $_.RelativePath })
    foreach ($key in $Plan.FolderKeys.psbase.Keys) {
        if (-not $remoteFolders.ContainsKey($key)) { throw 'Nach dem Upload fehlt mindestens ein erwarteter Zielordner.' }
    }
}

function Invoke-SpRecycleExtras {
    param([object]$Context, [object]$Remote, [object]$Plan)

    # Dedizierter Zielordner: Fremdinhalte und bei leerer Quelle alle Inhalte darunter
    # werden recycelt. Der Zielordner selbst ist nicht Teil dieser Inventur.
    $extraFiles = @($Remote.Files | Where-Object { -not $Plan.FileKeys.ContainsKey((Get-PathKey ([string]$_.RelativePath))) } |
        Sort-Object @{ Expression = { @(([string]$_.RelativePath).Split([char]47)).Count }; Descending = $true }, RelativePath)
    foreach ($file in $extraFiles) {
        $uri = New-AbsoluteUri ($Context.SiteUrl + "/_api/web/GetFileByServerRelativeUrl('$(ConvertTo-ODataLiteral ([string]$file.ServerRelativeUrl))')/recycle()")
        Assert-SpTarget $Context
        [void](Invoke-SpRequest -Method Post -SiteUrl $Context.SiteUrl -Uri $uri)
    }

    $extraFolders = @($Remote.Folders | Where-Object { -not $Plan.FolderKeys.ContainsKey((Get-PathKey ([string]$_.RelativePath))) } |
        Sort-Object @{ Expression = { @(([string]$_.RelativePath).Split([char]47)).Count }; Descending = $true }, RelativePath)
    foreach ($folder in $extraFolders) {
        $uri = New-AbsoluteUri ($Context.SiteUrl + "/_api/web/GetFolderByServerRelativeUrl('$(ConvertTo-ODataLiteral ([string]$folder.ServerRelativeUrl))')/recycle()")
        Assert-SpTarget $Context
        [void](Invoke-SpRequest -Method Post -SiteUrl $Context.SiteUrl -Uri $uri)
    }
    return [pscustomobject]@{ RecycledFiles = $extraFiles.Count; RecycledFolders = $extraFolders.Count }
}

function New-MutexName {
    param([object]$Config)
    # Wie in Python: GUID bindet die Sperre ans Ziel, auch bei URL-/Pfad-Aliasen.
    return 'Global\PPSI_VisioSharePointMirror_' + ([guid]$Config.TargetFolderUniqueId).ToString('N').ToUpperInvariant()
}

$config = $null
$mutex = $null
$mutexAcquired = $false
$runPath = $null
$exitCode = 1

try {
    $config = Read-Configuration $ConfigurationPath
    Initialize-Log $config.LogPath
    Write-Log INFO 'Spiegel-Lauf gestartet.'

    $mutex = New-Object Threading.Mutex($false, (New-MutexName $config))
    try { $mutexAcquired = $mutex.WaitOne(0, $false) }
    catch [Threading.AbandonedMutexException] { $mutexAcquired = $true }
    if (-not $mutexAcquired) { throw 'Fuer dieses SharePoint-Ziel laeuft bereits ein Spiegel-Lauf.' }

    # 1. Vorab alles pruefen, was SharePoint-Aenderungen verhindern muss.
    $source = @(Get-SourceInventory $config.SourcePath)
    $plan = New-MirrorPlan $source
    Write-Log INFO ("{0} Visio-Datei(en) gefunden." -f $source.Count)

    $sp = Get-SpContext $config
    Assert-SpPathLengths -Context $sp -Inventory $source -Folders $plan.Folders
    Write-Log INFO 'SharePoint-Ziel und Windows-Anmeldung wurden bestaetigt.'

    # 2. Erst vollstaendig lokal konvertieren, dann SharePoint anfassen.
    $runPath = Join-Path ([IO.Path]::GetTempPath()) ('PPSI-VisioSharePointMirror-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($runPath)
    $artifacts = @(Convert-ToPdf -Inventory $source -RunPath $runPath)
    Write-Log INFO ("{0} PDF-Datei(en) erzeugt." -f $artifacts.Count)

    if (-not (Test-SameInventory -Expected $source -Actual @(Get-SourceInventory $config.SourcePath))) {
        throw 'Die Quelle hat sich waehrend der Konvertierung geaendert; SharePoint blieb unveraendert.'
    }

    # 3. Nur nach erfolgreicher Konvertierung Ordner anlegen und PDFs hochladen.
    Assert-SpTarget $sp
    Ensure-SpFolders -Context $sp -Folders $plan.Folders
    Send-SpPdfs -Context $sp -Artifacts $artifacts
    Write-Log INFO ("{0} PDF-Datei(en) nach SharePoint hochgeladen." -f $artifacts.Count)

    # 4. Nur nach erfolgreichem Upload Ziel inventarisieren und Ueberzaehliges recyceln.
    if (-not (Test-SameInventory -Expected $source -Actual @(Get-SourceInventory $config.SourcePath))) {
        throw 'Die Quelle hat sich waehrend des Uploads geaendert; Papierkorbaktionen wurden unterdrueckt.'
    }
    Assert-SpTarget $sp
    $remote = Get-SpTargetInventory $sp
    Assert-SpTarget $sp
    Assert-SpRemoteInventory -Context $sp -Remote $remote -Plan $plan
    if (-not (Test-SameInventory -Expected $source -Actual @(Get-SourceInventory $config.SourcePath))) {
        throw 'Die Quelle hat sich waehrend der SharePoint-Inventur geaendert; Papierkorbaktionen wurden unterdrueckt.'
    }
    Assert-SpTarget $sp
    $result = Invoke-SpRecycleExtras -Context $sp -Remote $remote -Plan $plan
    Write-Log INFO ("Spiegelabgleich abgeschlossen: {0} Datei(en) und {1} Ordner recycelt." -f $result.RecycledFiles, $result.RecycledFolders)
    $exitCode = 0
}
catch {
    $message = if ([string]::IsNullOrWhiteSpace($_.Exception.Message)) { 'Unbekannter Fehler im Spiegel-Lauf.' } else { $_.Exception.Message }
    try { Write-Log ERROR $message } catch { [Console]::Error.WriteLine($message) }
    $exitCode = 1
}
finally {
    if ($runPath -and [IO.Directory]::Exists($runPath)) {
        try { [IO.Directory]::Delete($runPath, $true) }
        catch {
            $exitCode = 1
            try { Write-Log ERROR 'Das temporaere Laufverzeichnis konnte nicht vollstaendig entfernt werden.' } catch {}
        }
    }
    if ($mutexAcquired -and $mutex) { try { $mutex.ReleaseMutex() } catch {} }
    if ($mutex) { $mutex.Dispose() }
}

exit $exitCode
