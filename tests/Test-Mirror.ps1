#requires -Version 5.1
#requires -PSEdition Desktop
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Load definitions only: never execute the production entry point, COM or a real request.
$productionPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'production\Invoke-VisioSharePointMirror.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($productionPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join [Environment]::NewLine) }
foreach ($definition in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    . ([scriptblock]::Create($definition.Extent.Text))
}

$script:SpHeaders = @{ Accept = 'application/json;odata=verbose'; 'X-FORMS_BASED_AUTH_ACCEPTED' = 'f' }
$script:SpTimeoutSeconds = 300
$script:PassedCount = 0
$script:FailedCount = 0
$script:RestCalls = @()
$script:RestResponder = { throw 'Unexpected REST request.' }
$script:SiteUrl = 'https://sharepoint.example.invalid/sites/test'
$script:TargetRootUrl = '/sites/test/Documents/Mirror'
$script:TargetId = [guid]'11111111-1111-4111-8111-111111111111'
$script:Context = [pscustomobject]@{ SiteUrl = $script:SiteUrl; TargetRootUrl = $script:TargetRootUrl; TargetRootId = $script:TargetId }
$script:ReturnedTargetId = $script:TargetId
$script:ReturnedTargetUrl = $script:TargetRootUrl

# The exception exposes the same Response.StatusCode shape without opening a socket.
Add-Type -TypeDefinition @'
public sealed class MirrorTestHttpException : System.Exception {
    public MirrorTestHttpResponse Response { get; private set; }
    public MirrorTestHttpException(int status) : base("Mock HTTP failure") {
        Response = new MirrorTestHttpResponse { StatusCode = status };
    }
}
public sealed class MirrorTestHttpResponse { public int StatusCode { get; set; } }
'@

function Assert-MirrorTrue {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-MirrorEqual {
    param([AllowNull()][object]$Expected, [AllowNull()][object]$Actual, [string]$Message)
    if ($Expected -cne $Actual) { throw ('{0} Expected [{1}], received [{2}].' -f $Message, $Expected, $Actual) }
}

function Assert-MirrorThrows {
    param([scriptblock]$Body, [string]$Message)
    $caught = $false
    try { & $Body | Out-Null } catch { $caught = $true }
    Assert-MirrorTrue $caught $Message
}

function Invoke-MirrorTest {
    param([string]$Name, [scriptblock]$Body)
    $script:RestCalls = @()
    $script:RestResponder = { throw 'Unexpected REST request.' }
    $script:ReturnedTargetId = $script:TargetId
    $script:ReturnedTargetUrl = $script:TargetRootUrl
    try {
        & $Body
        $script:PassedCount++
        Write-Output "PASS: $Name"
    }
    catch {
        $script:FailedCount++
        Write-Output "FAIL: $Name"
        Write-Output ('  {0}' -f $_.Exception.Message)
    }
}

function Invoke-RestMethod {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri, [hashtable]$Headers, [AllowNull()][object]$Body,
        [string]$ContentType, [switch]$UseDefaultCredentials, [int]$TimeoutSec)
    $call = [pscustomobject]@{
        Method = $Method; Uri = $Uri; Headers = $Headers.Clone(); Body = $Body
        ContentType = $ContentType; UseDefaultCredentials = [bool]$UseDefaultCredentials; TimeoutSec = $TimeoutSec
    }
    $script:RestCalls += $call
    return (& $script:RestResponder $call)
}

function Get-MirrorMockResponse {
    param([object]$Call)
    if ($Call.Uri.EndsWith('/_api/contextinfo')) {
        return [pscustomobject]@{ d = [pscustomobject]@{ GetContextWebInformation = [pscustomobject]@{ FormDigestValue = 'test-digest' } } }
    }
    if ($Call.Method -eq 'Get' -and $Call.Uri -match '/lists/getbytitle\(') {
        return [pscustomobject]@{ d = [pscustomobject]@{ ForceCheckout = $false; RootFolder = [pscustomobject]@{ ServerRelativeUrl = '/sites/test/Documents' } } }
    }
    if ($Call.Method -eq 'Get' -and $Call.Uri -match '/GetFolderByServerRelativeUrl\(') {
        return [pscustomobject]@{ d = [pscustomobject]@{ UniqueId = [string]$script:ReturnedTargetId; ServerRelativeUrl = $script:ReturnedTargetUrl } }
    }
    if ($Call.Method -eq 'Post' -and $Call.Uri.EndsWith('/recycle()')) { return 'recycled' }
    throw ('Unexpected REST request: {0} {1}' -f $Call.Method, $Call.Uri)
}

function New-MirrorRemoteItem {
    param([string]$Path)
    return [pscustomobject]@{ RelativePath = $Path; ServerRelativeUrl = $script:TargetRootUrl + '/' + $Path }
}

Invoke-MirrorTest 'GET and contextinfo POST do not fetch a digest' {
    $script:RestResponder = { param($call) return [pscustomobject]@{ Value = 'ok' } }
    $result = Invoke-SpRequest -Uri ($script:SiteUrl + '/_api/web') -SiteUrl $script:SiteUrl
    Assert-MirrorEqual 'ok' $result.Value 'GET response was changed.'
    Assert-MirrorEqual 1 $script:RestCalls.Count 'GET made an extra request.'
    Assert-MirrorTrue (-not $script:RestCalls[0].Headers.ContainsKey('X-RequestDigest')) 'GET received a digest.'
    $script:RestResponder = { param($call) Get-MirrorMockResponse $call }
    Assert-MirrorEqual 'test-digest' (Get-SpDigest $script:SiteUrl) 'Digest response was not read.'
    Assert-MirrorEqual 2 $script:RestCalls.Count 'Digest request recursed.'
    Assert-MirrorEqual 'Post' $script:RestCalls[1].Method 'contextinfo must be POST.'
    Assert-MirrorTrue (-not $script:RestCalls[1].Headers.ContainsKey('X-RequestDigest')) 'contextinfo received a digest.'
}

Invoke-MirrorTest 'Write POST keeps the binary body and uses the fresh digest' {
    $script:RestResponder = {
        param($call)
        if ($call.Uri.EndsWith('/_api/contextinfo')) { return Get-MirrorMockResponse $call }
        return [pscustomobject]@{ Uploaded = $true }
    }
    $bytes = [byte[]]@(0, 1, 127, 128, 255)
    $result = Invoke-SpRequest -Method Post -SiteUrl $script:SiteUrl -Uri ($script:SiteUrl + '/_api/upload') -Body $bytes -ContentType 'application/octet-stream'
    Assert-MirrorTrue $result.Uploaded 'Write response was lost.'
    Assert-MirrorEqual 2 $script:RestCalls.Count 'Write must first fetch its digest.'
    Assert-MirrorTrue $script:RestCalls[0].Uri.EndsWith('/_api/contextinfo') 'Digest was not fetched before writing.'
    $write = $script:RestCalls[1]
    Assert-MirrorEqual 'test-digest' $write.Headers['X-RequestDigest'] 'Write digest missing.'
    Assert-MirrorTrue ($write.Body -is [byte[]]) 'Binary body changed type.'
    Assert-MirrorEqual '0,1,127,128,255' ($write.Body -join ',') 'Binary body changed bytes.'
    Assert-MirrorEqual 'application/octet-stream' $write.ContentType 'Upload content type changed.'
    foreach ($call in $script:RestCalls) {
        Assert-MirrorTrue $call.UseDefaultCredentials 'Windows credentials were not requested.'
        Assert-MirrorEqual 300 $call.TimeoutSec 'Configured timeout was not used.'
        Assert-MirrorEqual 'application/json;odata=verbose' $call.Headers.Accept 'Accept header missing.'
        Assert-MirrorEqual 'f' $call.Headers['X-FORMS_BASED_AUTH_ACCEPTED'] 'Windows auth header missing.'
    }
    Assert-MirrorTrue (-not $script:SpHeaders.ContainsKey('X-RequestDigest')) 'Shared headers were mutated.'
}

Invoke-MirrorTest 'Only an explicitly allowed HTTP 404 is suppressed' {
    $script:RestResponder = { throw (New-Object MirrorTestHttpException(404)) }
    Assert-MirrorTrue ($null -eq (Invoke-SpRequest -Uri ($script:SiteUrl + '/_api/missing') -AllowNotFound)) 'Allowed GET 404 was not empty.'
    Assert-MirrorThrows { Invoke-SpRequest -Uri ($script:SiteUrl + '/_api/missing') } 'Unexpected GET 404 was suppressed.'
    $script:RestResponder = { throw (New-Object MirrorTestHttpException(403)) }
    Assert-MirrorThrows { Invoke-SpRequest -Uri ($script:SiteUrl + '/_api/forbidden') -AllowNotFound } 'HTTP 403 was suppressed.'
    $script:RestResponder = { throw (New-Object System.InvalidOperationException('Mock non-HTTP error')) }
    Assert-MirrorThrows { Invoke-SpRequest -Uri ($script:SiteUrl + '/_api/failure') -AllowNotFound } 'Non-HTTP failure was suppressed.'
    $plainError = New-Object Management.Automation.ErrorRecord((New-Object System.InvalidOperationException('Mock error')), 'Mock', ([Management.Automation.ErrorCategory]::ConnectionError), $null)
    Assert-MirrorTrue ($null -eq (Get-HttpStatus $plainError)) 'An exception without Response must safely return no HTTP status.'
}

Invoke-MirrorTest 'Target acceptance requires the configured GUID and path' {
    $script:RestResponder = { param($call) Get-MirrorMockResponse $call }
    $config = [pscustomobject]@{ SharePointSiteUrl = $script:SiteUrl; LibraryName = 'Documents'; TargetFolderPath = 'Mirror'; TargetFolderUniqueId = $script:TargetId }
    $context = Get-SpContext $config
    Assert-MirrorEqual $script:TargetId $context.TargetRootId 'Valid target GUID was not retained.'
    Assert-MirrorEqual $script:TargetRootUrl $context.TargetRootUrl 'Valid target path was not retained.'
    $script:ReturnedTargetId = [guid]'22222222-2222-4222-8222-222222222222'
    Assert-MirrorThrows { Get-SpContext $config } 'A different target GUID was accepted.'
    $script:ReturnedTargetId = $script:TargetId
    $script:ReturnedTargetUrl = '/sites/test/Documents/Other'
    Assert-MirrorThrows { Get-SpContext $config } 'A different target path was accepted.'
    Assert-MirrorEqual 0 @($script:RestCalls | Where-Object Method -eq 'Post').Count 'Target validation made a write.'
}

Invoke-MirrorTest 'Configuration rejects root sources and unsafe target paths' {
    $fixturePath = Join-Path ([IO.Path]::GetTempPath()) ('mirror-config-test-' + [guid]::NewGuid().ToString('N') + '.json')
    $encoding = New-Object Text.UTF8Encoding($false)
    $config = [ordered]@{
        SourcePath = 'C:\MirrorTestSource'; SharePointSiteUrl = $script:SiteUrl; LibraryName = 'Documents'
        TargetFolderPath = 'Mirror'; TargetFolderUniqueId = [string]$script:TargetId; LogPath = 'C:\MirrorTestLogs\mirror.log'
    }
    try {
        [IO.File]::WriteAllText($fixturePath, ($config | ConvertTo-Json), $encoding)
        Assert-MirrorEqual 'Mirror' (Read-Configuration $fixturePath).TargetFolderPath 'Valid fixture was rejected.'
        foreach ($path in @('', '.', '/', '../Other', 'Mirror/../Other', '\Mirror', 'Mirror//Child', 'Mirror#Other')) {
            $config.TargetFolderPath = $path
            [IO.File]::WriteAllText($fixturePath, ($config | ConvertTo-Json), $encoding)
            Assert-MirrorThrows { Read-Configuration $fixturePath } ('Unsafe target path was accepted: ' + $path)
        }
        $config.TargetFolderPath = 'Mirror'
        foreach ($path in @('C:\', '\\server\share')) {
            $config.SourcePath = $path
            [IO.File]::WriteAllText($fixturePath, ($config | ConvertTo-Json), $encoding)
            Assert-MirrorThrows { Read-Configuration $fixturePath } ('A source root was accepted: ' + $path)
        }
    }
    finally { [IO.File]::Delete($fixturePath) }
}

Invoke-MirrorTest 'Plan handles empty input, normalizes keys and rejects file-folder collisions' {
    $empty = New-MirrorPlan -Inventory @()
    Assert-MirrorEqual 0 $empty.FileKeys.Count 'Empty plan contains files.'
    Assert-MirrorEqual 0 $empty.FolderKeys.Count 'Empty plan contains folder keys.'
    Assert-MirrorEqual 0 @($empty.Folders).Count 'Empty plan contains folders.'
    $plan = New-MirrorPlan -Inventory @([pscustomobject]@{ TargetRelativePath = 'Quelle/A/Test.pdf' })
    Assert-MirrorTrue $plan.FileKeys.ContainsKey('quelle/a/test.pdf') 'File key is not normalized.'
    Assert-MirrorTrue $plan.FolderKeys.ContainsKey('quelle/a') 'Folder key is not normalized.'
    Assert-MirrorEqual 'Quelle|Quelle/A' ($plan.Folders -join '|') 'Parents must precede their child folders.'
    $valuesPlan = New-MirrorPlan -Inventory @([pscustomobject]@{ TargetRelativePath = 'Values/Sub/A.pdf' })
    Assert-MirrorEqual 'Values|Values/Sub' ($valuesPlan.Folders -join '|') 'Folder named Values shadowed the hashtable property.'
    Assert-MirrorThrows {
        New-MirrorPlan -Inventory @([pscustomobject]@{ TargetRelativePath = 'Quelle/A.pdf' }, [pscustomobject]@{ TargetRelativePath = 'Quelle/a.pdf/Child.pdf' })
    } 'A file-folder collision was accepted.'
}

Invoke-MirrorTest 'Remote validation requires every expected PDF and folder inside the target' {
    $plan = New-MirrorPlan -Inventory @([pscustomobject]@{ TargetRelativePath = 'Quelle/A/Current.pdf' })
    $remote = [pscustomobject]@{ Files = @(New-MirrorRemoteItem 'Quelle/A/Current.pdf'); Folders = @((New-MirrorRemoteItem 'Quelle'), (New-MirrorRemoteItem 'Quelle/A')) }
    Assert-SpRemoteInventory -Context $script:Context -Remote $remote -Plan $plan
    $remote.Files = @()
    Assert-MirrorThrows { Assert-SpRemoteInventory -Context $script:Context -Remote $remote -Plan $plan } 'Missing expected PDF was accepted.'
    $remote.Files = @(New-MirrorRemoteItem 'Quelle/A/Current.pdf')
    $remote.Folders = @(New-MirrorRemoteItem 'Quelle')
    Assert-MirrorThrows { Assert-SpRemoteInventory -Context $script:Context -Remote $remote -Plan $plan } 'Missing expected folder was accepted.'
    $remote.Folders += New-MirrorRemoteItem 'Quelle/A'
    $remote.Files += [pscustomobject]@{ RelativePath = 'External.pdf'; ServerRelativeUrl = '/sites/test/Documents/Outside/External.pdf' }
    Assert-MirrorThrows { Assert-SpRemoteInventory -Context $script:Context -Remote $remote -Plan $plan } 'An entry outside the target was accepted.'
    $remote.Files = @(New-MirrorRemoteItem 'Quelle/A/Current.pdf')
    $remote.Folders += [pscustomobject]@{ RelativePath = ''; ServerRelativeUrl = $script:TargetRootUrl }
    Assert-MirrorThrows { Assert-SpRemoteInventory -Context $script:Context -Remote $remote -Plan $plan } 'The target root was accepted as a recyclable child.'
    $keysPlan = New-MirrorPlan -Inventory @([pscustomobject]@{ TargetRelativePath = 'Keys/Sub/A.pdf' })
    $keysRemote = [pscustomobject]@{ Files = @(New-MirrorRemoteItem 'Keys/Sub/A.pdf'); Folders = @(New-MirrorRemoteItem 'Keys') }
    Assert-MirrorThrows { Assert-SpRemoteInventory -Context $script:Context -Remote $keysRemote -Plan $keysPlan } 'Folder named Keys masked the missing Keys/Sub folder.'
}

Invoke-MirrorTest 'Recycle preserves expected paths and removes deeper extra folders first' {
    $script:RestResponder = { param($call) Get-MirrorMockResponse $call }
    $plan = New-MirrorPlan -Inventory @([pscustomobject]@{ TargetRelativePath = 'Quelle/A/Current.pdf' })
    $remote = [pscustomobject]@{
        Files = @((New-MirrorRemoteItem 'Quelle/A/Current.pdf'), (New-MirrorRemoteItem 'Quelle/A/Old.pdf'), (New-MirrorRemoteItem 'Orphan/Deep/Old.pdf'))
        Folders = @((New-MirrorRemoteItem 'Quelle'), (New-MirrorRemoteItem 'Quelle/A'), (New-MirrorRemoteItem 'Orphan'), (New-MirrorRemoteItem 'Orphan/Deep'))
    }
    Assert-SpRemoteInventory -Context $script:Context -Remote $remote -Plan $plan
    $result = Invoke-SpRecycleExtras -Context $script:Context -Remote $remote -Plan $plan
    Assert-MirrorEqual 2 $result.RecycledFiles 'Wrong count of recycled files.'
    Assert-MirrorEqual 2 $result.RecycledFolders 'Wrong count of recycled folders.'
    $recycles = @($script:RestCalls | Where-Object { $_.Method -eq 'Post' -and $_.Uri.EndsWith('/recycle()') })
    Assert-MirrorEqual 4 $recycles.Count 'Wrong number of recycle requests.'
    Assert-MirrorEqual 0 @($recycles | Where-Object { $_.Uri -match '/Current\.pdf|/Quelle(?:/A)?\x27\)/recycle' }).Count 'An expected path was recycled.'
    $folderCalls = @($recycles | Where-Object { $_.Uri -match '/GetFolderByServerRelativeUrl\(' })
    Assert-MirrorTrue ($folderCalls[0].Uri -match '/Orphan/Deep\x27\)/recycle\(\)$') 'Child folder was not recycled first.'
    Assert-MirrorTrue ($folderCalls[1].Uri -match '/Orphan\x27\)/recycle\(\)$') 'Parent folder was not recycled last.'
}

Invoke-MirrorTest 'Recycle stops immediately after a failed request or changed target identity' {
    $plan = New-MirrorPlan -Inventory @()
    $remote = [pscustomobject]@{ Files = @((New-MirrorRemoteItem 'Old/A.pdf'), (New-MirrorRemoteItem 'Old/B.pdf')); Folders = @(New-MirrorRemoteItem 'Old') }
    $script:RestResponder = {
        param($call)
        if ($call.Uri.EndsWith('/recycle()')) { throw (New-Object MirrorTestHttpException(500)) }
        return Get-MirrorMockResponse $call
    }
    Assert-MirrorThrows { Invoke-SpRecycleExtras -Context $script:Context -Remote $remote -Plan $plan } 'Recycle failure was suppressed.'
    Assert-MirrorEqual 1 @($script:RestCalls | Where-Object { $_.Uri.EndsWith('/recycle()') }).Count 'Recycling continued after the failed request.'
    $script:RestCalls = @()
    $script:ReturnedTargetId = [guid]'22222222-2222-4222-8222-222222222222'
    $script:RestResponder = { param($call) Get-MirrorMockResponse $call }
    Assert-MirrorThrows { Invoke-SpRecycleExtras -Context $script:Context -Remote $remote -Plan $plan } 'Changed target identity did not stop recycling.'
    Assert-MirrorEqual 0 @($script:RestCalls | Where-Object { $_.Uri.EndsWith('/recycle()') }).Count 'Recycling started despite the changed target identity.'
}

Invoke-MirrorTest 'An empty source recycles child entries and preserves the target root' {
    $script:RestResponder = { param($call) Get-MirrorMockResponse $call }
    $plan = New-MirrorPlan -Inventory @()
    $remote = [pscustomobject]@{ Files = @(New-MirrorRemoteItem 'Old/Old.pdf'); Folders = @(New-MirrorRemoteItem 'Old') }
    Assert-SpRemoteInventory -Context $script:Context -Remote $remote -Plan $plan
    $result = Invoke-SpRecycleExtras -Context $script:Context -Remote $remote -Plan $plan
    Assert-MirrorEqual 1 $result.RecycledFiles 'Empty source did not recycle the child file.'
    Assert-MirrorEqual 1 $result.RecycledFolders 'Empty source did not recycle the child folder.'
    $recycles = @($script:RestCalls | Where-Object { $_.Method -eq 'Post' -and $_.Uri.EndsWith('/recycle()') })
    Assert-MirrorEqual 2 $recycles.Count 'Empty source produced an unexpected recycle request.'
    foreach ($call in $recycles) {
        Assert-MirrorTrue $call.Uri.Contains($script:TargetRootUrl + '/') 'Recycle request addressed the target root or an outside path.'
    }
}

Write-Output ('Mirror tests: {0} passed, {1} failed.' -f $script:PassedCount, $script:FailedCount)
if ($script:FailedCount -gt 0) { exit 1 }
exit 0
