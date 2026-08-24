# Ausschliesslich deterministische In-Memory-Adapter fuer Validate-/Simulate-Tests.
# Der Produktiveinstieg laedt diese Datei nicht.

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
