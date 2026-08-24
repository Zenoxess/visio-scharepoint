Set-StrictMode -Version 2.0

# This file documents the future side-effect boundaries. The public scaffold does
# not dot-source this file and exposes no Execute mode. Every adapter deliberately
# throws so that it cannot accidentally become a productive path.

function Invoke-VssSourceInventoryAdapter {
    <#
    .SYNOPSIS
        Future read adapter for recursively inventorying a configured source root.
    .OUTPUTS
        A future implementation should return immutable candidate records containing
        relative path, size and UTC modification time. It must also report whether the
        complete inventory succeeded; deletion logic may never run after a partial scan.
        FileNamePattern must be compiled and every match evaluated with the validated
        Source.RegexTimeoutMilliseconds value; an infinite regex timeout is forbidden.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Configuration)

    throw 'Adapter stub only: source inventory is not implemented.'
}

function Invoke-VssStableFileStagingAdapter {
    <#
    .SYNOPSIS
        Future local staging adapter for one source candidate.
    .OUTPUTS
        A future implementation should return a local immutable copy plus SHA-256 and
        must defer files whose size or timestamp changes during the stability probes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Candidate,
        [Parameter(Mandatory = $true)][object]$Configuration
    )

    throw 'Adapter stub only: stable file staging is not implemented.'
}

function Invoke-VssVisioConversionAdapter {
    <#
    .SYNOPSIS
        Future isolated Visio-to-PDF conversion adapter.
    .OUTPUTS
        A future implementation should return a staged PDF record. COM must be isolated
        in a serial STA worker with a timeout; macros and automatic refresh are policy
        inputs from the validated configuration.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$StagedSource,
        [Parameter(Mandatory = $true)][object]$Configuration
    )

    throw 'Adapter stub only: Visio conversion is not implemented.'
}

function Invoke-VssMicrosoftGraphAdapter {
    <#
    .SYNOPSIS
        Future Microsoft Graph v1 reconciliation and upload adapter.
    .OUTPUTS
        A future implementation should return the DriveItem ID, eTag and destination
        path only after a successful commit. It consumes the ID-based target fields and
        resolves authentication from the selected non-secret identity configuration.
        SiteId, DriveId and TargetFolderId are opaque values: each must be escaped as
        one URI segment (or passed to an SDK parameter) and must never be concatenated
        into an unescaped request path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$PdfArtifact,
        [Parameter(Mandatory = $true)][object]$Configuration
    )

    throw 'Adapter stub only: Microsoft Graph access is not implemented.'
}

function Invoke-VssSharePointRestAdapter {
    <#
    .SYNOPSIS
        Future SharePoint REST reconciliation and upload adapter.
    .OUTPUTS
        A future implementation should return the server-relative result only after a
        successful commit. It consumes SiteUrl, LibraryName and TargetFolderPath and
        must never infer Microsoft Graph IDs or authentication behavior.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$PdfArtifact,
        [Parameter(Mandatory = $true)][object]$Configuration
    )

    throw 'Adapter stub only: SharePoint REST access is not implemented.'
}

function Save-VssSyncStateAdapter {
    <#
    .SYNOPSIS
        Future atomic persistence adapter for successful sync state transitions.
    .NOTES
        A future implementation must write only after the remote commit succeeds and
        must use an atomic replace strategy. Validate and DryRun never call this stub.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$State,
        [Parameter(Mandatory = $true)][object]$Configuration
    )

    throw 'Adapter stub only: sync state persistence is not implemented.'
}
