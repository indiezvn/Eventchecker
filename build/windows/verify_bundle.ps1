param(
    [Parameter(Mandatory = $true)]
    [string]$BundleRoot,
    [string]$SourcePath = "Log_checker.py",
    [string]$ExpectedSeries = "2.5.0"
)

$ErrorActionPreference = "Stop"

function Get-ReleaseMarker([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Release marker file is missing: $Path"
    }

    $text = Get-Content -Raw -LiteralPath $Path
    $markers = @(
        [regex]::Matches($text, 'v(?<series>\d+\.\d+\.\d+)\((?<build>\d+)\)') |
            ForEach-Object { $_.Value } |
            Sort-Object -Unique
    )
    if ($markers.Count -ne 1) {
        throw "Inconsistent version markers in $Path`: $($markers -join ', ')"
    }
    return [string]$markers[0]
}

function Resolve-BundleFile([string]$Root, [string]$RelativePath) {
    # PyInstaller 6 stores onedir support files under _internal; older builds
    # placed the same files directly beside the executable.
    $candidates = @(
        (Join-Path $Root $RelativePath),
        (Join-Path (Join-Path $Root "_internal") $RelativePath)
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    throw "Bundle file is missing: $RelativePath (checked root and _internal)"
}

$sourceMarker = Get-ReleaseMarker $SourcePath
$bundleSourcePath = Resolve-BundleFile $BundleRoot $SourcePath
$bundleUpdaterPath = Resolve-BundleFile $BundleRoot "remote_update.py"
$bundleMarker = Get-ReleaseMarker $bundleSourcePath
$expectedPattern = "^v$([regex]::Escape($ExpectedSeries))\(\d+\)$"

if ($sourceMarker -notmatch $expectedPattern) {
    throw "Source marker is not a $ExpectedSeries release: $sourceMarker"
}
if ($bundleMarker -ne $sourceMarker) {
    throw "Built bundle is stale: source=$sourceMarker bundle=$bundleMarker"
}
if (-not (Test-Path -LiteralPath (Join-Path $BundleRoot "EventInspector.exe") -PathType Leaf)) {
    throw "Built bundle is missing EventInspector.exe"
}

$bundletoolPath = Resolve-BundleFile $BundleRoot "services_checker\bundletool-all-1.18.1.jar"
$keystorePath = Resolve-BundleFile $BundleRoot "services_checker\my-key.keystore"
$bundletoolLength = (Get-Item -LiteralPath $bundletoolPath).Length
$keystoreLength = (Get-Item -LiteralPath $keystorePath).Length
if ($bundletoolLength -lt 1000000) {
    throw "Bundled bundletool is unexpectedly small: $bundletoolLength bytes"
}
if ($keystoreLength -lt 1) {
    throw "Bundled keystore is empty: $keystorePath"
}

Write-Host "Bundle validated: $bundleMarker"
