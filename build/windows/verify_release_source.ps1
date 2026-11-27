param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path,
    [string]$ExpectedSeries = "2.5.0",
    [string]$GithubEnvPath = "",
    [string]$GithubOutputPath = "",
    [switch]$PrintVersion
)

$ErrorActionPreference = "Stop"

$sourcePath = Join-Path $ProjectRoot "Log_checker.py"
if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
    throw "Release source is missing: $sourcePath"
}

$source = Get-Content -Raw -LiteralPath $sourcePath
$markers = @(
    [regex]::Matches($source, 'v(?<series>\d+\.\d+\.\d+)\((?<build>\d+)\)') |
        ForEach-Object { $_.Value } |
        Sort-Object -Unique
)

if ($markers.Count -ne 1) {
    throw "Release source contains inconsistent version markers: $($markers -join ', ')"
}

$marker = [string]$markers[0]
$pattern = "^v$([regex]::Escape($ExpectedSeries))\((?<build>\d+)\)$"
$match = [regex]::Match($marker, $pattern)
if (-not $match.Success) {
    throw "Wrong release source: expected v$ExpectedSeries(build), found $marker"
}

$releaseVersion = "$ExpectedSeries.$($match.Groups['build'].Value)"
$markerLine = "EVENTINSPECTOR_RELEASE_MARKER=$marker"
$versionLine = "EVENTINSPECTOR_RELEASE_VERSION=$releaseVersion"

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
if ($GithubEnvPath) {
    [System.IO.File]::AppendAllText($GithubEnvPath, "$markerLine`n$versionLine`n", $utf8NoBom)
}
if ($GithubOutputPath) {
    [System.IO.File]::AppendAllText($GithubOutputPath, "release_marker=$marker`nrelease_version=$releaseVersion`n", $utf8NoBom)
}

if ($PrintVersion) {
    Write-Output $releaseVersion
    exit 0
}

Write-Host "Release source validated: $marker"
