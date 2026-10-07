param([switch]$NoPause)

$ErrorActionPreference = 'Stop'
trap {
    Write-Host ''
    Write-Host 'The content index could not be refreshed.' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ''
    if (-not $NoPause) { Read-Host 'Press Enter to close' }
    exit 1
}

$tools = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $tools
$assets = Join-Path $root 'assets'
$configPath = Join-Path $assets 'site-config.json'
$manifest = Join-Path $assets 'content-manifest.json'
$config = if (Test-Path -LiteralPath $configPath) { Get-Content -Raw -LiteralPath $configPath | ConvertFrom-Json } else { $null }
$excludedFiles = @('README.md', 'LICENSE.md')
$excludedFolders = @('.git', '.github', '.obsidian', 'assets', 'tools')
if ($config.excludedFiles) { $excludedFiles = @($config.excludedFiles) }
if ($config.excludedFolders) { $excludedFolders = @($config.excludedFolders) }

$rootUri = [Uri]($root.TrimEnd('\') + '\')
$items = Get-ChildItem -LiteralPath $root -Recurse -File |
    Where-Object {
        if ($_.Extension -notin @('.md', '.pdf')) { return $false }
        if ($_.Name -in $excludedFiles) { return $false }
        $relativeNative = $_.FullName.Substring($root.Length).TrimStart('\')
        $segments = $relativeNative -split '[\\/]'
        return -not ($segments | Where-Object { $_ -in $excludedFolders })
    } |
    ForEach-Object {
        $relative = [Uri]::UnescapeDataString($rootUri.MakeRelativeUri([Uri]$_.FullName).ToString())
        $type = if ($_.Extension -eq '.pdf') { 'pdf' } else { 'markdown' }
        $first = if ($type -eq 'markdown') { Get-Content -LiteralPath $_.FullName -TotalCount 1 } else { '' }
        $title = if ($first -match '^#\s+(.+)$') { $Matches[1].Trim() } else { $_.BaseName }
        $folder = if ($relative.Contains('/')) { $relative.Substring(0, $relative.LastIndexOf('/')) } else { '' }
        [ordered]@{ title = $title; path = $relative; folder = $folder; type = $type }
    } | Sort-Object path

if (-not (Test-Path -LiteralPath $assets -PathType Container)) { New-Item -ItemType Directory -Path $assets | Out-Null }
$json = ConvertTo-Json -InputObject @($items) -Depth 4
[IO.File]::WriteAllText($manifest, $json, [Text.UTF8Encoding]::new($false))
Write-Host "Content index refreshed: $(@($items).Count) readable documents."
Write-Host "Manifest: $manifest"
if ($config.homeDocument -and -not (@($items).path -contains $config.homeDocument)) {
    Write-Warning "The configured home document '$($config.homeDocument)' is not in the index."
}
if (-not $NoPause) { Write-Host ''; Read-Host 'Press Enter to close' }
