param([string]$Image = 'ghcr.io/project-osrm/osrm-backend', [string]$MapFile = '')
$ErrorActionPreference = 'Stop'
$MotorRoot = $PSScriptRoot
$DataRoot = Join-Path $MotorRoot 'data'
New-Item -ItemType Directory -Force -Path $DataRoot | Out-Null
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) { throw '請先安裝並啟動 Docker Desktop（Linux containers）。' }
if ($MapFile) {
    Copy-Item -LiteralPath $MapFile -Destination (Join-Path $DataRoot 'taiwan.osm.pbf')
} elseif (-not (Test-Path (Join-Path $DataRoot 'taiwan.osm.pbf'))) {
    Write-Host '首次下載台灣 OpenStreetMap 地圖，檔案較大，需等待下載完成。'
    Invoke-WebRequest -Uri 'https://download.geofabrik.de/asia/taiwan-latest.osm.pbf' -OutFile (Join-Path $DataRoot 'taiwan.osm.pbf')
}
& docker pull $Image
if ($LASTEXITCODE -ne 0) { throw 'OSRM 映像下載失敗。' }
$PinnedImage = (& docker image inspect --format '{{index .RepoDigests 0}}' $Image).Trim()
if (-not $PinnedImage) { throw '無法取得 OSRM 映像 digest。' }
Copy-Item (Join-Path $MotorRoot 'motorcycle.lua') (Join-Path $DataRoot 'motorcycle.lua') -Force
$Volume = "${DataRoot}:/data"
foreach ($CommandArgs in @(
    @('osrm-extract','-p','/data/motorcycle.lua','/data/taiwan.osm.pbf'),
    @('osrm-partition','/data/taiwan.osrm'),
    @('osrm-customize','/data/taiwan.osrm')
)) {
    & docker run --rm -v $Volume $PinnedImage @CommandArgs
    if ($LASTEXITCODE -ne 0) { throw ('機車路網建置失敗：' + $CommandArgs[0]) }
}
@{image=$PinnedImage;profile_sha256=(Get-FileHash (Join-Path $MotorRoot 'motorcycle.lua')).Hash} |
    ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $MotorRoot 'build_info.json')
Write-Host '建置完成。下一步執行 start_motorcycle.ps1。'
