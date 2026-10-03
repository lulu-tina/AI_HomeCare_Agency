$ErrorActionPreference = 'Stop'
$MotorRoot = $PSScriptRoot
$BuildInfoPath = Join-Path $MotorRoot 'build_info.json'
if (-not (Test-Path $BuildInfoPath)) { throw '請先執行 build_motorcycle.ps1。' }
$BuildInfo = Get-Content -Raw -Encoding UTF8 $BuildInfoPath | ConvertFrom-Json
if ((Get-FileHash (Join-Path $MotorRoot 'motorcycle.lua')).Hash -ne $BuildInfo.profile_sha256) {
    throw '機車設定已變更，請重新執行 build_motorcycle.ps1。'
}
$Existing = & docker ps -a --filter 'name=^careflow-osrm-motorcycle$' --format '{{.Names}}'
if ($Existing) { throw 'careflow-osrm-motorcycle 容器已存在；請用 docker start careflow-osrm-motorcycle 啟動。重新建圖後需先停止並移除舊容器再執行本檔。' }
$DataRoot = Join-Path $MotorRoot 'data'
& docker run -d --name careflow-osrm-motorcycle -p '127.0.0.1:5001:5000' -v "${DataRoot}:/data:ro" $BuildInfo.image osrm-routed --algorithm mld /data/taiwan.osrm
if ($LASTEXITCODE -ne 0) { throw '機車服務啟動失敗，請檢查 Docker 和 port 5001。' }
Write-Host '機車路網服務已啟動：http://127.0.0.1:5001。請在 CareFlow 選 OSRM 機車路網。'
