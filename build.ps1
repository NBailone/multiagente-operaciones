# Build script for Sistema_Automatizacion (ONEFOLDER portable build)
# Usage: .\build.ps1
#
# 1. Runs PyInstaller with ui_app.spec
# 2. Copies the portable python39/ runtime next to the exe (sidecar, not bundled)
#    so PaddleOCR works without inflating every build.

$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$distDir     = Join-Path $projectRoot "dist\Sistema_Automatizacion"

Write-Host "==> [1/3] Running PyInstaller..." -ForegroundColor Cyan
python -m PyInstaller --noconfirm ui_app.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

Write-Host "==> [2/3] Copying python39/ sidecar next to exe..." -ForegroundColor Cyan
if (-not (Test-Path (Join-Path $projectRoot "python39\python.exe"))) {
    throw "python39\python.exe not found in project root. PaddleOCR needs this portable runtime."
}
Copy-Item -Path (Join-Path $projectRoot "python39") -Destination $distDir -Recurse -Force

Write-Host "==> [3/3] Verifying OCR engines..." -ForegroundColor Cyan
$tess = Join-Path $distDir "_internal\engines\tesseract\tesseract.exe"
& $tess --version | Select-Object -First 1
if ($LASTEXITCODE -ne 0) { throw "Tesseract sidecar check failed" }

$paddle = Join-Path $distDir "python39\python.exe"
& $paddle -c "from paddleocr import PaddleOCR; print('PaddleOCR import ok')"

Write-Host ""
Write-Host "Build OK -> $distDir" -ForegroundColor Green
