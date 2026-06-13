# ──────────────────────────────────────────────────────────
# Setup PaddleOCR sidecar portable
# ──────────────────────────────────────────────────────────
# EJECUTAR EN PowerShell (como administrador si es necesario)

Write-Host "=== Setup PaddleOCR Sidecar ===" -ForegroundColor Cyan

# 1. Ir a la carpeta del proyecto
Set-Location -LiteralPath "C:\Users\Nico\Desktop\Antigravity\Multiagente"

# 2. Crear carpeta engines
New-Item -ItemType Directory -Path ".\engines\paddleocr" -Force | Out-Null

# 3. Descargar Python 3.12 portable (embeddable) a la carpeta
Write-Host "[1/5] Descargando Python 3.12..." -ForegroundColor Yellow
$pythonUrl = "https://www.python.org/ftp/python/3.12.9/python-3.12.9-embed-amd64.zip"
$zipPath = "$env:TEMP\python312.zip"
Invoke-WebRequest -Uri $pythonUrl -OutFile $zipPath

# 4. Extraer Python 3.12
Write-Host "[2/5] Extrayendo Python 3.12 en .\engines\paddleocr\python\" -ForegroundColor Yellow
Expand-Archive -Path $zipPath -DestinationPath ".\engines\paddleocr\python" -Force

# 5. Activar pip en el embeddable Python
Write-Host "[3/5] Activando pip..." -ForegroundColor Yellow

# El embeddable no trae pip. Descargamos get-pip.py
$getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
$getPipPath = "$env:TEMP\get-pip.py"
Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipPath

# Editar python._pth para permitir site-packages
$pthFile = ".\engines\paddleocr\python\python._pth"
$pthContent = Get-Content -Path $pthFile -Raw
$pthContent = $pthContent.Replace("#import site", "import site")
Set-Content -Path $pthFile -Value $pthContent

# Instalar pip
& ".\engines\paddleocr\python\python.exe" $getPipPath

# 6. Instalar PaddlePaddle CPU + PaddleOCR
Write-Host "[4/5] Instalando PaddlePaddle CPU + PaddleOCR (esto tarda varios minutos)..." -ForegroundColor Yellow
& ".\engines\paddleocr\python\python.exe" -m pip install paddlepaddle paddleocr -q

# 7. Verificar instalación
Write-Host "[5/5] Verificando instalación..." -ForegroundColor Yellow
$result = & ".\engines\paddleocr\python\python.exe" -c "from paddleocr import PaddleOCR; ocr = PaddleOCR(lang='en', use_angle_cls=False); print('✓ PaddleOCR listo')" 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "=== INSTALACIÓN COMPLETA ===" -ForegroundColor Green
    Write-Host "PaddleOCR disponible en .\engines\paddleocr\" -ForegroundColor Green
} else {
    Write-Host "=== ERROR ===" -ForegroundColor Red
    Write-Host $result -ForegroundColor Red
    Write-Host "La primera descarga de modelos puede fallar si no hay internet." -ForegroundColor Yellow
}

# Limpiar descargas temporales
Remove-Item $zipPath -ErrorAction SilentlyContinue
Remove-Item $getPipPath -ErrorAction SilentlyContinue

Write-Host "`nListo. Abrí el programa y andá a Ajustes → 🤖 OCR para verificar." -ForegroundColor Cyan
