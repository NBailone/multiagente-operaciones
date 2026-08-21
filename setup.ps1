# ──────────────────────────────────────────────────────────────
# setup.ps1 — Bootstrap para clone fresco del proyecto
# ──────────────────────────────────────────────────────────────
# Reconstruye todo lo que Git no versiona (binarios de terceros):
#   1. Dependencias Python (requirements.txt)
#   2. python39/          → Python embeddable + PaddleOCR
#   3. engines/tesseract/ → Tesseract-OCR portable + idioma español
#   4. poppler/           → Poppler para conversión PDF→imagen
#
# Uso:  .\setup.ps1        (desde la raíz del proyecto)
# Es idempotente: lo que ya está instalado y funciona se salta.
# Puede pedir permiso de administrador (UAC) para instalar Tesseract.
# ──────────────────────────────────────────────────────────────

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot

# Versiones pinneadas (verificadas y probadas). Si falla el download,
# se usa el último release disponible con un aviso.
$PYTHON_VERSION = "3.9.7"
$TESSERACT_TAG  = "v5.4.0.20240606"
$POPLER_TAG     = "v26.02.0-0"
$SUMATRA_VERSION = "3.6.1"

$tmp = Join-Path $env:TEMP "multiagente-setup"
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

function Invoke-Download {
    param([string]$Url, [string]$OutFile)
    Write-Host "    descargando $Url" -ForegroundColor DarkGray
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
}

function Get-LatestGithubAssetUrl {
    param([string]$Repo, [string]$AssetPattern)
    $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest"
    $asset = $rel.assets | Where-Object { $_.name -match $AssetPattern } | Select-Object -First 1
    if ($null -eq $asset) { throw "No asset matching '$AssetPattern' in latest release of $Repo" }
    return @{ Url = $asset.browser_download_url; Tag = $rel.tag_name }
}

function Invoke-NativeQuiet {
    # Ejecuta un proceso sin que su salida a stderr dispare ErrorActionPreference=Stop
    param([string]$FilePath, [string[]]$ArgumentList = @())
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try { & $FilePath @ArgumentList 2>&1 | Out-Null } catch { }
    finally { $ErrorActionPreference = $prev }
    return $LASTEXITCODE
}

function Test-PaddleOk {
    $py = Join-Path $projectRoot "python39\python.exe"
    if (-not (Test-Path $py)) { return $false }
    return ((Invoke-NativeQuiet $py -ArgumentList @("-c", "from paddleocr import PaddleOCR")) -eq 0)
}

function Test-TesseractOk {
    $exe = Join-Path $projectRoot "engines\tesseract\tesseract.exe"
    if (-not (Test-Path $exe)) { return $false }
    return ((Invoke-NativeQuiet $exe -ArgumentList "--version") -eq 0)
}

function Test-PopplerOk {
    $exe = Join-Path $projectRoot "poppler\Library\bin\pdftoppm.exe"
    if (-not (Test-Path $exe)) { return $false }
    return ((Invoke-NativeQuiet $exe -ArgumentList "-v") -eq 0)
}

function Test-SumatraOk {
    return (Test-Path (Join-Path $projectRoot "engines\sumatra\SumatraPDF.exe"))
}

Write-Host "=== Setup Multiagente ===" -ForegroundColor Cyan

# ── [1/5] Dependencias Python ─────────────────────────────────
Write-Host "`n[1/5] Dependencias Python (requirements.txt)" -ForegroundColor Yellow
python -c "import customtkinter, openpyxl, xlrd, win32com" *> $null
if ($LASTEXITCODE -ne 0) {
    python -m pip install -r (Join-Path $projectRoot "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
} else {
    Write-Host "    ya instaladas, salteando" -ForegroundColor DarkGray
}

# ── [2/5] python39/ — runtime portable con PaddleOCR ──────────
Write-Host "`n[2/5] python39/ (Python embeddable + PaddleOCR)" -ForegroundColor Yellow
if (Test-PaddleOk) {
    Write-Host "    ya instalado, salteando" -ForegroundColor DarkGray
} else {
    $p39dir = Join-Path $projectRoot "python39"
    New-Item -ItemType Directory -Path $p39dir -Force | Out-Null

    # 2a. Descargar y extraer el embeddable
    $zip = Join-Path $tmp "python-$PYTHON_VERSION-embed-amd64.zip"
    Invoke-Download "https://www.python.org/ftp/python/$PYTHON_VERSION/python-$PYTHON_VERSION-embed-amd64.zip" $zip
    Expand-Archive -Path $zip -DestinationPath $p39dir -Force

    # 2b. Habilitar site-packages en el ._pth
    $pth = Join-Path $p39dir "python39._pth"
    (Get-Content $pth -Raw).Replace("#import site", "import site") | Set-Content $pth

    # 2c. Instalar pip (get-pip.py principal; fallback a la versión para 3.9)
    $getpip = Join-Path $tmp "get-pip.py"
    try {
        Invoke-Download "https://bootstrap.pypa.io/get-pip.py" $getpip
        & "$p39dir\python.exe" $getpip --no-warn-script-location *> $null
    } catch { }
    if (-not (Test-Path (Join-Path $p39dir "Scripts\pip.exe"))) {
        Invoke-Download "https://bootstrap.pypa.io/pip/3.9/get-pip.py" $getpip
        & "$p39dir\python.exe" $getpip --no-warn-script-location *> $null
    }

    # 2d. Instalar PaddleOCR (tarda varios minutos)
    Write-Host "    instalando paddlepaddle + paddleocr (varios minutos)..." -ForegroundColor DarkGray
    & "$p39dir\python.exe" -m pip install --no-warn-script-location paddlepaddle paddleocr
    if (-not (Test-PaddleOk)) { throw "PaddleOCR no quedó funcional en python39/" }
}

# ── [3/5] engines/tesseract/ — OCR portable ───────────────────
Write-Host "`n[3/5] engines/tesseract/ (Tesseract-OCR)" -ForegroundColor Yellow
if (Test-TesseractOk) {
    Write-Host "    ya instalado, salteando" -ForegroundColor DarkGray
} else {
    # 3a. URL pinneado; si falla, último release con aviso
    $tag = $TESSERACT_TAG
    $url = "https://github.com/UB-Mannheim/tesseract/releases/download/$tag/tesseract-ocr-w64-setup-$($tag.TrimStart('v')).exe"
    try {
        Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing | Out-Null
    } catch {
        Write-Host "    [AVISO] versión pinneada ($tag) no disponible, usando el último release" -ForegroundColor Magenta
        $latest = Get-LatestGithubAssetUrl -Repo "UB-Mannheim/tesseract" -AssetPattern "w64-setup"
        $url = $latest.Url
        $tag = $latest.Tag
    }

    # 3b. Instalación silenciosa dentro del proyecto (NSIS: /S /D= al final)
    $installer = Join-Path $tmp "tesseract-setup.exe"
    Invoke-Download $url $installer
    $targetDir = Join-Path $projectRoot "engines\tesseract"
    Write-Host "    instalando silenciosamente en $targetDir" -ForegroundColor DarkGray
    Start-Process -FilePath $installer -ArgumentList "/S", "/D=$targetDir" -Wait -PassThru | Out-Null

    if (-not (Test-TesseractOk)) { throw "Tesseract no quedó funcional en engines/tesseract/" }

    # 3c. Idioma español (el installer trae solo eng+osd)
    $spa = Join-Path $targetDir "tessdata\spa.traineddata"
    if (-not (Test-Path $spa)) {
        Invoke-Download "https://github.com/tesseract-ocr/tessdata_fast/raw/main/spa.traineddata" $spa
    }
}

# ── [4/5] poppler/ — conversión PDF → imagen ──────────────────
Write-Host "`n[4/5] poppler/ (conversión PDF)" -ForegroundColor Yellow
if (Test-PopplerOk) {
    Write-Host "    ya instalado, salteando" -ForegroundColor DarkGray
} else {
    $tag = $POPLER_TAG
    $url = "https://github.com/oschwartz10612/poppler-windows/releases/download/$tag/Release-$($tag.TrimStart('v')).zip"
    try {
        Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing | Out-Null
    } catch {
        Write-Host "    [AVISO] versión pinneada ($tag) no disponible, usando el último release" -ForegroundColor Magenta
        $latest = Get-LatestGithubAssetUrl -Repo "oschwartz10612/poppler-windows" -AssetPattern "^Release-.*\.zip$"
        $url = $latest.Url
        $tag = $latest.Tag
    }

    $zip = Join-Path $tmp "poppler.zip"
    Invoke-Download $url $zip
    $extractTo = Join-Path $tmp "poppler-extract"
    if (Test-Path $extractTo) { Remove-Item $extractTo -Recurse -Force }
    Expand-Archive -Path $zip -DestinationPath $extractTo -Force

    # Normalizar estructura: <raíz>/poppler/Library/bin/pdftoppm.exe
    $bin = Get-ChildItem $extractTo -Recurse -Filter "pdftoppm.exe" | Select-Object -First 1
    if ($null -eq $bin) { throw "pdftoppm.exe not found in poppler zip" }
    $dest = Join-Path $projectRoot "poppler"
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    # Copiar desde la carpeta que contiene Library/
    $libraryRoot = Split-Path (Split-Path $bin.FullName)
    Copy-Item -Path $libraryRoot -Destination $dest -Recurse
    Remove-Item $extractTo -Recurse -Force
}

# ── [5/5] engines/sumatra/ — impresión PDF con N copias ───────
Write-Host "`n[5/5] engines/sumatra/ (SumatraPDF para copias múltiples)" -ForegroundColor Yellow
if (Test-SumatraOk) {
    Write-Host "    ya instalado, salteando" -ForegroundColor DarkGray
} else {
    # URL pinneado; si falla, descubrir la versión actual desde la página oficial
    $url = "https://www.sumatrapdfreader.org/dl/rel/$SUMATRA_VERSION/SumatraPDF-$SUMATRA_VERSION-64.zip"
    try {
        Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing | Out-Null
    } catch {
        Write-Host "    [AVISO] versión pinneada ($SUMATRA_VERSION) no disponible, buscando la actual..." -ForegroundColor Magenta
        $page = Invoke-WebRequest -Uri "https://www.sumatrapdfreader.org/download-free-pdf-viewer" -UseBasicParsing
        $match = [regex]::Match($page.Content, '/dl/rel/[\d\.]+/SumatraPDF-[\d\.]+-64\.zip')
        if (-not $match.Success) { throw "No se pudo determinar la URL de SumatraPDF; descargalo a mano en engines\sumatra\" }
        $url = "https://www.sumatrapdfreader.org$($match.Value)"
    }

    $zip = Join-Path $tmp "sumatra.zip"
    Invoke-Download $url $zip
    $extractTo = Join-Path $tmp "sumatra-extract"
    if (Test-Path $extractTo) { Remove-Item $extractTo -Recurse -Force }
    Expand-Archive -Path $zip -DestinationPath $extractTo -Force

    # El zip trae el exe con nombre versionado (ej: SumatraPDF-3.6.1-64.exe); normalizar
    $exe = Get-ChildItem $extractTo -Recurse -Filter "SumatraPDF*.exe" | Select-Object -First 1
    if ($null -eq $exe) { throw "SumatraPDF.exe not found in zip" }
    $destDir = Join-Path $projectRoot "engines\sumatra"
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    Copy-Item $exe.FullName -Destination (Join-Path $destDir "SumatraPDF.exe") -Force
    Remove-Item $extractTo -Recurse -Force
}

# ── Verificación final ────────────────────────────────────────
Write-Host "`n=== Verificación final ===" -ForegroundColor Cyan
$tessVersion = (& (Join-Path $projectRoot "engines\tesseract\tesseract.exe") --version | Select-Object -First 1)
Write-Host "Tesseract : $tessVersion"
& (Join-Path $projectRoot "python39\python.exe") -c "from paddleocr import PaddleOCR; print('PaddleOCR: import ok')"
$prev = $ErrorActionPreference; $ErrorActionPreference = "Continue"
try {
    $popplerVersion = (& (Join-Path $projectRoot "poppler\Library\bin\pdftoppm.exe") -v 2>&1 | Select-Object -First 1)
    Write-Host "Poppler   : $popplerVersion"
} catch { Write-Host "Poppler   : [sin salida]" }
finally { $ErrorActionPreference = $prev }
Write-Host ("Sumatra   : " + $(if (Test-SumatraOk) { "ok (engines\sumatra)" } else { "[no instalado: las copias múltiples irán una por una]" }))

Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "Setup OK. Siguiente paso: configurar .env (clave de encriptación) y ejecutar 'python app.py'" -ForegroundColor Green

