# Design: PaddleOCR Integration

## Technical Approach

Dispatcher en `pdf_a_texto(ruta, engine=None)` que lee `config["ocr"]["engine"]` y bifurca entre Tesseract (ruta actual, sin regresión) y PaddleOCR (nuevo pipeline: preproc RGB + singleton + flatten). UI expone selector en nueva pestaña OCR. Sidecar detection para ambos engines en `./engines/`.

## Architecture Decisions

| Decisión | Opciones | Elegido | Razón |
|----------|----------|---------|-------|
| Preproc separado | Unificar vs bifurcar | Bifurcar | PaddleOCR necesita RGB, Tesseract binarizado; mezclar degrada accuracy |
| Singleton PaddleOCR | global vs recrear | Module-level lazy | 300-500 MB RAM — cargar una sola vez, solo si se usa |
| Sidecar detection | registry vs path check | Path check + sys.path | Simple, sin registry overhead; `sys.path` para import PaddleOCR empaquetado |
| Output flatten | PaddleOCR raw vs `"\n".join()` | Join reading order | `extraer_datos()` espera texto plano con saltos de línea |
| Config key | `config["ocr"]["engine"]` | `"tesseract"` por defecto | Sin cambios en comportamiento actual |

## Data Flow

```
PDF → convert_from_path() → PIL img
                              │
                    ┌─────────┴─────────┐
                    │                   │
            _preprocess_tesseract  _preprocess_paddle
            (gray→denoise→OTSU)   (resize→RGB)
                    │                   │
              _ocr_tesseract()     _ocr_paddle()
              (pytesseract)        (PaddleOCR lazy singleton)
                    │                   │
                    └─────────┬─────────┘
                              │
                    pdf_a_texto() dispatcher
                              │
                    extraer_datos() ← text plano
```

## File Changes

| File | Acción | Descripción |
|------|--------|-------------|
| `procesar_tickets.py` | Modify | Dual preproc, engine wrappers, dispatcher, Paddle singleton |
| `ui_app.py` | Modify | Nueva tab OCR con selector y status indicators |
| `ui_config.json` | Modify | Nueva sección `"ocr"` con `"engine"` key |

## procsar_tickets.py — Componentes Nuevos

```
OCR_ENGINE_TESSERACT = "tesseract"
OCR_ENGINE_PADDLE    = "paddleocr"

_paddle_ocr = None  # singleton, lazy init

def _iniciar_paddle() -> bool
    # try: import paddleocr → PaddleOCR(lang='en', use_angle_cls=False)
    # maneja ImportError, retorna bool

def _preprocess_tesseract(img: PIL.Image) -> PIL.Image
    # extraído de _preprocess_image() actual — binary pipeline

def _preprocess_paddle(img: PIL.Image) -> PIL.Image
    # resize opcional, convertir a RGB (nunca binarizar)

def _ocr_tesseract(img_proc: PIL.Image) -> str
    # pytesseract.image_to_string(img_proc, lang="spa")

def _ocr_paddle(img_rgb: PIL.Image) -> str
    # _paddle_ocr.ocr(np.array(img_rgb)) → flatten: "\n".join(line[1][0] for line in result[0])

def pdf_a_texto(ruta_pdf, engine=None, poppler_path=None) -> str
    # si engine=None, lee config (default "tesseract")
    # dispatchea a _ocr_tesseract o _ocr_paddle
```

## ui_app.py — OCR Tab

```
tab_names append: ("ocr", "OCR")

Atributos:
  self._ocr_engine_var = ctk.StringVar(value=config["ocr"]["engine"])
  self._ocr_lbl_tesseract = CTkLabel("✓ Tesseract detectable" | "✗ No encontrado")
  self._ocr_lbl_paddle    = CTkLabel("✓ PaddleOCR instalado" | "✗ No instalado")

Widgets en _ajustes_tab_ocr():
  - CTkLabel "Engine OCR por defecto"
  - CTkOptionMenu(values=["tesseract", "paddleocr"], variable=self._ocr_engine_var)
  - CTkLabel "Estado Tesseract:" + self._ocr_lbl_tesseract
  - CTkLabel "Estado PaddleOCR:" + self._ocr_lbl_paddle

Guardado en _guardar_ajustes():
  self.config["ocr"] = {"engine": self._ocr_engine_var.get()}
```

## Sidecar Detection

| Path | Qué verifica | Fallback |
|------|-------------|----------|
| `./engines/tesseract/tesseract.exe` | `os.path.isfile()` → set `pytesseract.tesseract_cmd` | System Tesseract (`$PATH` o `C:\Program Files\...`) |
| `./engines/paddleocr/` | `os.path.isdir()` → `sys.path.insert(0, path)` | `import paddleocr` desde site-packages |

Detection runs on tab load y al abrir panel Carga de Datos, mostrando estado visual.

## Interfaces / Contracts

```
# pdf_a_texto engine param
engine: str | None = None  # "tesseract" | "paddleocr" | None (usa config)

# ui_config.json nueva sección
{
  "ocr": {
    "engine": "tesseract"  # default, persistido
  }
}

# PaddleOCR output flatten
result[0] = [[bbox, (text, conf)], ...] → "\n".join([line[1][0] for line in result[0]])
```

## Testing Strategy

| Capa | Qué probar | Cómo |
|------|-----------|------|
| Unit | `_preprocess_paddle()` | Input PIL → output es RGB, mismo size |
| Unit | Paddle flatten | Mock `_paddle_ocr.ocr()` → verificar texto plano con saltos |
| Integration | `pdf_a_texto(engine="tesseract")` | Output bit-identical al actual |
| Integration | Engine selector en config | Cambiar engine vía dict → llamar `pdf_a_texto()` → verifica dispatch |
| Manual | UI tab OCR | Abrir Ajustes → ver tab, cambiar engine, guardar, recargar |

## Migration / Rollout

No migration required. Engine default `"tesseract"` — cero impacto en flujo actual. Usuario activa PaddleOCR explícitamente desde Ajustes.

## Open Questions

- [ ] PaddleOCR `lang='en'` funciona para español AGD? ¿Probar con `lang='es'`? PaddleOCR no tiene modelo español dedicado — `en` suele funcionar bien con caracteres latinos.
- [ ] Resize óptimo para PaddleOCR: ¿mantener 300 DPI o reducir a 640px de ancho?
