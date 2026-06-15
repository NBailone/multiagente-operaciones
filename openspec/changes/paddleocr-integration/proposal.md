# Proposal: PaddleOCR Integration as Optional OCR Engine

## Intent

Add PaddleOCR as configurable Tesseract alternative for Spanish AGD ticket OCR, improving accuracy on noisy scans and enabling future GPU acceleration.

## Scope

### In Scope
- Dispatcher in `procesar_tickets.py`: `pdf_a_texto(ruta, engine=None)` reads config or param
- Dual preprocessing: `_preprocess_tesseract()` (binary, unchanged), `_preprocess_paddle()` (RGB, never binary)
- PaddleOCR singleton, lazy init (~300–500 MB RAM)
- Output flatten: `"\n".join(line[1][0] ... )` → feeds `extraer_datos()`
- UI: OCR tab in Ajustes — CTkOptionMenu + install-status label
- Config: `ui_config.json` section `config["ocr"]["engine"]` default `"tesseract"`
- Dependency check: `try: import paddleocr` → warn if missing

### Out of Scope
- Auto-detect best engine, parallel engines, GPU config, EasyOCR
- PaddleOCR bbox for field-level extraction
- Bundling PaddlePaddle models in PyInstaller build

## Capabilities

### New Capabilities
- `ocr-engine`: Engine selection, dual preprocessing, PaddleOCR singleton, output flatten

### Modified Capabilities
None — no existing OCR specs.

## Approach

1. Extract current preproc into `_preprocess_tesseract()`. Add `_preprocess_paddle()` — RGB only, optional resize
2. Engine wrappers: `_ocr_tesseract(img_proc)` and `_ocr_paddle(img_rgb)`
3. Lazy init singleton: `PaddleOCR(lang='en', use_angle_cls=False)`
4. `pdf_a_texto()` reads `config["ocr"]["engine"]` if `engine=None`, dispatches to wrapper
5. New OCR tab in Ajustes — CTkOptionMenu for engine, CTkLabel for install status
6. `config["ocr"]["engine"]` saved to `ui_config.json` on change, loaded at startup

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `procesar_tickets.py` | Modified | Dispatcher, dual preproc, singleton, flatten |
| `ui_app.py` | Modified | New OCR tab, engine selector, dep check |
| `ui_config.json` | Modified | New `"ocr"` section with `"engine"` key |
| `Sistema_Automatizacion.spec` | Modified | Conditional hidden-imports for PaddlePaddle |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| 300–500 MB RAM on Paddle use | High | Lazy init; document as known constraint |
| Binary image fed to Paddle hurts accuracy | Med | Strict preproc separation per engine |
| First-run model download needs internet | High | Clear UI error; fallback to Tesseract |
| PyInstaller freeze of PaddlePaddle fragile | Med | Exclude from default build; document limitation |

## Rollback Plan

Set engine to `"tesseract"`, remove PaddleOCR paths. Revert `.spec` if build breaks.

## Dependencies

- `pip install paddlepaddle paddleocr` (~800 MB–1.2 GB disk)
- Python 3.8–3.12 (project is 3.14 — verify compat)
- Shapely + MSVC build tools on Windows

## Success Criteria

- [ ] Engine selection persists across app restart
- [ ] PaddleOCR produces same `extraer_datos()` fields as Tesseract (text may differ)
- [ ] Tesseract path bit-identical when engine=`"tesseract"` (no regression)
- [ ] PaddleOCR singleton loads only on first Paddle OCR call, not at startup
- [ ] Engine switch via config takes effect on next `pdf_a_texto()` call
