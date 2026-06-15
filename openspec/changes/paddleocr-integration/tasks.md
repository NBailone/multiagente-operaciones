# Tasks: PaddleOCR Integration

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~180 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Delivery strategy | single-pr |
| Decision needed before apply | No |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

## Phase 1: Configuración base

- [ ] 1.1 **Agregar sección OCR en ui_config.json** — `ui_config.json`: añadir `"ocr": {"engine": "tesseract"}` al root del JSON. Dependencia: ninguna.

## Phase 2: Motor OCR en procesar_tickets.py

- [ ] 2.1 **Constantes y detección de sidecars** — `procesar_tickets.py`: agregar `OCR_ENGINE_DEFAULT = "tesseract"`, `OCR_ENGINE_TESSERACT`, `OCR_ENGINE_PADDLE`, `PADDLE_SIDECAR`, `TESSERACT_SIDECAR`. Lógica: si `./engines/tesseract/tesseract.exe` existe, redirigir `pytesseract.tesseract_cmd`. Si `./engines/paddleocr/` existe, agregar a `sys.path`. Dependencia: 1.1.

- [ ] 2.2 **Refactor preprocesado** — `procesar_tickets.py`: renombrar `_preprocess_image` → `_preprocess_tesseract` (misma lógica). Agregar `_preprocess_paddle(img)`: convertir a RGB, resize opcional, nunca binarizar. Dependencia: 2.1.

- [ ] 2.3 **Wrapper _ocr_tesseract** — `procesar_tickets.py`: crear `_ocr_tesseract(img_proc) → str` que llama `pytesseract.image_to_string(img_proc, lang="spa")`. Dependencia: 2.2.

- [ ] 2.4 **Singleton PaddleOCR + wrapper** — `procesar_tickets.py`: variable módulo `_paddle_ocr = None`. Función `_iniciar_paddle() → bool`: `import paddleocr`, instancia `PaddleOCR(lang='en', use_angle_cls=False)`, maneja `ImportError`. Función `_ocr_paddle(img_rgb) → str`: llama `_paddle_ocr.ocr(np.array(img_rgb))`, aplana resultado con `"\n".join(...)`. Dependencia: 2.3.

- [ ] 2.5 **Dispatcher en pdf_a_texto** — `procesar_tickets.py`: modificar firma `pdf_a_texto(ruta_pdf, engine=None, poppler_path=None)`. Si `engine=None`, leer `config["ocr"]["engine"]` desde `ui_config.json`. Dispatcher llama a `_ocr_tesseract` o `_ocr_paddle` según engine. Dependencia: 2.4, 1.1.

## Phase 3: UI — selector OCR en Ajustes

- [ ] 3.1 **Nueva tab OCR en Ajustes** — `ui_app.py`: agregar `("ocr", "OCR")` a `tab_names`. Crear método `_ajustes_tab_ocr(parent)`: `CTkOptionMenu` con valores `["tesseract", "paddleocr"]`, labels de estado "Disponible"/"No encontrado" para cada engine. Llamar al método en `_panel_ajustes`. Dependencia: 1.1.

- [ ] 3.2 **Guardar selección OCR** — `ui_app.py`: en `_guardar_ajustes()`, persistir `self.config["ocr"] = {"engine": self._ocr_engine_var.get()}`. Dependencia: 3.1.

- [ ] 3.3 **Pasar engine al worker OCR** — `ui_app.py`: en `_cargar_datos_worker` (~L7568), leer `self.config.get("ocr", {}).get("engine", "tesseract")` y pasarlo como `engine=` a `procesar_tickets.pdf_a_texto(ruta, engine=...)`. Dependencia: 3.2, 2.5.
