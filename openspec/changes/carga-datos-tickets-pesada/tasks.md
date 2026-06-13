# Tasks: Carga de datos desde tickets de pesada

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~350 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |

Decision needed before apply: Yes
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

---

## Phase 1: Foundation — Sidebar y Navegación

### Tarea 1: Botón "Cargar Datos" en sidebar
**Archivos**: `ui_app.py` (L388-L402, L471-L485, L638-L658)
**Depende de**: —
**Descripción**: Insertar nuevo botón `self.btn_cargar_datos` entre Planillas (idx 2) y Correos (actual idx 3, pasa a 4). Agregar `"cargar-datos"` al `idx_map` en `_cambiar_panel_forzado`. Actualizar lista `botones` en `_marcar_nav_activo`. Agregar entrada `"cargar-datos"` en `logs_por_panel`.
**Criterios de aceptación**:
- [x] Botón "Cargar Datos" visible entre "Completar Planillas" y "Enviar Correos"
- [x] Click en botón cambia a panel correcto y destaca el botón activo
- [x] Correos pasa a índice 4, Backup a 5, Ajustes a 6
- [x] Consola se guarda/restaura al cambiar al nuevo panel
**Notas**: Tocar 3 métodos: `_crear_sidebar`, `_marcar_nav_activo`, `_cambiar_panel_forzado`. No romper índices existentes.

### Tarea 2: Remitente Balanza en Ajustes > Correo
**Archivos**: `ui_app.py` (~L6342-L6375)
**Depende de**: —
**Descripción**: Agregar `self._ent_remitente_balanza` como `_ajustes_row` al final de `_ajustes_tab_correo`, con label "Remitente Balanza:" valor inicial de `self.config.get("correo", {}).get("remitente_balanza", "")`. En `_guardar_ajustes`, agregar `"remitente_balanza": self._ent_remitente_balanza.get().strip()` al dict `correo_cfg`.
**Criterios de aceptación**:
- [x] Campo "Remitente Balanza" visible en Ajustes > Correo
- [x] Valor se persiste en `ui_config.json` bajo `correo.remitente_balanza`
- [x] Valor se restaura al abrir Ajustes nuevamente

---

## Phase 2: Panel Cargar Datos

### Tarea 3: Shell del panel `_panel_cargar_datos`
**Archivos**: `ui_app.py` (+~90 lines, nuevo método entre backup y ajustes)
**Depende de**: Tarea 1
**Descripción**: Crear `_panel_cargar_datos(self)` con:
- Toolbar: `CTkButton "Buscar en Correo"`, `CTkButton "Seleccionar PDFs"`, `CTkProgressBar` (indeterminate)
- `CTkScrollableFrame` conteniendo `ttk.Treeview` con columnas `("ticket", "patente", "semi", "conductor", "dni", "neto_ocr", "neto_cont", "tara_ocr", "tara_cont", "estado")`
- `CTkButton "Escribir en CONTENEDORES"` (disabled por defecto)
- Conectar `self.lbl_titulo_panel.configure(text="Carga de Datos")` en `_cambiar_panel_forzado`
**Criterios de aceptación**:
- [x] Panel se renderiza al hacer click en "Cargar Datos"
- [x] TreeView vacío con columnas correctas
- [x] Botón "Escribir en CONTENEDORES" deshabilitado
- [x] Toolbar muestra ambos botones de fuente + progress bar

---

## Phase 3: OCR y Procesamiento

### Tarea 4: Worker OCR `_cargar_datos_worker`
**Archivos**: `ui_app.py` (+~60 lines, nuevo método)
**Depende de**: Tarea 3
**Descripción**: Crear `_cargar_datos_worker(self, fuente, rutas)` que:
1. Para cada PDF en `rutas`, llama `procesar_tickets.pdf_a_texto(ruta)` y `procesar_tickets.extraer_datos(texto)`
2. Convierte claves del dict (Patente Camion→patente, Patente Acoplado→semi, Conductor→conductor, DNI extraído de Conductor, Peso Neto→neto, Peso Tara→tara, Merc./Permiso→permiso)
3. Por cada resultado, dispara matching de CONTENEDOR (Tarea 5)
4. Pushea tupla `(ticket_data, cont_data)` a `self.log_queue` con prefijo `_OCR_RESULT_`
5. Pushea `("_OCR_DONE_", resultados)` al final
**Criterios de aceptación**:
- [x] Worker corre en `threading.Thread` daemon
- [x] PDF inválido no corta el batch (loggea error, sigue)
- [x] `self.tarea_activa` se setea correctamente
- [x] Botones se deshabilitan durante ejecución
**Notas**: Usar `import procesar_tickets` directo. No modificar `procesar_tickets.py`.

### Tarea 5: Matching CONTENEDOR `_match_contenedor`
**Archivos**: `ui_app.py` (+~35 lines, nuevo método)
**Depende de**: Tarea 4
**Descripción**: Crear `_match_contenedor(self, permiso_ticket)` que:
1. Toma los últimos 5 caracteres alfanuméricos de `permiso_ticket`
2. Busca archivos `*CONTENEDORES*.*xls*` en `self._resolver_ruta("planillas_carga", "Desktop")` y subdirectorios
3. Para cada archivo, abre con `openpyxl`, lee celda DATOS R5 C7 (PE cell), compara últimos 5 chars
4. Retorna `(ruta_archivo, workbook, worksheet)` o `None`
**Criterios de aceptación**:
- [x] Match exacto por últimos 5 dígitos retorna archivo correcto
- [x] Sin match retorna `None` (sin crash)
- [x] Archivo bloqueado por Excel loggea error y retorna `None`
**Notas**: Usar `openpyxl.load_workbook(read_only=True)` para solo lectura.

### Tarea 6: Lectura de DATOS y comparación campo a campo
**Archivos**: `ui_app.py` (+~40 lines, dentro del worker)
**Depende de**: Tarea 5
**Descripción**: Dentro del worker, tras obtener CONTENEDOR match, leer valores esperados de la hoja DATOS:
- Localizar el bloque `k` (0..4) que coincide con el permiso en PE (R{k*13+20}, C1)
- Extraer: Neto esperado de PESO CARGA (R{k*13+28}, C+6), Tara esperado de TARA CONT (R{k*13+29}, C+6)
- Crear dict `cont_data` con `{neto_esperado, tara_esperado}`
- Devolver ambos dicts para que Tarea 7 los compare
**Criterios de aceptación**:
- [x] Lee correctamente valores de PESO CARGA y TARA CONT por coordenadas
- [x] Contenedor sin bloque matching loggea advertencia (no crash)

---

## Phase 4: Validación y Escritura

### Tarea 7: Población del TreeView con validación visual
**Archivos**: `ui_app.py` (+~45 lines, en `_poll_log_queue` + helper)
**Depende de**: Tarea 4, Tarea 6
**Descripción**: En el main thread (`self.after(0, ...)`), consumir mensajes `_OCR_RESULT_` de `log_queue`. Por cada resultado:
- Insertar fila en TreeView con: número ticket, patente, semi, conductor, DNI, neto_ocr, neto_cont, tara_ocr, tara_cont, estado
- Comparar neto_ocr vs neto_cont y tara_ocr vs tara_cont
- Aplicar tags `tag_green` (match) o `tag_red` (mismatch) a las celdas
- Estado: "✅ Ok" (todo match) o "❌ Diferencia" (algún mismatch)
- Al recibir `_OCR_DONE_`, habilitar "Escribir en CONTENEDORES" si hay al menos un "✅ Ok"
**Criterios de aceptación**:
- [x] TreeView muestra datos OCR por fila
- [x] Celdas neto/tara en verde si coinciden, rojo si no
- [x] Estado "✅ Ok" para match completo, "❌ Diferencia" si no
- [x] Botón escribir se habilita solo cuando hay al menos un "✅ Ok"
**Notas**: Usar `ttk.Treeview.tag_configure` para colores. No hay test framework — probar con 2 PDFs conocidos.

### Tarea 8: Escritura en CONTENEDORES `_escribir_en_contenedores`
**Archivos**: `ui_app.py` (+~30 lines)
**Depende de**: Tarea 7
**Descripción**: Crear método que:
1. Recolecta filas con "✅ Ok" del TreeView
2. Para cada una, abre CONTENEDOR con `openpyxl`, escribe Neto→PESO CARGA (R{k*13+28}, C+6) y Tara→TARA CONT (R{k*13+29}, C+6)
3. Guarda con `wb.save()`
4. Loggea resultado por fila
5. Botón "Escribir en CONTENEDORES" llama a worker thread
**Criterios de aceptación**:
- [x] "Escribir" actualiza celdas PESO CARGA y TARA CONT en DATOS
- [x] Archivo bloqueado muestra error "está bloqueado. Cerralo e intentá de nuevo"
- [x] Solo filas "✅ Ok" se escriben (las "❌ Diferencia" se saltan)

---

## Phase 5: Fuentes de Datos

### Tarea 9: Fuente Correo — "Buscar en Correo"
**Archivos**: `ui_app.py` (+~50 lines, nuevo método + popup)
**Depende de**: Tarea 3
**Descripción**: Al clickear "Buscar en Correo":
1. Verificar que `self.config["correo"]["remitente_balanza"]` no esté vacío; si vacío, mostrar warning "Configurá el Remitente Balanza en Ajustes > Correo primero"
2. Mostrar popup CTkToplevel pidiendo cantidad de emails a revisar (default 20)
3. En worker thread: conectar IMAP vía `_imap_conectar()`, buscar últimos N mails, filtrar por `FROM remitente_balanza`, detectar adjuntos PDF
4. Mostrar lista con checkboxes (reutilizar patrón de `_mail_tree`)
5. Al confirmar, descargar PDFs seleccionados a temp y pasar rutas a `_cargar_datos_worker`
**Criterios de aceptación**:
- [x] Sin remitente configurado → warning, no IMAP
- [x] Popup pide cantidad de mails
- [x] Solo muestra mails del remitente balanza con adjuntos PDF
- [x] PDFs descargados se pasan al OCR worker

### Tarea 10: Fuente Archivo — "Seleccionar PDFs"
**Archivos**: `ui_app.py` (+~15 lines)
**Depende de**: Tarea 3
**Descripción**: Al clickear "Seleccionar PDFs", abrir `filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])`. Si hay archivos seleccionados, pasarlos directamente a `_cargar_datos_worker(fuente="local", rutas=seleccionados)`.
**Criterios de aceptación**:
- [x] Filedialog se abre con filtro PDF
- [x] Múltiple selección permitida
- [x] Selección vacía no inicia worker
- [x] Log muestra cantidad de PDFs seleccionados

---

## Phase 6: Polish

### Tarea 11: Detección de dependencias al abrir panel
**Archivos**: `ui_app.py` (+~20 lines, dentro de `_panel_cargar_datos`)
**Depende de**: Tarea 3
**Descripción**: Al inicio de `_panel_cargar_datos`, verificar:
1. Tesseract: `pytesseract.get_tesseract_version()` — si falla, mostrar label inline: "Tesseract-OCR no encontrado. Descargalo de https://github.com/UB-Mannheim/tesseract/wiki"
2. Poppler: `pdf2image.pdfinfo_from_path` con PDF dummy o try/except — si falla: "Poppler no encontrado. Descargalo de https://github.com/oschwartz10612/poppler-windows/releases/"
3. Si alguna falta, deshabilitar botones de fuente y mostrar mensaje
**Criterios de aceptación**:
- [x] Sin Tesseract → mensaje inline + botones deshabilitados
- [x] Sin Poppler → mensaje inline + botones deshabilitados
- [x] Ambos presentes → panel opera normal
**Notas**: No crear archivos dummy. Intentar `pytesseract.get_tesseract_version()` y capturar `TesseractNotFoundError`.

---

## Summary

| Phase | Tareas | Foco |
|-------|--------|------|
| Phase 1: Foundation | T1, T2 | Sidebar + Settings |
| Phase 2: Panel Shell | T3 | Panel layout |
| Phase 3: Core Pipeline | T4, T5, T6 | OCR + Matching + DATOS |
| Phase 4: Validation & Write | T7, T8 | TreeView + Excel write |
| Phase 5: Sources | T9, T10 | Correo + Archivo |
| Phase 6: Polish | T11 | Dependency detection |
| **Total** | **11 tareas** | |

### Orden de implementación
T1→T2 (paralelo), luego T3, luego T4+T5+T6 (secuencial), luego T7, luego T8, luego T9+T10 (paralelo tras T3), luego T11 (al final de T3).

### Decisiones
- **Delivery**: Single PR (~350 líneas, bajo 400)
- **¿Preguntar antes de apply?**: Sí — el usuario debe confirmar el plan de tareas y la decisión de single PR
