# Proposal: Carga de datos desde tickets de pesada

## Intent

Eliminar carga manual de Neto/Tara en planillas CONTENEDORES. El sistema obtiene tickets PDF (correo IMAP o archivo local), extrae datos vía OCR, valida contra el CONTENEDOR, y escribe en hoja DATOS si coinciden.

## Scope

### In Scope
- Botón "Cargar Datos" en sidebar (entre Planillas y Correos) + panel propio
- Fuentes: IMAP (filtro por remitente configurable) y selector local de PDFs
- OCR: extraer Patente Camión, Patente Semi, Conductor, DNI, Neto, Tara, Permiso
- Asociar ticket a `*CONTENEDORES*.*xls*` por dígitos del permiso en carpeta/nombre
- Tabla comparativa con validación visual (verde/rojo)
- Escribir Neto→PESO CARGA, Tara→TARA CONT en hoja DATOS
- Campo "Remitente Balanza" en Ajustes > Correo

### Out of Scope
- Imágenes sueltas (solo PDF)
- Múltiples hojas DATOS por CONTENEDOR
- Reprocesamiento histórico

## Capabilities

### New Capabilities
- `ticket-pesada-ocr`: OCR y parseo de tickets de pesada PDF
- `ticket-pesada-correo`: descarga de tickets desde IMAP con filtro por remitente
- `ticket-pesada-validacion`: validación contra datos del CONTENEDOR y escritura en hoja DATOS

### Modified Capabilities
- `app-ui`: nuevo botón en sidebar y panel `_panel_cargar_datos`

## Approach

Reutilizar `_imap_conectar()`, `_mail_descargar_worker()`, lectores Excel existentes, y `procesar_tickets.py` como librería. OCR en worker thread. Validación campo a campo pinta TreeView. Escritura con `openpyxl` sobre celda destino. Remitente balanza en `self.config["correo"]["remitente_balanza"]`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `ui_app.py` | Modified | Sidebar: botón idx 2.5 + `_marcar_nav_activo` + `_panel_cargar_datos()` + campo en `_ajustes_tab_correo` |
| `procesar_tickets.py` | Unmodified | Import como librería (`pdf_a_texto`, `extraer_datos`) |
| `requirements.txt` | Modified | `pytesseract`, `pdf2image` |
| — | External | Tesseract-OCR + poppler (binarios) |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| OCR impreciso | Medium | Mostrar preview del texto extraído |
| Tesseract/Poppler faltantes | Medium | Detectar al abrir panel, mostrar error |
| Formato PDF variante | Low | El usuario puede corregir campos en tabla |

## Rollback Plan

No se modifican archivos hasta confirmación "Escribir en CONTENEDORES". Reversión manual desde Excel. UI se revierte con `git checkout ui_app.py`.

## Dependencies

- `pytesseract`, `pdf2image` (pip)
- Tesseract-OCR + idioma spa + Poppler (binarios sistema)

## Success Criteria

- [ ] Tabla muestra datos OCR extraídos de correo y PDFs locales
- [ ] Validación colorea verde/rojo cada campo contra CONTENEDOR
- [ ] Escritura actualiza PESO CARGA y TARA CONT en hoja DATOS
- [ ] Remitente Balanza se guarda en config y persiste
