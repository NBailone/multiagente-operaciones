# Design: Carga de datos desde tickets de pesada

## Technical Approach

Reuse `procesar_tickets.py` (`pdf_a_texto`, `extraer_datos`) for OCR in a daemon worker thread. Reuse `_imap_conectar()` for IMAP fetching filtered by a new `remitente_balanza` config field. Associate tickets to `*CONTENEDORES*.*xls*` by matching permit digits. Display extracted vs. expected fields in a tag-based TreeView (green=ok, red=mismatch). Write `Neto→PESO CARGA` and `Tara→TARA CONT` to DATOS cells via `openpyxl`. All I/O happens in the worker; UI updates via `self.after(0, ...)`, logs via `self.log_queue.put(...)`.

## Architecture Decisions

| Decision | Options | Choice | Rationale |
|----------|---------|--------|-----------|
| Sidebar insertion | (a) New button between Planillas and Correos (b) Sub-tab (c) Floating button | **Sidebar at index 3** | Consistent with existing nav; shifts Correos→4, Backup→5, Ajustes→6. 3-site change: `_crear_sidebar`, `_marcar_nav_activo`, `_cambiar_panel` |
| Worker threading | (a) `threading.Thread` + queue (b) `asyncio` (c) `concurrent.futures` | **`threading.Thread` + `log_queue`** | Follows existing app pattern — printing, backup, guarda all use this. No new infrastructure needed |
| CONTENEDOR matching | (a) Last-5-digit of permiso vs P.E. cell (b) User picks from dropdown (c) Folder name heuristic | **Last-N-digit match, fallback to dropdown** | Proposal spec. PE cell at DATOS R5 C7. Extract numeric suffix, compare. If no match or ambiguous, user selects manually |
| OCR integration | (a) Direct import (b) Subprocess call (c) Refactor into class | **Direct import** | `pdf_a_texto` and `extraer_datos` are pure functions — zero refactor needed. Only import `procesar_tickets` |
| Excel cell writing | (a) `openpyxl` direct cell write (b) Full row rewrite (c) xlwings | **`openpyxl` cell assignment** | We know the exact cell coordinates (label at col 1, value at col 7 per contenedor block). Minifies diff |
| Config storage | (a) New config section (b) Add to existing `correo` section | **`config["correo"]["remitente_balanza"]`** | It's an email filter — belongs with other mail config. Persisted in `_guardar_ajustes` |

## Data Flow

```
PDF source (IMAP or local picker)
       │
       ▼
Worker thread:
  ┌─────────────────────────────────────────┐
  │ 1. OCR: pdf_a_texto → extraer_datos    │
  │    → dict: patente, semi, neto, tara,  │
  │            permiso, conductor, dni      │
  │ 2. Find CONTENEDOR file:                │
  │    - Scan planillas_carga subdirs       │
  │    - Match last N digits vs P.E. cell   │
  │ 3. Read DATOS sheet:                    │
  │    - Locate matching contenedor block   │
  │    - Read expected values               │
  │ 4. Push result dict → queue             │
  └──────────────┬──────────────────────────┘
                 │ queue
                 ▼
Main thread (self.after):
  ┌─────────────────────────────────────────┐
  │ 5. Populate TreeView:                   │
  │    - One row per ticket                 │
  │    - Fields: Patente, Semi, Conductor,  │
  │      DNI, Neto (OCR vs Excel), Tara     │
  │    - Green/red tags per cell            │
  │ 6. "Escribir en CONTENEDORES" button    │
  │    → worker thread writes cells         │
  └─────────────────────────────────────────┘
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `ui_app.py` | Modify | +_panel_cargar_datos (~200 lines), sidebar/panel shift, +_cargar_datos_worker, +remitente_balanza field in `_ajustes_tab_correo` |
| `requirements.txt` | Modify | Add `pytesseract`, `pdf2image` |
| — | External | Tesseract-OCR + poppler binaries (system) |

## Interfaces / Contracts

### New UI panel: `_panel_cargar_datos(self, parent)`

Widget tree:
```
CTkFrame (parent)
 ├── Toolbar frame (CTkFrame)
 │    ├── CTkButton "Buscar en Correo" → IMAP fetch
 │    ├── CTkButton "Seleccionar PDFs" → filedialog
 │    └── CTkProgressBar (indeterminate during OCR)
 ├── CTkScrollableFrame (results)
 │    └── CTkTreeview (columns: campo, ticket, contenedor, estado)
 └── CTkButton "Escribir en CONTENEDORES" (disabled until validated)
```

### New worker: `_cargar_datos_worker(self, fuente: str, rutas: list[str])`

Signature matches the existing pattern:
```python
def _cargar_datos_worker(self, fuente, rutas):
    """fuente: "correo" or "local". rutas: list of PDF paths."""
    # 1. For each PDF: OCR → dict
    # 2. Find CONTENEDOR match by permiso digit match
    # 3. Read DATOS for expected values
    # 4. Push (index, ticket_data, cont_data) to log_queue
    # 5. Final push: ("_DONE_", resultados)
```

### Config key
```python
self.config["correo"]["remitente_balanza"] = "balanza@example.com"
```

### CONTENEDOR matching heuristic
```python
def _match_contenedor(self, permiso_ticket: str) -> str | None:
    """Match by last 5 alphanumeric chars of permiso against PE cell (R5 C7 in DATOS).
    Returns file path or None."""
```

### Excel write coordinates
DATOS sheet layout per contenedor block (offset k = 0..4, col = 1+12*k):
- Label at R{k*13+20}, value at same row, col+6
- PESO CARGA label at R{k*13+28}, value at col+6 → Neto
- TARA CONT label at R{k*13+29}, value at col+6 → Tara

## Testing Strategy

No existing test framework. Verification by manual integration:

| Layer | What | How |
|-------|------|-----|
| Unit | `pdf_a_texto` + `extraer_datos` | Already tested at `procesar_tickets.py` CLI |
| Integration | Panel flow | Load 2 known PDFs → verify TreeView shows correct OCR vs CONTENEDOR values |
| Integration | Excel write | Click "Escribir" → open Excel → verify PESO CARGA and TARA CONT cells updated |
| Smoke | IMAP filter | Configure remitente_balanza → "Buscar en Correo" → verify only matching sender's PDFs appear |
| Manual | Edge cases | Mismatched permit, missing CONTENEDOR, invalid PDF, no Tesseract binary |

## Migration / Rollout

No migration. Pure additive UI change. Install Tesseract-OCR (with `spa` language pack) and poppler on target machine. `pip install pytesseract pdf2image`. Detect missing binaries at panel open and show inline error with download link.

## Open Questions

None.
