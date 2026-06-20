# Design: Planilla de Carga

## Technical Approach

Add a "Planilla de Carga" button to the planillas toolbar that opens a multi-select popup. The popup scans the Desktop for folders matching the batch naming convention, presents checkboxes, and on confirmation copies the "planilla de cargo" sheet from each folder's `Contenedores.xlsx` into a new workbook. Processing runs in a daemon thread with log queue progress, cancel support, and an error summary modal on completion.

Implementation lives entirely in `ui_app.py` following the existing `_popup_agregar_guarda` / `_backup_pendrive_worker` pattern: modal CTkToplevel → checkbox scroll → thread worker → log_queue → done callback.

## Architecture Decisions

| Option | Tradeoff | Decision |
|--------|----------|----------|
| New class in `utils/` vs methods on `App` | Extract: clean separation, but requires passing `log_queue`, `after()`, config. Inline: follows existing popup+worker pattern in ui_app.py. | **Methods on `App`** — matches every other popup in the codebase. Extract later if pattern repeats. |
| `openpyxl` copy_sheet vs manual cell copy | `copy_sheet` preserves formatting/merges/widths. Manual copy loses formatting. | **`openpyxl` copy_sheet** — formatting matters for operational sheets. |
| Scrollable checkboxes in popup vs Treeview | Treeview adds complexity ( ttk styling, no native checkbox). CTkCheckBox in CTkScrollableFrame is the existing pattern (`_refrescar_lista_backup`). | **CTkCheckBox + CTkScrollableFrame** — proven pattern at line 3590. |
| Process all vs user-selected folders | Process all: faster but no control. User-selected: matches backup pattern. | **User-selected** — popup with checkboxes, "Seleccionar todo" helper. |

## Data Flow

    ┌───────────────┐     ┌──────────────────────┐     ┌──────────────────┐
    │ btn_planilla  │────▶│ _popup_planilla_carga │────▶│ Worker thread    │
    │ de_carga      │     │ (scan Desktop,        │     │ (for each folder │
    │ (line ~1819)  │     │  checkboxes)          │     │  copy sheet)     │
    └───────────────┘     └──────────────────────┘     └────────┬─────────┘
                                                                │
                                                     log_queue.put() × N
                                                                │
                                                     self.after(0, cb)
                                                                │
                                                     ┌──────────▼──────────┐
                                                     │ _poll_log_queue     │
                                                     │ → lbl_estado + log  │
                                                     └──────────┬──────────┘
                                                                │ done
                                                     ┌──────────▼──────────┐
                                                     │ _planilla_carga_done│
                                                     │ → error summary     │
                                                     │   modal if errors   │
                                                     └─────────────────────┘

**Per-folder worker flow:**
1. List Desktop folders → filter by `DD_MM_YYYY_*` pattern
2. For each selected folder:
   a. Find `Contenedores.xlsx` (case-insensitive match)
   b. Open with openpyxl, find sheet "planilla de cargo" (case-insensitive)
   c. Copy sheet to new workbook
   d. Build output filename from folder name parts
   e. Save in same folder as source
   f. Log success/error

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `ui_app.py` ~line 1818 | Modify | Add `btn_planilla_carga` button after `btn_editar_excels` |
| `ui_app.py` (after `_popup_editar_excels` ~line 3925) | Modify | Add `_popup_planilla_carga()` — scan Desktop, build checkbox popup |
| `ui_app.py` (after popup method) | Modify | Add `_planilla_carga_ejecutar(seleccionadas)` — set tarea_activa, launch thread |
| `ui_app.py` (after ejecutar) | Modify | Add `_planilla_carga_worker(seleccionadas)` — thread target, openpyxl logic |
| `ui_app.py` (after worker) | Modify | Add `_planilla_carga_done(errores)` — re-enable UI, show summary if errors |

No new files. No changes to `utils/`.

## Interfaces / Contracts

```python
# New methods on App class:

def _popup_planilla_carga(self):
    """Scan Desktop for batch folders, show multi-select popup."""

def _planilla_carga_ejecutar(self, seleccionadas):
    """Validate selection, set tarea_activa, launch worker thread."""

def _planilla_carga_worker(self, carpetas):
    """Daemon thread: for each folder, copy 'planilla de cargo' sheet."""

def _planilla_carga_done(self, errores):
    """Callback on main thread: re-enable UI, show error summary modal."""

# Filename generation:
# Input:  "18_06_2026_P1_PERMISO_CONTAINER_COMPANY"
# Parts:  ["18","06","2026","P1","PERMISO","CONTAINER","COMPANY"]
# Output: "PLANILLA DE CARGA_P1_PERMISO_CONTAINER_COMPANY.xlsx"

# Sheet matching: case-insensitive lookup in wb.sheetnames
#   target = "planilla de cargo"
#   match = next((s for s in wb.sheetnames if s.strip().lower() == target), None)
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Manual | Button appears after "Editar Excels" | Visual verification in toolbar |
| Manual | Desktop scan finds matching folders | Create test folders with correct naming, verify popup shows them |
| Manual | Checkboxes toggle correctly | Select/deselect, verify "Seleccionar todo" works |
| Manual | Sheet copy preserves formatting | Open output file, compare with source sheet visually |
| Manual | Filename generation from folder name | Verify `parts[5:]` extraction produces correct output name |
| Manual | Error handling: missing Contenedores.xlsx | Folder without file → logged as error, other folders continue |
| Manual | Error handling: sheet not found | Contenedores without "planilla de cargo" sheet → logged as error |
| Manual | Cancel support | Start processing, cancel → remaining folders skipped |
| Manual | File locked by Excel → retry dialog | Open output in Excel, re-run → retry/cancel dialog appears |
| Manual | Empty Desktop → no folders | Popup shows "No se encontraron carpetas" message |

No automated test infrastructure exists (per `openspec/config.yaml`). Manual verification only.

## Migration / Rollout

No migration required. This is a purely additive feature — new button, new popup, new worker. No existing behavior changes. No config schema changes. No data migration.

## Open Questions

- [ ] Should the output file be opened automatically after creation, or just logged? (Proposal says copy only — follow that.)
- [ ] What if a `PLANILLA DE CARGA_*.xlsx` already exists in the target folder? Overwrite silently or warn? (Suggest: overwrite — same batch re-run scenario.)
- [ ] Should "Seleccionar todo" be a checkbox at top, or select-all/deselect-all buttons? (Follow backup pattern: CTkCheckBox for each item, no select-all — keep simple.)
