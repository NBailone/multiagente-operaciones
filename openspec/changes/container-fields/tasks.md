# Tasks: Container Fields for FLEXI/ISO Tickets

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~40 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

---

## Phase 1: OCR Pipeline — `procesar_tickets.py`

- [ ] **1.1** Append `"Contenedor"` and `"Tara Contenedor"` to the end of `CAMPOS` list (L89). *File: `procesar_tickets.py`*
- [ ] **1.2** Add regex extraction in `extraer_datos()` after L635: `datos["Contenedor"] = buscar(r"Sigla Contenedor[:\s]+([A-Z]+\s*\d+[\s-]*\d*)")` and `re.search(...)` for tare weight with `Cert.Verif.INTI` lookahead. *File: `procesar_tickets.py`*
- [ ] **1.3** Add `"Contenedor": 20` and `"Tara Contenedor": 16` to the `anchos` dict in `crear_excel()` (L831). *File: `procesar_tickets.py`*

## Phase 2: UI Worker — `ui_app.py`

- [ ] **2.1** In `_cargar_datos_worker()` (L7525-7560): extract `datos.get("Contenedor", "")` and `datos.get("Tara Contenedor", "")`, convert tare to float (try/except), add `"contenedor"` and `"tara_contenedor"` to `ticket_data` dict between `"tara"` and `"permiso"`. *File: `ui_app.py`*

## Phase 3: UI Display — TreeView + Comparison

- [ ] **3.1** Add `"contenedor"` and `"tara_contenedor"` columns to TreeView in `_panel_cargar_datos()` (L6684-6688), between `"tara"` and `"permiso"`, with headers `"Contenedor"` / `"Tara Cont."` and widths 100 / 90. *File: `ui_app.py`*
- [ ] **3.2** Add container fields to values tuples in `_procesar_resultado_ocr()`: no-match branch (L6069) and match branch (L6113-6118), between `tara` and `permiso` positions. *File: `ui_app.py`*
- [ ] **3.3** Add `"Contenedor"` and `"Tara Contenedor"` to comparison data dicts (L6126-6155) in `ticket`, `contenedor`, and `ok` sub-dicts between `"Tara (kg)"` and `"Permiso"`. Add comparison booleans `ok_contenedor` and `ok_tara_contenedor` (L6091-6097). Add both fields to `campos` list (L6207-6208). *File: `ui_app.py`*

---

## Summary

| Phase | Tasks | Focus |
|-------|-------|-------|
| Phase 1 | 3 | OCR regex + CAMPOS + Excel widths |
| Phase 2 | 1 | Worker data propagation |
| Phase 3 | 3 | TreeView columns + values + comparison popup |
| **Total** | **7** | |

### Implementation Order
1.1 → 1.2 → 1.3 (sequential in `procesar_tickets.py`), then 2.1 (worker), then 3.1 → 3.2 → 3.3 (UI). Phases 1-3 are strictly ordered — each depends on the prior.

### Review Workload Forecast
- **Estimated changed lines**: ~40
- **400-line budget risk**: Low
- **Chained PRs recommended**: No
- **Delivery strategy**: single-pr
- **Decision needed before apply**: No

### Next Step
Ready for implementation (`sdd-apply`).
