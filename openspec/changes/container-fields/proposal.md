# Proposal: Container Fields for FLEXI/ISO Tickets

## Intent

AGD weighing tickets for FLEXI/ISO containerized cargo include container number ("Sigla Contenedor:") and container tare weight ("Tara Contenedor:"). These fields are currently ignored. The system must extract and expose them so they can be validated against CONTENEDOR Excel data and written to the output.

## Scope

### In Scope
- Add "Contenedor" and "Tara Contenedor" to the `CAMPOS` list in `procesar_tickets.py`
- Add regex extraction in `extraer_datos()` for both fields
- Add column widths to the `anchos` dict in `crear_excel()`
- Add fields to `ticket_data` dict in `_cargar_datos_worker()`
- Add columns to TreeView in `_panel_cargar_datos()`
- Add fields to comparison popup in `_abrir_comparacion()`

### Out of Scope
- Transport type detection (terrestrial vs. FLEXI/ISO) — container fields are universal; they'll be empty for terrestrial tickets
- Changes to CONTENEDOR Excel matching logic (`_leer_datos_contenedor`, `_match_contenedor`)
- Writing to CONTENEDOR DATOS sheet (would be covered by a future change)
- Modifications to `_cargar_datos_escribir()`

## Capabilities

### New Capabilities
- None

### Modified Capabilities
- `ticket-pesada-ocr`: Extend field extraction from 7 to 9 fields (add `Contenedor` and `Tara Contenedor`)
- `app-ui`: Extend "Cargar Datos" panel to display and compare the 2 new fields

## Approach

Add the fields universally to `CAMPOS` and `extraer_datos()` in `procesar_tickets.py`. For terrestrial tickets, the regex won't match so fields remain empty — no branching needed.

**Container number** (`Sigla Contenedor:`): Simple label + value on same line → regex `Sigla Contenedor[:\s]+(.+)`.

**Container tare** (`Tara Contenedor:`): The value appears *before* the label — a floating number (e.g. `2.100`) between `Vencimiento:` and `Cert.Verif.INTI`. Extract via lookahead: regex `(\d+[.,]\d{3})\s+Cert\.?Verif\.?INTI` to capture the weight preceding "Cert.Verif.INTI".

In `ui_app.py`, propagate the new fields through `ticket_data`, TreeView columns, and the comparison popup dict.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `procesar_tickets.py` | Modified | Add 2 fields to `CAMPOS` (L62-89), 2 regex lines in `extraer_datos()` (after L635), 2 entries in `anchos` dict (L822-832) |
| `ui_app.py` | Modified | `_cargar_datos_worker` ticket_data dict (L7555-7560), TreeView columns (L6684), headers/anchors (L6686-6688), `_procesar_resultado_ocr` values dict + comparison popup (L6126-6154) |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Container tare regex false positive on terrestrial ticket numbers | Low | The pattern `\d+[.,]\d{3}` between `Vencimiento:` and `Cert.Verif.INTI` is specific enough |
| OCR misreads container sigla (ambiguous chars like 0/O, 1/I) | Medium | Apply existing `corregir_patente()`-like normalization if needed in future |

## Rollback Plan

Revert with `git checkout procesar_tickets.py ui_app.py`. The two new fields default to empty strings for all existing data — no migration needed.

## Dependencies

None. Pure Python regex addition within existing code.

## Success Criteria

- [ ] OCR extraction populates "Contenedor" and "Tara Contenedor" from FLEXI/ISO ticket PDFs
- [ ] Fields appear in the Excel output columns
- [ ] Fields appear in the "Cargar Datos" TreeView and comparison popup
- [ ] Terrestrial tickets show empty container fields (no regression)
