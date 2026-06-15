# Design: Container Fields for FLEXI/ISO Tickets

## Technical Approach

Extend the OCR and UI pipelines to capture two new fields (`Contenedor`, `Tara Contenedor`) from AGD weigh tickets for FLEXI/ISO containerized cargo. Add at the end of `CAMPOS` to preserve backward compatibility with existing Excel files; terrestrial tickets naturally yield empty strings since their OCR text lacks the container labels.

## Architecture Decisions

### Decision: Append to CAMPOS (no insert)

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Append at end | Columns shift right in existing Excel; no data corruption | ✅ Chosen — existing positional code iterates `CAMPOS` by name, not index |
| Insert in middle | Logical grouping; changes column order in existing files | ❌ Rejected — could mix up columns in archived sheets |

### Decision: Container number regex — direct label match

`extraer_datos()` uses the existing `buscar()` helper for label+value on the same line. Pattern `Sigla Contenedor[:\s]+([A-Z]+\s*\d+[\s-]*\d*)` handles formats like `MSMU 258531-2` (4 letters, space, 6 digits, hyphen, check digit). The character class `[A-Z]` avoids matching garbage OCR.

### Decision: Container tare regex — lookahead without label

The tare value (`2.100`) appears *before* its label `Tara Contenedor:`. A regex anchored to the surrounding context (`Cert.Verif.INTI`) works reliably:

- Pattern `(\d+[.,]\d{3})\s+Cert\.?Verif\.?INTI\s*-\s*Balanza\s+Egr`
- The `\.?` after `Cert` tolerates OCR reading `Cert.Verif.INTI` or `CertVerifINTI`
- The group captures `2.100` (dot as thousands sep, meaning 2100 kg) — consistent with how all weights in the system use `.` as thousands separator

### Decision: No transport type branching

Container regexes run unconditionally for every ticket. Terrestrial tickets lack the labels, so `buscar()` returns `""` naturally. Zero branching reduces complexity and risk.

## Data Flow

```
PDF → OCR text → extraer_datos()
                    │
      ┌─────────────┼─────────────┐
      │             │             │
      ▼             ▼             ▼
  datos[...]    datos[         datos[
  (7 existing   "Contenedor"]  "Tara Contenedor"]
   fields)      "MSMU 258531-2"  "2.100"
      │             │             │
      └─────────────┼─────────────┘
                    │
                    ▼
          ticket_data dict
          (ui_app worker)
                    │
          ┌─────────┴─────────┐
          ▼                   ▼
   TreeView row      _cargar_datos_comparacion
   (display)         (double-click popup)
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `procesar_tickets.py` | Modify | Add 2 CAMPOS (L89), 2 regex lines (L636), 2 anchos (L831) |
| `ui_app.py` | Modify | Add 2 fields in worker tree_data, TreeView columns, values tuple, comparison dict, campos list |

## Interfaces / Contracts

### CAMPOS addition (procesar_tickets.py, L89)

```python
CAMPOS = [
    ...  # existing 27 fields
    "Contenedor",        # 28 — ISO/FLEXI container number
    "Tara Contenedor",   # 29 — container tare weight (kg)
]
```

### Regex extraction (procesar_tickets.py, after L635)

```python
# Número de contenedor ISO/FLEXI: "Sigla Contenedor: MSMU 258531-2"
datos["Contenedor"] = buscar(r"Sigla Contenedor[:\s]+([A-Z]+\s*\d+[\s-]*\d*)")

# Tara Contenedor: valor antes del label
m_tc = re.search(r'(\d+[.,]\d{3})\s+Cert\.?Verif\.?INTI\s*-\s*Balanza\s+Egr', text, re.IGNORECASE)
if m_tc:
    datos["Tara Contenedor"] = m_tc.group(1)
```

### Excel column widths (procesar_tickets.py, L831)

```python
"Contenedor": 20,
"Tara Contenedor": 16,
```

### Worker data (ui_app.py, L7555)

```python
contenedor_str  = datos.get("Contenedor", "")
tara_cont_str   = datos.get("Tara Contenedor", "")

# Convert tara to float (consistent with neto/tara pattern)
try:
    tara_contenedor = float(tara_cont_str.replace('.', '').replace(',', '.')) if tara_cont_str else 0
except (ValueError, AttributeError):
    tara_contenedor = 0

ticket_data = {
    "archivo": stem,
    "patente": patente, "semi": semi,
    "conductor": conductor, "dni": dni,
    "neto": neto, "tara": tara,
    "contenedor": contenedor_str,        # NEW — raw string
    "tara_contenedor": tara_contenedor,  # NEW — float
    "permiso": permiso,
}
```

### TreeView columns (ui_app.py, L6684-6688)

Insert `"contenedor"` and `"tara_contenedor"` after `"tara"` and before `"permiso"`:

```python
columns = ("archivo", "patente", "semi", "conductor", "dni",
           "neto", "tara", "contenedor", "tara_contenedor", "permiso", "estado")
headers = ("Archivo", "Patente", "Semirremolque", "Conductor", "DNI",
           "Neto", "Tara", "Contenedor", "Tara Cont.", "Permiso", "Estado")
anchos  = (180, 80, 80, 100, 80, 70, 70, 100, 90, 100, 100)
```

### Values tuples (ui_app.py, L6069 match / L6113-6118 no-match)

**No-match branch** (L6069):
```python
valores = (archivo, patente, semi, conductor, dni,
           neto_ocr, tara_ocr, contenedor_str, tara_cont_str, permiso, estado)
```

**Match branch** (L6113-6118):
```python
valores = (
    f"📄 {archivo}",
    patente, semi, conductor, dni,
    f"{neto_ocr:.0f}", f"{tara_ocr:.0f}",
    contenedor_str, f"{tara_cont_str:.0f}",
    permiso, estado,
)
```

Note: `tara_cont_str` is already a float in ticket_data; `contenedor_str` stays as raw string.

### Comparison dict (ui_app.py, L6126-6155)

Add between "Tara (kg)" and "Permiso" in all three sub-dicts:

```python
"ticket": {
    ...
    "Tara (kg)": f"{tara_ocr:.0f}",
    "Contenedor": contenedor_str,
    "Tara Contenedor": f"{tara_cont_str:.0f}",
    "Permiso": permiso,
},
"contenedor": {
    ...
    "Tara (kg)": f"{tara_cont:.0f}",
    "Contenedor": camion_match.get("contenedor", ""),
    "Tara Contenedor": str(camion_match.get("tara_contenedor", "")),
    "Permiso": pe_val,
},
"ok": {
    ...
    "Tara (kg)": ok_tara,
    "Contenedor": ok_contenedor,
    "Tara Contenedor": ok_tara_contenedor,
    "Permiso": ok_permiso,
},
```

### Comparison booleans (ui_app.py, L6091-6097)

```python
ok_contenedor      = _raw_txt(contenedor_str) == _raw_txt(camion_match.get("contenedor", ""))
ok_tara_contenedor = _comparar_num(tara_cont_str, camion_match.get("tara_contenedor", 0))
```

Since `_leer_datos_contenedor` is out of scope for this change, `camion_match` will not yet contain `contenedor`/`tara_contenedor` keys. The `get()` returns `""` / `0`, so:
- FLEXI/ISO tickets: OCR has values, Excel side is empty → `ok = False` → shown red (correct — data gap is visible)
- Terrestrial tickets: both sides empty → `ok = True` → shown green (no difference)

### Campos list for comparison popup (ui_app.py, L6207-6208)

```python
campos = ["Patente", "Semirremolque", "Conductor", "DNI",
           "Neto (kg)", "Tara (kg)", "Contenedor", "Tara Contenedor", "Permiso"]
```

## Error Handling

| Condition | Behavior |
|-----------|----------|
| OCR text lacks `Sigla Contenedor:` | `datos["Contenedor"]` = `""` (via `buscar()` default) |
| OCR text lacks `Cert.Verif.INTI` | `datos["Tara Contenedor"]` = `""` (dict default) |
| Container tare not parseable as float | `tara_contenedor` = `0` (try/except in worker) |
| Terrestrial ticket (no container data at all) | Both fields empty/0 — all comparisons return `True` |
| TreeView column count mismatch | Insert wrapped in existing `try/except Exception: pass` |

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | Container number regex | Run `extraer_datos()` with known OCR sample; assert `datos["Contenedor"] == "MSMU 258531-2"` |
| Unit | Container tare regex | Run `extraer_datos()` with text containing `2.100 Cert.Verif.INTI`; assert `datos["Tara Contenedor"] == "2.100"` |
| Unit | Terrestrial ticket regression | Run with non-container text; assert both new fields empty, existing 27 fields unchanged |
| Integration | TreeView column layout | Insert a mock ticket with container data; verify no `TclError` from column count mismatch |
| Manual | Comparison popup | Open comparison for a FLEXI/ISO ticket; verify new rows appear between "Tara (kg)" and "Permiso" |

## Migration / Rollout

No migration required. New fields default to empty strings/0 for all existing data. Excel files created before this change lack the columns, but `crear_excel()` reads `CAMPOS` dynamically — new files will include them.

## Open Questions

None.
