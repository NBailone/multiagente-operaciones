# Design: Usability Fixes (7 items)

## 1. Unified Guarda Logic (Issue 7) — Architectural Design

### Problem

Four code paths exist for writing the guarda name to Excel files, with inconsistent matching logic:

| Path | File | Lines | Matching | Bug |
|------|------|-------|----------|-----|
| Super-auto .xls | `ui_app.py` | 3126-3133 | `isinstance(val, str) and "GUARDA" in val.strip().upper()` | Skips float cells silently |
| Super-auto .xlsx | `ui_app.py` | 3158-3170 | `"GUARDA" in str(val).strip().upper()` | OK |
| Manual .xls | `ui_app.py` | 3840-3846 | `val and isinstance(val, str) and "GUARDA" in val.strip().upper()` | Skips float cells silently |
| Manual .xlsx | `ui_app.py` | 3868-3883 | `str(val).strip().upper() == "GUARDA" or str(val).strip().upper().startswith("GUARDA")` | OK but redundant logic |

Root causes:
- `.xls` paths use `isinstance(val, str)` guard that fails when xlrd returns a float for cells that visually contain text (xlrd infers type from cell format)
- Row 1 is searched (super-auto) when data starts at row 2
- Manual .xlsx uses verbose `== "GUARDA" or .startswith("GUARDA")` instead of `"GUARDA" in str(val).upper()`
- No shared helper — any fix must be applied 4 times

### Solution: Shared Helper Function

Extract a single helper method `_escribir_guarda_en_archivo` that handles both `.xls` and `.xlsx` internally. Both super-auto and manual paths call this one function.

### Helper Signature and Contract

```python
def _escribir_guarda_en_archivo(
    self, ruta: str, guarda_nombre: str, log_func: Optional[Callable] = None
) -> bool:
    """Write guarda_nombre to column H in the CHOFER sheet of an Excel file.

    Algorithm:
        1. Detect format: .xls → xlrd + xlutils, .xlsx → openpyxl
        2. Find sheet whose name contains "CHOFER" (case-insensitive)
        3. For row in range(2, 16) (rows 2–15 inclusive):
            a. Read cell value from column G (7 openpyxl / 6 xlrd)
            b. Convert to str, strip whitespace, uppercase
            c. If "GUARDA" is a substring of the result → match found
            d. Write guarda_nombre to column H (8 openpyxl / 7 xlrd), same row
            e. Break loop
        4. Save workbook
        5. Return True if written, False if no match found

    Args:
        ruta: Absolute path to .xls or .xlsx file
        guarda_nombre: Name string to write into the cell
        log_func: Optional callable for logging messages

    Returns:
        True if guarda name was written to column H, False if:
        - No "CHOFER" sheet found
        - No cell in column G rows 2-15 contains "GUARDA"
    """
```

### Corner Cases and Edge Cases

| Case | Behavior |
|------|----------|
| **No "GUARDA" found in any cell** | Returns False. Caller logs "Guarda no hallada" and continues. No file is modified. |
| **Multiple cells contain "GUARDA"** | Only the first match (lowest row number) is used. Break after write. |
| **Merged cells in openpyxl (.xlsx)** | Follow existing pattern: if cell value is None, check `merged_cells.ranges` and read from the top-left cell of the merged range. |
| **File locked by another process** | Let the existing retry wrapper handle it (caller already has `_abrir_excel_seguro` and retry logic). |
| **xlrd returns float for text cell (.xls)** | `str(val).strip().upper()` handles this correctly — no `isinstance` guard needed. |
| **Cell is empty/None** | `str(None)` → `"None"` which fails the `"GUARDA" in` check. `str("")` → `""` which also fails. Safe. |
| **Cell value is "GUARDAS" or "GUARDA:"** | `"GUARDA" in "GUARDAS"` → True. `"GUARDA" in "GUARDA:"` → True. Correct — these are variants of the same label. |
| **Row count less than 15** | xlrd: `cell_value()` returns `""` for out-of-range rows (already handled by `if row < rs.nrows` check). openpyxl: returns None. Both fail the match check. |

### Column Mapping (Critical)

| Library | Column G (search) | Column H (write) |
|---------|-------------------|-------------------|
| xlrd (0-indexed) | 6 | 7 |
| openpyxl (1-indexed) | 7 | 8 |

### Save Behavior

- `.xlsx`: Use `self._guardar_excel_seguro(wb, ruta)` (existing method via openpyxl)
- `.xls`: Use `wb_w.save(ruta)` then `rb.release_resources()` (existing xlutils pattern)

### Where to Insert the Helper

Place the helper method in the class between the existing methods that use it. It should go near `_super_aplicar_guarda` (line ~3189) or right before it.

### Caller Changes

**Super-auto path** (`_super_aplicar_guarda`, lines 3108-3181):
Replace the entire try/except body for EACH file with:
```python
if not self._escribir_guarda_en_archivo(ruta, self._super_guarda, self.log_queue.put):
    self.log_queue.put(f"[...]   ⚠ 'Guarda' no hallada en {archivo[:50]}")
else:
    aplicados += 1
```

**Manual path** (`_guarda_worker`, lines 3819-3890):
Replace the entire try/except body for EACH file with the same call:
```python
self._escribir_guarda_en_archivo(ruta_contenedores, guarda_elegido)
```

### Row Indexing Note

The current code uses `range(1, 16)` which includes row 1. Row 1 is typically headers. The unified helper uses `range(2, 16)` (rows 2-15) to match the data area only. This matches the user's confirmation that rows vary from 2 to 15.

---

## 2. Other Items — Mechanical Changes

### Item 1: Remove "PLANILLA DE CARGA:" line from email body

- **File**: `ui_app.py`
- **Line**: 5402
- **Change**: Delete the line `cuerpo += "PLANILLA DE CARGA:\n"`
- **Result**: Email body starts with "Se adjuntan las planillas de carga correspondientes:" followed directly by the file list
- **Risk**: None. Cosmetic change only.

### Item 2: Add Servicio ATA ($) configurable field

**A. Add UI field** in `_ajustes_tab_valores()` (after line 6302, following `_ent_precio_carpeta` pattern):
```python
self._ent_servicio_ata = self._ajustes_row(
    parent, "Servicio ATA ($):",
    str(self._cfg_obtener("valores", "servicio_ata", 65000)),
    extra="Valor del servicio ATA por contenedor.",
    width=120,
)
```

**B. Add save logic** in `_guardar_ajustes()` (inside the `valores` dict, after `precio_carpeta`):
```python
try:
    servicio_ata = int(self._ent_servicio_ata.get().strip())
except ValueError:
    servicio_ata = 65000
```
Then add `"servicio_ata": servicio_ata` to the `self.config["valores"]` dict.

**C. Replace 4 hardcoded `* 60000` occurrences:**

| Location | Line | Current | Replacement |
|----------|------|---------|-------------|
| `planillas_core()` — `d1` dict | 4077 | `"servicio": cant_final * 60000` | `"servicio": cant_final * int(self._cfg_obtener("valores", "servicio_ata", 65000))` |
| `planillas_core()` — `valores_check_1` | 4220 | `d1["cant_final"] * 60000` | `d1["cant_final"] * int(self._cfg_obtener("valores", "servicio_ata", 65000))` |
| `_completar_cobro()` — `servicio_ata` | 4424 | `d1["cant_final"] * 60000` | `d1["cant_final"] * int(self._cfg_obtener("valores", "servicio_ata", 65000))` |
| `_completar_cobro()` — `servicio_ata_2` | 4447 | `d2["cant_final"] * 60000` | `d2["cant_final"] * int(self._cfg_obtener("valores", "servicio_ata", 65000))` |

**D. Add default** to `ui_config.json`:
```json
"servicio_ata": 65000
```
Inside the `valores` object, after `"precio_carpeta": 49000`.

**Note**: Consider extracting `self._get_servicio_ata()` helper or computing once per batch to avoid 4x config reads, but since `_cfg_obtener` is an in-memory dict lookup, performance impact is negligible.

### Item 3: Increase "Correos Despachados" dialog height

- **File**: `ui_app.py`
- **Line**: 5512
- **Change**: `h = min(180 + n * 36, 480)` → `h = min(180 + n * 36, 560)`
- **Rationale**: With ~10 items, height reaches 540px, so 560 gives room. The "Cerrar" button (packed after the items list at line 5566) becomes visible without scrolling.
- **Risk**: None. Dialog was already using dynamic height calculation.

### Item 4: Eliminate Drive tab from Ajustes

**All removals in `ui_app.py`:**

| What | Lines | Change |
|------|-------|--------|
| Tab name entry in tab_names list | 5998 | Delete `("drive", "💾  Drive"),` |
| Tab content function call | 6041 | Delete `self._ajustes_tab_drive(self._ajustes_frames["drive"])` |
| `_ajustes_tab_drive()` method body | 6347-6386 | Delete entire method OR replace with `pass` plus a comment noting the method is intentionally empty |
| `_drive_verificar_conexion()` method | 6388+ | Delete entire method (called only from `_ajustes_tab_drive()`, which is removed) |
| `_drive_autenticar()` method | 3454-3495 | Delete entire method (called only from Drive backup path, which is also removed) |
| `_drive_buscar_o_crear_carpeta()` method | 3497-3507 | Delete entire method |
| `_drive_subir_sobrescrito()` method | 3509-3519 | Delete entire method |
| `_backup_drive_iniciar()` method | 3434-3452 | Delete entire method |
| `_backup_drive_worker()` method | 3522-3591 | Delete entire method |
| Google Drive save block in `_guardar_ajustes()` | 6558-6566 | Delete entire block (lines setting `self.config["google_drive"]`) |

**In `ui_config.json`:**

| What | Lines | Change |
|------|-------|--------|
| `google_drive` section | 58-62 | Delete the entire `"google_drive": { ... },` block |

**Cleanup consideration**: The `_backup_done()` method (line 3593) references `self.btn_backup_drive.configure(...)` — this is handled by Item 5 (button removal). Since the button is being removed in Item 5, that reference disappears naturally.

### Item 5: Remove dead Backup Drive button

**All removals in `ui_app.py`:**

| What | Lines | Change |
|------|-------|--------|
| Button creation + pack | 3248-3255 | Delete `self.btn_backup_drive = ctk.CTkButton(...)` and `.pack(...)` |
| Disable during backup pendrive | 3345 | Delete `self.btn_backup_drive.configure(state="disabled")` |
| `_backup_done()` reference | 3595 | Delete `self.btn_backup_drive.configure(text="📀  Backup", state="disabled", fg_color=Palette.BG_HOVER)` |

**Ordering constraint**: Item 5 must be applied before Item 4 (Drive elimination), or at minimum the `btn_backup_drive` references must be removed before or simultaneously with removing the methods that create/configure it.

### Item 6: Add xlutils dependency

**`ui_app.py` — `_instalar_deps_ui()`** (line 37):
Add `("xlutils", "xlutils")` to the deps list:
```python
deps = [
    ("customtkinter", "customtkinter"),
    ("openpyxl", "openpyxl"),
    ("xlrd", "xlrd"),
    ("win32com", "pywin32"),
    ("xlutils", "xlutils"),       # <-- add
]
```

**`requirements.txt`** (line 1, after `python-dotenv`):
```text
xlutils>=1.2.0
```

---

## 3. Change Dependencies and Ordering

```
Item 6 (xlutils) ─────→ Item 1 (email) ──┐
                                          ├──→ All independent, any order
Item 2 (ATA config) ─────────────────────┘

Item 5 (dead button) ──→ Item 4 (Drive tab)
                              ↑
                        Must be after Item 5
                              ↓
Item 3 (dialog height) ──→ Independent

Item 7 (Guarda helper) ──→ Independent (but replaces 4 code paths)
```

**Safe apply order**: 6, 1, 2, 3, 5, 4, 7. However, Items 5 and 4 touch overlapping areas (both reference `btn_backup_drive`). Apply them together in one focused pass.

---

## 4. Risks

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| `.xls` files with no row limit set in xlrd (reading beyond nrows) | Low | Crash | Already handled: `if row < rs.nrows` guard in current code, retained in helper |
| Merged cells in .xls (xlrd): current code does NOT handle merged cells in .xls, only .xlsx | Low | Silent miss | Document as known limitation. xlrd merged cell API differs from openpyxl. Add TODO comment. |
| Config files with `google_drive` key survive removal | Medium | No crash, dead config | `_cargar_config()` loads JSON into dict, extra keys are ignored. Cleanup is cosmetic. |
| User's running config has `btn_backup_drive` referenced from `_backup_done` callback | Low | AttributeError if code runs before removal | Remove button AND references in same commit. |
| `servicio_ata` default (65000) differs from current hardcoded 60000 — user gets new default | Medium | Changed behavior | Explicitly noted in proposal. User confirmed 65000 is correct. |

---

## 5. Unresolved Decisions / Assumptions

1. **Guarda row range**: Assumed rows 2-15 inclusive. Confirmed by user. No template inspection performed.
2. **Column G always contains "GUARDA" text, not label in another column**: Assumed based on user confirmation. If templates vary, the helper's `"GUARDA" in str(val)` check is broad enough to catch most variants.
3. **Only one guarda per file**: The `break` after first match assumes a single guarda per file. If templates ever require multiple guarda entries, this logic needs revisiting.
4. **Drive tab removal is wanted**: User confirmed via exploration. Entire Drive feature is removed, not just the tab.
5. **xlutils version**: Added as `>=1.2.0` (latest stable). If the user's environment has restrictions, pin accordingly.
