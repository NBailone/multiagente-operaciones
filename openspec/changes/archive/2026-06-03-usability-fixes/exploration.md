# Exploration: usability-fixes

## Summary

7 bugs/usability issues explored. 6 trivial, 1 needs design discussion.

---

## Issue 1: Email body "PLANILLA DE CARGA:" line

- **File**: `ui_app.py:5402`
- **Root cause**: Hardcoded extra line in email body string
- **Fix**: Delete `cuerpo += "PLANILLA DE CARGA:\n"`
- **Risk**: Low. One-line deletion.

## Issue 2: ATA y TARES hardcoded 60000

- **Files**: `ui_app.py:4077, 4220, 4424, 4447`
- **Root cause**: 4 occurrences of `* 60000` hardcoded. Should be 65000 and configurable from Ajustes > Valores.
- **Fix**: Follow existing `precio_carpeta` pattern — add field to Ajustes UI, read from config, replace 4 occurrences.
- **Risk**: Low.

## Issue 3: Dialog "Correos Despachados" button cutoff

- **File**: `ui_app.py:5507-5572` (`_correos_popup_confirmacion`)
- **Root cause**: Window height max 480px — with many items, the button gets pushed below visible area
- **Fix**: Increase max height or restructure to place button outside scrollable area
- **Risk**: Low.

## Issue 4: Drive tab in Ajustes

- **File**: `ui_app.py:5998, 6347-6386`
- **Finding**: The Drive tab already exists with Client ID, Secret, Folder ID fields. If user wanted it removed — that's the fix. If they want ADDITIONAL Drive features — that's a feature request.
- **Fix**: Confirm with user whether to remove the tab or add functionality.
- **Risk**: Low.

## Issue 5: Dead Backup button

- **File**: `ui_app.py:377-378, 3248-3255`
- **Root cause**: Button created with `state="disabled"` and no `command`. Placeholder for Drive backup that was never implemented.
- **Fix**: Remove the dead button (recommended — user confirmed Drive backup won't be done).
- **Risk**: Low.

## Issue 6: xlutils missing dependency

- **Files**: `ui_app.py:33-37` (auto-install), `requirements.txt`, `ui_app.py:3112, 3823` (runtime import)
- **Root cause**: `xlutils` used at runtime for `.xls` files but not in auto-install deps or requirements.txt
- **Fix**: Add to `_instalar_deps_ui()` and `requirements.txt`
- **Risk**: Low.

## Issue 7: Guarda not being added to Excel

- **Files**: `ui_app.py:3108-3181` (super-auto), `ui_app.py:3819-3890` (manual)
- **Root cause discovered**:
  - `.xls` paths have `isinstance(val, str)` check that fails if xlrd returns non-string type → guarda write silently skipped
  - Hardcoded row range 1-15 may miss guarda cells beyond row 15
  - 4 separate code paths with inconsistent matching logic
- **Fix**: Unify matching across all paths, expand row range, remove isinstance restriction
- **Risk**: Medium. Needs actual Excel template inspection to confirm column mapping.
