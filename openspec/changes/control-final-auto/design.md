# Design: Control Final Auto Mode

## Technical Approach

Add a CTkSwitch toggle next to `btn_control_final` in the toolbar. When ON, clicking the button triggers auto-scan of Desktop folders instead of the file dialog. The auto-scan finds folders containing Excel `*CONTENEDORES*` files, presents a checkbox popup, then gathers PDFs from selected folders and delegates to the existing `_control_final_worker`.

## Architecture Decisions

### Decision: Switch placement in toolbar

**Choice**: Wrap `btn_control_final` + CTkSwitch in a `CTkFrame` packed into toolbar.

**Alternatives considered**: Place switch in sidebar (like Súper Auto) — rejected because user explicitly wants it next to the button. Place switch inside button area — not possible with CTkButton.

**Rationale**: Keeps toggle visually associated with its button. Frame wrapping preserves horizontal toolbar layout without breaking existing `pack(side="left")` flow.

### Decision: Filename-based detection for auto-scan

**Choice**: Detect PDFs by filename pattern (scan prefix, `\d{2}AR\d+`), not by reading text content.

**Alternatives considered**: Read PDF text in scanner (like worker does) — rejected because it duplicates worker logic and is slow for discovery. Rely solely on worker classification — rejected because worker needs Excel present to be useful; we must filter folders first.

**Rationale**: The auto-scan is a fast discovery pass to find candidate folders. The worker handles the real classification. Filename patterns are sufficient to determine "this folder likely has relevant PDFs."

### Decision: Popup shows only folders with Excel

**Choice**: Filter out folders that don't contain an Excel matching `*CONTENEDORES*` (case-insensitive).

**Alternatives considered**: Show all folders with any Excel — rejected because user spec says "only folders WITH contenedores Excel." Show all Desktop folders — too noisy.

**Rationale**: Folders without the Contenedores Excel are useless for Control Final. Pre-filtering reduces cognitive load.

### Decision: Collect all PDFs from selected folders

**Choice**: Gather every `*.pdf` from selected folders (not just pattern-matched ones).

**Alternatives considered**: Only pass pattern-matched PDFs — rejected because the worker handles classification; some PDFs may have non-standard names but valid content.

**Rationale**: The worker already classifies PDFs by reading text. Letting it see all PDFs in a valid folder is more robust than filename heuristics.

## Data Flow

    Desktop/
    ├── Folder_A/
    │   ├── CONTENEDORES.xlsx  ← matches Excel filter
    │   ├── scan_001.pdf       ← matches scan prefix
    │   └── 26AR12345.pdf      ← matches \d{2}AR\d+
    ├── Folder_B/
    │   └── random.pdf         ← no Excel → EXCLUDED
    └── Folder_C/
        ├── CONTENEDORES.xlsx
        └── note.pdf           ← no pattern match but ALL pdfs collected

    btn_control_final click
         │
         ├── [Auto OFF] → file dialog → _control_final_worker(pdfs, excels)
         │
         └── [Auto ON] → _control_final_auto_scan()
                              │
                              ▼
                    Walk Desktop 1 level deep
                    Find folders with *CONTENEDORES* Excel
                              │
                              ▼
                    _control_final_auto_popup(folders)
                              │
                              ▼
                    User selects folders (checkboxes)
                              │
                              ▼
                    Collect all PDFs + Excel from selected
                              │
                              ▼
                    _control_final_worker(pdfs, excels)

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `ui_app.py` ~line 7522 | Modify | Wrap `btn_control_final` in CTkFrame, add CTkSwitch |
| `ui_app.py` ~line 8208 | Modify | Add `_control_final_switch_toggle` method, update command routing |
| `ui_app.py` new | Create | `_control_final_auto_scan()` — walk Desktop, find candidate folders |
| `ui_app.py` new | Create | `_control_final_auto_popup(folders)` — checkbox popup for folder selection |
| `ui_app.py` new | Create | `_control_final_auto_todas()` / `_control_final_auto_ninguna()` — popup helpers |

## Interfaces / Contracts

```python
# New instance variables (in __init__ or toolbar setup)
self._cf_auto_var = ctk.BooleanVar(value=False)
self._cf_auto_switch = ctk.CTkSwitch(...)  # next to btn_control_final

# New methods
def _control_final_switch_toggle(self):
    """Toggle auto mode. Changes btn text hint."""

def _control_final_seleccionar(self):
    """MODIFIED: route to auto-scan or file dialog based on switch state."""

def _control_final_auto_scan(self):
    """Walk Desktop 1 level deep. Find folders with *CONTENEDORES* Excel.
    Returns list of dicts: [{name, path, excel_path, pdf_count}]
    Calls _control_final_auto_popup with results."""

def _control_final_auto_popup(self, folders):
    """CTkToplevel popup. Checkboxes per folder. Todas/Ninguna buttons.
    On confirm: collect PDFs + Excel, call _control_final_worker."""

def _control_final_auto_todas(self):
    """Select all checkboxes in popup."""

def _control_final_auto_ninguna(self):
    """Deselect all checkboxes in popup."""
```

## File Detection Logic

### Excel detection (folder filter)
```python
# Case-insensitive glob: *CONTENEDORES*.{xls,xlsx}
import glob
excel_pattern = os.path.join(folder_path, "*[Cc][Oo][Nn][Tt][Ee][Nn][Ee][Dd][Oo][Rr][Ee][Ss]*")
matches = glob.glob(excel_pattern) + glob.glob(excel_pattern.upper())
```

### PDF filename patterns
```python
import re
SCAN_PREFIX = re.compile(r'^scan', re.IGNORECASE)
MIC_PATTERN = re.compile(r'\d{2}AR\d+', re.IGNORECASE)

def _is_candidate_pdf(filename):
    name = os.path.splitext(filename)[0]
    return bool(SCAN_PREFIX.match(name) or MIC_PATTERN.search(name))
```

### Folder walking
```python
desktop = self._resolver_ruta("planillas_carga", "Desktop")
for entry in os.scandir(desktop):
    if entry.is_dir() and not entry.name.startswith('.'):
        # Check for CONTENEDORES Excel
        # Count candidate PDFs (informational, not filter)
```

## Popup Design

Follows the `_controlar_coordinacion` popup pattern (line 7850):

- **Title**: "Auto-scan Control Final"
- **Subtitle**: "Carpetas con CONTENEDORES encontradas en el Desktop"
- **Scroll frame** with grid columns: checkbox | Folder name | Excel ✅/❌ | PDF count
- **Buttons**: "Todas" | "Ninguna" | "Procesar seleccionados" | "Cancelar"
- **Behavior**: Only folders with Excel are checkable. PDF count is informational.
- **On confirm**: Gather all `*.pdf` from each selected folder + the Excel file → `_control_final_worker(pdfs, excels)`

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Manual | Switch toggles correctly | Visual: button state changes |
| Manual | Auto-scan finds correct folders | Place test folders on Desktop, verify popup contents |
| Manual | Worker receives correct files | Log output shows PDF/Excel counts |
| Manual | Empty Desktop shows warning | No folders with Excel → messagebox |

No automated test infrastructure exists in this project.

## Edge Cases

- **Desktop empty or no folders with Excel**: Show `messagebox.showinfo("Sin resultados", "No se encontraron carpetas con CONTENEDORES en el Desktop")`
- **Folder has Excel but no PDFs**: Show in popup with PDF count = 0, still checkable (worker will report error)
- **Switch OFF mid-process**: Not possible — switch state read once at button click
- **Multiple Excel files in one folder**: Use first match, log warning if multiple
- **Folder names with special characters**: `os.scandir` handles Unicode; no special treatment needed
- **Locked Excel files**: Worker already handles this (existing behavior)
- **`tarea_activa` guard**: Check before auto-scan, same as manual flow (line 8229)

## Migration / Rollout

No migration required. This is a UI-only additive change. The switch defaults to OFF, preserving existing behavior.

## Open Questions

- [ ] Should the switch state persist across sessions? (Recommend: no, default OFF each launch)
- [ ] Should we show folder path or just folder name in popup? (Recommend: name with tooltip for full path)
