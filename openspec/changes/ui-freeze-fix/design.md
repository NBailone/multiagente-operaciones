# Design: UI Freeze Fix — Threading Improvement

## Technical Approach

Move all Desktop filesystem scans (`os.listdir`, `os.scandir`) off the main thread into a shared `_scan_desktop_folders()` helper executed via `threading.Thread(daemon=True)`. Results are pushed back to the main thread via `self.after(0, callback)`. Additionally: cap poll burst processing, add timeout to `preguntar_reintentar`, and snapshot CTkStringVar before thread start.

## Architecture Decisions

| Decision | Choice | Alternatives | Rationale |
|----------|--------|-------------|-----------|
| Scan execution | `threading.Thread` daemon | `concurrent.futures`, `asyncio` | Matches existing codebase pattern (22 existing `threading.Thread` calls). No new dependencies. |
| Shared helper location | Method on the app class (`_scan_desktop_folders`) | Standalone function in `utils/` | Needs `self._resolver_ruta()` for path resolution. Keeping it as a method avoids passing config state. |
| Result delivery | `self.after(0, callback)` | Queue-based | Existing pattern (45+ `self.after(0, ...)` calls). Simpler than queue for one-shot results. |
| Poll cap | 20 msg/tick | 100 (current) | Reduces main-thread blocking during burst log scenarios. |
| preguntar_reintentar timeout | 300s | 30s (spec), no timeout (current) | Long-running Excel operations may need retries; 300s prevents deadlock without premature cancellation. |
| Var snapshot | Read before `Thread.start()`, pass as arg | `self.after(0, ...)` + Event | Simpler, no deadlock risk, matches existing `_mail_ejecutar_buscar` pattern (line 2365). |

## Data Flow

```
User clicks refresh / opens popup
         │
         ▼
┌─────────────────────┐
│ Disable button(s)   │  ← main thread
│ Show "Buscando..."  │
└────────┬────────────┘
         │
         ▼  threading.Thread(daemon=True)
┌─────────────────────┐
│ _scan_desktop_folders│  ← background thread
│  os.scandir Desktop │
│  filter CONTENEDORES│
│  return list[dict]  │
└────────┬────────────┘
         │
         ▼  self.after(0, callback)
┌─────────────────────┐
│ Re-enable buttons   │  ← main thread
│ Populate UI widgets │
│ Stale check: panel  │
│ still active?       │
└─────────────────────┘
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `ui_app.py` ~line 961 | Modify | `_imp_escanear_carpetas`: disable button, launch thread, callback populates scroll frame |
| `ui_app.py` ~line 8290 | Modify | `_control_final_auto_scan`: launch thread, callback shows popup or message |
| `ui_app.py` ~line 3783 | Modify | `_popup_agregar_guarda`: launch thread, callback shows popup |
| `ui_app.py` (new, near scan helpers) | Create | `_scan_desktop_folders(self, config_key="planillas_carga")` — shared scan helper |
| `ui_app.py` ~line 6710 | Modify | `_poll_log_queue`: change `100` → `20`, add adaptive throttle (idle 3 ticks → 500ms) |
| `ui_app.py` ~line 177 | Modify | Add `self._poll_idle_count = 0` state var |
| `utils/excel_utils.py` line 34 | Modify | `resultado.wait(timeout=300)` |
| `ui_app.py` ~line 2357 | Modify | Snapshot `cantidad` before `Thread.start()`, pass as arg to `_mail_worker` |

## Interfaces / Contracts

```python
def _scan_desktop_folders(self, config_key: str = "planillas_carga") -> list[dict]:
    """Scan Desktop for folders containing CONTENEDORES Excel.
    
    Thread-safe: no tkinter widget access. Call from background thread.
    Returns: [{"name": str, "path": str, "excel_path": str, "pdf_count": int}]
    """

def _imp_escanear_carpetas(self):
    """Original method refactored: disables refresh button, launches
    _scan_desktop_folders in thread, callback populates _imp_scroll_carpetas."""

def _mail_worker(self, cantidad: int):
    """Modified: receives `cantidad` as parameter instead of reading from widget."""
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Manual | UI stays responsive during scan | Click refresh, move window, interact with other panels while scan runs |
| Manual | Poll throttle activates | Generate burst logs, observe idle transition to 500ms |
| Manual | preguntar_reintentar timeout | Mock dialog fail to show, verify returns False after timeout |
| Manual | Mail worker var snapshot | Start mail download, verify no TclError |

## Migration / Rollout

No migration required. All changes are internal threading improvements with no data model or API changes.

## Open Questions

- [ ] None — all requirements are clear from spec and exploration.
