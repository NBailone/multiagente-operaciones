# Tasks: UI Freeze Fix — Threading Improvement

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~120 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Delivery strategy | Single PR |

## Tasks

### Phase 1: Shared Helper

- [x] 1.1 Add `_scan_desktop_folders(self, pattern, callback)` method — threading.Thread that scans Desktop 1 level deep with `os.scandir`, filters by pattern, calls `self.after(0, callback, results)` when done. Disables a passed button during scan, re-enables after.

### Phase 2: Refactor Scans (HIGH impact)

- [x] 2.1 Refactor `_imp_escanear_carpetas()` — replace synchronous `os.listdir` with `_scan_desktop_folders()` call. Move scroll frame population to callback `_imp_poblar_carpetas(results)`. Disable refresh button during scan.
- [x] 2.2 Refactor `_control_final_auto_scan()` — replace `os.scandir` loop with `_scan_desktop_folders()` call. Move popup trigger to callback.
- [x] 2.3 Refactor `_popup_agregar_guarda()` — replace `os.listdir` with `_scan_desktop_folders()` call. Move popup trigger to callback. Disable button during scan.

### Phase 3: Medium/Low fixes

- [x] 3.1 Cap `_poll_log_queue` max messages from 100 to 20 per tick (one constant change)
- [x] 3.2 Add `timeout=300` to `threading.Event.wait()` in `utils/excel_utils.py` `preguntar_reinterruptar()`
- [x] 3.3 Snapshot `self._mail_entry_cantidad.get()` before thread start in `_mail_worker`, pass as arg

### Phase 4: Verification

- [ ] 4.1 Manual test: Impresión panel loads without freeze
- [ ] 4.2 Manual test: Control Final auto-scan doesn't freeze UI
- [ ] 4.3 Manual test: Popup agregar guarda doesn't freeze UI
- [ ] 4.4 Manual test: Log queue doesn't cause jank during heavy output
