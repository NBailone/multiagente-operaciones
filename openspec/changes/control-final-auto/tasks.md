# Tasks: Control Final — Auto Mode

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 400-500 |
| 400-line budget risk | Medium |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 → PR 2 → PR 3 |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: stacked-to-main|feature-branch-chain|size-exception|pending
400-line budget risk: Medium

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Add switch to toolbar, store config | PR 1 | Base: main; include switch toggle logic |
| 2 | Implement auto-scan folder discovery | PR 2 | Base: PR 1 branch; Desktop scan logic |
| 3 | Create folder selection popup | PR 3 | Base: PR 2 branch; checkbox grid, All/Ninguna buttons |
| 4 | Integrate auto-scan flow with worker | PR 4 | Base: PR 3 branch; combine PDFs/Excel, call worker |

## Phase 1: Foundation / Infrastructure

- [x] 1.1 Add `self._cf_auto_var = ctk.BooleanVar(value=False)` in __init__ or toolbar setup
- [x] 1.2 Add `self._cf_auto_switch = ctk.CTkSwitch(...)` next to btn_control_final in toolbar
- [x] 1.3 Add `self._cf_auto_switch.pack(side="left", padx=2, pady=5)` after btn_control_final
- [x] 1.4 Add `self._cf_auto_switch.configure(text="Auto")` to match Súper Auto style
- [x] 1.5 Add `self._cf_auto_switch.configure(command=self._control_final_switch_toggle)`
- [x] 1.6 Add `self._cf_auto_var.trace_add("write", self._control_final_switch_toggle)` for auto-updates

## Phase 2: Core Implementation

- [x] 2.1 Implement `_control_final_switch_toggle(self)` method to update button hint text
- [x] 2.2 Modify `_control_final_seleccionar(self)` to route based on switch state
- [x] 2.3 Add `self._cf_auto_var` to ui_config.json persistence
- [x] 2.4 Add `self._cf_auto_var.set(config.get("control_final_auto", False))` on startup

## Phase 3: Auto-Scan Logic

- [x] 3.1 Implement `_control_final_auto_scan(self)` to walk Desktop 1 level deep
- [x] 3.2 Add Excel detection logic: `*CONTENEDORES*.{xls,xlsx}` case-insensitive
- [x] 3.3 Add PDF count logic: count candidate PDFs by filename patterns
- [x] 3.4 Return list of dicts: `[{name, path, excel_path, pdf_count}]`
- [x] 3.5 Call `_control_final_auto_popup(folders)` with results

## Phase 4: Folder Selection Popup

- [x] 4.1 Implement `_control_final_auto_popup(self, folders)` CTkToplevel
- [x] 4.2 Follow _controlar_coordinacion popup pattern (line 7850)
- [x] 4.3 Create grid columns: checkbox | Folder name | Excel ✅/❌ | PDF count
- [x] 4.4 Add "Todas" and "Ninguna" buttons with `_control_final_auto_todas()` / `_control_final_auto_ninguna()`
- [x] 4.5 Implement confirm handler to collect all PDFs + Excel from selected folders
- [x] 4.6 Call `_control_final_worker(pdfs, excels)` with combined lists

## Phase 5: Helper Methods

- [x] 5.1 Implement `_control_final_auto_todas(self)` to select all checkboxes
- [x] 5.2 Implement `_control_final_auto_ninguna(self)` to deselect all checkboxes
- [x] 5.3 Add error handling for empty Desktop/no folders with Excel
- [x] 5.4 Add warning for multiple Excel files in one folder

## Phase 6: Testing & Verification

- [ ] 6.1 Manual test: Switch toggles correctly, button hint updates
- [ ] 6.2 Manual test: Auto-scan finds correct folders on Desktop
- [ ] 6.3 Manual test: Popup shows folder list with checkboxes
- [ ] 6.4 Manual test: Todas/Ninguna buttons work correctly
- [ ] 6.5 Manual test: Worker receives correct files from auto-scan
- [ ] 6.6 Manual test: Empty Desktop shows warning message
