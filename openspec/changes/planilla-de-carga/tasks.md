# Tasks: Planilla de Carga

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 200-300 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: stacked-to-main
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Add button to UI toolbar | PR 1 | Line ~1819, after btn_editar_excels |
| 2 | Implement popup with folder scanning | PR 1 | _popup_planilla_carga() method |
| 3 | Implement worker thread logic | PR 1 | Background processing, progress, error handling |

## Phase 1: UI Foundation

- [ ] 1.1 Add "Planilla de Carga" button after btn_editar_excels at line 1819 in ui_app.py
- [ ] 1.2 Import required modules (threading, queue, openpyxl) if not already imported
- [ ] 1.3 Add button styling matching existing toolbar buttons

## Phase 2: Popup Implementation

- [ ] 2.1 Implement _popup_planilla_carga() method to scan Desktop for matching folders
- [ ] 2.2 Create checkbox popup with folder selection (FR-003)
- [ ] 2.3 Add "Seleccionar todo" functionality
- [ ] 2.4 Implement folder filtering by pattern DD_MM_YYYY_BATCH_TYPE_PERMISO_CONTAINER_COMPANY (FR-002)

## Phase 3: Excel Processing Logic

- [ ] 3.1 Implement folder processing worker method _planilla_carga_worker()
- [ ] 3.2 Add logic to find Contenedores.xlsx in selected folders (FR-004)
- [ ] 3.3 Implement case-insensitive sheet lookup for "planilla de carga" (FR-004)
- [ ] 3.4 Generate output filename: PLANILLA DE CARGA_{parts[5:]}.xlsx (FR-005)
- [ ] 3.5 Copy sheet to new workbook and save in same folder (FR-006)

## Phase 4: Background Processing & UI

- [ ] 4.1 Implement threading with progress updates and log queue
- [ ] 4.2 Add cancel support for background processing (FR-007)
- [ ] 4.3 Implement modal popup with progress and error summary (FR-008)
- [ ] 4.4 Add auto-close on completion with error summary modal

## Phase 5: Integration & Testing

- [ ] 5.1 Test folder scanning with sample folders
- [ ] 5.2 Test Excel file finding and sheet copying
- [ ] 5.3 Test filename generation logic
- [ ] 5.4 Test background thread progress updates
- [ ] 5.5 Test error handling (missing files, locked files, cancel)

## Requirements Coverage

- FR-001: UI button added in Phase 1
- FR-002: Folder scanning in Phase 2
- FR-003: Checkbox popup in Phase 2
- FR-004: Excel file finding in Phase 3
- FR-005: Filename generation in Phase 3
- FR-006: Sheet copying in Phase 3
- FR-007: Background processing in Phase 4
- FR-008: Modal popup with cancel in Phase 4

## Dependencies

- Phase 1 → Phase 2 (button enables popup)
- Phase 2 → Phase 3 (selected folders enable processing)
- Phase 3 → Phase 4 (worker thread processes folders)
- All phases require existing utils/excel_utils.py and utils/excel_reader.py

## Complexity Estimates

- 1.1: Simple (UI button)
- 1.2: Simple (import)
- 1.3: Simple (styling)
- 2.1: Medium (popup logic)
- 2.2: Medium (UI components)
- 2.3: Simple (select all)
- 2.4: Medium (pattern matching)
- 3.1: Complex (threading + Excel)
- 3.2: Medium (file finding)
- 3.3: Medium (case-insensitive search)
- 3.4: Simple (string manipulation)
- 3.5: Medium (Excel operations)
- 4.1: Complex (background processing)
- 4.2: Medium (cancel support)
- 4.3: Complex (modal UI)
- 4.4: Simple (auto-close logic)
- 5.1-5.5: Medium (testing)

Total estimated lines: ~250-300
400-line budget risk: Low