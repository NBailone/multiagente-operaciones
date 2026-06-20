# Tasks: Planilla de Carga

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 200-350 |
| 400-line budget risk | Medium |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 → PR 2 → PR 3 |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Add button and basic popup structure | PR 1 | Base branch; includes tests/docs |
| 2 | Implement folder scanning and checkbox UI | PR 2 | Immediate parent/base branch boundary; depends on PR 1 |
| 3 | Implement Excel processing helper | PR 3 | Immediate parent/base branch boundary; depends on PR 2 |
| 4 | Implement background thread processing | PR 4 | Immediate parent/base branch boundary; depends on PR 3 |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: Medium

## Phase 1: Infrastructure / Foundation

- [ ] 1.1 Add button "📋 Planilla de Carga" after line 1818 in ui_app.py
- [ ] 1.2 Create _popup_planilla_carga() method with basic structure
- [ ] 1.3 Add imports for regex and background threading if needed

## Phase 2: Core Implementation

- [ ] 2.1 Implement folder scanning with regex pattern ^\d{2}_\d{2}_\d{4}_\d+_.*$
- [ ] 2.2 Create scrollable frame with CTkCheckBox per folder
- [ ] 2.3 Add "Generar" and "Cancelar" buttons with state management
- [ ] 2.4 Implement folder selection state tracking

## Phase 3: Excel Processing

- [ ] 3.1 Create _extraer_planilla_carga() helper method
- [ ] 3.2 Implement Contenedores*.xlsx file detection
- [ ] 3.3 Load Excel with openpyxl and find "planilla de carga" sheet
- [ ] 3.4 Copy sheet to new workbook and generate filename
- [ ] 3.5 Save file in same folder with proper naming

## Phase 4: Background Processing

- [ ] 4.1 Create _generar_planillas_carga_thread() method
- [ ] 4.2 Implement sequential folder processing
- [ ] 4.3 Add progress queue for UI updates
- [ ] 4.4 Collect and display errors per folder
- [ ] 4.5 Show summary dialog on completion

## Phase 5: Testing / Verification

- [ ] 5.1 Write unit tests for _extraer_planilla_carga()
- [ ] 5.2 Test folder scanning regex patterns
- [ ] 5.3 Test UI button state management
- [ ] 5.4 Test background thread error handling
- [ ] 5.5 Verify integration between components

## Phase 6: Documentation / Cleanup

- [ ] 6.1 Add docstrings for new methods
- [ ] 6.2 Update comments with Spanish explanations
- [ ] 6.3 Remove temporary code and debug statements
- [ ] 6.4 Add type hints where appropriate