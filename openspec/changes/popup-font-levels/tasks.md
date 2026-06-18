# Tasks: Popup Font Levels

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 200-250 lines |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 → PR 2 → PR 3 → PR 4 |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Constants and helper methods | PR 1 | Add FONT_LEVEL_SCALES, FONT_BASE_SIZES, _get_font_sizes(), _get_popup_geometry() |
| 2 | Tickets popup with scaling | PR 2 | Modify _abrir_comparacion() with font scaling and override selector |
| 3 | Coordinación popup with scaling | PR 3 | Modify _abrir_comparacion_coordinacion() with font scaling and override selector |
| 4 | Final popup and settings UI | PR 4 | Modify _abrir_comparacion_final() and add font level selector in Ajustes tab |

## Phase 1: Foundation / Infrastructure

- [x] 1.1 Add FONT_LEVEL_SCALES and FONT_BASE_SIZES constants to constants/palette.py
- [x] 1.2 Add _get_font_sizes(level) helper method to ui_app.py App class
- [x] 1.3 Add _get_popup_geometry(base_w, base_h, level) helper method to ui_app.py App class

## Phase 2: Core Implementation

- [x] 2.1 Modify _abrir_comparacion() to use scaled fonts and geometry
- [x] 2.2 Add per-popup font level override selector to Tickets popup
- [ ] 2.3 Test Tickets popup at all 3 font levels manually

## Phase 3: Integration / Wiring

- [x] 3.1 Modify _abrir_comparacion_coordinacion() to use scaled fonts and geometry
- [x] 3.2 Add per-popup font level override selector to Coordinación popup
- [ ] 3.3 Test Coordinación popup at all 3 font levels manually

## Phase 4: Testing / Verification

- [x] 4.1 Modify _abrir_comparacion_final() to use scaled fonts and geometry
- [x] 4.2 Add per-popup font level override selector to Final popup
- [x] 4.3 Add font level selector to Ajustes tab (CTkOptionMenu)
- [ ] 4.4 Test Final popup at all 3 font levels manually
- [ ] 4.5 Test Settings persistence across app restarts
- [ ] 4.6 Test override resets on popup close
- [ ] 4.7 Test Level 3 popup fits within 1920x1080 screen bounds

## Phase 5: Cleanup / Documentation

- [x] 5.1 Add font_level config key to ui_config.json via _guardar_config()
- [x] 5.2 Update _cargar_config() to read font_level with default 1
- [x] 5.3 Update _guardar_ajustes() to save font_level from settings
- [x] 5.4 Verify default level 1 matches current behavior exactly
- [x] 5.5 Document font scaling behavior in code comments
