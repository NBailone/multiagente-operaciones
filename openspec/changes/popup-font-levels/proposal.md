# Proposal: Popup Font Levels

## Intent

Users report difficulty reading comparison popup content at standard font sizes, especially on high-DPI displays. There is no way to adjust readability without resizing the entire app. This adds 3 discrete font scale levels (small/medium/large) with a persistent default and per-popup override.

## Scope

### In Scope
- Config key `font_level` in `ui_config.json` (default: 1)
- Settings UI selector in Ajustes tab for default font level
- Per-popup override selector (top-right corner) that temporarily overrides the default
- Scale factor applies to: popup window geometry, all font sizes, scrollable frame width
- 3 comparison popups: Control de Tickets, Control de Coordinación, Control Final

### Out of Scope
- Font scaling for the main application window (future enhancement)
- Custom font size slider (only 3 discrete levels)
- Per-user profiles or font family changes

## Capabilities

### New Capabilities
- `popup-font-levels`: Configurable font size scaling for comparison popup windows with persistent defaults and per-popup override

### Modified Capabilities
None — this is an additive UI enhancement.

## Approach

1. **Scale constants**: Define `FONT_LEVEL_SCALES = {1: 1.0, 2: 1.25, 3: 1.5}` and derived font size maps (`DATA_FONTS = {1: 11, 2: 14, 3: 16}`, `HEADER_FONTS = {1: 12, 2: 15, 3: 18}`)
2. **Config**: Add `font_level` to `ui_config.json` via existing `_cargar_config()` / `_guardar_config()` (lines ~80-120 in `ui_app.py`)
3. **Settings UI**: CTkOptionMenu in Ajustes tab bound to `font_level` config key
4. **Popup override**: Each `_abrir_comparacion*()` method adds a CTkOptionMenu in the top-right frame. Changing it re-renders the popup content with new scale. Override is ephemeral (not persisted)
5. **Rendering**: Pass scale factor to a helper that computes `base_size * scale` for each font, and `base_geometry * scale` for window size. Apply to all CTkLabel font= params and window.geometry()

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `ui_app.py` lines 6431-6530 | Modified | `_abrir_comparacion()` — add scale-aware geometry and font computation |
| `ui_app.py` lines 7494-7590 | Modified | `_abrir_comparacion_coordinacion()` — same pattern |
| `ui_app.py` lines 8380-8480 | Modified | `_abrir_comparacion_final()` — same pattern |
| `ui_app.py` Ajustes tab | Modified | Add font level selector (CTkOptionMenu) |
| `ui_config.json` | Modified | New `font_level` key (default: 1) |
| `ui_app.py` utils section | New | `_get_font_sizes(level)` and `_get_popup_geometry(base_w, base_h, level)` helpers |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Monolithic ui_app.py grows further | Low | Helpers are small (~20 lines); extraction to utils/ deferred to future refactor |
| Popup exceeds screen at level 3 | Low | Cap at 1140×720; test on 1920×1080 minimum |
| Config migration for existing users | None | Default level 1 = current behavior; no migration needed |

## Rollback Plan

Revert the `ui_app.py` changes and remove `font_level` from `ui_config.json`. Level 1 defaults match current behavior exactly, so no data loss.

## Dependencies

None — pure UI addition using existing CTkinter widgets and config system.

## Success Criteria

- [ ] Default `font_level: 1` produces identical output to current behavior
- [ ] Settings selector persists choice across app restarts
- [ ] Per-popup override applies immediately and resets on popup close
- [ ] Level 3 popup does not exceed 1920×1080 screen bounds
- [ ] All 3 popups scale consistently with the same level
