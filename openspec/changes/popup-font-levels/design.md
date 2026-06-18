# Design: Popup Font Levels

## Technical Approach

Add font scale configuration (3 discrete levels: small/medium/large) to the existing `ui_config.json` config system. A helper function computes scaled font sizes and popup geometry from a base level. Each comparison popup gets an ephemeral override selector. The default is applied globally via Settings (Ajustes tab). Level 1 matches current behavior exactly — zero migration risk.

## Architecture Decisions

### Decision: Scale via multiplication factor, not absolute sizes

**Choice**: `FONT_LEVEL_SCALES = {1: 1.0, 2: 1.25, 3: 1.5}` with base sizes mapped per role (data=11, header=12)
**Alternatives**: Hardcoded absolute font sizes per level; single multiplier applied uniformly
**Rationale**: Role-based mapping keeps visual hierarchy consistent across levels. A single multiplier would make headers too large at level 3 while data stays small.

### Decision: Config key at top level, not nested under a section

**Choice**: `self.config["font_level"]` (integer 1–3, default 1)
**Alternatives**: `self.config["ui"]["font_level"]` nested section
**Rationale**: Existing pattern uses flat keys for simple values (`window_geo`). This is a single integer — a section adds noise. Follows the `sash_*` flat key pattern.

### Decision: Per-popup override via CTkOptionMenu, ephemeral

**Choice**: Each popup renders a CTkOptionMenu in a top-right frame; changing it re-renders content. Override does not persist.
**Alternatives**: Override saved per-popup type; slider instead of discrete levels
**Rationale**: Ephemeral keeps it simple — users try a size, it resets next open. Saving per-popup adds config complexity for marginal benefit. Discrete levels match the proposal scope (3 levels, not arbitrary).

### Decision: Helper functions in ui_app.py, not extracted to utils/

**Choice**: `_get_font_sizes(level)` and `_get_popup_geometry(base_w, base_h, level)` as methods on `App`
**Alternatives**: Standalone functions in `constants/palette.py` or new `utils/font_scale.py`
**Rationale**: They access `self.config` and CTk widget creation. Extracting to standalone would require passing config as arg. Per `openspec/config.yaml` rules, prefer extraction when clean — but these are 15-line helpers tightly coupled to the UI class. Defer extraction to a future monolith refactor.

## Data Flow

    ┌──────────────┐     ┌───────────────────┐     ┌──────────────────┐
    │ Ajustes tab  │────▶│ config["font_level"]│────▶│ _cargar_config() │
    │ (CTkOption   │     │ (persisted JSON)   │     │ (on app start)   │
    │  Menu)       │     └───────────────────┘     └────────┬─────────┘
    └──────────────┘                                        │
                                                            ▼
    ┌──────────────────────────────────────────────────────────────┐
    │  _abrir_comparacion / _coordinacion / _final                 │
    │                                                              │
    │  1. Read level: self._popup_font_override or config default  │
    │  2. Compute fonts: _get_font_sizes(level) → {data: N, hdr: M}│
    │  3. Compute geometry: _get_popup_geometry(base_w, base_h, lv)│
    │  4. Render popup with scaled fonts and geometry              │
    │  5. Override selector (CTkOptionMenu) re-renders on change   │
    └──────────────────────────────────────────────────────────────┘

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `constants/palette.py` ~line 58 | Modify | Add `FONT_LEVEL_SCALES = {1: 1.0, 2: 1.25, 3: 1.5}` and `FONT_BASE_SIZES = {"data": 11, "header": 12}` after `FONT_MONO` |
| `ui_app.py` ~line 6598 (after `_cfg_obtener_rutas`) | Modify | Add `_get_font_sizes(self, level)` and `_get_popup_geometry(self, bw, bh, level)` helper methods |
| `ui_app.py` 6431–6522 (`_abrir_comparacion`) | Modify | Replace hardcoded `size=11/12` with scaled values; scale geometry; add override CTkOptionMenu in top-right frame |
| `ui_app.py` 7494–7593 (`_abrir_comparacion_coordinacion`) | Modify | Same pattern as Tickets popup |
| `ui_app.py` 8380–8519 (`_abrir_comparacion_final`) | Modify | Same pattern as Tickets popup |
| `ui_app.py` ~9580 (`_panel_ajustes`) | Modify | Add `("apariencia", "🎨  Apariencia")` to `tab_names` list |
| `ui_app.py` ~9645 (`_ajustes_builders`) | Modify | Add `"apariencia": self._ajustes_tab_apariencia` entry |
| `ui_app.py` ~9680 (after `_ajustes_tab_ocr`) | Modify | New `_ajustes_tab_apariencia(self, parent)` method — CTkOptionMenu with values `["Pequeño (1)", "Mediano (2)", "Grande (3)"]` |
| `ui_app.py` ~10347 (`_guardar_ajustes`) | Modify | Add `self.config["font_level"] = self._font_level_menu.get()` block before `_guardar_config()` call |

## Interfaces / Contracts

```python
# constants/palette.py — new constants
FONT_LEVEL_SCALES = {1: 1.0, 2: 1.25, 3: 1.5}
FONT_BASE_SIZES = {"data": 11, "header": 12}

# ui_app.py — new helper methods
def _get_font_sizes(self, level: int) -> dict:
    """Returns {"data": scaled_data_size, "header": scaled_header_size}"""
    scale = FONT_LEVEL_SCALES.get(level, 1.0)
    return {
        "data": round(FONT_BASE_SIZES["data"] * scale),
        "header": round(FONT_BASE_SIZES["header"] * scale),
    }

def _get_popup_geometry(self, base_w: int, base_h: int, level: int) -> tuple[int, int]:
    """Returns (scaled_w, scaled_h) capped at 1140×720."""
    scale = FONT_LEVEL_SCALES.get(level, 1.0)
    w = min(round(base_w * scale), 1140)
    h = min(round(base_h * scale), 720)
    return w, h

# Config key
# self.config["font_level"]  →  int (1, 2, or 3), default 1
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Manual | Default level 1 matches current output | Visual comparison: open each popup at level 1, verify identical to pre-change |
| Manual | Level 3 scales correctly | Open each popup at level 3, verify fonts larger and geometry scaled |
| Manual | Override resets on popup close | Change override in popup, close and reopen — should use default |
| Manual | Settings persistence | Set level 2 in Ajustes, restart app, verify level 2 loads |
| Manual | Screen bounds | Level 3 on 1920×1080 — popup must not exceed screen |

No automated test infrastructure exists (per `openspec/config.yaml`). Manual verification only.

## Migration / Rollout

No migration required. Default `font_level: 1` produces identical output to current behavior. Existing `ui_config.json` files without this key will use the default via `self.config.get("font_level", 1)`.

## Open Questions

- [ ] Should the override selector be visible to all users or only when a non-default level is set? (Proposal says always visible — following that.)
