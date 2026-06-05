# Tasks: Maritime Folder Naming Support

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~61 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |
| Chain strategy | size-exception |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | End-to-end maritime folder naming | PR 1 | All six hunks ship together; incomplete without the Phase 1 tuple extension |

## Phase 1: Excel Reader Tuple Extension

Foundation. 9 → 11 tuple; indices 0–8 untouched.

- [x] 1.1 Extend `_leer_xls_antiguo` (`ui_app.py` L4617-4717) — init `puerto_salida=""`, `peso_flexi=""`; add `elif` for `"PUERTO SALIDA"` and `"PESO FLEXI"` (label+1 read; `.strip()` for puerto, float → int → str for peso); return 11 elements at L4717. Spec: detect transport type.
- [x] 1.2 Extend `_leer_xlsx_moderno` (`ui_app.py` L4796-4891) — mirror of 1.1. Spec: detect transport type.

**Done when**: 11-tuple returned; `datos[0..8]` unchanged.

## Phase 2: Classification Helper

- [x] 2.1 Add `_clasificar_tipo_transporte(self, puerto_salida, peso_flexi) -> str` (`ui_app.py` ~L2622, before `_mail_nombre_carpeta`). Returns `"TERRESTRE"` | `"ISO"` | `"FLEXI"`: `None` / `""` / `"-"` after `.strip()` → TERRESTRE; set + weight in `{None, "", 0, 0.0, "0"}` → ISO; set + weight > 0 → FLEXI. Spec: detect transport type.

**Done when**: empty → TERRESTRE, `"TRP"`+0 → ISO, `"EXOLGAN"`+1200 → FLEXI.

## Phase 3: Folder Naming Callers

- [x] 3.1 Update `_mail_nombre_carpeta` (`ui_app.py` L2623-2683) — keep L2635 unpack at 0..3; read `puerto_salida, peso_flexi = datos[9], datos[10]`; call `self._clasificar_tipo_transporte(...)`; swap the `"TERRESTRE"` literal in `partes`; `suffix = frac if tipo == "TERRESTRE" else puerto_salida`; replace the `if frac: partes.append(frac)` block with `partes.append(suffix)`. Spec: build folder name with type-appropriate suffix.
- [x] 3.2 Update `_mail_procesar_comparte` (`ui_app.py` L2685-2881) — `_extraer_datos` returns 7-tuple (adds `puerto`, `peso`); A/B unpacks gain `puerto_a, peso_a` and `puerto_b, peso_b`; classify A and B independently; replace both `"TERRESTRE"` literals (L2848, L2859); tails use `frac_a or puerto_a` (L2849) and `frac_b or puerto_b` (L2860). Spec: build folder name + share case classifies each CONTENEDORES independently.

**Done when**: terrestrial `_F<n>` + TERRESTRE; ISO `_TRP` + ISO; FLEXI `_EXOLGAN` + FLEXI; share case → two folders with their own suffix.

## Phase 4: Worker Safety and Verification

- [x] 4.1 Fix `_planillas_worker` (`ui_app.py` L3988) — `..., guarda = datos` → `..., guarda, *_ = datos` so the 11-element tuple unpacks. Spec: tuple tolerates extra fields.
- [x] 4.2 Manual smoke run — one mail per scenario: terrestrial, maritime ISO, maritime FLEXI, share case, CHOFER-missing fallback, "Completar Planillas" regression. Assert folder tail and third segment; confirm no `ValueError: too many values to unpack`. Covers all spec scenarios.

**Done when**: all six scenarios pass; planillas worker no longer raises.
