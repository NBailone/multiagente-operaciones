# Design: Maritime Folder Naming Support

## Technical Approach

Extend the CHOFER-sheet reader tuple from 9 to 11 elements (appending
`puerto_salida` and `peso_flexi`), classify transport type from those
two fields, and branch the folder-naming suffix in the two mail
callers. `_planillas_worker` tolerates the extra fields via `*_` unpack.
No new modules, no new public API — change is contained to `ui_app.py`.

## Architecture Decisions

### Decision: Helper lives as a method on `App` class

| Option | Pros | Cons | Verdict |
|--------|------|------|---------|
| Method on `App` (`_clasificar_tipo_transporte`) | Co-located with the 2 callers; no new import; matches the style of `_leer_xls_antiguo` / `_leer_xlsx_moderno` | Hard to unit-test in isolation | **Chosen** |
| Function in `utils/excel_reader.py` | Pure, easy to test; aligns with ongoing extraction | Requires touching `utils/__init__.py` and the top-level import; mixes a tiny pure helper with extraction-stage modules | Rejected |
| Inline local function in each caller | Smallest diff | Duplicates classification across `_mail_nombre_carpeta` and `_mail_procesar_comparte` | Rejected |

**Rationale**: ~10-line helper, no external deps, consumed only by `App` methods. Extracting to `utils/` adds indirection without payoff; moving it later is a 5-line job.

### Decision: Append new fields to the END of the tuple (indexes 9, 10)

**Rationale**: zero risk to the 5 existing positional unpacks (returns
at L4717 and L4891, plus consumers at L2635, L2729, L3988). Backward
compatible by construction.

### Decision: `*_` discard in `_planillas_worker`

**Rationale**: worker doesn't need the new fields. `*_` is
self-documenting ("trailing ignored") and matches project style.

## Data Flow

```
Excel .xls / .xlsx (CHOFER sheet)
  │
  ▼  _leer_xls_antiguo / _leer_xlsx_moderno  (L4617, L4796)
  │   find "PUERTO SALIDA" / "PESO FLEXI" labels (case-insensitive, label+1)
  │   return 11-tuple
  │
  ├──▶ _mail_nombre_carpeta  (L2623)
  │      unpack 0..3; read [9], [10]
  │      _clasificar_tipo_transporte() → tipo
  │      replace literal "TERRESTRE"; tail = frac | puerto
  │
  ├──▶ _mail_procesar_comparte._extraer_datos  (L2722)
  │      return 7-tuple; caller classifies A and B independently
  │      replace 2 "TERRESTRE" literals; tails per type
  │
  └──▶ _planillas_worker  (L3988)
         f, pe, carp, dest, bl, trans, pat, prec, guarda, *_ = datos
```

## File Changes

| File | Action | Lines | Description |
|------|--------|-------|-------------|
| `ui_app.py` | Modify | ~10 | `_leer_xls_antiguo` (L4617-4717): init `puerto_salida=""`, `peso_flexi=""`; add elif branches for `"PUERTO SALIDA"` and `"PESO FLEXI"` (label+1 read; `.strip()` for puerto, float→int→str coercion for peso); extend return at L4717 to 11 elements. |
| `ui_app.py` | Modify | ~10 | `_leer_xlsx_moderno` (L4796-4891): mirror of above; elif in the same loop, return at L4891. |
| `ui_app.py` | Add | ~12 | New method `_clasificar_tipo_transporte(self, puerto_salida, peso_flexi) -> str` placed next to `_mail_nombre_carpeta` (~L2622). Returns `"TERRESTRE"`, `"ISO"`, or `"FLEXI"`. |
| `ui_app.py` | Modify | ~8 | `_mail_nombre_carpeta` (L2623-2683): keep unpack 0..3, add `puerto_salida, peso_flexi = datos[9], datos[10]`; call helper; replace literal `"TERRESTRE"` in `partes`; `suffix = frac if tipo == "TERRESTRE" else puerto_salida`; `partes.append(suffix)` instead of `if frac: partes.append(frac)`. |
| `ui_app.py` | Modify | ~20 | `_mail_procesar_comparte` (L2685-2881): `_extraer_datos` returns 7-tuple (adds `puerto`, `peso`); unpack A and B include `puerto_a, peso_a`; classify A and B independently; replace both `"TERRESTRE"` literals (L2848, L2859); tails use `frac_a/puerto_a` (L2849) and `frac_b/puerto_b` (L2860) per shipment type. |
| `ui_app.py` | Modify | 1 | `_planillas_worker` (L3988): `... , guarda = datos` → `... , guarda, *_ = datos`. |

## Interfaces / Contracts

```python
def _clasificar_tipo_transporte(self, puerto_salida, peso_flexi) -> str:
    """Return "TERRESTRE" | "ISO" | "FLEXI".

    puerto_salida: None / "" / "-" (after .strip()) → TERRESTRE
    puerto_salida set + peso in {None, "", 0, 0.0, "0", "0.0"} → ISO
    puerto_salida set + peso > 0                           → FLEXI
    peso < 0 treated as 0 (ISO).  Float peso: int if integral else str.
    """
```

## Testing Strategy

| Layer | What | Approach |
|-------|------|----------|
| Manual — terrestrial | Empty `Puerto Salida` | Mail-download; assert folder ends in `_F<n>`, segment 3 is `TERRESTRE`. |
| Manual — maritime ISO | `Puerto="TRP"`, `Peso=0` | Folder ends in `_TRP`, segment 3 is `ISO`. |
| Manual — maritime FLEXI | `Puerto="EXOLGAN"`, `Peso=1200` | Folder ends in `_EXOLGAN`, segment 3 is `FLEXI`. |
| Manual — share case | Terrestrial + maritime in same mail | Two folders, each with its own suffix. |
| Manual — fallback | CHOFER sheet missing | Folder from temp-folder basename, warning logged. |
| Manual — regression | "Completar Planillas" on changed Excel | `_planillas_worker` does not raise. |

No automated tests — matches the proposal's "no new automated tests" decision.

## Migration / Rollout

No migration required. Tuple extension is backward compatible.

## Open Questions

- **None blocking.** Edge cases (`None`, `"-"`, `0`, `"0"`, whitespace, negative) are handled in the helper contract above.
