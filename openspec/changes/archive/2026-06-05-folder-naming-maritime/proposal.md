# Proposal: Maritime Folder Naming Support

## Intent

Mail-download flow hardcodes `TERRESTRE` and ends every folder with a
shipment fraction. Maritime loads (ISO / FLEXI) need a port-of-origin
suffix. Derive the type from CONTENEDORES.

## Scope

### In Scope

- Read `Puerto Salida` and `Peso Flexi` in `_leer_xlsx_moderno` (L4780) and
  `_leer_xls_antiguo` (L4601); append to **end** of the 9-element tuple
  (indices 9, 10). Positional access 0–8 keeps working.
- Branch folder name in `_mail_nombre_carpeta` (L2623): `TERRESTRE` ends in
  fraction, `MARITIMO ISO` / `MARITIMO FLEXI` end in port.
- Update `_mail_procesar_comparte` (L2685): two `TERRESTRE` literals →
  dynamic; `_extraer_datos` forwards the new fields.

### Out of Scope

- "Completar Planillas" (SOBRES, COBRO, PC) and its `TERRES` literal.
- `CHILE` terminal literal at end of folder name.
- New GUI / settings — type is derived from the Excel.
- New automated tests — manual sample run per type.

## Capabilities

### New Capabilities

- `mail-folder-naming`: derives transport type from CONTENEDORES (`Puerto
  Salida` + `Peso Flexi`) and builds the folder name in the right format.

### Modified Capabilities

- None — no existing `openspec/specs/` files.

## Approach

1. Extend tuple in both Excel readers 9 → 11.
2. Detect: `puerto_salida` set & ≠ `"-"` → maritime; `peso_flexi` in
   `{0, "", None}` → ISO, else FLEXI.
3. Same prefix; swap suffix (fraction ↔ port).
4. `_planillas_worker` (L3988): `*_` unpack, or it raises
   `ValueError: too many values to unpack`.

## Affected Areas

| Area | Impact |
|------|--------|
| `ui_app.py:_leer_xlsx_moderno` (L4780) | Tuple 9→11; read new labels |
| `ui_app.py:_leer_xls_antiguo` (L4601) | Same |
| `ui_app.py:_mail_nombre_carpeta` (L2623) | Type branch; suffix swap |
| `ui_app.py:_mail_procesar_comparte` (L2685) | `TERRESTRE` literals → dynamic; expand `_extraer_datos` |
| `ui_app.py:_planillas_worker` (L3988) | `*_` unpack — safety only |

## Risks

| Risk | Lik | Mitigation |
|------|-----|------------|
| `_planillas_worker` breaks on tuple grow | High | `*_` unpack in this change |
| `Puerto Salida` whitespace / `"-"` / `None` | Med | `.strip()`; empty/`"-"` → terrestrial |
| Wrong sheet/column in legacy `.xls` | Med | Reuse existing case-insensitive label search |
| Fraction and port both populated oddly | Low | Maritime wins if `Puerto Salida` set |

## Rollback Plan

Revert the five `ui_app.py` hunks. Tuple back to 9, `TERRESTRE` literal
restored, planillas worker to 9-name unpack. No data migration.

## Dependencies

None new. Existing `openpyxl` / `xlrd` reads.

## Success Criteria

- [ ] Maritime `TRP` + `Peso Flexi=0` → `_TRP`, `ISO`
- [ ] Maritime `EXOLGAN` + `Peso Flexi=1200` → `_EXOLGAN`, `FLEXI`
- [ ] Terrestrial (empty `Puerto Salida`) → `_F1…_Fn`, `TERRESTRE`
- [ ] `_planillas_worker` runs without unpack error
- [ ] No new error logs in mail worker for either type
