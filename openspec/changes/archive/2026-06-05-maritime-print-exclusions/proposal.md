# Proposal: maritime-print-exclusions

## Intent

Skip ATA/tares document printing for maritime folders (ISO/FLEXI) even when "Servicio ATA / Recibo ATA" is checked. Currently all folders print ATA sheets regardless of transport type.

## Scope

### In Scope
- Helper `_detectar_tipo_carpeta()` to parse transport type from folder name
- Skip ATA section in panel manual path (`_imp_worker`)
- Skip ATA section in súper auto path (`_super_imprimir`)

### Out of Scope
- UI changes to print dialog
- Changes to other print options (factura, packing list, etc.)
- Diagnostic window changes
- Changes to folder naming or spec-level print capability

## Capabilities

### New Capabilities
None — pure implementation change, no new spec-level behavior.

### Modified Capabilities
None — existing "document printing" capability contracts unchanged (same options, same flows; only runtime logic changes).

## Approach

1. Add `_detectar_tipo_carpeta(nombre: str) -> str` helper parsing folder name by `split("_")[2]`. Returns `TERRESTRE` (default), `ISO`, or `FLEXI`.
2. In `_imp_worker` (panel manual, ~L1524): before the `if opciones.get("servicio_ata"):` block, check type and skip if maritime.
3. In `_super_imprimir` (súper auto, ~L3150): before the `if hacer_recibo and sobres:` ATA block, check type and skip if maritime.
4. Old-format folders without a recognizable third segment default to TERRESTRE.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `ui_app.py` | Modified | +1 helper function, 2 guard conditions in print paths (~20 lines) |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Old-format folder name parsed as wrong type | Low | Default to TERRESTRE preserves existing behavior |
| Missing one of two print paths | Med | Both panel manual and súper auto must be modified; verify both |
| Folder name with underscores in carrier/dest garbles split | Low | Only index 2 is read; extra underscores in later segments don't affect it |

## Rollback Plan

Revert `ui_app.py` via `git checkout -- ui_app.py` or revert the merge commit.

## Dependencies

None.

## Success Criteria

- [ ] Maritime folders (ISO/FLEXI) skip ATA/tares printing even when "Servicio ATA / Recibo ATA" is checked
- [ ] Terrestrial folders print ATA sheets as before (existing behavior preserved)
- [ ] Both print paths (panel manual and súper auto) respect the exclusion
- [ ] Old-format folders without a type segment default to TERRESTRE and print ATA normally
