# Proposal: Usability Fixes

## Intent

Seven usability issues in daily operations degrade UX or cause silent data loss. The Guarda Excel bug never writes the guarda name, forcing manual corrections. Fixing all seven in one change avoids context-switching on the same file (`ui_app.py`).

## Scope

### In Scope

1. Remove "PLANILLA DE CARGA:" line from CARGA TERRESTRE email
2. Add editable "Servicio ATA ($)" to Ajustes > Valores (default 65000), replace 4 hardcoded `* 60000`
3. Increase "Correos despachados" dialog height from 480 to 560
4. Eliminate Drive tab, `_ajustes_tab_drive()`, Drive fields in `_guardar_ajustes()`, and `google_drive` from `ui_config.json` defaults
5. Remove dead `btn_backup_drive` button
6. Add `xlutils` to `_instalar_deps_ui()` and `requirements.txt`
7. Fix Guarda Excel: unify matching, expand search to rows 2-15, fix `.xls` isinstance, write to column H

### Out of Scope

- New Drive features (deferred indefinitely)
- Template restructuring or migration

## Capabilities

### New Capabilities

None — all changes are bug fixes and config additions within existing functionality.

### Modified Capabilities

None — no spec-level behavior changes.

## Approach

All fixes target `ui_app.py` plus `requirements.txt` and `ui_config.json`. Each is independent, reviewable as one commit.

1. **Email**: Delete `cuerpo += "PLANILLA DE CARGA:\n"` at line 5402.
2. **ATA**: Follow `precio_carpeta` pattern — add tkinter field, read from config, replace 4 `* 60000`.
3. **Dialog**: Change `maxheight` from 480 to 560.
4. **Drive**: Remove tab creation, method, config save/load fields, default section.
5. **Dead button**: Remove `btn_backup_drive` and references.
6. **xlutils**: One-line additions in deps installer and requirements.txt.
7. **Guarda**: Unify 4 paths to `"GUARDA" in str(val).strip().upper()`, search rows 2-15, write to column H.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `ui_app.py` | Modified | 7 edits, ~50 lines |
| `requirements.txt` | Modified | Add `xlutils` |
| `ui_config.json` | Modified | Remove `google_drive` default |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Drive removal breaks saved configs with `google_drive` key | Low | Config load handles missing keys gracefully |
| Guarda row range expansion matches unexpected cells | Medium | Test with real templates; bounded to rows 2-15 |
| Missed `* 60000` occurrence | Low | Grep confirms exactly 4 |

## Rollback Plan

Single commit. Revert to restore all files. ATA config is additive — removing field falls back to default.

## Dependencies

None.

## Success Criteria

- [ ] Email body no longer contains "PLANILLA DE CARGA:"
- [ ] "Servicio ATA ($)" field appears in Ajustes and persists
- [ ] "Correos despachados" "Cerrar" button visible without scrolling
- [ ] No Drive references in Ajustes UI or config defaults
- [ ] No `btn_backup_drive` in UI or codebase
- [ ] `xlutils` installs automatically on dep update
- [ ] Guarda name written to column H on correct Excel row
