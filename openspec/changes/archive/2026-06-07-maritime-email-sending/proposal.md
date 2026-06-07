# Proposal: Maritime Email Sending

## Intent

Extend `_correos_core` to send emails for ISO/FLEXI (maritime) folders — individual per-folder and grupal for planillas — matching existing terrestrial support with maritime-appropriate content, attachments, and folder matching.

## Scope

### In Scope

- Individual email for ISO/FLEXI folders: subject without "MIC", body matches subject, attachments = Contenedores.xlsx + get\*.pdf, SALIDA/SALIDAS based on PDF count, suffix extracted from ISO/FLEXI folder names
- New grupal email for maritime PLANILLA DE CARGA\*.xlsx files: subject singular/plural by count, plain-text body, same 14 recipients as CARGA TERRESTRE
- Terrestrial grupal exclusion: only scan TERRESTRE folders for CARGA TERRESTRE email
- Regex updates at L5377, L5461, L5481 to match ISO/FLEXI folder name patterns
- Existing terrestrial individual email unchanged

### Out of Scope

- Maritime COMPARTIDO (paired folder) emails — deferred
- New UI, configuration, or recipient management
- Changes to mail-folder-naming spec or transport classifier

## Capabilities

### New Capabilities

- `maritime-email-individual`: Individual email sending for ISO/FLEXI folders — attachment selection, subject/body construction, suffix extraction
- `maritime-email-grupal`: Grupal email for maritime PLANILLA DE CARGA files — aggregation, subject, body, and attachment logic

### Modified Capabilities

- None

## Approach

1. **Maritime detection**: Use `_clasificar_tipo_transporte` (L2641) or folder-name patterns to classify each folder as terrestrial vs maritime
2. **Individual maritime**: Branch in per-folder loop — ISO/FLEXI folders use maritime subject/body/attachment logic. Singular "SALIDA" if 1 get\*.pdf, plural "SALIDAS" if 2+. Extract suffix from folder name (non-TERRESTRE regex)
3. **Maritime grupal**: New scan over ISO/FLEXI folders for PLANILLA DE CARGA\*.xlsx. Build subject, body, send to same 14 recipients
4. **Terrestrial grupal filter**: Add type check before including folders in CARGA TERRESTRE
5. **Regex updates**: L5377, L5461, L5481 — add ISO/FLEXI patterns alongside existing TERRESTRE ones

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `ui_app.py:_correos_core` | Modified | Maritime branches, regex updates, new grupal maritime block |
| `specs/maritime-email-individual/spec.md` | New | Spec for individual maritime email |
| `specs/maritime-email-grupal/spec.md` | New | Spec for grupal maritime email |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Regex changes break terrestrial email | Medium | Changes localized to 3 lines; test both terrestrial and maritime flows |
| Folders with mixed types in grupal | Low | Filter by transport type before aggregation |

## Rollback Plan

- Revert all changes to `_correos_core` (localized to one function)
- Restore original regex patterns at L5377, L5461, L5481
- No database or config migrations — pure code revert

## Success Criteria

- [ ] ISO/FLEXI folders send individual emails with correct subject, body, and attachments (Contenedores.xlsx + get\*.pdf)
- [ ] SALIDA/SALIDAS correctly reflects count of get\*.pdf attachments
- [ ] Maritime grupal email aggregates PLANILLA DE CARGA\*.xlsx from all ISO/FLEXI folders
- [ ] CARGA TERRESTRE grupal email excludes ISO/FLEXI folders
- [ ] Existing terrestrial individual email behavior is preserved
