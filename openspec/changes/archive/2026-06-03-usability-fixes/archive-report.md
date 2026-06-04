# Archive Report: usability-fixes

**Archived**: 2026-06-03
**Status**: CLOSED — Fully implemented and verified

## Summary

Seven usability fixes implemented and verified across 3 files. All 8 tasks completed.

## Artifacts

| Artifact | Status | Notes |
|----------|--------|-------|
| exploration.md | Read | Explored 7 usability issues; 6 trivial, 1 needed design discussion |
| proposal.md | Read | Defined scope, approach, rollback plan |
| design.md | Read | Unified Guarda architecture + mechanical changes |
| tasks.md | Read | 8 tasks across 4 phases, all marked complete |
| verify-report.md | Read | **0 CRITICAL, 0 WARNING, 0 SUGGESTION** |

## Key Decisions During Implementation

1. **xlutils -> win32com for .xls Guarda writes** (Task T1): Originally planned to add xlutils dependency, but xlutils cannot preserve formatting in existing `.xls` files. Switched to Excel COM (win32com) which was already a dependency. xlutils was NOT added to `requirements.txt` or the deps installer. Documentation in design.md was updated accordingly.

2. **Dialog height adjustment** (Task T6): Implementation uses `min(220 + n * 40, 640)` vs design's `min(180 + n * 36, 560)`. Expanded formula based on observed template item sizes for better visibility of the "Cerrar" button.

## Files Modified

| File | Changes |
|------|---------|
| `ui_app.py` | Guarda helper creation + 4 call sites collapsed to 2; email line removed; ATA config field added + 4 hardcoded values replaced; dialog maxheight increased; dead button removed; Drive tab + all Drive methods removed |
| `ui_config.json` | Removed `google_drive` default block; added `ata_tares: 65000` default |
| `requirements.txt` | No changes (xlutils not needed after COM pivot) |

## Verification

| Severity | Count |
|----------|-------|
| CRITICAL | 0 |
| WARNING | 0 |
| SUGGESTION | 0 |

All checksets pass. Implementation matches spec with documented deviations (xlutils pivot, dialog formula).

## User Confirmation

User confirmed everything works perfectly.

## Archive Location

`openspec/changes/archive/2026-06-03-usability-fixes/`
