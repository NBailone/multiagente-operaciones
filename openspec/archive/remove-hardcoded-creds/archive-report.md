# Archive Report: remove-hardcoded-creds

**Date**: 2026-06-03
**Status**: ARCHIVED — change complete, all 9 tasks verified

---

## Summary

Removed hardcoded encryption key `"ENCRYPTION_KEY_PLACEHOLDER"` and credential defaults (`USUARIO`, `PASSWORD`) from source code. Moved secret management to `.env` loaded via `python-dotenv` with auto-generation on first run.

## Files Changed (final)

| File | Action | Lines | Detail |
|------|--------|-------|--------|
| `constants/palette.py` | Modified | -2 | Removed `USUARIO`, `PASSWORD` lines 7-8 |
| `constants/__init__.py` | Modified | -1 | Removed `USUARIO, PASSWORD` from re-export line |
| `ui_app.py` | Modified | ~+30 net | Added `import dotenv`, `.env` loading block, `_ensure_secret_key()`, replaced 9 default refs, replaced 4 hardcoded key refs, added decrypt-failure guards |
| `.env.example` | Created | +5 | Template with `MULTIAGENTE_SECRET_KEY` and generation instructions |
| `requirements.txt` | Created | +1 | `python-dotenv>=1.0.0` |

## Verification Result

| Severity | Count |
|----------|-------|
| CRITICAL | 0 |
| WARNING  | 1 |
| SUGGESTION | 2 |

**W1**: Design doc used `os.urandom(32).hex()` but code uses `secrets.token_hex(32)` — resolved during archive (design.md updated to match code).

**S1**: Success criteria checklist in tasks.md had 3 unchecked items that were verified — resolved during archive.

**S2**: Import structure in design.md File Changes table did not match actual implementation — resolved during archive.

## Artifact Inventory

All artifacts are located in `openspec/archive/remove-hardcoded-creds/`:

| Artifact | Topic Key | Description |
|----------|-----------|-------------|
| Exploration | `sdd/remove-hardcoded-creds/explore` | Initial codebase investigation and approaches |
| Proposal | `sdd/remove-hardcoded-creds/proposal` | Problem scope, intent, success criteria |
| Design | `sdd/remove-hardcoded-creds/design` | Architecture decisions, data flow, edge cases |
| Tasks | `sdd/remove-hardcoded-creds/tasks` | 9 tasks, dependency graph, apply order |
| Archive Report | `sdd/remove-hardcoded-creds/archive-report` | This document |

## Key Decisions

1. **Secret key source**: `os.environ["MULTIAGENTE_SECRET_KEY"]` read from `.env` via `python-dotenv` at each call site (no attribute caching needed — `os.environ` is already a process-level dict).
2. **Key auto-generation**: `secrets.token_hex(32)` generates a 64-char hex key on first run if `.env` is missing; written to `.env` silently.
3. **Decrypt-failure handling**: If decrypted value still starts with `enc::`, return `""` (empty) instead of garbage — user re-enters password in Settings.
4. **`.env` path**: Same `base_dir` logic as `config_file` — `sys.executable` when frozen, `__file__` dirname otherwise.
5. **No migration**: Old encrypted blobs remain in `ui_config.json` but become undecryptable — user re-enters password once.

## Lessons Learned

- The `os.urandom(32).hex()` vs `secrets.token_hex(32)` discrepancy in the design doc was caught by verify — both produce identical results, but the code block in tasks.md correctly specified `secrets.token_hex(32)` from the start, indicating the design decision table was written from memory rather than cross-referenced against the task spec.
- The import structure deviation (single-line vs. 3-line imports) was a cosmetic difference caused by the code auto-formatter re-sorting imports — worth noting that design docs should either enforce a formatter or be reviewed post-apply for accuracy.
- All 9 tasks were implemented correctly with zero CRITICAL findings, confirming the SDD workflow successfully caught edge cases like decrypt-failure handling and frozen .exe paths during the design phase.

## Rollback

Revert palette.py, `__init__.py`, and ui_app.py changes. Delete `.env.example` and `requirements.txt` created by this change. Delete `.env` if present. Old encrypted blobs in `ui_config.json` remain intact — the rollback restores the old hardcoded key so decryption works again.
