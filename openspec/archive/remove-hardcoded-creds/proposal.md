# Proposal: Remove Hardcoded Credentials

## Intent

The encryption key `"ENCRYPTION_KEY_PLACEHOLDER"` is hardcoded in source code
(ui_app.py, two locations). Anybody with repo access can decrypt stored
passwords. Move the key to a `.env` file loaded via `python-dotenv` and purge all
credential defaults from source constants.

## Scope

### In Scope

- Add `python-dotenv` as dependency
- Create `.env.example` with `MULTIAGENTE_SECRET_KEY=<generate-a-secure-key>`
- Remove `USUARIO`, `PASSWORD` from `constants/palette.py` and its re-export in
  `constants/__init__.py`
- Remove hardcoded `"ENCRYPTION_KEY_PLACEHOLDER"` fallback from
  `ui_app.py:_cfg_obtener` (line 5763) and `_cfg_obtener_correo` (line 5769)
- Load `MULTIAGENTE_SECRET_KEY` from `os.getenv(...)` at startup; generate and
  persist one if none exists
- Update `_cfg_obtener_correo()` defaults to `""` instead of `USUARIO`/`PASSWORD`
- If decryption with new key fails, treat as unconfigured — user re-enters password in Settings
- Keep `DESTINATARIOS_*` in `palette.py` (non-secret, used as Settings defaults)

### Out of Scope

- Encrypt/decrypt algorithm itself (stays as-is)
- `IMAP_SERVER`, `PUERTO_IMAP` (org config, not secret)
- `keyring` or OS-level credential storage
- Removing `DESTINATARIOS_*` from `palette.py` (are Settings defaults, not secrets)

## Capabilities

### New Capabilities

- `env-config`: environment-driven secret management via `.env` + `python-dotenv`

### Modified Capabilities

None — pure infrastructure refactor, no spec-level behavior change.

## Approach

1. Add `python-dotenv` to dependencies (create `requirements.txt` or append to
   existing). Add `import dotenv; dotenv.load_dotenv()` at startup before config
   loads.
2. Generate a random 32-byte hex key on first run if `MULTIAGENTE_SECRET_KEY` is
   unset; persist it into `.env` automatically.
3. Replace all uses of `"ENCRYPTION_KEY_PLACEHOLDER"` literal with
   `os.environ["MULTIAGENTE_SECRET_KEY"]`.
4. When decrypting with the new key fails, treat credential as unconfigured —
   user re-enters password in the Settings tab.
5. Remove `USUARIO`, `PASSWORD` from palette.py; replace fallback defaults with
   `""` in `_cfg_obtener_correo()` calls.
6. Stop importing `USUARIO`, `PASSWORD` in `constants/__init__.py`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `constants/palette.py` | Modified | Remove `USUARIO`, `PASSWORD` |
| `constants/__init__.py` | Modified | Remove `USUARIO`, `PASSWORD` re-exports |
| `ui_app.py` | Modified | 4 import lines, 2 hardcoded key refs, config defaults, add migration + .env loading |
| `.env.example` | Created | Template with `MULTIAGENTE_SECRET_KEY` |
| `requirements.txt` (new) | Created | Add `python-dotenv` |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Old encrypted passwords become undecryptable | Medium | Decryption failure treated as unconfigured — user re-enters password in Settings. Low impact: single field. |
| `.env` not loaded early enough | Low | Call `load_dotenv()` at module top, before `App.__init__` |
| `.env` missing on frozen .exe | Low | Key is generated at runtime if missing; .env ships alongside .exe |

## Rollback Plan

Revert palette.py, `__init__.py`, and ui_app.py changes. Revert to old
hardcoded key. The `.env` file can be deleted — it's not read without the code
change. No data migration is destructive because old blobs remain in
`ui_config.json` undamaged.

## Dependencies

- `python-dotenv` (PyPI)
- Existing codebase has no `requirements.txt` — may need to create one

## Success Criteria

- [x] No hardcoded `USUARIO`, `PASSWORD`, or `"ENCRYPTION_KEY_PLACEHOLDER"` remain
  in source
- [x] Existing encrypted passwords in `ui_config.json` failure is graceful — user prompted to re-enter
- [x] Fresh install without `.env` auto-generates a secret key and works
- [x] `.env` is listed in `.gitignore` (confirmed: already present)
- [x] `git diff --stat` shows <= 3 files modified + 1 created
