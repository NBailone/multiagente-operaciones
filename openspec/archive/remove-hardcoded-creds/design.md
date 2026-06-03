# Design: Remove Hardcoded Credentials

## Technical Approach

Extract the encryption key `"ENCRYPTION_KEY_PLACEHOLDER"` from source into a `.env` file loaded at startup via `python-dotenv`. Generate a random 64-character hex key on first run. When the new key cannot decrypt existing stored blobs, treat the credential as unconfigured (return `""`) rather than surfacing the raw `enc::...` blob. No migration, no fallback to the old key.

---

## Architecture Decisions

### Decision: .env loading point

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Top of `ui_app.py` before imports | Simplest but `dotenv` import happens mid-file | Rejected — Python executes imports first anyway |
| After stdlib imports, before `class App` | Clean, no chicken-egg problem | **Chosen** |
| Inside `App.__init__` before `_cargar_config` | Works but delays env var visibility to any module-level code | Rejected — too late |

**Chosen**: Module-level block at line ~100, after all imports, just before `class App(ctk.CTk)`. This guarantees `os.environ["MULTIAGENTE_SECRET_KEY"]` is set before `App.__init__` calls `_cargar_config()`.

### Decision: Key generation and caching

| Option | Tradeoff | Decision |
|--------|----------|----------|
| `os.environ` only, no caching | One lookup per encrypt/decrypt (negligible cost) | **Chosen** — env is already an in-process dict |
| `App._encryption_key` attribute | Extra attribute, must be set before parent calls | Rejected — no measurable benefit |
| Module-level constant | Can't mutate if key is generated at runtime | Rejected — const pattern doesn't fit |

**Chosen**: Read from `os.environ["MULTIAGENTE_SECRET_KEY"]` at each call site. No attribute caching needed — `os.environ` is already a cached process-level dict.

### Decision: Decrypt failure handling

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Return `enc::...` blob as-is | Dialog shows garbage, user blocked | Rejected — breaks UX |
| Return `""` (empty) | Treated as "no password", user re-enters in Settings | **Chosen** |
| Show error dialog on startup | Extra UI complexity, blocks flow | Rejected — too invasive |

**Chosen**: In `_cfg_obtener` and `_cfg_obtener_correo`, if `_decrypt_val` returns a value that still starts with `enc::` (meaning decrypt failed), return the default (`""`) instead. The calling code already handles empty passwords gracefully — the IMAP connection fails with a retry, the master password check is skipped (`_pw_inicio_valida = True`), and Settings shows empty fields.

### Decision: `.env` path for frozen .exe

| Option | Tradeoff | Decision |
|--------|----------|----------|
| `load_dotenv()` with cwd | May not match .exe location | Rejected — unreliable |
| Explicit path from `sys.executable` | Same pattern as `config_file`, always correct | **Chosen** |

**Chosen**: Use the same `base_dir` logic as `config_file` (line 109-112): `sys.executable` dirname when frozen, `__file__` dirname otherwise. Pass explicit path to `dotenv.load_dotenv()`.

### Decision: Key auto-generation into `.env`

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Only read, never write | Admin must create `.env` manually | Rejected — first-run friction |
| Generate and write `.env` silently | Fully automatic, 64-char hex key via `secrets.token_hex(32)` | **Chosen** |
| Generate and print instructions | User has to copy-paste | Rejected — unnecessary step |

**Chosen**: Module-level function `_ensure_secret_key()` that checks `os.environ.get("MULTIAGENTE_SECRET_KEY")`, generates if missing, writes to `.env` on the same determined path, and sets `os.environ["MULTIAGENTE_SECRET_KEY"]` for the current process.

### Decision: requirements.txt scope

**Chosen**: Single dependency `python-dotenv>=1.0.0`. No pinning of existing dependencies — they are auto-installed by `_instalar_deps_ui()` and this change does not modify them.

---

## Data Flow

### Startup Sequence

```
ui_app.py module loads
        │
        ▼
  Determine .env path (sys.executable if frozen, __file__ otherwise)
        │
        ▼
  dotenv.load_dotenv(path)     ← reads MULTIAGENTE_SECRET_KEY if present
        │
        ▼
  _ensure_secret_key()          ← generates random key if env var is missing
        │                           writes to .env file
        │                           sets os.environ["MULTIAGENTE_SECRET_KEY"]
        ▼
  class App.__init__
        │
        ▼
  _cargar_config()              ← reads ui_config.json (may contain enc:: blobs from old key)
        │
        ▼
  _cfg_obtener / _cfg_obtener_correo     ← tries decrypt with env key
        │
        ├── Decrypt succeeds → use decrypted value
        └── Decrypt fails (output still starts with "enc::") → return ""
```

### Encryption Flow (Settings Save)

```
User fills Settings form
        │
        ▼
  _guardar_ajustes()
        │
        ├── master password set?   → encrypt with master password
        └── no master password?    → encrypt with os.environ["MULTIAGENTE_SECRET_KEY"]
        │
        ▼
  _encrypt_val(plaintext, key)   → enc::<base64(salt+xor_ciphertext)>
        │
        ▼
  Store in self.config["correo"]["password"]
```

### Decryption Flow (Mail Fetch)

```
_imap_conectar()
        │
        ▼
  _cfg_obtener_correo("password", "")
        │
        ├── key = _master_pw_cache or os.environ["MULTIAGENTE_SECRET_KEY"]
        │
        ▼
  _decrypt_val(enc::blob, key)
        │
        ├── Success → returned to imaplib.login()
        └── Failure → return "" → imaplib.login("", "") fails → user re-enters in Settings
```

---

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `constants/palette.py` | Modify | Remove `USUARIO = ""` and `PASSWORD = ""` (lines 7-8) |
| `constants/__init__.py` | Modify | Remove `USUARIO, PASSWORD` from the re-export import (line 2) |
| `ui_app.py` | Modify | ~12 lines changed: `import dotenv` added, `from constants import` split into 3 lines (IMAP_SERVER, PUERTO_IMAP, DESTINATARIOS_*), module-level .env loading block inserted (~22 lines), 9 USUARIO/PASSWORD defaults replaced with `""`, 4 hardcoded key refs replaced with `os.environ[...]`, decrypt-failure guards added |
| `.env.example` | Create | Template with `MULTIAGENTE_SECRET_KEY=<your-64-char-hex-key>` |
| `requirements.txt` | Create | `python-dotenv>=1.0.0` |

---

## Interfaces / Contracts

No new public interfaces. One internal function:

```
def _ensure_secret_key(dotenv_path: str) -> None:
    """Generate and persist MULTIAGENTE_SECRET_KEY if not set.

    Checks os.environ after dotenv.load_dotenv(). If still missing,
    generates secrets.token_hex(32), writes to .env, sets os.environ.
    If .env is not writable, logs warning and continues with in-memory only.
    """
```

---

## Edge Cases

| Scenario | Behavior | Detail |
|----------|----------|--------|
| **Fresh install, no `.env`** | Key auto-generated first run | `_ensure_secret_key()` creates `.env` alongside the app |
| **Existing user upgrades** | Old encrypted passwords become undecryptable | `_cfg_obtener*` returns `""` on decrypt failure — user re-enters in Settings |
| **`.env` missing on frozen .exe** | Key generated, `.env` written next to .exe | Same path logic as `config_file` via `sys.executable` |
| **`.env` not writable** | Key in memory only, warning logged | Passwords re-enterable each session until `.env` is fixed |
| **User clears password in Settings** | Empty string stored | Already handled — `_encrypt_val("", key)` returns `""` |
| **User clears master password** | `_master_pw_cache` set to `""`, env key becomes fallback | `key = pw1 if pw1 else os.environ["MULTIAGENTE_SECRET_KEY"]` |
| **`MULTIAGENTE_SECRET_KEY` contains newlines/spaces** | Handled by `python-dotenv` value parsing | Key is hex-only (no whitespace concerns) |

---

## Migration / Rollout

No migration required. Old encrypted blobs in `ui_config.json` remain intact but undecryptable with the new key. User re-enters password via Settings once. Rollback: revert the 4 modified files and delete `.env` — old blobs were never modified.

---

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | `_ensure_secret_key()` generates valid 64-char hex key | Assert `len(key) == 64` and all hex chars |
| Unit | `_cfg_obtener` returns default when decrypt fails with env key | Mock `_decrypt_val` to return `enc::...`, assert result is `""` |
| Unit | `.env` not found → key auto-generated | Remove `.env`, load module, assert key exists in environ |
| Integration | Old `enc::` blob from `ui_config.json` fails gracefully | Craft a blob encrypted with old key, load app, assert password field shows empty |
| Manual | Frozen .exe loads `.env` next to exe | Build with PyInstaller, run without `.env`, verify key is generated |

---

## Open Questions

None.
