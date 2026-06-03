# Exploration: remove-hardcoded-creds

## Current State

The application manages credentials through two layers:

### Layer 1 — Hardcoded defaults in `constants/palette.py`

| Variable | Value | Sensitive? |
|----------|-------|------------|
| `USUARIO` | `""` (empty) | No (placeholder) |
| `PASSWORD` | `""` (empty) | No (placeholder) |
| `IMAP_SERVER` | `"imap.empresa.com"` | No (config, not secret) |
| `PUERTO_IMAP` | `143` | No (config, not secret) |
| `DESTINATARIOS_GRUPAL` | 14 real emails | Mildly (org data) |
| `DESTINATARIOS_INDIVIDUAL` | 3 real emails | Mildly (org data) |
| Fallback encryption key | `"ENCRYPTION_KEY_PLACEHOLDER"` | **Yes** |

### Layer 2 — Runtime config (`ui_config.json`)

- Contains actual active credentials (encrypted password, server, recipients)
- Password stored as `enc::<base64(salt+encrypted)>`
- Loaded by `App._cargar_config()` at startup

### Credential Flow

```
constants/palette.py          ui_config.json
  (hardcoded defaults)          (runtime values)
        |                              |
        v                              v
  constants/__init__.py                |
        |                              |
        +--- re-exported --------------+
                    |
                    v
           ui_app.py (imports at lines 72-74)
                    |
           _cfg_obtener_correo(key, default)
                    |
          Checks self.config["correo"][key]
                    |
        +-----------+-----------+
        |                       |
    Has value?             Not found?
        |                       |
    Decrypt if              Return default
    key == password         (from constants)
        |
    Return value
```

### Encryption Details (ui_app.py lines 6569-6603)

- Custom XOR + PBKDF2-HMAC-SHA256
- Format: `enc::<base64(salt+ciphertext)>`
- Key sources: master password (user-configured) or hardcoded fallback `"ENCRYPTION_KEY_PLACEHOLDER"`
- **Weakness**: `dklen=len(plaintext)` leaks password length; no authenticated encryption

### Key Finding

The constants in `palette.py` are **fallbacks only**. When `ui_config.json` has the values (which it always does after initial setup), the constants are never reached. They exist solely as defaults when a config key is missing.

## Approaches Considered

1. **Remove hardcoded defaults only** — Minimal changes. The app already reads from `ui_config.json`. Low effort.
2. **Environment variables via `.env`** — Industry standard. Requires `python-dotenv`, migration path for existing encrypted passwords. Medium effort.
3. **Separate credentials file + `.gitignore`** — Simple, keeps secrets out of VCS. Low-Medium effort.
4. **OS keyring (`keyring` library)** — Most secure, OS-native. Cross-platform issues for `.exe` distribution. High effort.

## Affected Files

- `constants/palette.py` — hardcoded values live here
- `constants/__init__.py` — re-exports
- `ui_app.py` — imports at lines 72-74, usage in `_cfg_obtener_correo()` (line 5766), encrypt/decrypt (lines 6569-6603), IMAP connection (line 2168), SMTP recovery (line 5787), email composition (lines 5319-5364)
- `ui_config.json` — runtime config with actual values

## Risks

- Removing `DESTINATARIOS_*` breaks Settings tab defaults
- Existing encrypted passwords tied to hardcoded key — changing the key breaks decryption
- `dist/ui_config.json` shipped with frozen executable must stay in sync
- Password recovery flow depends on the same decryption pipeline
