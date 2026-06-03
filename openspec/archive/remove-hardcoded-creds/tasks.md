# Tasks: Remove Hardcoded Credentials

## Review Workload Forecast

| Metric | Value |
|--------|-------|
| Estimated lines changed | ~50 (2 removed, ~48 added/modified across 5 files) |
| Estimated files changed | 5 (3 modified, 2 created) |
| Chained PRs recommended | No |
| 400-line budget risk | Low (<50 lines) |
| Decision needed before apply | No |

---

## Task Dependency Graph

```
T1 ──┐
T2 ──┤   T5 ── T6 ── T7 ── T8 ── T9
T3 ──┤     ↑ (all in ui_app.py, sequential)
T4 ──┘

Legend:
  ──  sequential (must wait for predecessor)
  ──┐  parallel (independent files)
```

---

## Task List

### [x] T1: Create `requirements.txt` with `python-dotenv` dependency

**Spec ref**: Proposal — Dependencies section: "Add python-dotenv as dependency" / "Existing codebase has no requirements.txt — may need to create one"

**File**: `requirements.txt` (new file in project root)

**Action**: Create `requirements.txt` containing:

```
python-dotenv>=1.0.0
```

No other dependencies — `_instalar_deps_ui()` handles existing deps via auto-install.

**Verification**: `git diff --stat` shows 1 file created, 1 line. `cat requirements.txt` shows the expected content.

**Blocks**: Nothing (T5 reads this at build time, not at file-change time).

---

### [x] T2: Create `.env.example` template

**Spec ref**: Proposal — Scope: "Create .env.example with `MULTIAGENTE_SECRET_KEY=<generate-a-secure-key>`"

**File**: `.env.example` (new file in project root)

**Action**: Create `.env.example` containing:

```
# Secret key used for encrypting/decrypting stored credentials.
# Generate a 64-character hex key:
#   python -c "import secrets; print(secrets.token_hex(32))"
MULTIAGENTE_SECRET_KEY=<your-64-char-hex-key>
```

**Verification**: File exists in project root. The key name `MULTIAGENTE_SECRET_KEY` matches what `_ensure_secret_key()` and `_cfg_obtener*` will read. Comment explains how to generate.

**Blocks**: Nothing (documentation-only artifact).

---

### [x] T3: Remove `USUARIO` and `PASSWORD` from `constants/palette.py`

**Spec ref**: Proposal — Scope: "Remove USUARIO, PASSWORD from constants/palette.py"

**Design ref**: File Changes table — "Remove USUARIO = \"\" and PASSWORD = \"\" (lines 7-8)"

**File**: `constants/palette.py`, lines 7-8

**Action**: Delete both lines:

```python
# Line 7 (delete)
USUARIO = ""
# Line 8 (delete)
PASSWORD = ""
```

Leave `IMAP_SERVER` and `PUERTO_IMAP` intact (non-secret config, in-scope for keeping).

**Verification**: `grep -n "USUARIO\|PASSWORD" constants/palette.py` returns nothing. `IMAP_SERVER` is still present.

**Blocks**: T4 (import removal depends on palette having removed the vars).

---

### [x] T4: Remove `USUARIO, PASSWORD` from `constants/__init__.py` re-export

**Spec ref**: Proposal — Scope: "Remove re-export in constants/__init__.py"

**Design ref**: File Changes table — "Remove USUARIO, PASSWORD from the re-export import (line 2)"

**File**: `constants/__init__.py`, line 2

**Action**: Change line 2 from:

```python
from .palette import USUARIO, PASSWORD, IMAP_SERVER, PUERTO_IMAP
```

to:

```python
from .palette import IMAP_SERVER, PUERTO_IMAP
```

**Verification**: `constants/__init__.py` no longer exports `USUARIO` or `PASSWORD`. `IMAP_SERVER` and `PUERTO_IMAP` still export. Running `python -c "from constants import USUARIO"` fails with `ImportError`.

**Blocks**: T5 (ui_app.py import removal can proceed once the name no longer exists upstream — but the downstream import removal happens in T5).

---

### [x] T5: Clean up `ui_app.py` imports — remove `USUARIO` / `PASSWORD` and add `import dotenv`

**Spec ref**: Proposal — Scope: "Remove USUARIO, PASSWORD from ... ui_app.py imports" and "Add import dotenv"

**Design ref**: Architecture Decision: `.env` loading point — "Module-level block at line ~100, after all imports, just before class App"

**File**: `ui_app.py`, lines 73-74

**Action**:

a) Change line 73 from:
```python
from constants import USUARIO, PASSWORD, IMAP_SERVER, PUERTO_IMAP
```
to:
```python
from constants import IMAP_SERVER, PUERTO_IMAP
```

b) Add `import dotenv` among the stdlib imports at the top of the file (after line 28 `import shutil`, before the `_instalar_deps_ui` function block). Suggested location: after line 28.

**Verification**: `grep -n "from constants" ui_app.py | grep "USUARIO\|PASSWORD"` returns nothing. `grep -n "^import dotenv" ui_app.py` shows the import present before `_instalar_deps_ui`.

**Blocks**: T7, T8 (downstream references to `USUARIO`/`PASSWORD` and hardcoded key will cause `NameError` at module import if not cleaned up first).

---

### [x] T6: Add module-level `.env` loading and `_ensure_secret_key()` function

**Spec ref**: Proposal — Approach steps 1-3: "load_dotenv at startup", "generate random key on first run", "replace all uses of hardcoded key"

**Design ref**: Architecture decisions for loading point, key generation, .env path, and auto-generation. Data Flow diagram (startup sequence).

**File**: `ui_app.py`

**Action**: Insert a new module-level block between line 74 (imports end) and line 81 (class OutputRedirector). The block must:

a) Determine `.env` path using the same `base_dir` logic as lines 109-112:
   - If frozen (`getattr(sys, 'frozen', False)`): `base_dir = os.path.dirname(sys.executable)`
   - Otherwise: `base_dir = os.path.dirname(__file__)`
   - `dotenv_path = os.path.join(base_dir, ".env")`

b) Call `dotenv.load_dotenv(dotenv_path)` to populate `os.environ`.

c) Define `_ensure_secret_key(dotenv_path: str) -> None`:
   - Check `os.environ.get("MULTIAGENTE_SECRET_KEY")`
   - If missing: generate `os.urandom(32).hex()` (64-char hex key)
   - Write `MULTIAGENTE_SECRET_KEY=<generated_key>` to the `.env` file
   - Set `os.environ["MULTIAGENTE_SECRET_KEY"]` for the current process
   - If `.env` is not writable, log a warning via `print()` and continue with in-memory key only (so the next session still auto-generates)

d) Call `_ensure_secret_key(dotenv_path)` immediately after the function definition.

The complete block should look like:

```python
# ── .env loading for secret key ─────────────────────────────────────────
_frozen = getattr(sys, 'frozen', False)
_base_dir = os.path.dirname(sys.executable) if _frozen else os.path.dirname(__file__)
_dotenv_path = os.path.join(_base_dir, ".env")

dotenv.load_dotenv(_dotenv_path)


def _ensure_secret_key(dotenv_path: str) -> None:
    key = os.environ.get("MULTIAGENTE_SECRET_KEY")
    if key:
        return
    import secrets
    key = secrets.token_hex(32)
    try:
        with open(dotenv_path, "w", encoding="utf-8") as f:
            f.write(f"MULTIAGENTE_SECRET_KEY={key}\n")
    except OSError:
        print(f"[WARN] Could not write {dotenv_path} — key in memory only for this session.")
    os.environ["MULTIAGENTE_SECRET_KEY"] = key


_ensure_secret_key(_dotenv_path)
```

**Verification**: 
- `python -c "exec(open('ui_app.py').read()); import os; assert len(os.environ['MULTIAGENTE_SECRET_KEY']) == 64"` (after import)
- When `.env` does not exist, it is created with a valid 64-char hex key
- When `.env` exists with `MULTIAGENTE_SECRET_KEY`, it is read and no new key is generated
- `MULTIAGENTE_SECRET_KEY` is present in `os.environ` before `class App` definition

**Blocks**: T8 (key replacement needs `os.environ["MULTIAGENTE_SECRET_KEY"]` to be available; T9 (decrypt failure handling) is independent). If T8 runs before T6, the key lookup will crash at import time.

---

### [x] T7: Replace all `USUARIO` / `PASSWORD` default references with `""` in `ui_app.py`

**Spec ref**: Proposal — Scope: "Update `_cfg_obtener_correo()` defaults to `\"\"` instead of `USUARIO`/`PASSWORD`"

**Design ref**: The design does not enumerate individual call sites but states the intent that `USUARIO`/`PASSWORD` defaults become `""`.

**File**: `ui_app.py`

**Action**: Replace every occurrence of `USUARIO` and `PASSWORD` as a default argument with its empty-string equivalent. Based on grep results, the affected lines are:

| Line | Current | Replace with |
|------|---------|--------------|
| 2173 | `self._cfg_obtener_correo("usuario", USUARIO)` | `self._cfg_obtener_correo("usuario", "")` |
| 2174 | `self._cfg_obtener_correo("password", PASSWORD)` | `self._cfg_obtener_correo("password", "")` |
| 5321 | `self._cfg_obtener_correo("usuario", USUARIO)` | `self._cfg_obtener_correo("usuario", "")` |
| 5346 | `self._cfg_obtener_correo("usuario", USUARIO)` | `self._cfg_obtener_correo("usuario", "")` |
| 5363 | `self._cfg_obtener_correo("usuario", USUARIO)` | `self._cfg_obtener_correo("usuario", "")` |
| 5408 | `self._cfg_obtener_correo("usuario", USUARIO)` | `self._cfg_obtener_correo("usuario", "")` |
| 5409 | `self._cfg_obtener_correo("password", PASSWORD)` | `self._cfg_obtener_correo("password", "")` |
| 6077 | `self._cfg_obtener_correo("usuario", USUARIO)` | `self._cfg_obtener_correo("usuario", "")` |
| 6079 | `self._cfg_obtener_correo("password", PASSWORD)` | `self._cfg_obtener_correo("password", "")` |

**Verification**: `grep -n "USUARIO\|PASSWORD" ui_app.py` returns nothing. All 9 default-reference call sites have been updated.

**Blocks**: Nothing. This is a mechanical text replacement and can be done independently of T8 and T9.

---

### [x] T8: Replace all `"ENCRYPTION_KEY_PLACEHOLDER"` literals with `os.environ["MULTIAGENTE_SECRET_KEY"]`

**Spec ref**: Proposal — Scope: "Remove hardcoded `\"ENCRYPTION_KEY_PLACEHOLDER\"` fallback from `ui_app.py:_cfg_obtener` (line 5763) and `_cfg_obtener_correo` (line 5769) ... Replace all uses", and Approach step 3: "Replace all uses of `\"ENCRYPTION_KEY_PLACEHOLDER\"` literal with `os.environ[\"MULTIAGENTE_SECRET_KEY\"]`"

**Design ref**: Data Flow diagrams (startup, encryption, decryption). Architecture Decision: Key generation and caching.

**File**: `ui_app.py`

**Action**: Replace 4 occurrences of `"ENCRYPTION_KEY_PLACEHOLDER"` with `os.environ["MULTIAGENTE_SECRET_KEY"]`:

| Line | Current | Replace with |
|------|---------|--------------|
| 5763 | `self._decrypt_val(val, "ENCRYPTION_KEY_PLACEHOLDER")` | `self._decrypt_val(val, os.environ["MULTIAGENTE_SECRET_KEY"])` |
| 5769 | `key = self._master_pw_cache if self._master_pw_cache else "ENCRYPTION_KEY_PLACEHOLDER"` | `key = self._master_pw_cache if self._master_pw_cache else os.environ["MULTIAGENTE_SECRET_KEY"]` |
| 6462 | `key = pw1 if pw1 else "ENCRYPTION_KEY_PLACEHOLDER"` | `key = pw1 if pw1 else os.environ["MULTIAGENTE_SECRET_KEY"]` |
| 6540 | `self._encrypt_val(pw1, "ENCRYPTION_KEY_PLACEHOLDER")` | `self._encrypt_val(pw1, os.environ["MULTIAGENTE_SECRET_KEY"])` |

**Verification**: `grep -n "ENCRYPTION_KEY_PLACEHOLDER" ui_app.py` returns nothing. `grep -n "os.environ\[.MULTIAGENTE_SECRET_KEY.\]" ui_app.py` shows 4 occurrences, one at each of the 4 original key use locations.

**Blocks**: T6 must run first (ensures `os.environ["MULTIAGENTE_SECRET_KEY"]` is populated before module-level code executes and the App class is defined). Exception: if imports are at the top and `dotenv.load_dotenv()` runs before class definition, the key exists in environ before any method runs.

---

### [x] T9: Add decrypt-failure handling in `_cfg_obtener` and `_cfg_obtener_correo`

**Spec ref**: Proposal — Approach step 4: "When decrypting with the new key fails, treat credential as unconfigured — user re-enters password in the Settings tab."

**Design ref**: Architecture Decision: Decrypt failure handling. Data Flow diagrams (decryption for mail fetch, startup). Edge Cases table — "Existing user upgrades: _cfg_obtener* returns \"\" on decrypt failure".

**File**: `ui_app.py`, lines 5760-5771

**Action**:

a) In `_cfg_obtener` (lines 5760-5764), after the decrypt call on line 5763, add a guard that returns `default` if the decrypted value still starts with `"enc::"`:

```python
def _cfg_obtener(self, seccion, clave, default):
    val = self.config.get(seccion, {}).get(clave, default)
    if seccion == "seguridad" and clave == "password" and val != default:
        val = self._decrypt_val(val, os.environ["MULTIAGENTE_SECRET_KEY"])
        if val.startswith("enc::"):
            return default
    return val
```

b) In `_cfg_obtener_correo` (lines 5766-5771), after the decrypt call on line 5770, add the same guard:

```python
def _cfg_obtener_correo(self, clave, default):
    val = self._cfg_obtener("correo", clave, default)
    if clave == "password" and val != default:
        key = self._master_pw_cache if self._master_pw_cache else os.environ["MULTIAGENTE_SECRET_KEY"]
        val = self._decrypt_val(val, key)
        if val.startswith("enc::"):
            return default
    return val
```

**Verification**: 

- Unit-like check: mock `_decrypt_val` to return `"enc::garbage..."` — assert `_cfg_obtener` returns the `default` parameter instead.
- Unit-like check: mock `_decrypt_val` to return `"my_real_password"` — assert `_cfg_obtener` returns `"my_real_password"` as before.
- The `_cfg_obtener_correo("password", "")` returns `""` when decrypt fails with either the master password or the env key.

**Blocks**: Nothing at file level — this is a logic change inside existing functions, independent of other tasks other than T8 (which updates the key reference on line 5769). Must run after T8 (or be merged with T8 in the same edit block) to avoid double-editing the same lines.

---

## Task Summary

| ID | Description | File(s) | Parallel? | Blocks |
|----|-------------|---------|-----------|--------|
| T1 | Create requirements.txt | `requirements.txt` | T1-T4 can run in parallel | Nothing |
| T2 | Create .env.example | `.env.example` | T1-T4 can run in parallel | Nothing |
| T3 | Remove USUARIO, PASSWORD from palette.py | `constants/palette.py` | T1-T4 can run in parallel | T4 |
| T4 | Remove re-exports from __init__.py | `constants/__init__.py` | Depends on T3 | T5 |
| T5 | Clean up ui_app.py imports | `ui_app.py` | Depends on T4 | T7, T8 |
| T6 | Add .env loading + _ensure_secret_key | `ui_app.py` | Depends on T5 | T8 |
| T7 | Replace USUARIO/PASSWORD defaults with "" in ui_app.py | `ui_app.py` | Depends on T5 | Nothing (mechanical) |
| T8 | Replace hardcoded key with os.environ read | `ui_app.py` | Depends on T6 | T9 (or merge) |
| T9 | Add decrypt-failure handling | `ui_app.py` | Depends on T8 | Nothing |

## Apply Order

**Batch 1 (parallel, independent files)**:
1. T1 — requirements.txt
2. T2 — .env.example
3. T3 — constants/palette.py

**Batch 2 (sequential, single file)**:
4. T4 — constants/__init__.py (depends on T3)

**Batch 3 (all in ui_app.py — do in one pass to avoid merge conflicts)**:
5. T5 — import cleanup (remove USUARIO/PASSWORD, add dotenv)
6. T6 — .env loading + _ensure_secret_key() (insert after imports)
7. T7 — replace USUARIO/PASSWORD defaults with ""
8. T8 — replace "ENCRYPTION_KEY_PLACEHOLDER" with os.environ
9. T9 — decrypt-failure handling guard

T5-T9 touch the same file (ui_app.py). The recommended apply approach is to make all 5 changes in a single editing session to avoid line-number drift and merge conflicts.

## Success Criteria Verification

After all tasks are applied, validate:

- [x] `grep -r "ENCRYPTION_KEY_PLACEHOLDER" constants/ ui_app.py` returns nothing
- [x] `grep -rn "USUARIO\|PASSWORD" constants/ ui_app.py` returns nothing
- [x] `.env.example` exists in project root with correct template
- [x] `requirements.txt` exists with `python-dotenv>=1.0.0`
- [x] `.env` is in `.gitignore` (confirm — design states it's already present)
- [x] Fresh run without `.env` generates one automatically
- [x] Old `enc::` blobs in `ui_config.json` fail gracefully (return "" instead of garbage)
