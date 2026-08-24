"""Post-build fixup: ensure _internal/ has the OpenSSL 3 DLLs matching this
CPython build.

Why: PyInstaller's binary analysis can collect a third-party libcrypto-3.dll
(e.g. the one shipped in engines/tesseract/) into _internal/ root. The bundled
_hashlib.pyd resolves that copy first, fails with WinError 127 (missing
procedure), and `import _hashlib` silently fails inside hashlib — which removes
hashlib.pbkdf2_hmac and breaks password encryption at runtime.

Usage: python fix_openssl_dlls.py [dist_dir]
Exits non-zero if verification fails.
"""

import ctypes
import hashlib
import importlib.util
import os
import shutil
import sys

DLL_NAMES = ("libcrypto-3.dll", "libssl-3.dll")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def dev_dlls_dir():
    base = getattr(sys, "base_prefix", sys.prefix)
    d = os.path.join(base, "DLLs")
    if not os.path.isdir(d):
        raise SystemExit(f"ERROR: dev DLLs dir not found: {d}")
    return d


def fix(dist_internal):
    src_dir = dev_dlls_dir()
    for name in DLL_NAMES:
        src = os.path.join(src_dir, name)
        dst = os.path.join(dist_internal, name)
        if not os.path.isfile(src):
            raise SystemExit(f"ERROR: source DLL missing: {src}")
        if os.path.isfile(dst) and sha256(src) == sha256(dst):
            print(f"  [ok] {name} already correct ({sha256(src)[:16]})")
            continue
        shutil.copy2(src, dst)
        print(f"  [fix] replaced {name}: {sha256(dst)[:16]}")


def verify(dist_internal):
    pyd = os.path.join(dist_internal, "_hashlib.pyd")
    if not os.path.isfile(pyd):
        raise SystemExit(f"ERROR: bundled _hashlib.pyd not found: {pyd}")
    for name in DLL_NAMES:
        dst = os.path.join(dist_internal, name)
        src = os.path.join(dev_dlls_dir(), name)
        if not os.path.isfile(dst) or sha256(dst) != sha256(src):
            raise SystemExit(f"ERROR: {name} does not match the CPython build")
    os.add_dll_directory(dist_internal)
    spec = importlib.util.spec_from_file_location("_hashlib", pyd)
    try:
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
    except ImportError as e:
        raise SystemExit(f"ERROR: bundled _hashlib.pyd failed to import: {e}")
    if not hasattr(mod, "pbkdf2_hmac"):
        raise SystemExit(
            "ERROR: bundled _hashlib imported but pbkdf2_hmac is missing "
            "(wrong libcrypto still being resolved)"
        )
    print("  [ok] bundled _hashlib imports with pbkdf2_hmac")


def main():
    dist = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "dist", "Sistema_Automatizacion",
    )
    internal = os.path.join(dist, "_internal")
    if not os.path.isdir(internal):
        raise SystemExit(f"ERROR: not found: {internal} (build first?)")
    print(f"Fixing OpenSSL DLLs in {internal}")
    fix(internal)
    verify(internal)
    print("OpenSSL DLLs OK")


if __name__ == "__main__":
    main()
