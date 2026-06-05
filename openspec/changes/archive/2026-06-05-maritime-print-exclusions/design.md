# Design: maritime-print-exclusions

## Technical Approach

Skip ATA/Recibo ATA printing for maritime folders (ISO/FLEXI) in both
print execution paths by parsing the transport type from the folder name,
which already encodes it from the mail-processing phase. When the type is
ISO or FLEXI, log a skip message and bypass the ATA sheet lookup + print.

## Architecture Decisions

### Decision: Folder-name parsing over Excel classification

| Option | Tradeoff |
|--------|----------|
| Parse `split("_")[2]` from `os.path.basename(ruta)` | Zero I/O, no Excel dependency |
| Open CONTENEDORES and call `_clasificar_tipo_transporte` | Duplicate-heavy, reads the same file twice per folder |

**Choice**: Parse folder name. The type at position 2 is set during
mail processing (`_mail_nombre_carpeta`) per the mail-folder-naming spec,
so re-reading Excel would be wasted work.

### Decision: Private method on `App` vs standalone function vs inline

**Choice**: Method `_detectar_tipo_carpeta(self, nombre: str) -> str`
on `App`. Both callers (`_imp_worker`, `_super_imprimir`) already have
`self` and the codebase uses private methods for all helpers.

**Alternatives considered**: Standalone function (acceptable but
inconsistent with project conventions), inline detection in each path
(duplicates the parsing and the `TERRESTRE` default).

### Decision: Guard wraps only the ATA-printing sub-block

No `continue` — the guard is placed inside the existing
`if opciones.get("servicio_ata"):` / `if hacer_recibo and sobres:`
blocks. This keeps the ATA detection and print loop together while
preserving the ability to add print steps after ATA without breakage.

### Decision: Log skip messages, don't stay silent

Both paths already log every print action. Adding a clear skip message
(e.g., "⏭ Saltando Recibo ATA: carpeta marítima (ISO)") makes the
behavior transparent during debugging.

## Data Flow

```
Folder name (os.path.basename) ──→ split("_") → partes[2]
                                              ↓
            "TERRESTRE" ◄── fallback (safe default if < 3 parts)
                                              ↓
                     ┌─── "ISO" / "FLEXI" ──→ skip ATA (log message)
                     │
                     └─── "TERRESTRE" ──────→ existing ATA lookup + print
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `ui_app.py` | Modify | +1 helper method `_detectar_tipo_carpeta` (~6 lines), +2 guard blocks in `_imp_worker` (~L1524, ~4 lines) and `_super_imprimir` (~L3150, ~4 lines) |

Total: ~14 lines added, 0 modified, 0 deleted.

## Helper Contract

```python
def _detectar_tipo_carpeta(self, nombre_carpeta: str) -> str:
    """Return transport type from folder name: TERRESTRE, ISO, or FLEXI.

    Parses the third segment (0-indexed, index 2) of the underscore-
    separated folder name set during mail processing. Old-format
    folders without enough segments default to TERRESTRE.
    """
    partes = nombre_carpeta.split("_")
    if len(partes) < 3:
        return "TERRESTRE"
    tipo = partes[2]
    if tipo not in ("TERRESTRE", "ISO", "FLEXI"):
        return "TERRESTRE"    # unknown type → safe default
    return tipo
```

### Guard in `_imp_worker` (panel manual, ~L1524)

```python
# 4. Servicio ATA / Recibo ATA
if opciones.get("servicio_ata"):
    tipo = self._detectar_tipo_carpeta(nombre)
    if tipo in ("ISO", "FLEXI"):
        self._log(f"  ⏭ Saltando Recibo ATA: carpeta marítima ({tipo})")
    else:
        # existing ATA lookup + print logic unchanged
        ...
```

### Guard in `_super_imprimir` (súper auto, ~L3150)

```python
# 4. Servicio ATA / Recibo ATA
if hacer_recibo and sobres:
    tipo = self._detectar_tipo_carpeta(nombre)
    if tipo in ("ISO", "FLEXI"):
        self.log_queue.put(f"[...]   ⏭ Saltando Recibo ATA: carpeta marítima ({tipo})")
    else:
        # existing ATA lookup + print logic unchanged
        ...
```

Note: `_super_imprimir` uses `log_queue.put` for thread-safe logging, while
`_imp_worker` uses `self._log` (called from background thread via `after`).

## Testing Strategy

Manual verification on real data — no test suite exists for these methods.

| Layer | What | Approach |
|-------|------|----------|
| Manual | ISO folder in both paths | Select folder with `ISO` at index 2, verify ATA skip logged |
| Manual | FLEXI folder in both paths | Same as ISO |
| Manual | TERRESTRE folder (existing) | Verify ATA still prints |
| Manual | Old-format folder (no type) | Verify ATA still prints (TERRESTRE default) |

## Migration / Rollout

No migration required. Change is purely runtime logic — existing folders
on disk are unchanged. Deploy by replacing `ui_app.py`.

## Open Questions

None.
