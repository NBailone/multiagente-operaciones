# Tasks: maritime-print-exclusions

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~20 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | size-exception |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Helper + guards + verify | PR 1 (single) | All phases in one PR, ~20 lines total |

## Phase 1: Helper — Add `_detectar_tipo_carpeta`

- [x] 1.1 Add `_detectar_tipo_carpeta(self, nombre_carpeta: str) -> str` method to `App` after `_detectar_impresoras` (~L1408, `ui_app.py`). Parse `nombre_carpeta.split("_")`, return `partes[2]` if ≥3 parts and value is `ISO`, `FLEXI`, or `TERRESTRE`; default to `TERRESTRE`.

## Phase 2: Guards — Skip ATA in both print paths

- [x] 2.1 In `_imp_worker` (~L1524, `ui_app.py`): inside `if opciones.get("servicio_ata"):`, call `self._detectar_tipo_carpeta(nombre)`. If result is `ISO` or `FLEXI`, log `⏭ Marítimo ({tipo}): omitiendo Recibo ATA` via `self._log(...)` and skip the ATA block via `else`.
- [x] 2.2 In `_super_imprimir` (~L3150, `ui_app.py`): inside `if hacer_recibo and sobres:`, call `self._detectar_tipo_carpeta(nombre)`. If result is `ISO` or `FLEXI`, log `⏭ Marítimo ({tipo}): omitiendo Recibo ATA` via `self.log_queue.put(...)` and skip the ATA block via `else`.

## Phase 3: Verification — Manual smoke test

- [ ] 3.1 Pick a real ISO folder (splits to `ISO` at index 2), print via panel manual — verify the skip message appears and no ATA sheet is sent.
- [ ] 3.2 Pick a real FLEXI folder — same check in both manual and súper auto paths.
- [ ] 3.3 Pick a TERRESTRE folder — verify ATA still prints normally (regression).
- [ ] 3.4 Pick an old-format folder (<3 underscore segments) — verify TERRESTRE default preserves existing ATA behavior.

## Implementation Order

Phase 1 first (the helper both guards depend on). Then Phase 2 in any order (both guard blocks are independent). Then Phase 3 for verification.

## Next Step

Ready for implementation (sdd-apply). Single PR, no chaining needed. Decision not required before apply (ask-on-risk resolves to No because workload is trivially under budget).
