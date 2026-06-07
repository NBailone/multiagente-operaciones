# Tasks: Maritime Email Sending

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~100–135 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-always |
| Chain strategy | pending |

```
Decision needed before apply: Yes
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low
```

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Full maritime email implementation | PR 1 | Single PR — all changes localized to `ui_app.py:_correos_core` |

## Phase 1: Regex Updates (L5377, L5461, L5481)

- [x] 1.1 Broaden COMPARTIDO fallback dest regex (L5377): replace `TERRESTRE_` with `(?:TERRESTRES?|ISO|FLEXI)_` pattern
- [x] 1.2 Broaden COMPARTIDO `_nombre_desde_pe` regex (L5461): same `(?:TERRESTRES?|ISO|FLEXI)_` alternation
- [x] 1.3 Broaden suffix extraction regex (L5481): from `TERRESTRES?_(.*)` to `(?:TERRESTRES?|ISO|FLEXI)_(.*)`

## Phase 2: Planilla List Split + Maritime Detection

- [x] 2.1 Rename `todas_las_planillas` → `terr_planillas` and add `mar_planillas = []` at initialization
- [x] 2.2 Gate planilla append by folder type: `_TERRESTRE_` → `terr_planillas`; `_ISO_`/`_FLEXI_` → `mar_planillas`
- [x] 2.3 Add `es_maritimo = bool(re.search(r"_(ISO|FLEXI)_", item))` flag in per-folder loop
- [x] 2.4 Add maritime file filter: when `es_maritimo`, collect `Contenedores.xlsx` + `get*.pdf` only (skip PLT\*.pdf, MIC\*.pdf, numbered \*.xlsx)

## Phase 3: Maritime Individual Email Branch

- [x] 3.1 Insert `elif es_maritimo and adjuntos_validos:` block before the terrestrial `else` (L5480)
- [x] 3.2 Count `get*.pdf` → singular `SALIDA` (1) or plural `SALIDAS` (2+)
- [x] 3.3 Extract suffix via combined regex; build subject/body as `SALIDA(S), PLANILLA COMPLETA DE EXPORTACIÓN_{sufijo}`
- [x] 3.4 Attach `Contenedores.xlsx` + all `get*.pdf`; send to `destinatarios_individual` (3 recipients)

## Phase 4: Maritime Grupal Email Block

- [x] 4.1 Add post-loop block after existing CARGA TERRESTRE: build subject `PLANILLA DE CARGA` (1 file) / `PLANILLAS DE CARGA` (2+)
- [x] 4.2 Build plain-text body with Spanish singular/plural forms (`adjunta la planilla` / `adjuntan las planillas`), listing each attached planilla filename
- [x] 4.3 Attach all `mar_planillas` entries (no master Excel); send to `destinatarios_grupal` (14 recipients)
- [x] 4.4 Add UI row via `_agregar_fila_correos("Grupal", ...)` for maritime grupal email
