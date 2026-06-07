# Design: Maritime Email Sending

## Technical Approach

Extend `_correos_core` with maritime branches in two places: the per-folder loop (individual emails) and the post-loop grupal block. Classify folders by matching `_ISO_` or `_FLEXI_` in the folder name string — avoids reading Contenedores.xlsx at email-send time. Terrestrial grupal filters planillas by folder type; maritime grupal runs as a parallel email with its own subject/body.

## Architecture Decisions

| Decision | Options | Choice | Rationale |
|----------|---------|--------|-----------|
| Maritime detection | (a) Read Excel → `_clasificar_tipo_transporte` (b) Folder name regex `_(ISO\|FLEXI)_` | **Folder name regex** | Excel may not exist at send time; folder name always embeds the type |
| Individual branch | (a) Insert maritime `elif` before terrestrial `else` (b) Extract method | **Inline `elif`** | Minimal diff, follows existing per-folder pattern, localized to one function |
| File filter | (a) Single unified filter (b) Separate filters per type | **Separate filters** | Attachments are disjoint — maritime uses Contenedores.xlsx + get\*.pdf, terrestrial uses PLT/MIC/Excel |
| Planilla collection | (a) One list + type tag per entry (b) Two lists | **Two lists** `terr_planillas` / `mar_planillas` | Each grupal block reads its own list; no filtering at send time |
| Terrestrial grupal filter | (a) Filter at send time (b) Filter at collection time | **Filter at collection time** | Simplest — gate the `append` in the folder loop by `_TERRESTRE_` check |
| Suffix regex | (a) Separate regex per type (b) Combined alternation | **Combined `(?:TERRESTRES?\|ISO\|FLEXI)_(.*)`** | Single pattern covers all three transport types at L5377, L5461, L5481 |
| Maritime grupal attachment | (a) Include CARGA TERRESTRE master (b) No master | **No master** | Maritime has no equivalent "master" Excel file |

## Data Flow

```
Folder iteration (per folder):
  │
  ├─ Name pattern check → es_maritimo = bool(re.search(r"_(ISO|FLEXI)_", item))
  │
  ├─ File filter
  │   ├─ es_maritimo: Contenedores.xlsx + get*.pdf (excludes PLT/MIC/Excel)
  │   └─ terrestrial:  PLT*.pdf + MIC*.pdf + numbered Excel (existing)
  │
  ├─ Planilla collection
  │   ├─ es_maritimo → mar_planillas.append(...)
  │   └─ terrestrial  → terr_planillas.append(...)
  │
  ├─ COMPARTIDO pair? → existing COMPARTIDO logic (unchanged)
  ├─ elif es_maritimo + adjuntos → Maritime individual email
  └─ elif terrestrial + adjuntos → Terrestrial individual email (existing)

Post-loop:
  terr_planillas? → CARGA TERRESTRE grupal (existing)
  mar_planillas?  → PLANILLA(S) DE CARGA grupal (new)
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `ui_app.py` | Modify | `_correos_core` — 5 surgical changes |

### Surgery 1 — Planilla collection bucket split (L5401-L5407)

Classify each `PLANILLA DE CARGA*` by folder type before appending:

```
terr_planillas = []   # renamed from todas_las_planillas
mar_planillas = []    # new

if "_TERRESTRE_" in item.upper():
    terr_planillas.append(...)
elif "_ISO_" in item.upper() or "_FLEXI_" in item.upper():
    mar_planillas.append(...)
```

### Surgery 2 — Maritime flag + file filter (insert before L5401)

Set `es_maritimo` in the per-folder loop. When true, collect maritime-relevant files (Contenedores.xlsx + `get*.pdf`) instead of the terrestrial filter. Contains both the filtered file list for individual email AND the planilla collection.

### Surgery 3 — Maritime individual branch (L5480)

Insert a new `elif es_maritimo and adjuntos_validos:` block before the existing `else`:

- **Attachments**: `Contenedores.xlsx` + all `get*.pdf` from the folder
- **Subject**: `SALIDA, PLANILLA COMPLETA DE EXPORTACIÓN_{sufijo}` (1 get\*.pdf) or `SALIDAS, ...` (2+). No MIC reference.
- **Body**: Plain text identical to subject
- **Suffix**: Extracted via combined regex `(?:TERRESTRES?|ISO|FLEXI)_(.*)` for consistency
- **Recipients**: `DESTINATARIOS_INDIVIDUAL` (3 addresses, same as terrestrial)

### Surgery 4 — Regex updates (L5377, L5461, L5481)

Broaden each from `TERRESTRES?` only to `(?:TERRESTRES?|ISO|FLEXI)`:

- **L5377**: COMPARTIDO fallback dest extraction — defensive, enables correct matching if maritime COMPARTIDO folders appear in the future
- **L5461**: COMPARTIDO `_nombre_desde_pe` helper — same combined pattern
- **L5481**: Individual suffix extraction — the primary change, now handles ISO/FLEXI folder names

### Surgery 5 — Maritime grupal block (after L5547)

New post-loop block parallel to the existing CARGA TERRESTRE block:

- **Subject**: `PLANILLA DE CARGA` (1 file) / `PLANILLAS DE CARGA` (2+)
- **Body**: `"Estimados,\n\nSe adjunta(n) la(s) planilla(s) de carga correspondiente(s):\n\n  • {name}\n\nSaludos cordiales."` with singular/plural Spanish forms
- **Attachments**: All `mar_planillas` entries. No master Excel (unlike CARGA TERRESTRE).
- **Recipients**: `DESTINATARIOS_GRUPAL` (14 addresses, same as CARGA TERRESTRE)
- **UI row**: Same `_agregar_fila_correos("Grupal", ...)` pattern

## Variables

| Name | Scope | Purpose |
|------|-------|---------|
| `terr_planillas` | Method | PLANILLA DE CARGA files from TERRESTRE folders (replaces `todas_las_planillas`) |
| `mar_planillas` | Method | PLANILLA DE CARGA files from ISO/FLEXI folders |
| `es_maritimo` | Per-folder | `bool` — `re.search(r"_(ISO\|FLEXI)_", item)` result |

## Testing Strategy

| Layer | What | How |
|-------|------|-----|
| Regex | Suffix extraction | Verify combined pattern captures correct group for ISO, FLEXI, TERRESTRE names |
| Unit | Maritime file filter | Contenedores.xlsx + get\*.pdf selected; PLT\*.pdf, MIC\*.pdf excluded |
| Unit | Subject/body | SALIDA vs SALIDAS based on get\*.pdf count; body == subject; no MIC in subject |
| Integration | Full `_correos_core` | Mock desktop with ISO + TERRESTRE folders; verify correct email creation per type |
| Integration | Grupal split | Mixed folder types produce separate grupal emails with correct attachments |

## Migration / Rollout

No migration required. Pure code change to one method. No config, database, or external dependency changes.

## Open Questions

None.
