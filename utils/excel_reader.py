"""
Lectores de archivos Excel — B/L lookup por carpeta.
Extraído de ui_app.py — Paso 4 de la refactorización modular.
"""


def buscar_bl_por_carpeta_xlsx(wb, nombre_hoja, carpeta_buscada):
    """Busca un B/L en la hoja actual (y la anterior) filtrando por nro. de carpeta."""
    idx_actual = wb.sheetnames.index(nombre_hoja)
    hojas_a_buscar = [nombre_hoja]
    if idx_actual > 0:
        hojas_a_buscar.append(wb.sheetnames[idx_actual - 1])
    for hn in hojas_a_buscar:
        ws = wb[hn]
        filas_carpeta = {}
        filas_bl = {}
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if val and isinstance(val, str):
                    val_up = val.strip().upper()
                    if "CARPETA" in val_up or val_up == "CARP.":
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt:
                            filas_carpeta[r] = str(nxt).strip()
                    elif val_up in ("BL", "B/L", "CONOCIMIENTO"):
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt:
                            filas_bl[r] = str(nxt).strip()
        for r_carp, carp_val in filas_carpeta.items():
            if carp_val == carpeta_buscada:
                mejor_dist = 9999
                mejor_bl = ""
                for r_bl, bl_val in filas_bl.items():
                    dist = abs(r_bl - r_carp)
                    if dist < mejor_dist:
                        mejor_dist = dist
                        mejor_bl = bl_val
                if mejor_bl:
                    return mejor_bl
    return ""


def buscar_bl_por_carpeta_xls(sheet, carpeta_buscada):
    """Busca un B/L en la hoja .xls filtrando por nro. de carpeta."""
    filas_carpeta = {}
    filas_bl = {}
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            if val and isinstance(val, str):
                val_up = val.strip().upper()
                if "CARPETA" in val_up or val_up == "CARP.":
                    nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                    if nxt:
                        filas_carpeta[r] = str(int(nxt)) if isinstance(nxt, float) else str(nxt).strip()
                elif val_up in ("BL", "B/L", "CONOCIMIENTO"):
                    nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                    if nxt:
                        filas_bl[r] = str(nxt).strip()
    for r_carp, carp_val in filas_carpeta.items():
        if carp_val == carpeta_buscada:
            mejor_dist = 9999
            mejor_bl = ""
            for r_bl, bl_val in filas_bl.items():
                dist = abs(r_bl - r_carp)
                if dist < mejor_dist:
                    mejor_dist = dist
                    mejor_bl = bl_val
            if mejor_bl:
                return mejor_bl
    return ""
