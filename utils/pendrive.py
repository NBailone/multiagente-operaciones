"""
Utilidades de sistema de archivos y pendrive.
Extraído de ui_app.py — Paso 2 de la refactorización modular.
"""
import os
import sys
import re
from datetime import datetime


def buscar_archivo_en_pendrive(nombre_archivo, ruta_base=None):
    """Busca un archivo en pendrives D-Z. Soporta rutas absolutas y relativas.
    Fallback: busca en la carpeta del programa (sys._MEIPASS / __file__)."""
    if ruta_base is None:
        ruta_base = os.path.join("TRABAJO", "01_PLANILLAS")
    ruta_relativa = os.path.join(ruta_base, nombre_archivo)

    # Si la ruta ya es absoluta (ej: F:\TRABAJO\...), verificar directo
    if os.path.isabs(ruta_relativa):
        if os.path.exists(ruta_relativa):
            return ruta_relativa
    else:
        for letra_ascii in range(ord('D'), ord('Z') + 1):
            unidad = f"{chr(letra_ascii)}:\\"
            posible_ruta = os.path.join(unidad, ruta_relativa)
            try:
                if os.path.exists(posible_ruta):
                    return posible_ruta
            except Exception:
                pass

    # Fallback: buscar en la carpeta del programa (raíz del proyecto, donde están los dorsos)
    raiz_app = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(__file__)))
    ruta_app = os.path.join(raiz_app, nombre_archivo)
    if os.path.exists(ruta_app):
        return ruta_app

    return None


def formatear_fecha_excel(d_val):
    """Convierte un valor de fecha de Excel a string DD/MM/YYYY."""
    if not d_val:
        return "No encontrada"
    if isinstance(d_val, datetime):
        return f"{d_val.day}/{d_val.month}/{d_val.year}"
    d_str = str(d_val).strip()
    match = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", d_str)
    if match:
        return f"{int(match.group(1))}/{int(match.group(2))}/{match.group(3)}"
    match_iso = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", d_str)
    if match_iso:
        return f"{int(match_iso.group(3))}/{int(match_iso.group(2))}/{match_iso.group(1)}"
    return d_str
