"""
Utilidades de email (adjuntar archivos a mensajes MIME).
Extraído de ui_app.py — Paso 2 de la refactorización modular.
"""
import os
import mimetypes
from email.mime.base import MIMEBase
from email import encoders


def adjuntar_archivo(msg, ruta):
    """Adjunta un archivo a un mensaje MIME (msg)."""
    nombre_archivo = os.path.basename(ruta)
    ctype, encoding = mimetypes.guess_type(ruta)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    try:
        with open(ruta, "rb", buffering=16384) as f:
            part = MIMEBase(maintype, subtype)
            part.set_payload(f.read())
            encoders.encode_base64(part)
        safe_name = nombre_archivo.replace('"', '\\"')
        part["Content-Disposition"] = f'attachment; filename="{safe_name}"'
        part.set_param("name", nombre_archivo, header="Content-Type")
        msg.attach(part)
    except Exception as e:
        raise RuntimeError(f"Error adjuntando {nombre_archivo}: {e}")
