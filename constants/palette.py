"""
Paleta de colores, fuentes y configuración por defecto del sistema.
Extraído de ui_app.py — Paso 1 de la refactorización modular.

Los valores de correo/servidor son placeholders de ejemplo; la
configuración real se carga desde ui_config.json (encriptada).
"""

# ── Configuración de correo por defecto ────────────────────────────────────
IMAP_SERVER = "imap.empresa.com"
PUERTO_IMAP = 143

DESTINATARIOS_GRUPAL = [
    "operaciones@empresa.com", "logistica@empresa.com",
    "despachos@empresa.com", "clientea@cliente-a.com",
    "clientea2@cliente-a.com", "clientea3@cliente-a.com",
    "clienteb@cliente-b.com", "clienteb2@cliente-b.com",
    "clienteb3@cliente-b.com", "clientea4@cliente-a.com",
    "clientec@cliente-c.com", "clienteb4@cliente-b.com",
    "clientea5@cliente-a.com", "clientec2@cliente-c.com",
]
DESTINATARIOS_INDIVIDUAL = [
    "operaciones@empresa.com",
    "despachos@empresa.com",
    "logistica@empresa.com",
]

# ── Paleta de Colores Profesional ───────────────────────────────────────────
class Palette:
    BG_MAIN        = "#102340"  # Deep blue background
    BG_SIDEBAR     = "#081326"  # Darker blue sidebar
    BG_CARD        = "#162f56"  # Slightly lighter blue
    BG_INPUT       = "#0d1c33"
    BG_HOVER       = "#1d3e70"
    BG_TABLE       = "#0a1324"
    BG_TABLE_ALT   = "#102340"
    BG_LOG         = "#060e1a"  # Very dark blue for terminal
    ACCENT         = "#308cff"  # Bright blue
    ACCENT_HOVER   = "#5caeff"
    ACCENT_DIM     = "#1a5bc2"
    SECONDARY      = "#00796B"  # Teal
    SECONDARY_HOVER = "#00695C"
    SUCCESS        = "#00e676"  # Neon green
    SUCCESS_BG     = "#003314"
    WARNING        = "#ff9800"  # Amber
    WARNING_HOVER  = "#e67e22"
    WARNING_BG     = "#4d2e00"
    ERROR          = "#ff4d4d"  # Neon red
    ERROR_BG       = "#4a0f0a"
    INFO           = "#00e5ff"  # Cyan
    INFO_BG        = "#00444d"
    TEXT_PRIMARY   = "#ffffff"  # White
    TEXT_SECONDARY = "#b3c6e0"  # Light blue-gray
    TEXT_MUTED     = "#6784a8"
    BORDER         = "#1e3a68"
    BORDER_ACTIVE  = "#308cff"
    DIVIDER        = "#142d54"
    WHITE          = "#ffffff"

# ── Fuentes ──────────────────────────────────────────────────────────────────
FONT_FAMILY = "Segoe UI"
FONT_MONO  = "Consolas"

# ── Font scale levels for comparison popups ──────────────────────────────────
FONT_LEVEL_SCALES = {1: 1.0, 2: 1.25, 3: 1.5}
FONT_BASE_SIZES = {"data": 11, "header": 12, "legend": 11}
