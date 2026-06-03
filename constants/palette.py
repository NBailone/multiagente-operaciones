"""
Paleta de colores, fuentes y credenciales por defecto del sistema.
Extraído de ui_app.py — Paso 1 de la refactorización modular.
"""

# ── Credenciales y configuración por defecto ────────────────────────────────
IMAP_SERVER = "imap.empresa.com"
PUERTO_IMAP = 143

DESTINATARIOS_GRUPAL = [
    "usuario@empresa.com", "usuario@empresa.com",
    "usuario@empresa.com", "usuario@empresa.com",
    "usuario@empresa.com", "usuario@empresa.com",
    "usuario@empresa.com", "usuario@empresa.com",
    "usuario@empresa.com", "usuario@empresa.com",
    "usuario@empresa.com", "usuario@empresa.com",
    "usuario@empresa.com", "usuario@empresa.com",
]
DESTINATARIOS_INDIVIDUAL = [
    "usuario@empresa.com",
    "usuario@empresa.com",
    "usuario@empresa.com",
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
