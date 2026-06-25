# Sistema de Automatización de Operaciones

Aplicación de escritorio para la gestión integral de operaciones logísticas: impresión documental, procesamiento de contenedores, control de datos con OCR, envío de correos y backups.

## Ruta rápida

1. Clonar el repositorio
2. Instalar dependencias: `pip install -r requirements.txt`
3. Ejecutar: `python ui_app.py`

## Módulos

| Módulo | Descripción |
|--------|-------------|
| **Descargar Mails** | Descarga y filtra correos del servidor IMAP |
| **Impresión Documental** | Genera e imprime sobres, permisos de exportación, hojas de ruta y recibos ATA |
| **Completar Planillas** | Completa planillas de carga (Excel) con datos de contenedores |
| **Controlar Datos** | Extrae datos de tickets de pesaje usando OCR (PaddleOCR o API de visión) |
| **Enviar Correos** | Envía correos con archivos adjuntos a destinatarios configurados |
| **Backup** | Respalda carpetas de trabajo a pendrive o ubicación alternativa |
| **Ajustes** | Configura rutas, credenciales, modelos OCR y valores del sistema |

## Configuración

1. Copiar `.env.example` a `.env` y generar una secret key:
   ```bash
   python -c "import secrets; print(secrets.token_hex(32))"
   ```

2. Editar `ui_config.json` con las rutas y credenciales de tu entorno.

3. Los datos sensibles (contraseñas, API keys) se almacenan encriptados con `enc::` prefix.

## OCR

El sistema soporta dos motores de OCR:

- **Local**: PaddleOCR (requiere `engines/paddleocr/`)
- **API Visión**: Modelos de IA vía OpenRouter (Gemini, Gemma, etc.)

Configurar en Ajustes > Método OCR.

## Empaquetar como .exe

```bash
pyinstaller Sistema_Automatizacion.spec
```

El ejecutable se genera en `dist/`. Incluir las carpetas `engines/`, `poppler/` y `python39/` junto al `.exe`.

## Estructura

```
├── ui_app.py              # Interfaz principal (CustomTkinter)
├── procesar_tickets.py    # Motor de OCR para tickets de pesaje
├── constants/             # Paleta de colores, fuentes y configuración
├── utils/                 # Utilidades (email, Excel, pendrive)
├── engines/               # Binarios OCR (PaddleOCR, Tesseract)
├── openspec/              # Artefactos SDD
└── ui_config.json         # Configuración de la aplicación
```

## Stack

- Python 3.9+
- CustomTkinter (UI moderna sobre Tkinter)
- OpenPyXL / xlrd (lectura/escritura Excel)
- PaddleOCR / Tesseract (reconocimiento óptico)
- pywin32 (impresión y COM en Windows)
- python-dotenv (gestión de variables de entorno)
