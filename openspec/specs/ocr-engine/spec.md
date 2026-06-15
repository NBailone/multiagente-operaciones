# Especificación: Motor OCR (ocr-engine)

## Propósito

Permitir al usuario elegir entre Tesseract y PaddleOCR como motor de reconocimiento óptico de caracteres para el procesamiento de tickets AGD. Ambos motores coexisten en el sistema; la selección persiste en `ui_config.json` y la ausencia de un motor sidecar deshabilita esa opción sin romper la aplicación.

## Requisitos

### R1: Selector de motor OCR
La pantalla de Ajustes DEBE mostrar un menú desplegable para seleccionar el motor OCR entre "Tesseract" (valor por defecto) y "PaddleOCR".

### R2: Persistencia de selección
El motor seleccionado DEBE persistir en `ui_config.json` bajo `config["ocr"]["engine"]`. El cambio DEBE tener efecto en la siguiente llamada a `pdf_a_texto()` sin reiniciar la aplicación.

### R3: Ruta Tesseract
Tesseract DEBE operar desde la carpeta sidecar `./engines/tesseract/` o desde una instalación del sistema. El preprocesamiento DEBE ser el existente (binarización → `image_to_string`). Esta ruta NO DEBE modificarse.

### R4: Ruta PaddleOCR
PaddleOCR DEBE cargarse exclusivamente desde la carpeta sidecar `./engines/paddleocr/`. DEBE recibir imagen en RGB (sin binarizar) y DEBE aplanar su salida estructurada a texto plano (`"\n".join(...)`) para alimentar `extraer_datos()`.

### R5: Inicialización diferida
El singleton de PaddleOCR DEBE inicializarse de forma perezosa (lazy): solo en la primera llamada OCR que use PaddleOCR, nunca al arrancar la aplicación.

### R6: Disponibilidad según sidecar
Si una carpeta sidecar no existe, la opción del motor correspondiente DEBE aparecer deshabilitada u oculta en la interfaz. La aplicación DEBE funcionar normalmente con el motor disponible restante.

### R7: Sin auto-descarga
El sistema NO DEBE descargar, instalar ni actualizar motores automáticamente bajo ninguna circunstancia.

## Escenarios

### Escenario: Selección y persistencia de motor
- DADO que ambos motores están disponibles (sidecars presentes)
- CUANDO el usuario selecciona "PaddleOCR" en el menú desplegable de Ajustes
- ENTONCES `ui_config.json["ocr"]["engine"]` DEBE contener `"paddleocr"`
- Y la siguiente llamada a `pdf_a_texto()` DEBE usar PaddleOCR

### Escenario: Sidecar PaddleOCR faltante
- DADO que la carpeta `./engines/paddleocr/` no existe
- CUANDO el usuario abre la pantalla de Ajustes
- ENTONCES el motor "PaddleOCR" DEBE aparecer deshabilitado u oculto
- Y el motor activo DEBE ser Tesseract

### Escenario: Ambos sidecars presentes, flujo completo
- DADO que ambas carpetas sidecar existen y la configuración indica `"tesseract"`
- CUANDO se ejecuta `pdf_a_texto()` sobre un ticket AGD
- ENTONCES el resultado DEBE contener los mismos campos que `extraer_datos()` espera
- Y el preprocesamiento DEBE ser el de Tesseract (binarizado)

### Escenario: Cambio de motor en caliente
- DADO que el motor activo es Tesseract y `ui_config.json["ocr"]["engine"]` cambia a `"paddleocr"`
- CUANDO se invoca `pdf_a_texto(engine=None)`
- ENTONCES el dispatcher DEBE leer la config y usar PaddleOCR

### Escenario: Sin sidecar Tesseract
- DADO que la carpeta `./engines/tesseract/` no existe y no hay instalación del sistema
- CUANDO el usuario abre la pantalla de Ajustes
- ENTONCES el motor "Tesseract" DEBE aparecer deshabilitado u oculto
- Y si PaddleOCR está disponible, DEBE ser el motor activo

## Fuera de alcance

- Detección automática del mejor motor
- OCR paralelo (ambos motores simultáneamente)
- Extracción de coordenadas (bbox) de PaddleOCR
- Soporte para EasyOCR u otros motores
- Configuración GPU / CUDA
- Empaquetado de modelos PaddlePaddle en el build de PyInstaller
