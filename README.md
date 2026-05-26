# preliminares-extraction

Herramientas para exportar el listado **En curso** de Peritaciones Diversos desde ePAC y enviarlo por Telegram.

## Scripts

### `exportar_en_curso_epac.py`
Lanza un navegador Playwright, hace login en ePAC, navega a Peritaciones Diversos, extrae todos los registros paginados de la pestaña "En curso" y los guarda en `estadisticas_preliminares.xlsx`.

```bash
python3 exportar_en_curso_epac.py [--headless] [--out RUTA] [--env FICHERO]
```

| Argumento | Descripción |
|-----------|-------------|
| `--headless` | Ejecuta sin ventana gráfica (recomendado en servidor) |
| `--out` | Ruta de salida del Excel. Por defecto: `~/preliminares-extraction/estadisticas_preliminares.xlsx` |
| `--env` | Ruta a un fichero `.env` alternativo |

### `bot_telegram.py`
Bot de Telegram con un botón persistente en el teclado. Al pulsarlo:

1. Desaparece el botón y aparece el mensaje "⏳ Generando Excel, espera un momento..."
2. Se ejecuta `exportar_en_curso_epac.py --headless` en segundo plano
3. Cuando termina, envía el Excel y restaura el botón

Si se pulsa mientras ya hay una generación en curso, el bot avisa y no lanza una segunda ejecución.

```bash
python3 bot_telegram.py
```

La primera vez hay que abrir el chat con el bot y pulsar **Start** (o enviar `/start`) para que aparezca el botón. A partir de ahí queda fijo.

## Instalación

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## Configuración

Copia `.env.example` a `.env` y rellena las credenciales:

```bash
cp .env.example .env
```

Las credenciales de ePAC se obtienen automáticamente de la BD (`softline_aseguradoras_claves_web`, `id_cia IN (42, 399)`). Si la BD no está disponible, se usan las variables `EPAC_URL`, `EPAC_USERNAME` y `EPAC_PASSWORD` del `.env`.

## Dependencias externas

Requiere acceso al repositorio `contacto_whatsapp` en `../contacto_whatsapp` para reutilizar `browser.py`, `config.py` y las páginas de ePAC.
