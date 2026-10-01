# Spotify Reportes Automatizados

Automatización en Python de la generación semanal/mensual de los reportes de
charts y market share de Spotify Latam, a partir de la fuente de datos de BigQuery.

## Objetivo

Reemplazar el proceso manual en Excel por scripts en Python que, a partir de la
fuente de datos (`Fuente_de_datos_BQ_Spotify...xlsx`), generen automáticamente:

- **Reporte_Chart_Top_Semanal_Spotify_Latam** (pestañas `Resumen Total` y `Detalle Tracks`)
- **Reporte_MS_MS_TOP_200_Spotify** (Market Share YTD, pestaña `% Market Share` + una pestaña por país)

Los dos reportes se recalculan con el histórico completo acumulado (`data/history/`),
no solo con la semana que se acaba de cargar — ver `src/history.py`.

## Estructura del repositorio

```
spotify-reportes-automatizados/
├── data/
│   ├── raw/          # Fuente de datos BQ (xlsx/csv) — NO se sube a git
│   ├── output/        # Reportes generados — NO se sube a git
│   └── history/
│       ├── seed/       # Histórico real ya extraído (2019-2026) — sí se sube a git
│       └── universal_data.db   # Base SQLite generada localmente — NO se sube a git
├── scripts/
│   └── sembrar_historico.py   # Siembra el histórico (correr una sola vez)
├── src/
│   ├── __init__.py
│   ├── config.py       # Rutas, nombres de columnas, países, constantes
│   ├── load_data.py    # Carga y limpieza de la fuente BQ
│   ├── history.py       # Persistencia histórica (SQLite)
│   ├── chart_semanal.py    # Genera Reporte_Chart_Top_Semanal
│   ├── market_share.py     # Genera Reporte_MS_MS_TOP_200
│   ├── main.py          # Orquesta la generación de ambos reportes (línea de comandos)
│   └── gui.py            # Ventana simple para generar los reportes sin usar la terminal
├── tests/
├── run_gui.py         # Punto de entrada para empaquetar gui.py con PyInstaller
├── build.bat           # Genera ReportesSpotifyLatam.exe (Windows)
├── requirements.txt
├── .gitignore
└── README.md
```

## Instalación (una sola vez)

```bash
python -m venv .venv
source .venv/bin/activate     # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m scripts.sembrar_historico   # siembra el histórico real (2019-2026)
```

## Uso semanal — línea de comandos

1. Coloca la fuente de datos más reciente en `data/raw/`.
2. Ejecuta:

```bash
python -m src.main --semana 25 --fuente data/raw/Fuente_de_datos_BQ_Spotify.xlsx
```

3. Los reportes se generan en `data/output/`, o en la carpeta que se pase
   con `--salida "D:/Reportes/Semana 38"`. Si la fecha de esa semana ya
   estaba en el histórico (por ejemplo si se corre dos veces por error con
   el mismo archivo), no se duplica: solo se regeneran los reportes.

## Uso semanal — sin terminal (GUI)

Para alguien que no vaya a usar la terminal ni git cada semana:

1. Genera el ejecutable una sola vez (requiere haber hecho la instalación de
   arriba primero): en Windows, haz doble clic en `build.bat` (o corre
   `build.bat` desde una terminal en la raíz del repo). Al terminar, queda
   `ReportesSpotifyLatam.exe` en la raíz del repo.
2. Cada semana: doble clic en `ReportesSpotifyLatam.exe` → elige el archivo
   fuente de la semana → escribe el número de semana → elige la carpeta donde
   dejar los reportes (la primera vez es `data/output`; después queda la
   última que hayas usado) → clic en "Generar reportes". Al terminar, muestra en qué carpeta quedaron los dos reportes.

**Importante:** no muevas `ReportesSpotifyLatam.exe` fuera de esta carpeta —
necesita quedarse junto a `data/` para leer y guardar el histórico. Si lo
mueves, copia también la carpeta `data/` junto a él.

## Uso mensual — reporte "Market Share Spotify Latam"

Al cerrar el mes, un reporte más: el mismo % de market share pero con los
meses como columnas (y una banda extra, el Top 20). **No hace falta cargar
nada nuevo**: el mes se arma con las semanas de BQ que ya se cargaron.

- Sin terminal: en la ventana, botón **"Reporte mensual..."** → elige el mes
  y el año → elige la carpeta → "Generar reporte mensual".
- Con terminal:

```bash
python -m scripts.generar_mensual --anio 2026 --mes 9
python -m scripts.generar_mensual --anio 2026 --mes 9 --salida "D:/Reportes"
```

Si al mes le falta alguna semana, el programa lo dice y genera el archivo
igual, pero **no guarda ese mes en el histórico** hasta que esté completo
(así un mes corto no se arrastra en silencio al trimestre y al año).

Las semanas del histórico sembrado (hasta la 33 de 2026) traen solo el % y no
se pueden sumar. Cuando aparezca la fuente de BQ de una de ellas:

```bash
python -m scripts.completar_semanas "C:/ruta/BQ_semana_32.xlsx"
```

`recargar_semanas.py` NO sirve para eso: salta a propósito las fechas que ya
vienen sembradas, para no numerar la misma semana dos veces.

La historia de mayo de 2017 a julio de 2026 viene sembrada desde la
plantilla; ver `claude/mensual_market_share_analisis.md` para las reglas de
cálculo y cómo se validaron.

## Tests

```bash
pytest tests/ -v
```

## Estado del proyecto

- [x] Mapear 1:1 las fórmulas/lógica de las plantillas actuales a Python
- [x] Histórico acumulado (`history.py`), sembrado con datos reales (2019-2026 Chart, 2025-2026 Market Share)
- [x] Validar reportes generados contra los reportes manuales existentes (323 valores reales comparados, 0 diferencias)
- [x] Ejecución sin terminal (GUI de escritorio, empaquetada con PyInstaller)
- [x] Reporte mensual de Market Share (`src/mensual.py`), validado contra la plantilla (272.504 celdas, 0 diferencias fuera del mes nuevo)
- [ ] Espacio de pruebas para el código de fechas de lanzamiento vía API de Spotify (`experiments/spotify_api/`)
