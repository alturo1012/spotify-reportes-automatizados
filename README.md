# Spotify Reportes Automatizados

Aplicativo en Python que genera los reportes de charts y market share de
Spotify Latam y los reportes BMAT, a partir de las fuentes semanales. Reemplaza
el proceso manual en Excel/VBA.

Versión de entrega: `v1.0` (rama `main`).

## Qué genera

| Proceso | Entrada | Salida |
|---|---|---|
| **Semanal Spotify** | El Excel de BigQuery de la semana (una sola fecha en `chart_date`) | `Reporte_Chart_Top_Semanal_Sem_NN.xlsx` y `Reporte_MS_TOP200_Sem_NN.xlsx` |
| **Mensual Spotify** | Nada nuevo: usa las semanas ya cargadas | `Market Share Spotify Latam a <mes> de <año>.xlsx` |
| **Semanal BMAT** | La carpeta con los `WK<semana>-<mercado>.xlsx` | 4 reportes `MS BMAT ...`, un intermedio `Top N BMAT ...` por mercado y la lista `Revisar clasificacion BMAT ...` |

Todos los reportes se recalculan con el histórico completo acumulado, no solo
con la semana que se acaba de cargar.

- **Chart Top Semanal**: hoja `Resumen Total` (serie histórica desde 2019 del
  conteo de tracks de Universal por banda 10/30/50/100/200, con semáforo sobre
  el objetivo del 30%, y debajo el listado de productos Universal de la semana
  con su fecha de lanzamiento) y hoja `Detalle Tracks` (Top 200 por país).
- **MS Top 200**: hoja `% Market Share` (YTD del año contra el anterior, 17
  países) y una hoja por país con el % semanal por banda 10/20/50/100/200 y
  sello.
- **Mensual**: hoja `Resumen`, 17 hojas de país y 17 hojas `XX-Det` (tracks,
  streams y %), con meses, trimestres, semestre y año desde mayo de 2017.
- **BMAT**: Colombia, Perú, Ecuador y Centroamérica (8 mercados), con bandas
  de 10 a 10.000 y 8 sellos.

## Estructura del repositorio

```
spotify-reportes-automatizados/
├── data/
│   ├── raw/                  # Fuentes de BQ — no se sube a git
│   ├── output/               # Reportes generados — no se sube a git
│   ├── history/
│   │   ├── seed/             # Histórico sembrado (CSV) — sí se sube a git
│   │   └── universal_data.db # Base SQLite local — no se sube a git
│   ├── release_date.db       # Fechas de lanzamiento ya conocidas (solo lectura)
│   ├── bmat_titularidad_compartida.xlsx  # Tracks BMAT con dueño compartido (editable)
│   └── preferencias.json     # Última carpeta de salida usada — se crea sola
├── scripts/
│   ├── sembrar_historico.py  # Crea la base a partir de data/history/seed/
│   ├── recargar_semanas.py   # Reconstruye el histórico cargando las semanas en orden
│   ├── completar_semanas.py  # Rellena tracks y streams de semanas ya sembradas
│   ├── generar_mensual.py    # Reporte mensual por línea de comandos
│   └── generar_bmat.py       # Reportes BMAT por línea de comandos
├── src/
│   ├── config.py             # Rutas, países, sellos, bandas, colores
│   ├── load_data.py          # Lectura de la fuente BQ y regla del dueño del track
│   ├── history.py            # Histórico acumulado (SQLite)
│   ├── chart_semanal.py      # Reporte Chart Top Semanal
│   ├── market_share.py       # Reporte MS Top 200
│   ├── mensual.py            # Cierre del mes
│   ├── mensual_reporte.py    # Excel del reporte mensual
│   ├── bmat_proceso.py       # Proceso semanal BMAT de punta a punta
│   ├── bmat_calculo.py       # Lectura de los WK, bandas y titularidad compartida
│   ├── bmat_clasificacion.py # Clasificación de tracks por ISRC y lista para revisar
│   ├── bmat.py               # Excel de los reportes BMAT
│   ├── spotify_release_dates.py  # Fechas de lanzamiento (caché, base local, API)
│   ├── preferencias.py       # Carpeta de salida recordada
│   ├── main.py               # Proceso semanal por línea de comandos
│   └── gui.py                # Ventana del aplicativo
├── tests/                    # 203 pruebas automáticas
├── run_gui.py                # Punto de entrada del ejecutable
├── build.bat                 # Genera ReportesSpotifyLatam.exe (Windows)
├── .env.example              # Plantilla de credenciales de Spotify
└── requirements.txt
```

## Instalación (una sola vez)

En Windows, desde la carpeta del repositorio:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m scripts.sembrar_historico
copy .env.example .env
```

Después abre `.env` y escribe `SPOTIFY_CLIENT_ID` y `SPOTIFY_CLIENT_SECRET`
(se obtienen en Spotify for Developers). Sin esas credenciales los reportes se
generan igual, pero las fechas de lanzamiento que no estén ya guardadas salen
vacías.

Por último, genera el ejecutable con `build.bat`. Queda
`ReportesSpotifyLatam.exe` en la raíz del repositorio.

**El histórico sembrado no llega hasta hoy.** La siembra trae Spotify hasta la
semana 33 de 2026, BMAT hasta la semana 35 de 2026 y el mensual hasta julio de
2026. En una instalación nueva hay dos caminos:

- copiar `data/history/universal_data.db` desde el equipo donde ya se venía
  usando (lo recomendado), o
- sembrar y volver a cargar las semanas posteriores, en orden, con
  `scripts/recargar_semanas.py` (Spotify) y el botón de BMAT semana por semana.

## Uso con la ventana

Doble clic en `ReportesSpotifyLatam.exe`. Arriba muestra hasta qué semana
llega el histórico.

- **Semanal**: elige el archivo de BigQuery, escribe el número de semana (solo
  nombra los archivos), elige la carpeta de salida y pulsa "Generar reportes".
- **Reporte mensual...**: elige mes y año, y la carpeta.
- **Reportes BMAT...**: elige la carpeta con los archivos WK de la semana y,
  si ya la corregiste, la lista "Revisar clasificacion BMAT".

Cada ventana recuerda la última carpeta de salida. Si el mensaje final trae
avisos, hay que leerlos: dicen qué quedó incompleto y cómo arreglarlo.

No muevas el `.exe` fuera de esta carpeta: necesita estar junto a `data/` y a
`.env`. Si lo mueves, lleva las dos cosas con él.

## Uso por línea de comandos

```powershell
# Semanal Spotify
python -m src.main --fuente "data/raw/BQ Spotify Semana 39.xlsx" --semana 39
python -m src.main --fuente "..." --semana 39 --salida "D:/Reportes/Semana 39"

# Mensual
python -m scripts.generar_mensual --anio 2026 --mes 9

# BMAT
python -m scripts.generar_bmat "C:/ruta/BMAT semana 39"
python -m scripts.generar_bmat "C:/ruta/BMAT semana 39" --revision "Revisar clasificacion BMAT sem 39 de 2026.xlsx"
```

Si la fecha de un archivo semanal ya está en el histórico, no se duplica: solo
se regeneran los reportes.

## Mantenimiento del histórico

El archivo `data/history/universal_data.db` es lo único que no se puede
regenerar por completo. **Hay que respaldarlo.**

| Situación | Qué correr |
|---|---|
| No existe la base, o se borró | `python -m scripts.sembrar_historico` y luego `recargar_semanas` con las fuentes posteriores a la siembra |
| Se cargó una semana fuera de orden o falta una en el medio | `python -m scripts.recargar_semanas "C:/fuentes/*.xlsx"` con todas las fuentes posteriores a la semana 33 de 2026 |
| A un mes le faltan streams de una semana sembrada | `python -m scripts.completar_semanas "C:/ruta/BQ_semana_NN.xlsx"` |
| Se corrigió un valor en un CSV de siembra | Borrar la base y volver a sembrar (la siembra no pisa lo que ya existe) |

`recargar_semanas` vacía y reconstruye el histórico de Spotify. No toca BMAT ni
las fechas de lanzamiento ya resueltas.

## Reglas de negocio

Confirmadas por el área. No se cambian sin preguntar.

- **Track compartido entre dos sellos**: cuenta una sola vez, para el dueño del
  producto. Los streams sí se reparten.
- **Sellos**: Som Livre y Altafonte van dentro de Sony.
- **Orden del listado de canciones**: por cantidad de países en Top 10, luego
  Top 30, Top 50, Top 100 y Top 200, de mayor a menor; desempata la fecha de
  lanzamiento, de la más antigua a la más reciente.
- **Objetivo de Universal**: 30% de cada banda (3, 9, 15, 30 y 60 tracks).
- **Mes de una semana**: el de su fecha de corte. El mes cierra el último
  jueves.
- **Cierre mensual**: los streams se suman, los tracks se promedian y el % se
  recalcula sobre el total del mes.
- **Año, en las filas de tracks del mensual**: `(H1 + Q3 + Q4) / 3`, igual que
  la plantilla original.
- **Fila Virgin**: se corrigió el error de la plantilla; su valor es el
  complemento de los otros seis sellos.
- **BMAT, titularidad compartida**: los streams de los tracks de
  `data/bmat_titularidad_compartida.xlsx` se reparten entre sus dos sellos.

## Pruebas

```powershell
pytest -q
```

Deben pasar las 203. Las de la ventana necesitan `tkinter` (viene con el
Python de Windows).

Después de cualquier cambio de código hay que volver a correr `build.bat`: el
ejecutable no se actualiza solo.

## Validación

- Chart Semanal: 32.990 valores de la serie histórica contra el reporte
  oficial de la semana 33 de 2026, sin diferencias.
- Market Share semanal: 171.850 valores de las hojas de país contra el reporte
  oficial, sin diferencias (aparte de la fila Virgin corregida).
- Conteo de tracks: 1.785 de 1.785 valores idénticos contra la plantilla YTD
  (3 semanas, 17 países, 5 bandas, 7 sellos).
- Streams: 1.760 de 1.785 idénticos; las 25 diferencias están en Portugal,
  España y Perú, en tracks con dueño ambiguo entre Sony y The Orchard.
- Mensual de agosto de 2026, calculado desde sus cuatro fuentes de BQ contra
  el reporte oficial: 528 de 595 celdas idénticas; las demás se explican por
  la fila Virgin corregida y por tracks con dueño ambiguo en Portugal y
  España.
- BMAT semana 35 de 2026, 11 mercados: tracks sin diferencias; las de streams
  corresponden a errores de las fórmulas manuales.

## Limitaciones conocidas

- El YTD de 2026 usa el promedio del % semanal porque faltan los streams
  crudos de las semanas 25 a 33. La diferencia medida contra la fórmula exacta
  es de 0,1 puntos en promedio. Desde 2027 vuelve solo a la fórmula exacta.
- Tracks con dueño ambiguo (dos sellos con aviso de copyright equivalente)
  pueden diferir levemente en streams, sobre todo en Portugal y España.
- En BMAT, "China (feat. J Balvin, Ozuna)" está clasificado como Sony pero la
  lista de titularidad lo reparte entre The Orchard y Universal. El programa
  usa la lista y muestra un aviso.
- BMAT cubre Colombia, Perú, Ecuador y Centroamérica. Los WK de Bolivia,
  Chile, Brasil, Argentina, España y México se ignoran.
- Los reportes no se suben solos a ningún lado.
