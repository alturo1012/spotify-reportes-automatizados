# Spotify Reportes Automatizados

Aplicativo en Python que genera los reportes de charts y market share de
Spotify Latam y los reportes BMAT, a partir de las fuentes semanales. Reemplaza
el proceso manual en Excel/VBA.

Versión de entrega: `v1.2` (rama `main`). La versión se ve en el título de la
ventana y abajo a la izquierda; es lo primero que hay que preguntar cuando
alguien reporta un problema.

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
│   │   ├── universal_data.db # Base SQLite local — no se sube a git
│   │   └── respaldos/        # Copias automáticas de la base — no se sube a git
│   ├── release_date.db       # Fechas de lanzamiento ya conocidas (solo lectura)
│   ├── bmat_titularidad_compartida.xlsx  # Tracks BMAT con dueño compartido (editable)
│   ├── logs/                 # Registro de errores (aplicativo.log) — no se sube a git
│   └── preferencias.json     # Última carpeta de salida usada — se crea sola
├── scripts/
│   ├── sembrar_historico.py  # Crea la base a partir de data/history/seed/
│   ├── recargar_semanas.py   # Reconstruye el histórico cargando las semanas en orden
│   ├── completar_semanas.py  # Rellena tracks y streams de semanas ya sembradas
│   ├── empaquetar_instalacion.py  # Zip para instalar en otro equipo con la base al día
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
│   ├── respaldo.py           # Respaldo automático de la base
│   ├── registro.py           # Registro de errores en data/logs/
│   ├── paquete.py            # Paquete de instalación
│   ├── version.py            # Número de versión del aplicativo
│   ├── main.py               # Proceso semanal por línea de comandos
│   └── gui.py                # Ventana del aplicativo
├── tests/                    # 237 pruebas automáticas
├── run_gui.py                # Punto de entrada del ejecutable
├── build.bat                 # Genera ReportesSpotifyLatam.exe (Windows)
├── .env.example              # Plantilla de credenciales de Spotify
├── requirements.txt          # Dependencias directas, versiones exactas
└── requirements-lock.txt     # Entorno completo validado (lo usa build.bat)
```

## Instalación (una sola vez)

En Windows, con **Python 3.11** (el aplicativo se validó con 3.11.9), desde la
carpeta del repositorio:

```powershell
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install -r requirements-lock.txt
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
2026. Por eso, para instalar en otro equipo lo normal es usar el paquete de
instalación (abajo), que lleva la base al día.

## Instalar en otro equipo: paquete de instalación

En el equipo donde ya se viene usando, después de correr `build.bat`:

```powershell
python -m scripts.empaquetar_instalacion
python -m scripts.empaquetar_instalacion --salida "D:/Entregas"
```

Deja en `data/output/` un `ReportesSpotifyLatam_v1.2_instalacion_<fecha>.zip`
(unos 19 MB) con el ejecutable, una copia verificada de la base actual, la
siembra, `release_date.db`, la lista de titularidad compartida de BMAT,
`.env.example` y un `LEEME_INSTALACION.txt` con los pasos y hasta qué semana
llega el histórico. En el otro equipo basta con descomprimirlo, crear el `.env`
y abrir el ejecutable; no hace falta Python.

No lleva los respaldos, el registro de errores, las preferencias de este equipo
ni el `.env`. Para incluir las credenciales de Spotify de este equipo se agrega
`--con-credenciales`; en ese caso, no compartir el zip por un canal público.

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

### BMAT: la lista de tracks para revisar

En el proceso manual, a cada track se le ponía a mano la columna "Disqueras"
(la major dueña o, si es independiente, su distribuidora). El aplicativo lo
hace solo: los tracks ya conocidos toman su clasificación por ISRC, y los
nuevos se clasifican con la clasificación más común de su distribuidora, que
acierta alrededor del 97%.

`Revisar clasificacion BMAT sem NN de AAAA.xlsx` trae solo esos tracks nuevos.
Los reportes de la semana ya salen completos aunque no se revise.

1. Empezar por los de prioridad **Alta** (en rojo): están en el Top 200 de
   algún mercado o su regla tiene menos de 90% de confianza.
2. Comparar "Disqueras asignada" y "Sello asignado" contra "Disquera
   original" y "Distribuidora original". Si está bien, no hacer nada.
3. Si está mal, escribir la clasificación correcta en la columna amarilla
   "Disqueras corregida": una major de la lista desplegable o, si es
   independiente, el nombre de su distribuidora.
4. Guardar y volver a generar la misma carpeta de WK, eligiendo la lista
   corregida en el campo opcional.

Las correcciones quedan guardadas: el track sale bien en las semanas
siguientes y no vuelve a la lista. Si no se revisa, queda guardada la
clasificación de la regla; las semanas ya generadas solo cambian si se
vuelven a generar. Como mínimo, revisar los de prioridad Alta cada semana.

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

### Las semanas se cargan en orden

El número de semana se asigna por orden de carga, así que el aplicativo se
detiene **antes de guardar** si una carga desordenaría el histórico:

- **La fuente es anterior a la última semana cargada:** se bloquea siempre y no
  se genera nada. La única forma de meter esa semana en su lugar es
  `recargar_semanas` (ver abajo).
- **Faltan semanas entre la última cargada y esta:** se bloquea y dice qué
  fechas faltan. Lo normal es cargar primero las que faltan. Si una semana de
  verdad no existe en BigQuery, la ventana pregunta si cargar igual; por
  terminal se agrega `--permitir-hueco`.

### Revisión de la fuente

Al cargar el archivo de BigQuery, el aplicativo revisa que esté completo y
avisa (sin detener nada) si a algún país le faltan datos, si un país trae
menos de las 200 posiciones o si hay filas sin streams o en cero. Un track
compartido entre dos sellos viene en dos filas con la misma posición: eso es
normal y no genera aviso.

### Año de BMAT

Los archivos `WK` no traen el año. El aplicativo lo deduce de la última semana
BMAT guardada: una semana mayor que la última es del mismo año, y una mucho
menor (por ejemplo `WK01` después de la `WK52`) es del año siguiente. Si el año
que asigna no es el actual, lo avisa. Para cargar una semana de otro año a
propósito se usa `--anio`.

## Versiones fijas

Las librerías están fijadas a la versión exacta con la que pasan las pruebas:

- `requirements.txt`: lo que el código usa directamente (pandas, openpyxl,
  spotipy, python-dotenv, python-dateutil), con `==`.
- `requirements-lock.txt`: el entorno completo, incluidas las dependencias de
  esas librerías, pytest y PyInstaller. Es lo que se instala y lo que usa
  `build.bat`, así que cualquier equipo arma el mismo ejecutable.

`build.bat` avisa si el `.venv` no es de Python 3.11.

**Para actualizar una librería a propósito:**

1. Cambiar su versión en `requirements.txt` e instalarla
   (`pip install pandas==X.Y.Z`).
2. Correr `pytest -q`. Si algo falla, volver a la versión anterior.
3. Si todo pasa, regenerar el lock con `pip freeze > requirements-lock.txt`,
   revisar que no se hayan colado librerías ajenas y volver a poner el
   encabezado de comentarios.
4. `build.bat` y subir los dos archivos juntos.

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

### Respaldos automáticos

Antes de cada proceso que escribe en la base (semanal con una fecha nueva,
mensual, BMAT, `recargar_semanas` y `completar_semanas`), el aplicativo guarda
una copia comprimida en `data/history/respaldos/`, por ejemplo
`universal_data_2026-10-09_153012_semanal.zip`. El final del nombre dice qué
proceso vino después de la copia. Se conservan las últimas 15 (unos 12 MB
cada una); las más viejas se borran solas.

Si un respaldo falla, el proceso sigue y la falla queda en el registro. La
excepción es `recargar_semanas`: si no puede respaldar, no vacía nada.

**Para restaurar:** cerrar el aplicativo, descomprimir el `.zip` elegido y
poner el `universal_data.db` que trae en `data/history/`, reemplazando el
actual. Conviene guardar aparte la base actual antes de reemplazarla.

Los respaldos protegen contra un error de carga, no contra perder el equipo:
hay que seguir copiando `data/history/` a otro lugar (Drive, OneDrive o un
disco externo) de vez en cuando.

### Registro de errores

Todo lo que pasa queda anotado en `data/logs/aplicativo.log`: cada proceso con
sus archivos de entrada y salida, los avisos mostrados, los respaldos y los
errores con su detalle técnico completo. Cuando algo falla, la ventana muestra
la ruta de ese archivo; es lo que hay que enviar para pedir ayuda. El archivo
rota solo y nunca pasa de unos 6 MB.

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

Deben pasar las 237. Las de la ventana necesitan `tkinter` (viene con el
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
