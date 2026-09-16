"""Configuración central: rutas y constantes del proyecto.

Ver claude/mapeo_logica_plantillas.md (en el Project de Claude) para el
detalle y la evidencia (fórmulas/macro reales) detrás de cada constante de
este archivo.
"""
import sys
from pathlib import Path


def _calcular_root_dir() -> Path:
    """Carpeta raíz del proyecto (donde viven `data/`, `src/`, etc.).

    Separado en una función (en vez de calcularlo directo al importar el
    módulo) para poder probarlo con pytest simulando `sys.frozen`, sin
    depender de si el proceso de test mismo está o no empaquetado.
    """
    if getattr(sys, "frozen", False):
        # Corriendo empaquetado como .exe (PyInstaller, ver build.bat / Paso
        # 7): los datos (histórico, fuente, reportes de salida) tienen que
        # vivir junto al .exe, NO en la carpeta temporal donde PyInstaller
        # descomprime el código cada vez que arranca (esa carpeta se borra
        # al cerrar la app -- si ROOT_DIR apuntara ahí, el histórico se
        # "perdería" cada vez). Por eso: no muevas el .exe fuera de la
        # carpeta del proyecto; si lo mueves, copia también `data/` con él.
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


ROOT_DIR = _calcular_root_dir()
RAW_DIR = ROOT_DIR / "data" / "raw"
OUTPUT_DIR = ROOT_DIR / "data" / "output"

# Base de datos EXTERNA de fechas de lanzamiento, traída del proyecto
# anterior del usuario (la que maneja release_date_management.py). Tiene la
# tabla `track (uri, release_date)`, donde `uri` es el id de track de
# Spotify sin el prefijo "spotify:track:".
#
# Se consulta ANTES de preguntarle a la API (ver
# spotify_release_dates.resolver_fechas_lanzamiento): si la fecha ya está
# ahí, nos ahorramos la llamada. Es OPCIONAL: si el archivo no está, todo
# sigue funcionando igual, solo que resolviendo todo contra la API.
#
# Se abre siempre en modo solo lectura -- es un archivo del usuario, este
# proyecto no le escribe nada.
RELEASE_DATE_DB = ROOT_DIR / "data" / "release_date.db"

# Nombres posibles de la hoja de datos dentro del archivo fuente de
# BigQuery, en orden de preferencia. Venía siempre como "Consulta1", pero
# desde la semana 36 de 2026 empezó a llegar como "spotify" -- y no hay
# garantía de cuál va a venir la próxima vez, así que se aceptan las dos
# (ver load_data.elegir_hoja; la comparación ignora mayúsculas y espacios).
# Si aparece un nombre nuevo, basta con agregarlo a esta lista.
HOJAS_FUENTE = ["Consulta1", "spotify"]

# Columnas tal cual vienen en la hoja de datos de la fuente BQ
SOURCE_COLUMNS = [
    "country",
    "country_alt",
    "chart_date",
    "is_latest_date",
    "artist",
    "song_name",
    "position",
    "stream_count",
    "ISRC",
    "label_group",
    "repertoire",
    "repertoire_group",
    "album_copyright",
    "label_name",
    "content_provider_name",
    "major_label",
    "artist_country",
    "region",
    "main_language",
]

# Países que tienen pestaña propia en el reporte de Market Share
PAISES_MS = [
    "CO", "PE", "EC", "CR", "GT", "PN", "HN", "SV", "NI",
    "DO", "AR", "CL", "BR", "MX", "SP", "PT", "VE",
]

# Orden real de los bloques de país en la pestaña "Resumen Total" del
# Reporte_Chart_Top_Semanal -- verificado 1:1 contra Reporte_Chart_Top
# Semanal Spotify Latam a Sem 24 de 2026.xlsm (fila 5). Es un orden
# distinto al de PAISES_MS (que se usa para Market Share) -- no es un
# error, la plantilla original de Chart Semanal simplemente los ordena
# distinto.
ORDEN_PAISES_CHART = [
    "CO", "PE", "EC", "PN", "CR", "GT", "SV", "HN", "NI", "DO",
    "VE", "MX", "SP", "CL", "AR", "BR", "PT",
]

# Nombre de país tal cual aparece en los encabezados combinados de esa
# misma pestaña (mayúsculas, sin tilde en "PERU"/"MEXICO" -- así está en el
# archivo real) -- verificado 1:1 contra el archivo real.
NOMBRE_PAIS_CHART = {
    "CO": "COLOMBIA", "PE": "PERU", "EC": "ECUADOR", "PN": "PANAMA",
    "CR": "COSTA RICA", "GT": "GUATEMALA", "SV": "SALVADOR", "HN": "HONDURAS",
    "NI": "NICARAGUA", "DO": "DOMINICANA", "VE": "VENEZUELA", "MX": "MEXICO",
    "SP": "ESPAÑA", "CL": "CHILE", "AR": "ARGENTINA", "BR": "BRASIL",
    "PT": "PORTUGAL",
}

# Orden real de los bloques de país en la pestaña resumen "% Market Share"
# del Reporte_MS_TOP200 -- verificado 1:1 contra PLANTILLA_SEMANAL_MS_TOP200.xlsx
# (fila 4, encabezados de bloque). Es un orden DISTINTO al de PAISES_MS (que
# sigue siendo el correcto para el orden de las pestañas por país -- verificado
# también contra los nombres de pestaña reales del mismo archivo) -- no es un
# error, la plantilla usa un orden distinto para esta pestaña resumen en
# particular (ej. Dominicana aparece en la posición 4, no la 10).
ORDEN_PAISES_MS_RESUMEN = [
    "CO", "PE", "EC", "DO", "CR", "GT", "PN", "SV", "HN", "NI",
    "AR", "CL", "BR", "MX", "SP", "PT", "VE",
]

# Nombre de país tal cual aparece en los encabezados de bloque de esa misma
# pestaña resumen -- verificado 1:1 contra el archivo real. OJO: "EL
# SALVADOR" acá (con "EL"), a diferencia de NOMBRE_PAIS_CHART que usa
# "SALVADOR" solo -- cada reporte usa el nombre tal cual viene en su propia
# plantilla real, no se unificaron a propósito.
NOMBRE_PAIS_MS_RESUMEN = {
    "CO": "COLOMBIA", "PE": "PERU", "EC": "ECUADOR", "DO": "DOMINICANA",
    "CR": "COSTA RICA", "GT": "GUATEMALA", "PN": "PANAMA", "SV": "EL SALVADOR",
    "HN": "HONDURAS", "NI": "NICARAGUA", "AR": "ARGENTINA", "CL": "CHILE",
    "BR": "BRASIL", "MX": "MEXICO", "SP": "ESPAÑA", "PT": "PORTUGAL",
    "VE": "VENEZUELA",
}

# Orden de sellos en las tablas de bloque de esa misma pestaña resumen --
# verificado 1:1 contra el archivo real (filas 7-13 de cada bloque). Es
# DISTINTO del orden de LABEL_GROUPS_MS (que sigue siendo el correcto para
# las filas 106-112 de cada pestaña de país individual, verificado también
# contra el archivo real) -- la plantilla real usa dos órdenes de sello
# distintos en dos lugares distintos, no es un error de transcripción.
ORDEN_LABELS_MS_RESUMEN = [
    "Universal", "Sony", "INgrooves", "Orchard", "Warner", "Indies", "Virgin",
]

# Mapeo país (nombre completo, tal cual llega en la columna "country" de la
# fuente BQ) -> código de 2 letras usado en las pestañas de los reportes.
# Confirmado contra los 17 países reales de PLANTILLA_SEMANAL_MS_TOP200.xlsx
# y PLANTILLA_SEMANAL_ChartTop.xlsm (mapeo_logica_plantillas.md, sección 0).
#
# Nota Chile: las plantillas usan "CH" en algunos archivos y "CL" en otros.
# Se deja "CL" (ISO real) como estándar del proyecto — ver "Pendientes" en
# el doc de mapeo si en algún momento se decide lo contrario.
COUNTRY_CODE_MAP = {
    "Colombia": "CO",
    "Peru": "PE",
    "Ecuador": "EC",
    "Costa Rica": "CR",
    "Guatemala": "GT",
    "Panama": "PN",
    "Honduras": "HN",
    "El Salvador": "SV",
    "Nicaragua": "NI",
    "Dominican Republic": "DO",
    "Argentina": "AR",
    "Chile": "CL",
    "Brazil": "BR",
    "Mexico": "MX",
    "Spain": "SP",
    "Portugal": "PT",
    "Venezuela": "VE",
}

# Orden de sellos tal cual aparece en las filas 106-112 de la pestaña de país
# en PLANTILLA_SEMANAL_MS_TOP200.xlsx. Es el orden que deben respetar las
# tablas de % Market Share para que las filas salgan en el mismo orden que
# en la plantilla original.
LABEL_GROUPS_MS = [
    "Universal",
    "INgrooves",
    "Virgin",
    "Sony",
    "Orchard",
    "Warner",
    "Indies",
]

# label_group tal cual viene de la fuente BQ trae 9 valores distintos, pero
# el reporte final agrupa a 7. Som Livre y Altafonte se cuentan como Indies
# (confirmado con el usuario y con el conteo de valores reales de la fuente).
# Som Livre y Altafonte van dentro de SONY, no de Indies.
#
# Estaban como "Indies" y la revisión del 16/09/2026 pidió validarlo. Se
# comprobó contra el reporte oficial de la semana 33, Brasil, banda 200
# (que es donde Som Livre pesa más, un 10,5% de los streams):
#
#     sello       oficial   como Indies   como Sony
#     Sony         0,2755      0,1528       0,2755
#     Indies       0,2241      0,3468       0,2241
#
# Con Som Livre y Altafonte dentro de Sony los dos valores dan EXACTO; como
# Indies se desvían 12 puntos. Confirmado también en España.
LABEL_GROUP_ALIASES = {
    "Som Livre": "Sony",
    "Altafonte": "Sony",
}


def normalizar_label_group(valor: str) -> str:
    """Aplica el agrupamiento real usado en los reportes (Som Livre/Altafonte -> Sony).

    Cualquier otro valor de label_group se devuelve tal cual viene de la fuente.
    """
    return LABEL_GROUP_ALIASES.get(valor, valor)


# Bandas del reporte de Chart Semanal (Resumen Total). OJO: son distintas a
# las del reporte de Market Share — confirmado contra los dos Excel reales.
BANDAS_CHART = [10, 30, 50, 100, 200]

# Bandas usadas en PLANTILLA_SEMANAL_MS_TOP200.xlsx (filas "Tracks TOP N" /
# "Streams TOP N" / "Streams (%) TOP N"). El resumen "% Market Share" (fila
# 99-112, ver calcular_ytd_por_pais) solo usa la de 200; las 5 se usan en la
# cuadrícula semanal de las pestañas individuales por país (ver
# market_share._escribir_pagina_pais).
BANDAS_MARKET_SHARE = [10, 20, 50, 100, 200]

# Nombres de mes en español, tal cual aparecen en la columna "mes" del
# histórico sembrado (chart_band_weekly). No se puede usar
# fecha.strftime("%B") para esto porque depende del locale del sistema
# donde se corra el script (en la mayoría de máquinas Windows/servidor da
# nombres en inglés, ej. "JUNE" en vez de "JUNIO") — bug real encontrado al
# probar el append de una semana nueva contra el histórico sembrado.
MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

# Abreviado en minúscula, para el título "Week Ending - dd mmm, aaaa" del
# listado de canciones (pestaña "Resumen Total", debajo de la serie
# histórica) -- igual al formato de la plantilla original.
MESES_ES_ABREV = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr",
    5: "may", 6: "jun", 7: "jul", 8: "ago",
    9: "sep", 10: "oct", 11: "nov", 12: "dic",
}

# Semáforo de participación de Universal en la serie histórica de "Resumen
# Total" (color de fondo de cada celda conteo_universal, según qué tan cerca
# está del objetivo de participación de Universal en esa banda). Confirmado
# con el usuario con el ejemplo de banda=10 (objetivo = 3 canciones, 30% de
# 10): 1-2 canciones -> rojo, 3 -> amarillo, 4-10 -> verde. El mismo
# porcentaje aplicado a las demás bandas (30/50/100/200) da exactamente
# 9/15/30/60 -- los mismos números que traía la leyenda de la plantilla
# original (fila 1-2) que antes no se había podido explicar; confirma que
# el criterio se generaliza igual a todas las bandas.
PCT_OBJETIVO_UNIVERSAL = 0.30

# Colores estándar de "Reglas de resaltado de celdas" de Excel (rojo/
# amarillo/verde suaves), para que el semáforo se vea como el de cualquier
# reporte de Excel normal.
COLOR_SEMAFORO_ROJO = "FFC7CE"
COLOR_SEMAFORO_AMARILLO = "FFEB9C"
COLOR_SEMAFORO_VERDE = "C6EFCE"

# Color de LETRA que acompaña a cada uno de esos rellenos en las "Reglas de
# resaltado de celdas" de Excel. Verificado 1:1 contra el formato condicional
# real del reporte MS TOP 200 (los dxf de la pestaña "% Market Share").
COLOR_TEXTO_SEMAFORO_ROJO = "9C0006"
COLOR_TEXTO_SEMAFORO_AMARILLO = "9C5700"
COLOR_TEXTO_SEMAFORO_VERDE = "006100"

# Semáforo de la POSICIÓN en el listado de canciones de "Resumen Total"
# (pedido en la reunión del 14/09/2026). Cada posición se escribe en la
# columna de la banda que le corresponde (una posición 25 va en la columna
# "top 30", ver chart_semanal._tier_de_posicion), así que el color se define
# por banda y no por rangos sueltos: así no quedan huecos sin pintar.
#
#   banda 10           -> verde   (el mejor tramo)
#   bandas 30 y 50     -> amarillo
#   bandas 100 y 200   -> rojo
#
# Mismos rellenos y colores de letra del semáforo del Market Share, para que
# los tres reportes se lean igual.
COLOR_POSICION_POR_BANDA = {
    10: (COLOR_SEMAFORO_VERDE, COLOR_TEXTO_SEMAFORO_VERDE),
    30: (COLOR_SEMAFORO_AMARILLO, COLOR_TEXTO_SEMAFORO_AMARILLO),
    50: (COLOR_SEMAFORO_AMARILLO, COLOR_TEXTO_SEMAFORO_AMARILLO),
    100: (COLOR_SEMAFORO_ROJO, COLOR_TEXTO_SEMAFORO_ROJO),
    200: (COLOR_SEMAFORO_ROJO, COLOR_TEXTO_SEMAFORO_ROJO),
}

# Tamaños de la pestaña resumen "% Market Share" y de las pestañas por país
# del reporte de Market Share (pedido en la misma reunión: "que tenga todos
# los bordes y aumentar un poco el tamaño de la letra y el ancho y largo de
# las celdas"). El default de Excel es letra 11, alto de fila ~15 y ancho de
# columna 11: esto sube un punto la letra y un par de puntos el resto -- lo
# justo para que se lea más cómodo sin que la hoja deje de caber en pantalla.
MS_FUENTE_TAMANO = 12
MS_ANCHO_COLUMNA = 13
MS_ALTO_FILA = 19
# Gris del borde de la cuadrícula (la usan los dos reportes: el Market Share
# y, desde la reunión del 14/09/2026, también el Chart Semanal). Más suave
# que el negro puro: marca la retícula sin competir con los colores del
# semáforo.
COLOR_BORDE_CUADRICULA = "808080"

# Tope de canciones del "listado de canciones" de "Resumen Total" (ver
# chart_semanal.construir_listado_canciones). None = sin tope, todas.
#
# Estuvo en 200 mientras el listado traía TODOS los sellos: una semana
# completa son 1000+ canciones y la hoja quedaba impracticable. Desde la
# revisión del 16/09/2026 el listado trae solo productos Universal -- unos
# 226 por semana --, así que el tope dejó de proteger de nada y pasó a
# esconder productos: con 200 el TOP 200 de Colombia mostraba 45 tracks
# cuando la serie histórica decía 46. Sin tope, el detalle cuadra con la
# serie, que es la comprobación que hizo el revisor.
TOP_N_LISTADO_CANCIONES = None

# ---------------------------------------------------------------------------
# Reporte BMAT (Promúsica Colombia)
#
# Es un reporte aparte de los de Spotify, con su propio universo de datos:
# BMAT/Promúsica consolida varias plataformas (ver BMAT_FUENTE), no solo
# Spotify, y usa sus propias bandas y su propia lista de sellos -- 8 en vez
# de 7 (aparece "ADA Music") y con otra grafía ("Ingrooves" en vez de
# "INgrooves", "Independientes" en vez de "Indies"). Por eso NO reutiliza
# LABEL_GROUPS_MS ni BANDAS_MARKET_SHARE: son universos distintos que
# coinciden en parte, y unificarlos escondería esa diferencia.
#
# Todo verificado contra "MS BMAT COL a Sem 23 de 2026.xlsx".
# ---------------------------------------------------------------------------
BMAT_PAIS = "CO"
BMAT_BANDAS = [10, 50, 100, 200, 1000, 3000, 5000, 10000]
BMAT_LABELS = [
    "Universal", "Ingrooves", "Virgin", "Sony",
    "The Orchard", "Warner", "ADA Music", "Independientes",
]
BMAT_SHEET_RESUMEN = "Resumen Promusica Colombia"
BMAT_SHEET_MARKET_SHARE = "Market Share Promusica Colombia"
BMAT_TITULO_RESUMEN = "TOP 1.000 \nPROMUSICA COLOMBIA"
BMAT_TITULO_MARKET_SHARE = "TOP 5.000 PROMUSICA\n COLOMBIA"
BMAT_FUENTE = "FUENTE ( Napster, GooglePlay, Spotify, Deezer)"
# Azul marino de los encabezados de ese archivo (distinto del de los
# reportes de Spotify, que usan 1F3864).
COLOR_BANNER_BMAT = "002060"

CHART_SHEET_RESUMEN = "Resumen Total"
CHART_SHEET_DETALLE = "Detalle Tracks"
MS_SHEET_PORCENTAJE = "% Market Share"

# Colores de los banners (título "TOP 200 WEEKLY MARKET SHARE" y encabezado
# de cada bloque de país) de la pestaña resumen "% Market Share" -- azul
# marino con texto blanco, el mismo estilo visual de PLANTILLA_SEMANAL_MS_TOP200.xlsx
# (esa plantilla usa un color de tema de Excel que no se pudo leer 1:1 vía
# openpyxl -- se usó un azul marino estándar equivalente).
COLOR_BANNER_MS_FONDO = "1F3864"
COLOR_BANNER_MS_TEXTO = "FFFFFF"
