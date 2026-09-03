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

# Columnas tal cual vienen en la hoja "Consulta1" de la fuente BQ
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
LABEL_GROUP_ALIASES = {
    "Som Livre": "Indies",
    "Altafonte": "Indies",
}


def normalizar_label_group(valor: str) -> str:
    """Aplica el agrupamiento real usado en los reportes (Som Livre/Altafonte -> Indies).

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

# Cuántas canciones como máximo se listan en el "listado de canciones" de
# "Resumen Total" (ver chart_semanal.construir_listado_canciones) -- con
# todas las canciones de una semana (pueden ser 1000+) la hoja queda
# enorme y poco práctica; el usuario pidió dejar solo las mejores 200 (ya
# ordenadas por cantidad de países y suma de posiciones).
TOP_N_LISTADO_CANCIONES = 200

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
