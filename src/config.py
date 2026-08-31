"""Configuración central: rutas y constantes del proyecto.
"""
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
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
# "Streams TOP N"). Hoy solo se usa la de 200 para el % Market Share, pero se
# deja la lista completa por si se necesitan las otras bandas más adelante.
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

CHART_SHEET_RESUMEN = "Resumen Total"
CHART_SHEET_DETALLE = "Detalle Tracks"
MS_SHEET_PORCENTAJE = "% Market Share"
