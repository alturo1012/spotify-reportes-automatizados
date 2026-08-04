"""Configuración central: rutas y constantes del proyecto."""
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

CHART_SHEET_RESUMEN = "Resumen Total"
CHART_SHEET_DETALLE = "Detalle Tracks"
MS_SHEET_PORCENTAJE = "% Market Share"
