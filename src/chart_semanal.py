"""Genera el Reporte_Chart_Top_Semanal_Spotify_Latam a partir de la fuente BQ.

TODO: mapear aquí la lógica exacta de PLANTILLA_SEMANAL / Top Semanal Spotify Latam:
- Cómo se arma "Resumen Total" (agregaciones por país/artista/label, etc.)
- Cómo se arma "Detalle Tracks" (ranking de tracks por país)
Revisar la plantilla original para replicar fórmulas y agrupaciones 1:1
antes de dar por buena la salida.
"""
from pathlib import Path
import pandas as pd

from . import config


def construir_resumen_total(df: pd.DataFrame) -> pd.DataFrame:
    """Placeholder: agregación de streams por país y artista."""
    resumen = (
        df.groupby(["country", "artist"], as_index=False)["stream_count"]
        .sum()
        .sort_values(["country", "stream_count"], ascending=[True, False])
    )
    return resumen


def construir_detalle_tracks(df: pd.DataFrame) -> pd.DataFrame:
    """Placeholder: detalle track por track, respetando el 'position' del chart."""
    cols = [
        "country", "chart_date", "position", "artist", "song_name",
        "stream_count", "label_group", "label_name",
    ]
    return df[cols].sort_values(["country", "position"])


def generar_reporte(df: pd.DataFrame, output_path: Path) -> Path:
    resumen = construir_resumen_total(df)
    detalle = construir_detalle_tracks(df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name=config.CHART_SHEET_RESUMEN, index=False)
        detalle.to_excel(writer, sheet_name=config.CHART_SHEET_DETALLE, index=False)

    return output_path
