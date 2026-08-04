"""Genera el Reporte_MS_MS_TOP_200_Spotify (Market Share YTD) a partir de la fuente BQ.

TODO: mapear aquí la lógica exacta de PLANTILLA_MES / MS TOP 200 Spotify YTD:
- Cómo se calcula "% Market Share" (por major_label / label_group, sobre streams totales)
- Qué agregación exacta va en cada pestaña de país (CO, PE, EC, ... )
Revisar la plantilla original para replicar fórmulas y agrupaciones 1:1
antes de dar por buena la salida.
"""
from pathlib import Path
import pandas as pd

from . import config


def calcular_market_share(df: pd.DataFrame, country: str | None = None) -> pd.DataFrame:
    """Placeholder: % de streams por major_label, opcionalmente filtrado por país."""
    data = df if country is None else df[df["country_alt"] == country]

    ms = (
        data.groupby("major_label", as_index=False)["stream_count"]
        .sum()
    )
    ms["market_share_pct"] = ms["stream_count"] / ms["stream_count"].sum()
    return ms.sort_values("market_share_pct", ascending=False)


def generar_reporte(df: pd.DataFrame, output_path: Path) -> Path:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        calcular_market_share(df).to_excel(
            writer, sheet_name=config.MS_SHEET_PORCENTAJE, index=False
        )
        for pais in config.PAISES_MS:
            hoja = calcular_market_share(df, country=pais)
            hoja.to_excel(writer, sheet_name=pais, index=False)

    return output_path
