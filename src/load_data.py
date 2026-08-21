"""Carga y limpieza de la fuente de datos BQ (export de BigQuery a Excel)."""
from pathlib import Path
import pandas as pd


from . import config


def load_source(path: Path, sheet_name: str = "Consulta1") -> pd.DataFrame:
    """Carga la hoja de datos crudos de la fuente BQ y valida columnas esperadas.

    Parameters
    ----------
    path: ruta al xlsx exportado de BigQuery (p. ej. Fuente_de_datos_BQ_Spotify...xlsx)
    sheet_name: nombre de la hoja con los datos (por defecto "Consulta1")
    """
    df = pd.read_excel(path, sheet_name=sheet_name)

    faltantes = set(config.SOURCE_COLUMNS) - set(df.columns)
    if faltantes:
        raise ValueError(f"Faltan columnas esperadas en la fuente: {faltantes}")

    df["chart_date"] = pd.to_datetime(df["chart_date"])
    return df


def filtrar_ultima_fecha(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve solo las filas marcadas como is_latest_date (semana vigente)."""
    return df[df["is_latest_date"] == True].copy()  # noqa: E712
