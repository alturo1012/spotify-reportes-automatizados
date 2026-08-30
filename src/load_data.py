"""Carga y limpieza de la fuente de datos BQ (export de BigQuery a Excel)."""
from pathlib import Path
import pandas as pd
 
from . import config
 
 
def load_source(path: Path, sheet_name: str = "Consulta1") -> pd.DataFrame:
    """Carga la hoja de datos crudos de la fuente BQ, valida columnas
    esperadas y deja el DataFrame listo para usar en los reportes:
    agrega `country_code` (2 letras) y normaliza `label_group` a las 7
    categorías reales de los reportes (ver config.COUNTRY_CODE_MAP y
    config.normalizar_label_group).
 
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
 
    df["country_code"] = df["country"].map(config.COUNTRY_CODE_MAP)
    paises_sin_mapeo = sorted(df.loc[df["country_code"].isna(), "country"].unique())
    if paises_sin_mapeo:
        raise ValueError(
            "Países en la fuente sin código mapeado en config.COUNTRY_CODE_MAP: "
            f"{paises_sin_mapeo}. Agrégalos al diccionario antes de continuar."
        )
 
    df["label_group"] = df["label_group"].map(config.normalizar_label_group)
 
    return df
 
 
def filtrar_ultima_fecha(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve solo las filas marcadas como is_latest_date (semana vigente)."""
    return df[df["is_latest_date"] == True].copy()  # noqa: E712