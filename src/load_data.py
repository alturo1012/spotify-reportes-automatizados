"""Carga y limpieza de la fuente de datos BQ (export de BigQuery a Excel)."""
from pathlib import Path
import pandas as pd
 
from . import config
 
 
def tracks_unicos(df: pd.DataFrame) -> pd.DataFrame:
    """Una fila por TRACK, a partir de la fuente que trae una fila por
    (track, participación de sello).

    POR QUÉ HACE FALTA (comentario de la revisión del 16/09/2026)
    ------------------------------------------------------------
    La fuente BQ parte un mismo track en varias filas cuando su market share
    está compartido: misma posición y mismo ISRC, con los streams ya
    repartidos según la participación. Ejemplos reales de la semana 33:

        CO pos 100  "+57"      Universal 16% + Universal 84%
        PT pos 182  "Maria Joana"  Universal 33% + Warner 67%
        PT pos  84  "Faz Bem"      Sony 50% + Universal 50%

    Contar filas infla el número de tracks (el track se cuenta dos veces).
    Verificado contra el reporte oficial de la semana 33: contando filas hay
    10 diferencias de 85 valores; contando tracks con esta función, 2.

    CRITERIOS
    ---------
    - Los streams SE SUMAN: el total del track es la suma de sus partes.
      (Para el Market Share NO se usa esta función: ahí cada sello se queda
      con su parte, que es justamente lo que la revisión pide.)
    - El track se le asigna al sello con participación ESTRICTAMENTE mayor.
      En empate exacto (50/50, que es el caso más común) no se le asigna a
      ninguno: `label_group` queda en None y el track no le suma a nadie.
      Criterio confirmado con el usuario el 16/09/2026.

    Devuelve las mismas columnas de entrada; `stream_count` sumado y
    `label_group` reemplazado por el sello dueño (o None).
    """
    if df.empty:
        return df.copy()

    claves = ["country_code", "position"]
    streams_por_sello = df.groupby(claves + ["label_group"])["stream_count"].sum()

    duenos = {}
    for clave, serie in streams_por_sello.groupby(level=claves):
        por_sello = serie.droplevel(claves)
        mayor = por_sello.max()
        # Estrictamente mayor: si dos sellos empatan en el máximo, nadie.
        duenos[clave] = None if (por_sello == mayor).sum() > 1 else por_sello.idxmax()

    totales = df.groupby(claves)["stream_count"].sum()

    # De cada track nos quedamos con la fila de mayor participación, que es
    # la que mejor describe al track (label_name, copyright, etc.).
    unicos = df.sort_values("stream_count", ascending=False).drop_duplicates(subset=claves).copy()
    llaves = list(zip(unicos["country_code"], unicos["position"]))
    unicos["stream_count"] = [totales[k] for k in llaves]
    # dtype=object para que los None sobrevivan: con el dtype de texto de
    # pandas 3 un None se convierte de vuelta en NaN, y NaN pasaría los
    # filtros de "sin sello" como si fuera un valor cualquiera.
    unicos["label_group"] = pd.Series(
        [duenos[k] for k in llaves], index=unicos.index, dtype=object
    )
    return unicos.sort_values(claves).reset_index(drop=True)


def elegir_hoja(path: Path) -> str:
    """Devuelve el nombre de la hoja de datos del archivo fuente.
    La exportación de BigQuery no siempre trae el mismo nombre de hoja:
    venía como "Consulta1" y en la semana 36 empezó a llegar como
    "spotify". En vez de casarse con uno, se prueban los nombres conocidos
    (config.HOJAS_FUENTE) sin importar mayúsculas ni espacios de más, y si
    no aparece ninguno se falla con un mensaje que dice qué hojas TIENE el
    archivo -- que es justo lo que hace falta para agregar la nueva.
    """
    hojas = pd.ExcelFile(path).sheet_names
    normalizadas = {str(h).strip().lower(): h for h in hojas}
    for candidata in config.HOJAS_FUENTE:
        hoja = normalizadas.get(candidata.strip().lower())
        if hoja is not None:
            return hoja
    raise ValueError(
        f"No se encontró la hoja de datos en {Path(path).name}. Se buscaron "
        f"{config.HOJAS_FUENTE} y el archivo tiene {hojas}. Si el nombre nuevo "
        "es correcto, agrégalo a config.HOJAS_FUENTE."
    )


def load_source(path: Path, sheet_name: str = None) -> pd.DataFrame:
    """Carga la hoja de datos crudos de la fuente BQ, valida columnas
    esperadas y deja el DataFrame listo para usar en los reportes:
    agrega `country_code` (2 letras) y normaliza `label_group` a las 7
    categorías reales de los reportes (ver config.COUNTRY_CODE_MAP y
    config.normalizar_label_group).
 
    Parameters
    ----------
    path: ruta al xlsx exportado de BigQuery (p. ej. Fuente_de_datos_BQ_Spotify...xlsx)
    sheet_name: nombre de la hoja con los datos. Por defecto (None) se
        detecta sola entre los nombres conocidos -- ver elegir_hoja.
    """
    if sheet_name is None:
        sheet_name = elegir_hoja(path)
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
    return df[df["is_latest_date"] == True].copy()  # noqa: E712z