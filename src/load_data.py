"""Carga y limpieza de la fuente de datos BQ (export de BigQuery a Excel)."""
import re
from pathlib import Path
import pandas as pd
 
from . import config
 
 
# Marcas de un aviso de copyright de verdad: el símbolo © o ℗ (o sus
# versiones en texto) o un año. Sirven para distinguirlo de una etiqueta
# genérica de sello -- ver dueno_del_track.
_MARCAS_COPYRIGHT = re.compile(r"©|℗|\(C\)|\(P\)|\b(?:19|20)\d{2}\b", re.IGNORECASE)


def dueno_del_track(filas: pd.DataFrame):
    """El sello DUEÑO de un track, entre las filas (una por participación)
    con que la fuente lo describe. None si no se puede determinar.

    LA REGLA, Y DE DÓNDE SALE
    -------------------------
    Tatiana y Alejandro la definieron así: "el track se contabiliza para la
    disquera propietaria del producto" -- no para la de mayor participación.
    En un 50/50 la participación no dice nada, pero el dueño sí está en la
    fuente: cuando un track viene compartido, UNA de las filas trae el aviso
    de copyright REAL y la otra una etiqueta genérica del sello. Ejemplos
    reales de Portugal, semana 33:

        Faz Bem     [Sony]      "Sony Music"                 <- genérica
                    [Universal] "© 2024 Universal Music Portugal, S.A. ..."  <- dueño
        PENSAR EM TI[Universal] "UMLA"                       <- genérica
                    [Virgin]    "© 2026 Mário Cotrim, distributed by ... Virgin Music Portugal"
        DÁKITI      [Universal] "Universal Music"            <- genérica
                    [Orchard]   "(C) 2020 Rimas Entertainment LLC ..."

    Así que el dueño es el sello de la fila cuyo `album_copyright` es un
    aviso de verdad. Verificado contra los 5 casos que ellos mismos
    clasificaron a mano: coincide en los 5, y deja el conteo del Chart de la
    semana 33 idéntico al informe oficial (0 diferencias en 85 valores).

    Si ninguna fila trae aviso real, o si hay empate entre varias, se
    devuelve None y el track no le suma a ningún sello -- mejor no contarlo
    que contárselo a quien no es.
    """
    sellos_presentes = set(filas["label_group"].dropna())
    if len(sellos_presentes) == 1:
        # Un solo sello en todas las filas (el track viene partido, pero
        # dentro de la misma disquera): no hay nada que decidir.
        return sellos_presentes.pop()
    if not sellos_presentes:
        return None

    if "album_copyright" not in filas.columns:
        return None

    reales = [
        (sello, texto)
        for sello, texto in zip(filas["label_group"], filas["album_copyright"])
        if isinstance(texto, str) and _MARCAS_COPYRIGHT.search(texto)
    ]
    if not reales:
        return None
    if len({sello for sello, _ in reales}) == 1:
        return reales[0][0]
    # Varios sellos con aviso real: gana el más específico (el más largo), y
    # si empatan en largo no se decide.
    reales.sort(key=lambda par: len(par[1]), reverse=True)
    if len(reales[0][1]) == len(reales[1][1]):
        return None
    return reales[0][0]


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
    - El track se le asigna al SELLO DUEÑO DEL PRODUCTO, no al de mayor
      participación -- ver `dueno_del_track`. Es la regla que confirmaron
      Tatiana y Alejandro: "el track se contabiliza para la disquera
      propietaria del producto; los streams sí se reparten".

    Devuelve las mismas columnas de entrada; `stream_count` sumado y
    `label_group` reemplazado por el sello dueño (o None si no se puede
    determinar).
    """
    if df.empty:
        return df.copy()

    claves = ["country_code", "position"]

    duenos = {}
    for clave, grupo in df.groupby(claves):
        duenos[clave] = dueno_del_track(grupo)

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
    return df[df["is_latest_date"] == True].copy()  # noqa: E712