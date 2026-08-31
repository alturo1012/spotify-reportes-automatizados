"""Genera el Reporte_MS_MS_TOP_200_Spotify (Market Share YTD) a partir del
histórico acumulado en src/history.py.

Fórmula real (verificada contra PLANTILLA_SEMANAL_MS_TOP200.xlsx, filas
99-112 — ver claude/mapeo_logica_plantillas.md sección 1):

    % Market Share YTD (sello, país) =
        suma streams_top200 del sello, semanas 1..N del año
        -------------------------------------------------------
        suma streams_top200 de los 7 sellos, semanas 1..N del año

Se compara el mismo número de semanas (1..N) entre el año en curso y el
año anterior. NO es un promedio de porcentajes semanales.
"""
from pathlib import Path
import pandas as pd

from . import config, history


def calcular_ytd_por_pais(
    country_code: str, anio_actual: int, hasta_semana: int
) -> pd.DataFrame:
    """% Market Share YTD de un país: un DataFrame con una fila por sello
    (en el orden de config.LABEL_GROUPS_MS), comparando anio_actual vs.
    anio_actual - 1, mismas semanas 1..N en ambos.

    N = min(hasta_semana, semanas realmente disponibles de anio_actual - 1).
    Esto evita comparar, por ejemplo, 25 semanas de 2026 contra solo 24 de
    2025 si el histórico del año anterior no llega tan lejos todavía — la
    comparación YTD deja de tener sentido si no son las mismas semanas.
    """
    anio_anterior = anio_actual - 1

    historico_pais_anterior = history.cargar_ms_label_weekly()
    historico_pais_anterior = historico_pais_anterior[
        (historico_pais_anterior["country_code"] == country_code)
        & (historico_pais_anterior["anio"] == anio_anterior)
    ]
    semana_max_anterior = (
        int(historico_pais_anterior["semana"].max())
        if not historico_pais_anterior.empty
        else 0
    )
    hasta_semana_efectiva = min(hasta_semana, semana_max_anterior)

    ytd_actual = history.query_ytd_ms(anio_actual, hasta_semana_efectiva)
    ytd_actual = ytd_actual[ytd_actual["country_code"] == country_code]
    ytd_anterior = history.query_ytd_ms(anio_anterior, hasta_semana_efectiva)
    ytd_anterior = ytd_anterior[ytd_anterior["country_code"] == country_code]

    streams_actual = ytd_actual.groupby("label_group")["streams_top200"].sum()
    streams_anterior = ytd_anterior.groupby("label_group")["streams_top200"].sum()

    total_actual = streams_actual.sum()
    total_anterior = streams_anterior.sum()

    filas = []
    for label in config.LABEL_GROUPS_MS:
        s_actual = float(streams_actual.get(label, 0.0))
        s_anterior = float(streams_anterior.get(label, 0.0))
        pct_actual = s_actual / total_actual if total_actual else 0.0
        pct_anterior = s_anterior / total_anterior if total_anterior else 0.0
        filas.append({
            "label_group": label,
            f"pct_YTD_{anio_actual}": pct_actual,
            f"pct_YTD_{anio_anterior}": pct_anterior,
            "g_l": pct_actual - pct_anterior,
        })
    return pd.DataFrame(filas)


def construir_resumen_pct(anio_actual: int, hasta_semana: int) -> pd.DataFrame:
    """Tabla larga (tidy) con el % Market Share YTD de los 17 países, para
    la pestaña resumen "% Market Share". Una fila por país/sello.
    """
    partes = []
    for country_code in config.PAISES_MS:
        tabla_pais = calcular_ytd_por_pais(country_code, anio_actual, hasta_semana)
        tabla_pais.insert(0, "country_code", country_code)
        partes.append(tabla_pais)
    return pd.concat(partes, ignore_index=True)


def generar_reporte(
    df_semana: pd.DataFrame, output_path: Path, guardar_en_historico: bool = True
) -> Path:
    """Genera el reporte de Market Share a partir del DataFrame de la semana
    nueva (el que devuelve load_data.load_source). Guarda esa semana en el
    histórico (a menos que ya se haya guardado antes, con
    guardar_en_historico=False) y calcula el YTD con todo el histórico
    acumulado hasta esa semana.
    """
    if guardar_en_historico:
        history.append_semana_ms(df_semana)

    fecha = pd.Timestamp(df_semana["chart_date"].unique()[0])
    anio_actual = fecha.year

    todo_el_historico = history.cargar_ms_label_weekly()
    hasta_semana = int(todo_el_historico.loc[todo_el_historico["anio"] == anio_actual, "semana"].max())

    resumen = construir_resumen_pct(anio_actual, hasta_semana)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name=config.MS_SHEET_PORCENTAJE, index=False)
        for country_code in config.PAISES_MS:
            tabla_pais = calcular_ytd_por_pais(country_code, anio_actual, hasta_semana)
            tabla_pais.to_excel(writer, sheet_name=country_code, index=False)

    return output_path

