"""Genera el Reporte_Chart_Top_Semanal_Spotify_Latam a partir del histórico
acumulado en src/history.py (pestaña "Resumen Total") y de la semana nueva
cargada (pestaña "Detalle Tracks").

Ver claude/mapeo_logica_plantillas.md sección 2 para el detalle completo.

ADVERTENCIA (sección 2.2 del mapeo): el conteo "Universal" que usa este
módulo se basa en `label_group == 'Universal'`. En el proceso manual
original, ese conteo salía de contar celdas coloreadas a mano semana a
semana (macro VBA `CountCcolor`), no de una fórmula sobre datos. Es la mejor
aproximación automatizable, pero no hay garantía matemática de que coincida
100% con el criterio manual — validar contra la primera semana real de uso
antes de confiar en el número (Paso 6 del plan).

"Detalle Tracks" NO lleva histórico acumulado a propósito (decisión
confirmada con el usuario): cada reporte muestra el detalle de la semana que
se acaba de cargar, no una serie histórica.
"""
from pathlib import Path
import pandas as pd

from . import config, history


def construir_resumen_total() -> pd.DataFrame:
    """Serie histórica completa (todo lo acumulado en chart_band_weekly):
    una fila por semana, con columnas por país x banda, en el mismo orden
    que la plantilla original (config.PAISES_MS x config.BANDAS_CHART).
    """
    df = history.cargar_chart_band_weekly()
    if df.empty:
        return pd.DataFrame(columns=["anio", "semana", "mes"])

    tabla = df.pivot_table(
        index=["anio", "semana", "mes"],
        columns=["country_code", "banda"],
        values="conteo_universal",
        aggfunc="first",
    )

    columnas_ordenadas = [
        (pais, banda)
        for pais in config.PAISES_MS
        for banda in config.BANDAS_CHART
        if (pais, banda) in tabla.columns
    ]
    tabla = tabla[columnas_ordenadas]
    tabla.columns = [f"{pais}_top{banda}" for pais, banda in tabla.columns]
    tabla = tabla.reset_index().sort_values(["anio", "semana"])
    return tabla


def construir_detalle_tracks(df_semana: pd.DataFrame) -> pd.DataFrame:
    """Detalle track por track (posición 1-200) de la semana que se acaba
    de cargar, por país. Sin histórico — ver advertencia del módulo.
    """
    columnas = [
        "country_code", "chart_date", "position", "artist", "song_name",
        "stream_count", "label_group", "label_name",
    ]
    return df_semana[columnas].sort_values(["country_code", "position"])


def generar_reporte(
    df_semana: pd.DataFrame, output_path: Path, guardar_en_historico: bool = True
) -> Path:
    """Genera el reporte de Chart Semanal a partir del DataFrame de la
    semana nueva (el que devuelve load_data.load_source). Guarda esa semana
    en el histórico (a menos que ya se haya guardado antes, con
    guardar_en_historico=False) y arma "Resumen Total" con TODO el
    histórico acumulado, y "Detalle Tracks" solo con la semana actual.
    """
    if guardar_en_historico:
        history.append_semana_chart(df_semana)

    resumen = construir_resumen_total()
    detalle = construir_detalle_tracks(df_semana)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name=config.CHART_SHEET_RESUMEN, index=False)
        detalle.to_excel(writer, sheet_name=config.CHART_SHEET_DETALLE, index=False)

    return output_path
