"""Pruebas de src/chart_semanal.py — usan una base SQLite temporal (nunca
tocan data/history/universal_data.db de verdad).

Corre con: pytest tests/test_chart_semanal.py -v
"""
import pandas as pd
import pytest

from src import chart_semanal, config, history


@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    """Redirige history.DB_PATH a un archivo temporal por cada test, para no
    tocar nunca la base real del proyecto."""
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")


def _csv_vacio(tmp_path, nombre, columnas):
    path = tmp_path / nombre
    pd.DataFrame(columns=columnas).to_csv(path, index=False)
    return path


def test_construir_resumen_total_pivotea_columnas_pais_banda(tmp_path):
    chart_csv = tmp_path / "seed_chart.csv"
    pd.DataFrame([
        {"anio": 2026, "semana": 1, "mes": "ENERO", "country_code": "CO", "banda": 10, "conteo_universal": 1},
        {"anio": 2026, "semana": 1, "mes": "ENERO", "country_code": "CO", "banda": 30, "conteo_universal": 5},
        {"anio": 2026, "semana": 1, "mes": "ENERO", "country_code": "PE", "banda": 10, "conteo_universal": 2},
    ]).to_csv(chart_csv, index=False)
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    history.seed_historico(chart_csv, ms_csv)

    resumen = chart_semanal.construir_resumen_total()

    assert list(resumen.columns[:3]) == ["anio", "semana", "mes"]
    assert "CO_top10" in resumen.columns and "PE_top10" in resumen.columns
    fila = resumen.iloc[0]
    assert fila["CO_top10"] == 1
    assert fila["CO_top30"] == 5
    assert fila["PE_top10"] == 2


def test_construir_resumen_total_vacio_no_falla(tmp_path):
    resumen = chart_semanal.construir_resumen_total()
    assert resumen.empty


def test_append_semana_chart_guarda_mes_en_espanol_no_en_ingles(tmp_path):
    # Regresión del bug real que encontramos: fecha.strftime("%B") depende
    # del locale del sistema y puede devolver "JUNE" en vez de "JUNIO",
    # rompiendo la continuidad con el histórico sembrado (que siempre está
    # en español). Verificado contra Reporte_Chart_Top Semanal Spotify Latam
    # a Sem 24 de 2026.xlsm, donde todas las filas usan mes en español.
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    history.seed_historico(chart_csv, ms_csv)

    df_semana = pd.DataFrame({
        "country_code": ["CO"],
        "label_group": ["Universal"],
        "position": [1],
        "stream_count": [1_000_000],
        "chart_date": pd.to_datetime(["2026-06-18"]),
    })
    history.append_semana_chart(df_semana)

    chart_df = history.cargar_chart_band_weekly()
    assert chart_df["mes"].iloc[0] == "JUNIO"


def test_construir_detalle_tracks_ordena_por_pais_y_posicion(tmp_path):
    df_semana = pd.DataFrame({
        "country_code": ["PE", "CO", "CO"],
        "chart_date": pd.to_datetime(["2026-06-18"] * 3),
        "position": [1, 2, 1],
        "artist": ["a", "b", "c"],
        "song_name": ["x", "y", "z"],
        "stream_count": [100, 200, 300],
        "label_group": ["Universal", "Sony", "Universal"],
        "label_name": ["UMG", "Sony Music", "UMG"],
    })
    detalle = chart_semanal.construir_detalle_tracks(df_semana)
    assert list(detalle["country_code"]) == ["CO", "CO", "PE"]
    assert list(detalle["position"]) == [1, 2, 1]


def test_conteo_universal_coincide_con_datos_reales_de_colombia_semana_24(tmp_path):
    # Validación Paso 6: estos son los 10 tracks REALES del Top 10 de Colombia
    # en la semana del 2026-06-11 (BASE Informe Sportify charts Semana 24),
    # tal cual venían en la fuente cruda (position, label_group). El conteo
    # "Universal" para TOP 10 en el reporte oficial (Reporte_Chart_Top Semanal
    # Spotify Latam a Sem 24 de 2026.xlsm, fila de la semana 24) es 1 -- y la
    # validación completa (17 países x 5 bandas = 85 valores, hecha aparte
    # contra BASE Informe Sportify charts Semana 24 y el reporte oficial)
    # confirmó que este método coincide exactamente en los 85 casos. Este
    # test deja un caso real fijo como regresión rápida.
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    history.seed_historico(chart_csv, ms_csv)

    top10_reales_co = [
        (1, "Warner"), (2, "Warner"), (3, "INgrooves"), (4, "INgrooves"),
        (5, "Warner"), (6, "Warner"), (7, "Warner"), (8, "Sony"),
        (9, "Universal"), (10, "Warner"),
    ]
    df_semana = pd.DataFrame({
        "country_code": ["CO"] * 10,
        "label_group": [lbl for _, lbl in top10_reales_co],
        "position": [pos for pos, _ in top10_reales_co],
        "stream_count": [1_000_000] * 10,
        "chart_date": pd.to_datetime(["2026-06-11"] * 10),
    })
    history.append_semana_chart(df_semana)

    resumen = chart_semanal.construir_resumen_total()
    assert resumen.iloc[0]["CO_top10"] == 1  # valor oficial real verificado


def test_generar_reporte_escribe_las_dos_pestanas(tmp_path):
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    history.seed_historico(chart_csv, ms_csv)

    df_semana = pd.DataFrame({
        "country_code": ["CO"],
        "chart_date": pd.to_datetime(["2026-06-18"]),
        "position": [1],
        "artist": ["a"],
        "song_name": ["x"],
        "stream_count": [1_000_000],
        "label_group": ["Universal"],
        "label_name": ["UMG"],
    })
    salida = tmp_path / "reporte.xlsx"
    chart_semanal.generar_reporte(df_semana, salida)

    assert salida.exists()
    hojas = pd.read_excel(salida, sheet_name=None)
    assert config.CHART_SHEET_RESUMEN in hojas
    assert config.CHART_SHEET_DETALLE in hojas
    assert len(hojas[config.CHART_SHEET_DETALLE]) == 1
