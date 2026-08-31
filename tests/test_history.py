"""Pruebas de src/history.py — usan una base SQLite temporal (nunca tocan
data/history/universal_data.db de verdad).

Corre con: pytest tests/test_history.py -v
"""
import pandas as pd
import pytest

from src import history


@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    """Redirige history.DB_PATH a un archivo temporal por cada test, para no
    tocar nunca la base real del proyecto."""
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")


def _csv_vacio(tmp_path, nombre, columnas):
    path = tmp_path / nombre
    pd.DataFrame(columns=columnas).to_csv(path, index=False)
    return path


def test_seed_historico_es_idempotente(tmp_path):
    chart_csv = tmp_path / "seed_chart.csv"
    ms_csv = tmp_path / "seed_ms.csv"
    pd.DataFrame([
        {"anio": 2025, "semana": 1, "mes": "ENERO", "country_code": "CO", "banda": 10, "conteo_universal": 3},
    ]).to_csv(chart_csv, index=False)
    pd.DataFrame([
        {"anio": 2025, "semana": 1, "country_code": "CO", "label_group": "Universal",
         "streams_top200": 18.86, "chart_date": "2025-01-02"},
    ]).to_csv(ms_csv, index=False)

    history.seed_historico(chart_csv, ms_csv)
    history.seed_historico(chart_csv, ms_csv)  # correr dos veces no debe duplicar

    assert len(history.cargar_chart_band_weekly()) == 1
    assert len(history.cargar_ms_label_weekly()) == 1


def test_append_semana_continua_la_numeracion_de_semana(tmp_path):
    chart_csv = tmp_path / "seed_chart.csv"
    ms_csv = tmp_path / "seed_ms.csv"
    pd.DataFrame([
        {"anio": 2026, "semana": 1, "mes": "ENERO", "country_code": "CO", "banda": 200, "conteo_universal": 50},
    ]).to_csv(chart_csv, index=False)
    pd.DataFrame([
        {"anio": 2026, "semana": 1, "country_code": "CO", "label_group": "Universal",
         "streams_top200": 10.0, "chart_date": "2026-01-01"},
    ]).to_csv(ms_csv, index=False)
    history.seed_historico(chart_csv, ms_csv)

    df_semana = pd.DataFrame({
        "country_code": ["CO", "CO"],
        "label_group": ["Universal", "Sony"],
        "position": [1, 2],
        "stream_count": [2_000_000, 1_000_000],
        "chart_date": pd.to_datetime(["2026-01-08", "2026-01-08"]),
    })
    history.append_semana_chart(df_semana)
    history.append_semana_ms(df_semana)

    chart_df = history.cargar_chart_band_weekly()
    ms_df = history.cargar_ms_label_weekly()

    # La semana sembrada era la 1; la nueva debe quedar como la 2, no repetirla.
    assert chart_df[chart_df.country_code == "CO"]["semana"].max() == 2
    assert ms_df[ms_df.country_code == "CO"]["semana"].max() == 2


def test_append_semana_ms_escala_streams_a_millones(tmp_path):
    # Regresión del bug real que encontramos: el histórico sembrado guarda
    # streams en millones, y append_semana_ms tiene que aplicar el mismo
    # factor para que sea comparable (verificado 1:1 contra un archivo real:
    # 12,012,244 streams crudos == 12.012244 en el histórico sembrado).
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    history.seed_historico(chart_csv, ms_csv)

    df_semana = pd.DataFrame({
        "country_code": ["CO"],
        "label_group": ["Universal"],
        "position": [1],
        "stream_count": [12_012_244],
        "chart_date": pd.to_datetime(["2026-06-18"]),
    })
    history.append_semana_ms(df_semana)

    ms_df = history.cargar_ms_label_weekly()
    valor = ms_df[(ms_df.country_code == "CO") & (ms_df.label_group == "Universal")]["streams_top200"].iloc[0]
    assert valor == pytest.approx(12.012244)
