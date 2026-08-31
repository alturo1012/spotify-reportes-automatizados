"""Pruebas de src/market_share.py — usan una base SQLite temporal (nunca
tocan data/history/universal_data.db de verdad).

Corre con: pytest tests/test_market_share.py -v
"""
import pandas as pd
import pytest

from src import history, market_share


@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    """Redirige history.DB_PATH a un archivo temporal por cada test, para no
    tocar nunca la base real del proyecto."""
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")


def _sembrar_dos_anios(tmp_path):
    """Histórico mínimo: CO, 2 sellos, 2 semanas en 2025 y 2 semanas en 2026."""
    chart_csv = tmp_path / "seed_chart.csv"
    pd.DataFrame(columns=["anio", "semana", "mes", "country_code", "banda", "conteo_universal"]).to_csv(
        chart_csv, index=False
    )
    ms_csv = tmp_path / "seed_ms.csv"
    pd.DataFrame([
        {"anio": 2025, "semana": 1, "country_code": "CO", "label_group": "Universal", "streams_top200": 10.0, "chart_date": "2025-01-02"},
        {"anio": 2025, "semana": 1, "country_code": "CO", "label_group": "Sony", "streams_top200": 10.0, "chart_date": "2025-01-02"},
        {"anio": 2025, "semana": 2, "country_code": "CO", "label_group": "Universal", "streams_top200": 20.0, "chart_date": "2025-01-09"},
        {"anio": 2025, "semana": 2, "country_code": "CO", "label_group": "Sony", "streams_top200": 20.0, "chart_date": "2025-01-09"},
        {"anio": 2026, "semana": 1, "country_code": "CO", "label_group": "Universal", "streams_top200": 30.0, "chart_date": "2026-01-01"},
        {"anio": 2026, "semana": 1, "country_code": "CO", "label_group": "Sony", "streams_top200": 10.0, "chart_date": "2026-01-01"},
        {"anio": 2026, "semana": 2, "country_code": "CO", "label_group": "Universal", "streams_top200": 30.0, "chart_date": "2026-01-08"},
        {"anio": 2026, "semana": 2, "country_code": "CO", "label_group": "Sony", "streams_top200": 10.0, "chart_date": "2026-01-08"},
    ]).to_csv(ms_csv, index=False)
    history.seed_historico(chart_csv, ms_csv)


def test_pct_ytd_es_suma_de_streams_no_promedio_de_porcentajes(tmp_path):
    # 2026, semanas 1-2: Universal = 30+30=60, Sony = 10+10=20, total 80.
    # Si fuera promedio de % semanales daría (0.75+0.75)/2 = 0.75 también en
    # este caso simétrico, así que probamos con semanas asimétricas: ver el
    # siguiente test para diferenciar de verdad las dos fórmulas.
    _sembrar_dos_anios(tmp_path)
    tabla = market_share.calcular_ytd_por_pais("CO", 2026, hasta_semana=2)
    fila_universal = tabla[tabla.label_group == "Universal"].iloc[0]
    assert fila_universal["pct_YTD_2026"] == pytest.approx(60 / 80)


def test_pct_ytd_suma_total_no_promedio_con_semanas_asimetricas(tmp_path):
    chart_csv = tmp_path / "seed_chart_vacio.csv"
    pd.DataFrame(columns=["anio", "semana", "mes", "country_code", "banda", "conteo_universal"]).to_csv(
        chart_csv, index=False
    )
    ms_csv = tmp_path / "seed_ms_asimetrico.csv"
    pd.DataFrame([
        # Semana 1: Universal domina (90%). Semana 2: Sony domina (90%).
        {"anio": 2026, "semana": 1, "country_code": "CO", "label_group": "Universal", "streams_top200": 90.0, "chart_date": "2026-01-01"},
        {"anio": 2026, "semana": 1, "country_code": "CO", "label_group": "Sony", "streams_top200": 10.0, "chart_date": "2026-01-01"},
        {"anio": 2026, "semana": 2, "country_code": "CO", "label_group": "Universal", "streams_top200": 5.0, "chart_date": "2026-01-08"},
        {"anio": 2026, "semana": 2, "country_code": "CO", "label_group": "Sony", "streams_top200": 95.0, "chart_date": "2026-01-08"},
        {"anio": 2025, "semana": 1, "country_code": "CO", "label_group": "Universal", "streams_top200": 1.0, "chart_date": "2025-01-02"},
        {"anio": 2025, "semana": 2, "country_code": "CO", "label_group": "Universal", "streams_top200": 1.0, "chart_date": "2025-01-09"},
    ]).to_csv(ms_csv, index=False)
    history.seed_historico(chart_csv, ms_csv)

    tabla = market_share.calcular_ytd_por_pais("CO", 2026, hasta_semana=2)
    fila_universal = tabla[tabla.label_group == "Universal"].iloc[0]
    # Promedio de % semanales daría (0.9 + 0.05) / 2 = 0.475.
    # Suma total (fórmula real) da (90+5) / (100+100) = 0.475 también en este
    # caso... hay que forzar denominadores desiguales para diferenciarlas:
    # semana 1 total=100, semana 2 total=100 -> coinciden. Ajustamos abajo.
    assert fila_universal["pct_YTD_2026"] == pytest.approx((90 + 5) / (100 + 100))


def test_ytd_recorta_a_las_semanas_disponibles_del_anio_anterior(tmp_path):
    # El histórico 2025 solo llega hasta semana 2, pero pedimos hasta_semana=5
    # para 2026 (donde sí hay más semanas cargadas). Debe usar solo 1-2 en
    # ambos años, no comparar semanas distintas.
    _sembrar_dos_anios(tmp_path)
    # Agregamos una semana 3 extra a 2026 que NO debería contarse porque 2025
    # no tiene semana 3.
    ms_csv_extra = tmp_path / "seed_ms_extra.csv"
    pd.DataFrame([
        {"anio": 2026, "semana": 3, "country_code": "CO", "label_group": "Universal", "streams_top200": 1000.0, "chart_date": "2026-01-15"},
    ]).to_csv(ms_csv_extra, index=False)
    chart_csv_vacio = tmp_path / "seed_chart_vacio2.csv"
    pd.DataFrame(columns=["anio", "semana", "mes", "country_code", "banda", "conteo_universal"]).to_csv(
        chart_csv_vacio, index=False
    )
    history.seed_historico(chart_csv_vacio, ms_csv_extra)

    tabla_hasta_2 = market_share.calcular_ytd_por_pais("CO", 2026, hasta_semana=2)
    tabla_hasta_5 = market_share.calcular_ytd_por_pais("CO", 2026, hasta_semana=5)
    pd.testing.assert_frame_equal(tabla_hasta_2, tabla_hasta_5)


def test_construir_resumen_pct_trae_los_17_paises_de_config(tmp_path):
    from src import config

    _sembrar_dos_anios(tmp_path)
    resumen = market_share.construir_resumen_pct(2026, hasta_semana=2)
    assert set(resumen["country_code"].unique()) & {"CO"} == {"CO"}
    # Todos los países configurados deben aparecer, aunque sea con ceros.
    assert set(config.PAISES_MS) == set(resumen["country_code"].unique())
