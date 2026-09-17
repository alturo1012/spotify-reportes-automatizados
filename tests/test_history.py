"""Pruebas de src/history.py — usan una base SQLite temporal (nunca tocan
data/history/universal_data.db de verdad).

Corre con: pytest tests/test_history.py -v
"""
import pandas as pd
import pytest

from src import config, history


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


def test_semana_ya_cargada_detecta_fecha_existente_y_ausente(tmp_path):
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = tmp_path / "seed_ms.csv"
    pd.DataFrame([
        {"anio": 2026, "semana": 1, "country_code": "CO", "label_group": "Universal",
         "streams_top200": 10.0, "chart_date": "2026-01-01"},
    ]).to_csv(ms_csv, index=False)
    history.seed_historico(chart_csv, ms_csv)

    assert history.semana_ya_cargada("2026-01-01") is True
    assert history.semana_ya_cargada("2026-01-08") is False


def test_append_semana_tracks_guarda_una_fila_por_track_y_numera_la_semana(tmp_path):
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    history.seed_historico(chart_csv, ms_csv)

    df_semana = pd.DataFrame({
        "country_code": ["CO", "PE"],
        "chart_date": pd.to_datetime(["2026-06-18"] * 2),
        "position": [1, 5],
        "artist": ["a", "b"],
        "song_name": ["x", "y"],
        "stream_count": [1_000_000, 500_000],
        "label_group": ["Universal", "Sony"],
        "label_name": ["UMG", "Sony Music"],
        "region": ["Latin", "Anglo"],
    })
    # Igual que append_semana_chart/append_semana_ms: cada llamada numera
    # "la siguiente" semana, no es idempotente por sí sola -- lo que evita
    # duplicar una semana es el chequeo previo de history.semana_ya_cargada
    # en main.py (ver test_main_corrido_dos_veces_con_la_misma_fuente...).
    history.append_semana_tracks(df_semana)

    tracks_df = history.cargar_chart_track_weekly()
    assert len(tracks_df) == 2
    assert set(tracks_df["semana"]) == {1}  # primera semana de esa base -> 1

    fila_co = tracks_df[tracks_df.country_code == "CO"].iloc[0]
    assert fila_co["artist"] == "a" and fila_co["song_name"] == "x"
    assert fila_co["region"] == "Latin"

    otra_semana = history.cargar_tracks_de_semana(2026, 1)
    assert len(otra_semana) == 2
    assert history.cargar_tracks_de_semana(2026, 2).empty  # esa semana no existe todavía


def test_append_semana_tracks_sin_columnas_opcionales_no_falla(tmp_path):
    # region/stream_count/label_group/label_name pueden faltar (algunos
    # callers de prueba no las incluyen) -- no debe tumbar el guardado.
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
    })
    history.append_semana_tracks(df_semana)

    tracks_df = history.cargar_chart_track_weekly()
    assert len(tracks_df) == 1
    assert pd.isna(tracks_df.iloc[0]["region"])


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


def test_append_semana_ms_bandas_calcula_el_pct_por_banda_y_sello(tmp_path):
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    history.seed_historico(chart_csv, ms_csv)

    # pos1 Universal=100, pos2 Sony=100  -> banda 10: Universal 100/200 = 0.5
    # pos15 Universal=50                 -> banda 20: 150/250 = 0.6
    df_semana = pd.DataFrame({
        "country_code": ["CO"] * 3,
        "chart_date": pd.to_datetime(["2026-08-20"] * 3),
        "position": [1, 2, 15],
        "label_group": ["Universal", "Sony", "Universal"],
        "stream_count": [100, 100, 50],
    })
    history.append_semana_ms_bandas(df_semana)

    df = history.cargar_ms_band_label_weekly()
    pct = df.set_index(["banda", "label_group"])["pct_streams"]
    assert pct[(10, "Universal")] == pytest.approx(0.5)
    assert pct[(20, "Universal")] == pytest.approx(0.6)
    assert pct[(200, "Sony")] == pytest.approx(100 / 250)
    # los 7 sellos configurados quedan siempre, aunque no aparezcan esa semana
    assert pct[(10, "Warner")] == 0.0
    assert len(df) == len(config.BANDAS_MARKET_SHARE) * len(config.LABEL_GROUPS_MS)


def test_seed_historico_siembra_el_pct_por_banda_y_es_idempotente(tmp_path):
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    bandas_csv = tmp_path / "seed_bandas.csv"
    pd.DataFrame([
        {"anio": 2026, "semana": 33, "chart_date": "2026-08-13", "country_code": "CO",
         "banda": 200, "label_group": "Universal", "pct_streams": 0.1234},
    ]).to_csv(bandas_csv, index=False)

    history.seed_historico(chart_csv, ms_csv, bandas_csv)
    history.seed_historico(chart_csv, ms_csv, bandas_csv)  # dos veces no duplica

    df = history.cargar_ms_band_label_weekly()
    assert len(df) == 1
    assert df.iloc[0]["pct_streams"] == pytest.approx(0.1234)


def test_semana_ya_cargada_mira_tambien_el_historico_por_banda(tmp_path):
    # Regresión: al sembrar ms_band_label_weekly hasta la semana 33 mientras
    # ms_label_weekly seguía en la 24, una fecha podía estar en una tabla y
    # no en la otra. Si el chequeo solo mirara ms_label_weekly, volver a
    # cargar esa semana la habría duplicado con otro número de semana.
    chart_csv = _csv_vacio(tmp_path, "seed_chart.csv",
                           ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"])
    ms_csv = _csv_vacio(tmp_path, "seed_ms.csv",
                        ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"])
    bandas_csv = tmp_path / "seed_bandas.csv"
    pd.DataFrame([
        {"anio": 2026, "semana": 33, "chart_date": "2026-08-13", "country_code": "CO",
         "banda": 200, "label_group": "Universal", "pct_streams": 0.1234},
    ]).to_csv(bandas_csv, index=False)
    history.seed_historico(chart_csv, ms_csv, bandas_csv)

    assert history.semana_ya_cargada("2026-08-13") is True   # solo está en ms_band_label_weekly
    assert history.semana_ya_cargada("2026-08-20") is False


def test_append_semana_chart_cuenta_tracks_no_filas(tmp_path):
    # Revisión del 16/09/2026: un track con market share compartido llega en
    # varias filas y antes contaba una vez por fila.
    filas = []
    for streams in (62660.8, 328969.2):      # CO pos 100 "+57", caso real
        filas.append({
            "country_code": "CO", "chart_date": pd.Timestamp("2026-09-10"), "position": 100,
            "artist": "KAROL G", "song_name": "+57", "stream_count": streams,
            "label_group": "Universal", "label_name": "UMG",
        })
    history.append_semana_chart(pd.DataFrame(filas))

    df = history.cargar_chart_band_weekly()
    conteo = df[(df.country_code == "CO") & (df.banda == 100)].conteo_universal.iloc[0]
    assert conteo == 1        # antes daba 2


def test_append_semana_chart_no_cuenta_el_track_sin_dueno_identificable(tmp_path):
    # Sin un copyright que diga de quién es el producto, el track no le
    # suma a ningún sello (ver load_data.dueno_del_track).
    filas = [{
        "country_code": "CO", "chart_date": pd.Timestamp("2026-09-10"), "position": 50,
        "artist": "A", "song_name": "Empatada", "stream_count": 1000.0,
        "label_group": sello, "label_name": sello,
    } for sello in ("Universal", "Sony")]
    history.append_semana_chart(pd.DataFrame(filas))

    df = history.cargar_chart_band_weekly()
    assert df[(df.country_code == "CO") & (df.banda == 50)].conteo_universal.iloc[0] == 0
