"""Pruebas de src/main.py (Paso 5: orquestación completa) — usan una base
SQLite temporal y una carpeta de salida temporal (nunca tocan
data/history/universal_data.db ni data/output/ de verdad).

Corre con: pytest tests/test_main.py -v
"""
import pandas as pd
import pytest

from src import config, history, main


@pytest.fixture(autouse=True)
def entorno_temporal(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")
    monkeypatch.setattr(config, "OUTPUT_DIR", tmp_path / "output")


def _fuente_minima(tmp_path, chart_date="2026-06-18"):
    """Un xlsx mínimo con las columnas que espera load_data.load_source."""
    filas = []
    for pais, artistas in [("Colombia", ["Artista A", "Artista B"]), ("Peru", ["Artista C"])]:
        for i, artista in enumerate(artistas, start=1):
            filas.append({
                "country": pais, "country_alt": pais, "chart_date": chart_date,
                "is_latest_date": True, "artist": artista, "song_name": f"Cancion {i}",
                "position": i, "stream_count": 1_000_000 * i, "ISRC": f"ISRC{i}",
                "label_group": "Universal", "repertoire": "Local", "repertoire_group": "Local",
                "album_copyright": "UMG", "label_name": "UMG", "content_provider_name": "UMG",
                "major_label": "Universal", "artist_country": pais, "region": "LATAM",
                "main_language": "es",
            })
    df = pd.DataFrame(filas)
    path = tmp_path / "fuente.xlsx"
    df.to_excel(path, sheet_name="Consulta1", index=False)
    return path


def _seed_vacio():
    import io
    chart_csv = io.StringIO()
    pd.DataFrame(columns=["anio", "semana", "mes", "country_code", "banda", "conteo_universal"]).to_csv(chart_csv, index=False)
    chart_csv.seek(0)
    ms_csv = io.StringIO()
    pd.DataFrame(columns=["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"]).to_csv(ms_csv, index=False)
    ms_csv.seek(0)
    return chart_csv, ms_csv


def test_main_genera_los_dos_reportes(tmp_path):
    history.seed_historico(*_seed_vacio())
    fuente = _fuente_minima(tmp_path)

    main.main(["--fuente", str(fuente), "--semana", "25"])

    assert (config.OUTPUT_DIR / "Reporte_Chart_Top_Semanal_Sem_25.xlsx").exists()
    assert (config.OUTPUT_DIR / "Reporte_MS_TOP200_Sem_25.xlsx").exists()

    ms_df = history.cargar_ms_label_weekly()
    assert len(ms_df) > 0  # la semana quedó guardada en el histórico


def test_main_guarda_el_detalle_track_por_track_en_el_historico(tmp_path):
    # Desde el ajuste de "empezar a guardar de ahora en adelante": cada
    # corrida normal de main.py debe dejar el detalle track por track en
    # chart_track_weekly, no solo los conteos agregados.
    history.seed_historico(*_seed_vacio())
    fuente = _fuente_minima(tmp_path)

    main.main(["--fuente", str(fuente), "--semana", "25"])

    tracks_df = history.cargar_chart_track_weekly()
    assert len(tracks_df) == 3  # 2 tracks CO + 1 track PE en _fuente_minima
    assert set(tracks_df["country_code"]) == {"CO", "PE"}


def test_main_corrido_dos_veces_con_la_misma_fuente_no_duplica_la_semana(tmp_path):
    # Regresión de un problema real que se anticipó en el plan (Paso 5):
    # _proxima_semana() siempre calcula "la siguiente", así que sin este
    # chequeo, correr main.py dos veces con el mismo archivo (por error, o
    # por reintentar tras una falla) generaría una semana fantasma extra en
    # vez de detectar que esa fecha ya estaba.
    history.seed_historico(*_seed_vacio())
    fuente = _fuente_minima(tmp_path)

    main.main(["--fuente", str(fuente), "--semana", "25"])
    main.main(["--fuente", str(fuente), "--semana", "25"])

    ms_df = history.cargar_ms_label_weekly()
    chart_df = history.cargar_chart_band_weekly()
    tracks_df = history.cargar_chart_track_weekly()
    assert ms_df["semana"].max() == 1
    assert chart_df["semana"].max() == 1
    assert tracks_df["semana"].max() == 1
    assert len(tracks_df) == 3  # tampoco se duplicó el detalle track por track


def test_main_falla_con_mensaje_claro_si_no_existe_el_archivo(tmp_path):
    with pytest.raises(SystemExit, match="No se encontró el archivo fuente"):
        main.main(["--fuente", str(tmp_path / "no_existe.xlsx"), "--semana", "25"])
