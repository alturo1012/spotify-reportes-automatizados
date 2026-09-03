"""Pruebas de src/gui.py -- solo de `generar()`, la función que conecta los
botones de la ventana con `main.py`. NO importan tkinter ni abren ninguna
ventana (este entorno no tiene pantalla), así que no prueban la parte
visual -- esa se prueba a mano una vez empaquetado el .exe (ver Paso 7 en
claude/plan_fusion_paso_a_paso.md).

Corre con: pytest tests/test_gui.py -v
"""
import pandas as pd
import pytest

from src import config, gui, history


@pytest.fixture(autouse=True)
def entorno_temporal(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")
    monkeypatch.setattr(config, "OUTPUT_DIR", tmp_path / "output")


def _fuente_minima(tmp_path, chart_date="2026-06-18"):
    filas = [{
        "country": "Colombia", "country_alt": "Colombia", "chart_date": chart_date,
        "is_latest_date": True, "artist": "Artista", "song_name": "Cancion",
        "position": 1, "stream_count": 1_000_000, "ISRC": "ISRC1",
        "label_group": "Universal", "repertoire": "Local", "repertoire_group": "Local",
        "album_copyright": "UMG", "label_name": "UMG", "content_provider_name": "UMG",
        "major_label": "Universal", "artist_country": "Colombia", "region": "LATAM",
        "main_language": "es",
    }]
    path = tmp_path / "fuente.xlsx"
    pd.DataFrame(filas).to_excel(path, sheet_name="Consulta1", index=False)
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


def test_generar_devuelve_las_rutas_de_los_dos_reportes_generados(tmp_path):
    history.seed_historico(*_seed_vacio())
    fuente = _fuente_minima(tmp_path)

    chart_out, ms_out = gui.generar(str(fuente), "25")

    assert chart_out.exists()
    assert ms_out.exists()
    assert chart_out.name == "Reporte_Chart_Top_Semanal_Sem_25.xlsx"
    assert ms_out.name == "Reporte_MS_TOP200_Sem_25.xlsx"


def test_generar_propaga_el_error_si_el_archivo_no_existe(tmp_path):
    with pytest.raises(SystemExit, match="No se encontró el archivo fuente"):
        gui.generar(str(tmp_path / "no_existe.xlsx"), "25")
