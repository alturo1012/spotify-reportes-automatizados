"""Pruebas de src/load_data.py — sobre todo la detección del nombre de la
hoja de datos, que la exportación de BigQuery cambió sin avisar.

Corre con: pytest tests/test_load_data.py -v
"""
import pandas as pd
import pytest

from src import config, load_data


def _fuente(tmp_path, hoja: str, nombre="fuente.xlsx"):
    """Un xlsx con las columnas que espera load_source, en la hoja indicada."""
    fila = {
        "country": "Colombia", "country_alt": "Colombia", "chart_date": "2026-09-10",
        "is_latest_date": True, "artist": "Artista", "song_name": "Cancion",
        "position": 1, "stream_count": 1_000_000, "ISRC": "ISRC1",
        "label_group": "Universal", "repertoire": "Local", "repertoire_group": "Local",
        "album_copyright": "UMG", "label_name": "UMG", "content_provider_name": "UMG",
        "major_label": "Universal", "artist_country": "Colombia", "region": "LATAM",
        "main_language": "es",
    }
    path = tmp_path / nombre
    pd.DataFrame([fila]).to_excel(path, sheet_name=hoja, index=False)
    return path


@pytest.mark.parametrize("hoja", ["Consulta1", "spotify", "Spotify", " SPOTIFY "])
def test_carga_la_fuente_con_cualquiera_de_los_nombres_de_hoja_conocidos(tmp_path, hoja):
    # La fuente venía siempre como "Consulta1" y en la semana 36 de 2026
    # llegó como "spotify" -- el .exe reventaba. Se aceptan los dos, sin
    # importar mayúsculas ni espacios de más.
    df = load_data.load_source(_fuente(tmp_path, hoja))
    assert len(df) == 1
    assert df["country_code"].iloc[0] == "CO"


def test_error_claro_cuando_no_esta_ninguna_hoja_conocida(tmp_path):
    # El mensaje tiene que decir qué hojas TIENE el archivo: es lo que hace
    # falta para agregar el nombre nuevo a config.HOJAS_FUENTE.
    path = _fuente(tmp_path, "Hoja3")
    with pytest.raises(ValueError, match="No se encontró la hoja de datos"):
        load_data.load_source(path)
    try:
        load_data.load_source(path)
    except ValueError as e:
        assert "Hoja3" in str(e)
        assert "Consulta1" in str(e)


def test_se_puede_seguir_pidiendo_una_hoja_a_mano(tmp_path):
    df = load_data.load_source(_fuente(tmp_path, "Consulta1"), sheet_name="Consulta1")
    assert len(df) == 1


def test_las_hojas_conocidas_estan_configuradas():
    assert "Consulta1" in config.HOJAS_FUENTE
    assert "spotify" in config.HOJAS_FUENTE
