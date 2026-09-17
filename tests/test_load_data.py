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


# --- tracks compartidos entre sellos (revisión del 16/09/2026) ---

def _fila(pos, sello, streams, nombre="Cancion", isrc="ISRC1", pais="Colombia",
          copyright="x"):
    return {
        "country": pais, "country_alt": pais, "chart_date": "2026-09-10",
        "is_latest_date": True, "artist": "Artista", "song_name": nombre,
        "position": pos, "stream_count": streams, "ISRC": isrc,
        "label_group": sello, "repertoire": "Local", "repertoire_group": "Local",
        "album_copyright": copyright, "label_name": "x", "content_provider_name": "x",
        "major_label": sello, "artist_country": pais, "region": "LATAM",
        "main_language": "es",
    }


def _fuente_filas(tmp_path, filas, nombre="compartidos.xlsx"):
    path = tmp_path / nombre
    pd.DataFrame(filas).to_excel(path, sheet_name="Consulta1", index=False)
    return load_data.load_source(path)


def test_tracks_unicos_suma_los_streams_del_track_partido(tmp_path):
    # Caso real: CO pos 100 "+57" llega partido Universal 16% + Universal 84%.
    df = _fuente_filas(tmp_path, [
        _fila(100, "Universal", 62660.8),
        _fila(100, "Universal", 328969.2),
    ])
    u = load_data.tracks_unicos(df)

    assert len(u) == 1                                   # una sola línea
    assert u["stream_count"].iloc[0] == pytest.approx(391630.0)
    assert u["label_group"].iloc[0] == "Universal"


def test_tracks_unicos_asigna_el_track_a_su_dueno_aunque_no_tenga_la_mayoria(tmp_path):
    # La regla que confirmaron Tatiana y Alejandro: manda la PROPIEDAD del
    # producto, no la participación. Caso real: PT pos 182 "Maria Joana" --
    # acá invertido a propósito, con Universal en mayoría (67%) y el
    # copyright de Warner: el track es de Warner igual.
    df = _fuente_filas(tmp_path, [
        _fila(182, "Universal", 44292.67, copyright="Universal Music"),
        _fila(182, "Warner", 22146.33,
              copyright="© 2023 Warner Music Portugal, Lda, ℗ 2023 Warner Music Portugal, Lda"),
    ])
    u = load_data.tracks_unicos(df)

    assert len(u) == 1
    assert u["label_group"].iloc[0] == "Warner"
    assert u["stream_count"].iloc[0] == pytest.approx(66439.0)


def test_tracks_unicos_resuelve_el_empate_por_el_copyright(tmp_path):
    # Caso real: PT pos 84 "Faz Bem", Sony 50% + Universal 50%. La
    # participación no decide; el copyright dice que el producto es de
    # Universal, así que el track se le cuenta a Universal.
    df = _fuente_filas(tmp_path, [
        _fila(84, "Sony", 48361.5, copyright="Sony Music"),
        _fila(84, "Universal", 48361.5,
              copyright="© 2024 Universal Music Portugal, S.A., ℗ 2024 Universal Music Portugal, S.A."),
    ])
    u = load_data.tracks_unicos(df)

    assert len(u) == 1
    assert u["label_group"].iloc[0] == "Universal"
    assert u["stream_count"].iloc[0] == pytest.approx(96723.0)


def test_tracks_unicos_sin_copyright_util_no_se_lo_asigna_a_nadie(tmp_path):
    # Si ninguna de las dos filas trae un aviso de copyright de verdad, no
    # hay forma de saber de quién es el producto: mejor no contarlo que
    # contárselo a quien no es.
    df = _fuente_filas(tmp_path, [
        _fila(91, "Universal", 100.0, copyright="UMLA"),
        _fila(91, "Indies", 100.0, copyright="Indies"),
    ])
    u = load_data.tracks_unicos(df)

    assert u["label_group"].iloc[0] is None


def test_tracks_unicos_no_toca_los_tracks_normales(tmp_path):
    df = _fuente_filas(tmp_path, [
        _fila(1, "Universal", 1000.0, nombre="A", isrc="I1"),
        _fila(2, "Sony", 500.0, nombre="B", isrc="I2"),
    ])
    u = load_data.tracks_unicos(df)

    assert len(u) == 2
    assert list(u["label_group"]) == ["Universal", "Sony"]
    assert list(u["stream_count"]) == [1000.0, 500.0]


def test_tracks_unicos_separa_por_pais(tmp_path):
    # La misma posición en dos países son dos tracks distintos.
    df = _fuente_filas(tmp_path, [
        _fila(5, "Universal", 100.0, pais="Colombia"),
        _fila(5, "Universal", 200.0, pais="Peru"),
    ])
    u = load_data.tracks_unicos(df)

    assert len(u) == 2
    assert sorted(u["stream_count"]) == [100.0, 200.0]
