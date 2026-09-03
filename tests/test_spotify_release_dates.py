"""Pruebas de src/spotify_release_dates.py — usan una base SQLite temporal
(nunca tocan data/history/universal_data.db de verdad) y un cliente FALSO
en vez de la API real de Spotify (no hay red disponible en pruebas, y no
queremos depender de credenciales reales para que los tests sean
deterministas).

Corre con: pytest tests/test_spotify_release_dates.py -v
"""
import pandas as pd
import pytest

from src import history, spotify_release_dates


@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")


class ClienteFalso:
    """Doble de prueba de SpotifyReleaseDateClient -- no llama a la red,
    cuenta cuántas veces se le pide cada cosa (para verificar que la caché
    realmente evita llamadas repetidas)."""

    def __init__(self, id_por_isrc: dict, fecha_por_id: dict):
        self.id_por_isrc = id_por_isrc
        self.fecha_por_id = fecha_por_id
        self.llamadas_busqueda = []
        self.llamadas_fecha = []

    def buscar_track_id_por_isrc(self, isrc):
        self.llamadas_busqueda.append(isrc)
        return self.id_por_isrc.get(isrc)

    def fechas_de_lanzamiento(self, track_ids):
        self.llamadas_fecha.append(list(track_ids))
        return {tid: self.fecha_por_id.get(tid) for tid in track_ids}


def test_resolver_fechas_lanzamiento_resuelve_y_cachea(tmp_path):
    cliente = ClienteFalso(
        id_por_isrc={"ISRC1": "trackA", "ISRC2": "trackB"},
        fecha_por_id={"trackA": "2024-05-10", "trackB": "2023-01-01"},
    )
    isrcs = pd.Series(["ISRC1", "ISRC2", "ISRC1"])  # ISRC1 repetido

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=cliente)

    assert list(fechas) == ["2024-05-10", "2023-01-01", "2024-05-10"]
    # 1 sola llamada de búsqueda por ISRC único (no 3, aunque ISRC1 se repite)
    assert sorted(cliente.llamadas_busqueda) == ["ISRC1", "ISRC2"]

    # Segunda corrida: todo ya en caché, no debería volver a llamar al cliente.
    cliente2 = ClienteFalso(id_por_isrc={}, fecha_por_id={})
    fechas2 = spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=cliente2)
    assert list(fechas2) == ["2024-05-10", "2023-01-01", "2024-05-10"]
    assert cliente2.llamadas_busqueda == []
    assert cliente2.llamadas_fecha == []


def test_resolver_fechas_lanzamiento_isrc_no_encontrado_da_none_y_no_reintenta(tmp_path):
    cliente = ClienteFalso(id_por_isrc={}, fecha_por_id={})  # ISRC no existe en Spotify
    isrcs = pd.Series(["ISRC_INEXISTENTE"])

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=cliente)
    assert fechas.iloc[0] is None
    assert cliente.llamadas_busqueda == ["ISRC_INEXISTENTE"]

    # Ya quedó cacheado como "no encontrado" -- no se vuelve a buscar.
    cliente2 = ClienteFalso(id_por_isrc={}, fecha_por_id={})
    spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=cliente2)
    assert cliente2.llamadas_busqueda == []


def test_resolver_fechas_lanzamiento_isrc_faltante_no_llama_al_cliente(tmp_path):
    isrcs = pd.Series([None, float("nan")])
    fechas = spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=None)
    assert list(fechas.isna()) == [True, True]


def test_resolver_fechas_lanzamiento_dos_isrc_mismo_track_id_una_sola_llamada_de_fecha(tmp_path):
    # Dos ISRC distintos (ej. versión explícita/limpia) pueden resolver al
    # mismo track_id -- no debería pedirse la fecha dos veces para ese id.
    cliente = ClienteFalso(
        id_por_isrc={"ISRC_A": "trackX", "ISRC_B": "trackX"},
        fecha_por_id={"trackX": "2022-02-02"},
    )
    isrcs = pd.Series(["ISRC_A", "ISRC_B"])

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=cliente)
    assert list(fechas) == ["2022-02-02", "2022-02-02"]
    assert cliente.llamadas_fecha == [["trackX"]]


def test_spotify_release_date_client_sin_credenciales_lanza_error(tmp_path, monkeypatch):
    monkeypatch.delenv("SPOTIFY_CLIENT_ID", raising=False)
    monkeypatch.delenv("SPOTIFY_CLIENT_SECRET", raising=False)
    with pytest.raises(RuntimeError):
        spotify_release_dates.SpotifyReleaseDateClient()
