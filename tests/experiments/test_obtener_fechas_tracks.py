
import pytest

from experiments.spotify_api.obtener_fechas_tracks import obtener_fechas_tracks


@pytest.mark.skip(reason="Pendiente: pegar el código real de la reunión primero")
def test_obtener_fechas_tracks_devuelve_fecha():
    resultado = obtener_fechas_tracks(["<track_id_o_isrc_de_prueba>"])
    assert resultado is not None
