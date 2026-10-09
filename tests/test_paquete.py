"""Pruebas del paquete de instalación (src/paquete.py)."""
import sqlite3
import zipfile

import pytest

from src import config, history, paquete
from src.version import VERSION


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    """Una carpeta de proyecto falsa con todo lo que el paquete busca."""
    r = tmp_path / "proyecto"
    (r / "data" / "history" / "seed").mkdir(parents=True)
    (r / "data" / "history" / "seed" / "seed_x.csv").write_text("a,b\n1,2\n")
    (r / "data" / "bmat_titularidad_compartida.xlsx").write_bytes(b"xlsx")
    (r / "data" / "release_date.db").write_bytes(b"db")
    (r / "data" / "preferencias.json").write_text('{"carpeta_salida": "C:/algo"}')
    (r / "data" / "logs").mkdir()
    (r / "data" / "logs" / "aplicativo.log").write_text("log")
    (r / ".env.example").write_text("SPOTIFY_CLIENT_ID=\n")
    (r / ".env").write_text("SPOTIFY_CLIENT_ID=secreto\n")
    (r / paquete.NOMBRE_EXE).write_bytes(b"MZ")
    monkeypatch.setattr(config, "ROOT_DIR", r)
    monkeypatch.setattr(config, "RELEASE_DATE_DB", r / "data" / "release_date.db")
    monkeypatch.setattr(history, "DB_PATH", r / "data" / "history" / "universal_data.db")
    conn = history.conectar()
    conn.execute("CREATE TABLE prueba (x INTEGER)")
    conn.execute("INSERT INTO prueba VALUES (7)")
    conn.commit()
    conn.close()
    return r


def test_arma_el_zip_con_lo_necesario_y_sin_lo_privado(raiz, tmp_path):
    ruta, estado = paquete.generar(tmp_path / "salida")
    assert ruta.name.startswith(f"ReportesSpotifyLatam_v{VERSION}_instalacion_")
    with zipfile.ZipFile(ruta) as z:
        nombres = set(z.namelist())
        leeme = z.read("ReportesSpotifyLatam/LEEME_INSTALACION.txt").decode("utf-8")
        z.extract("ReportesSpotifyLatam/data/history/universal_data.db", tmp_path / "x")
    c = "ReportesSpotifyLatam/"
    for esperado in [paquete.NOMBRE_EXE, "data/history/universal_data.db",
                     "data/history/seed/seed_x.csv", "data/release_date.db",
                     "data/bmat_titularidad_compartida.xlsx", ".env.example",
                     "LEEME_INSTALACION.txt", "data/output/", "data/raw/"]:
        assert c + esperado in nombres, esperado
    for privado in [".env", "data/preferencias.json", "data/logs/aplicativo.log"]:
        assert c + privado not in nombres, privado
    # la base viaja completa y legible
    conn = sqlite3.connect(tmp_path / "x" / "ReportesSpotifyLatam/data/history/universal_data.db")
    assert conn.execute("SELECT x FROM prueba").fetchone()[0] == 7
    conn.close()
    assert VERSION in leeme and "Spotify semanal" in leeme and "Copia .env.example" in leeme
    assert len(estado) == 3


def test_con_credenciales_incluye_el_env(raiz, tmp_path):
    ruta, _ = paquete.generar(tmp_path / "salida", con_credenciales=True)
    with zipfile.ZipFile(ruta) as z:
        assert "ReportesSpotifyLatam/.env" in z.namelist()


def test_sin_ejecutable_no_arma_nada(raiz, tmp_path):
    (raiz / paquete.NOMBRE_EXE).unlink()
    with pytest.raises(FileNotFoundError, match="build.bat"):
        paquete.generar(tmp_path / "salida")
    assert not (tmp_path / "salida").exists() or not list((tmp_path / "salida").glob("*.zip"))


def test_sin_base_no_arma_nada(raiz, tmp_path):
    history.DB_PATH.unlink()
    with pytest.raises(FileNotFoundError, match="base"):
        paquete.generar(tmp_path / "salida")
