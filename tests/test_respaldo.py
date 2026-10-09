"""Pruebas del respaldo automático de la base (src/respaldo.py)."""
import logging
import sqlite3
import zipfile

import pandas as pd
import pytest

from src import config, history, mensual, registro, respaldo


def _base_con_una_fila(valor=1):
    conn = history.conectar()
    conn.execute("CREATE TABLE IF NOT EXISTS prueba (x INTEGER)")
    conn.execute("DELETE FROM prueba")
    conn.execute("INSERT INTO prueba VALUES (?)", (valor,))
    conn.commit()
    conn.close()


def _leer_respaldo(ruta, tmp_path):
    with zipfile.ZipFile(ruta) as z:
        assert z.namelist() == ["universal_data.db"]
        z.extract("universal_data.db", tmp_path / "extraido")
    conn = sqlite3.connect(tmp_path / "extraido" / "universal_data.db")
    try:
        return conn.execute("SELECT x FROM prueba").fetchone()[0]
    finally:
        conn.close()


def test_sin_base_no_hace_nada():
    assert not history.DB_PATH.exists()
    assert respaldo.respaldar("semanal") is None
    assert respaldo.listar() == []


def test_guarda_una_copia_comprimida_y_restaurable(tmp_path):
    _base_con_una_fila(42)
    ruta = respaldo.respaldar("semanal")
    assert ruta is not None and ruta.exists()
    assert ruta.parent == history.DB_PATH.parent / "respaldos"
    assert ruta.name.startswith("universal_data_") and ruta.name.endswith("_semanal.zip")
    assert _leer_respaldo(ruta, tmp_path) == 42


def test_la_copia_es_de_antes_del_cambio(tmp_path):
    """El sentido del respaldo: lo que había ANTES de que el proceso escriba."""
    _base_con_una_fila(1)
    ruta = respaldo.respaldar("bmat")
    _base_con_una_fila(2)
    assert _leer_respaldo(ruta, tmp_path) == 1


def test_conserva_solo_los_ultimos(monkeypatch):
    _base_con_una_fila()
    for _ in range(5):
        respaldo.respaldar("semanal", conservar=3)
    quedan = respaldo.listar()
    assert len(quedan) == 3
    # Sin archivos a medias
    assert not list(respaldo.carpeta_respaldos().glob("*.tmp"))


def test_dos_respaldos_en_el_mismo_segundo_no_se_pisan():
    _base_con_una_fila()
    a = respaldo.respaldar("semanal", conservar=10)
    b = respaldo.respaldar("semanal", conservar=10)
    assert a != b and a.exists() and b.exists()


def test_motivo_raro_queda_limpio_en_el_nombre():
    _base_con_una_fila()
    ruta = respaldo.respaldar("Antes de / recargar!")
    assert ruta.name.endswith("_antes-de---recargar.zip")


def test_si_falla_no_lanza_y_queda_en_el_registro(monkeypatch):
    _base_con_una_fila()
    registro.configurar()

    def _falla(*a, **k):
        raise OSError("disco lleno")

    monkeypatch.setattr(respaldo.zipfile, "ZipFile", _falla)
    assert respaldo.respaldar("semanal") is None
    for h in logging.getLogger().handlers:
        h.flush()
    texto = registro.ruta_log().read_text(encoding="utf-8")
    assert "No se pudo respaldar la base antes de: semanal" in texto
    assert "disco lleno" in texto


def test_el_mensual_respalda_antes_de_generar(monkeypatch, tmp_path):
    _base_con_una_fila()
    llamado = {}

    def _falso(anio, mes, salida=None):
        llamado["respaldos_al_entrar"] = len(respaldo.listar())
        return tmp_path / "x.xlsx", []

    monkeypatch.setattr(mensual, "_generar", _falso)
    mensual.generar(2026, 9)
    assert llamado["respaldos_al_entrar"] == 1
