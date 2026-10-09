"""Pruebas del registro de errores (src/registro.py)."""
import logging

import pytest

from src import config, mensual, registro


def _texto():
    for h in logging.getLogger().handlers:
        h.flush()
    return registro.ruta_log().read_text(encoding="utf-8")


def _manejadores_del_registro():
    return [h for h in logging.getLogger().handlers if getattr(h, registro._MARCA, False)]


def test_escribe_en_la_carpeta_de_logs():
    ruta = registro.configurar()
    assert ruta == config.LOG_DIR / "aplicativo.log"
    logging.getLogger("prueba").info("hola registro")
    assert "hola registro" in _texto()


def test_configurar_dos_veces_no_duplica_lineas():
    registro.configurar()
    registro.configurar()
    assert len(_manejadores_del_registro()) == 1
    logging.getLogger("prueba").info("una sola vez")
    assert _texto().count("una sola vez") == 1


def test_un_error_queda_con_el_detalle_tecnico():
    registro.configurar()
    try:
        1 / 0
    except ZeroDivisionError:
        logging.getLogger("prueba").exception("falló algo")
    texto = _texto()
    assert "falló algo" in texto and "ZeroDivisionError" in texto and "Traceback" in texto


def test_carpeta_sin_permiso_no_rompe(monkeypatch, tmp_path):
    archivo = tmp_path / "no_es_carpeta"
    archivo.write_text("x")
    monkeypatch.setattr(config, "LOG_DIR", archivo / "logs")   # no se puede crear
    registro.configurar()   # no lanza


def test_un_error_del_mensual_queda_registrado(monkeypatch):
    def _explota(*a, **k):
        raise ValueError("semana rara")

    monkeypatch.setattr(mensual, "_generar", _explota)
    with pytest.raises(ValueError):
        mensual.generar(2026, 9)
    texto = _texto()
    assert "Error en el reporte mensual 2026-09" in texto and "semana rara" in texto
