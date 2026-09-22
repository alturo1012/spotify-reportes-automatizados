"""Configuración común a TODAS las pruebas (pytest la carga sola, no hay que
importarla desde ningún lado).

Redirige la base de datos y la carpeta de salida a archivos temporales por
prueba, para que ninguna toque `data/history/universal_data.db` ni
`data/output/` de verdad. Cada archivo de pruebas
trae además su propio fixture equivalente; este de acá es la red de
seguridad: aunque alguno quede desactualizado, la base real sigue intacta.

Lo que NO hace falta hacer acá es aislar los CSV de siembra: eso lo resuelve
`history.seed_historico()` en el origen, con la regla de "todo o exactamente
lo que nombres" (ver su docstring). Una prueba que siembra un histórico
chico nombrando dos CSV obtiene esas dos tablas y nada más.
"""
import pytest

from src import config, history


@pytest.fixture(autouse=True)
def base_de_datos_temporal(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "conftest_universal_data.db")


@pytest.fixture(autouse=True)
def salida_temporal(tmp_path, monkeypatch):
    """Los reportes que generan las pruebas van a una carpeta temporal, nunca
    a data/output/. Antes test_recargar_semanas escribía ahí de verdad:
    pisaba los reportes reales de la semana 35/36 con datos de prueba, y si
    el usuario tenía uno abierto en Excel la prueba fallaba con
    PermissionError."""
    monkeypatch.setattr(config, "OUTPUT_DIR", tmp_path / "conftest_output")

