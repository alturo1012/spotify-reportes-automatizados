"""Configuración común a TODAS las pruebas (pytest la carga solo, no hay que
importarla desde ningún lado).

Aísla el histórico: ninguna prueba toca la base real del proyecto
(`data/history/universal_data.db`) ni los CSV de siembra reales de
`data/history/seed/`. La prueba que sí necesita los datos reales lo pide
explícitamente con el fixture `seeds_reales`.

Por qué está centralizado acá y no repetido en cada archivo de pruebas:
antes cada archivo traía su propio fixture de aislamiento, y al agregar el
tercer CSV de siembra (`seed_ms_band_label_weekly.csv`, 171.850 filas) hubo
que tocar los seis. Con que UNO quedara desactualizado, sus pruebas se
tragaban el CSV real sin avisar y fallaban de forma difícil de entender
(pasó de verdad: `semana_ya_cargada("2026-01-08")` devolvía True porque esa
fecha existe en el histórico real sembrado sin querer). Un solo lugar
elimina esa clase de error.
"""
import pytest

from src import history


@pytest.fixture(autouse=True)
def historico_aislado(tmp_path, monkeypatch):
    """Base temporal y CSV de siembra apuntando a archivos que no existen.

    Con los SEED_* apuntando a rutas inexistentes, `seed_historico()` omite
    esas tablas (ver su docstring), así que cada prueba arranca con el
    histórico que ella misma siembra y nada más.
    """
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")
    for constante in ("SEED_CHART_CSV", "SEED_MS_CSV", "SEED_MS_BANDAS_CSV"):
        monkeypatch.setattr(history, constante, tmp_path / f"sin_{constante.lower()}.csv")


@pytest.fixture
def seeds_reales(monkeypatch):
    """Devuelve los SEED_* a los archivos reales de `data/history/seed/`.

    Solo para las pruebas de validación que comparan contra los reportes
    oficiales y por lo tanto necesitan el histórico de producción.
    """
    monkeypatch.setattr(history, "SEED_CHART_CSV", history.SEED_DIR / "seed_chart_band_weekly.csv")
    monkeypatch.setattr(history, "SEED_MS_CSV", history.SEED_DIR / "seed_ms_label_weekly.csv")
    monkeypatch.setattr(history, "SEED_MS_BANDAS_CSV", history.SEED_DIR / "seed_ms_band_label_weekly.csv")
