"""Pruebas de las constantes/funciones de src/config.py.

Corre con: pytest tests/test_config.py
"""
from src import config


def test_country_code_map_tiene_los_17_paises_de_paises_ms():
    codigos_en_mapa = set(config.COUNTRY_CODE_MAP.values())
    assert set(config.PAISES_MS) == codigos_en_mapa


def test_normalizar_label_group_agrupa_som_livre_y_altafonte_en_indies():
    assert config.normalizar_label_group("Som Livre") == "Indies"
    assert config.normalizar_label_group("Altafonte") == "Indies"


def test_normalizar_label_group_deja_igual_los_demas_valores():
    for valor in ["Universal", "Sony", "Orchard", "Warner", "INgrooves", "Virgin", "Indies"]:
        assert config.normalizar_label_group(valor) == valor


def test_bandas_chart_y_market_share_son_distintas():
    # Confirmado contra las plantillas reales: no son la misma lista de bandas.
    assert config.BANDAS_CHART == [10, 30, 50, 100, 200]
    assert config.BANDAS_MARKET_SHARE == [10, 20, 50, 100, 200]


def test_label_groups_ms_tiene_las_7_categorias_en_orden_de_la_plantilla():
    assert config.LABEL_GROUPS_MS == [
        "Universal", "INgrooves", "Virgin", "Sony", "Orchard", "Warner", "Indies",
    ]