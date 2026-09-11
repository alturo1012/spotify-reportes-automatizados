"""Pruebas de src/bmat.py — el reporte MS BMAT (Promúsica Colombia).

Corre con: pytest tests/test_bmat.py -v
"""
import openpyxl
import pandas as pd
import pytest

from src import bmat, config, history


@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")


def _sembrar_bmat(tmp_path, filas):
    csv = tmp_path / "seed_bmat.csv"
    pd.DataFrame(filas).to_csv(csv, index=False)
    history.seed_historico(bmat_csv=csv)


def _fila(anio, semana, banda, label, tracks, streams, pct):
    return {"anio": anio, "semana": semana, "country_code": "CO", "banda": banda,
            "label_group": label, "tracks": tracks, "streams_millones": streams,
            "pct_streams": pct}


def _semana_completa(anio, semana, banda=10):
    """Una banda completa: Universal 30%, Sony 50% (le gana), el resto 20%."""
    reparto = {"Universal": (3, 3.0, 0.30), "Sony": (5, 5.0, 0.50),
               "Ingrooves": (2, 2.0, 0.20)}
    filas = []
    for label in config.BMAT_LABELS:
        t, s, p = reparto.get(label, (0, 0.0, 0.0))
        filas.append(_fila(anio, semana, banda, label, t, s, p))
    return filas


def test_seed_bmat_guarda_el_historico_y_es_idempotente(tmp_path):
    _sembrar_bmat(tmp_path, _semana_completa(2026, 23))
    _sembrar_bmat(tmp_path, _semana_completa(2026, 23))  # dos veces no duplica

    df = history.cargar_bmat_weekly()
    assert len(df) == len(config.BMAT_LABELS)
    assert set(df["label_group"]) == set(config.BMAT_LABELS)
    assert df["country_code"].unique().tolist() == ["CO"]


def test_estructura_de_las_dos_hojas(tmp_path):
    _sembrar_bmat(tmp_path, _semana_completa(2026, 23))
    salida = tmp_path / "bmat.xlsx"
    bmat.generar_reporte(salida)

    wb = openpyxl.load_workbook(salida)
    assert wb.sheetnames == [config.BMAT_SHEET_RESUMEN, config.BMAT_SHEET_MARKET_SHARE]

    # --- hoja de detalle: 8 bandas x 3 sub-tablas x (8 sellos + Total) ---
    res = wb[config.BMAT_SHEET_RESUMEN]
    assert res.cell(row=4, column=4).value == "Top 10"
    assert res.cell(row=37, column=4).value == "Top 50"
    assert res.cell(row=229, column=4).value == "Top 10.000"
    for etiqueta, fila in (("TRACKS", 6), ("Streams \n(Millones)", 16), ("% \nStreams", 26)):
        assert res.cell(row=fila, column=1).value == etiqueta
        assert [res.cell(row=fila + k, column=2).value for k in range(8)] == config.BMAT_LABELS
        assert res.cell(row=fila + 8, column=2).value == "Total"
    assert config.BMAT_FUENTE in str(res.cell(row=37, column=1).value)
    assert res.freeze_panes == "D4"

    # --- hoja resumen: 8 bandas de 8 sellos, sin fila de Total ---
    ms = wb[config.BMAT_SHEET_MARKET_SHARE]
    assert ms.cell(row=4, column=2).value == "TOP\n10"
    assert [ms.cell(row=4 + k, column=3).value for k in range(8)] == config.BMAT_LABELS
    assert ms.cell(row=13, column=2).value == "TOP\n50"
    assert ms.cell(row=67, column=2).value == "TOP\n10.000"
    assert ms.freeze_panes == "E4"
    assert "B4:B11" in {str(r) for r in ms.merged_cells.ranges}


def test_la_fila_total_suma_los_ocho_sellos(tmp_path):
    _sembrar_bmat(tmp_path, _semana_completa(2026, 23))
    salida = tmp_path / "bmat.xlsx"
    bmat.generar_reporte(salida)

    res = openpyxl.load_workbook(salida)[config.BMAT_SHEET_RESUMEN]
    col = 4  # primera semana
    assert res.cell(row=6 + 8, column=col).value == pytest.approx(10)     # tracks
    assert res.cell(row=16 + 8, column=col).value == pytest.approx(10.0)  # streams
    assert res.cell(row=26 + 8, column=col).value == pytest.approx(1.0)   # %


def test_semaforo_igual_al_del_reporte_ms(tmp_path):
    # Misma regla: rojo para quien le gana a Universal, verde si Universal
    # lidera. Acá Sony (50%) le gana a Universal (30%).
    _sembrar_bmat(tmp_path, _semana_completa(2026, 23))
    salida = tmp_path / "bmat.xlsx"
    bmat.generar_reporte(salida)

    ms = openpyxl.load_workbook(salida)[config.BMAT_SHEET_MARKET_SHARE]
    def color(cel):
        if not cel.fill or cel.fill.fill_type is None:
            return None
        return str(getattr(cel.fill.start_color, "rgb", ""))[-6:]

    fila = {lab: 4 + k for k, lab in enumerate(config.BMAT_LABELS)}
    assert color(ms.cell(row=fila["Sony"], column=5)) == config.COLOR_SEMAFORO_ROJO
    assert color(ms.cell(row=fila["Universal"], column=5)) is None
    assert color(ms.cell(row=fila["Ingrooves"], column=5)) is None


def test_semaforo_pinta_verde_cuando_universal_lidera(tmp_path):
    filas = [_fila(2026, 23, 10, lab, 0, 0.0, 0.0) for lab in config.BMAT_LABELS]
    filas[0].update(tracks=8, streams_millones=8.0, pct_streams=0.80)  # Universal
    filas[3].update(tracks=2, streams_millones=2.0, pct_streams=0.20)  # Sony
    _sembrar_bmat(tmp_path, filas)
    salida = tmp_path / "bmat.xlsx"
    bmat.generar_reporte(salida)

    ms = openpyxl.load_workbook(salida)[config.BMAT_SHEET_MARKET_SHARE]
    celda = ms.cell(row=4, column=5)  # Universal, banda 10, primera semana
    assert str(getattr(celda.fill.start_color, "rgb", ""))[-6:] == config.COLOR_SEMAFORO_VERDE


def test_las_semanas_salen_en_orden_cronologico(tmp_path):
    filas = (_semana_completa(2025, 52) + _semana_completa(2026, 1)
             + _semana_completa(2026, 23))
    _sembrar_bmat(tmp_path, filas)
    salida = tmp_path / "bmat.xlsx"
    bmat.generar_reporte(salida)

    res = openpyxl.load_workbook(salida)[config.BMAT_SHEET_RESUMEN]
    assert [res.cell(row=2, column=c).value for c in (4, 5, 6)] == [52, 1, 23]
    # el año se escribe una sola vez, en su primera columna
    assert res.cell(row=1, column=4).value == 2025
    assert res.cell(row=1, column=5).value == 2026
    assert res.cell(row=1, column=6).value is None


def test_sin_historico_genera_el_esqueleto_sin_reventar(tmp_path):
    salida = tmp_path / "bmat.xlsx"
    bmat.generar_reporte(salida)

    wb = openpyxl.load_workbook(salida)
    res = wb[config.BMAT_SHEET_RESUMEN]
    assert res.cell(row=4, column=4).value == "Top 10"          # el formato está
    assert res.cell(row=6, column=2).value == "Universal"
    assert res.cell(row=6, column=4).value is None              # pero sin datos
