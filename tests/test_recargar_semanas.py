"""Prueba de scripts/recargar_semanas.py — reconstruir el histórico cuando
una semana se saltó y las siguientes quedaron corridas de número.

Corre con: pytest tests/test_recargar_semanas.py -v
"""
import pandas as pd
import pytest

from scripts import recargar_semanas
from src import history, main


_SEED_REAL = history.seed_historico


@pytest.fixture(autouse=True)
def entorno_temporal(tmp_path, monkeypatch):
    """Base temporal y siembra VACÍA.

    El script llama a `seed_historico()` sin argumentos, que en producción
    carga el histórico real (hasta la semana 33 de 2026). En las pruebas eso
    traería 170.000 filas y ataría los números esperados al contenido de los
    CSV reales, así que se reemplaza por una siembra vacía: lo que se está
    probando es el ORDEN en que quedan las semanas, no la siembra.
    """
    monkeypatch.setattr(history, "DB_PATH", tmp_path / "test_universal_data.db")
    monkeypatch.setattr(history, "seed_historico", lambda *a, **k: _SEED_REAL(*(a or _seed_vacio())))


def _fuente(tmp_path, chart_date):
    filas = [{
        "country": "Colombia", "country_alt": "Colombia", "chart_date": chart_date,
        "is_latest_date": True, "artist": "Artista", "song_name": "Cancion",
        "position": 1, "stream_count": 1_000_000, "ISRC": "ISRC1",
        "label_group": "Universal", "repertoire": "Local", "repertoire_group": "Local",
        "album_copyright": "UMG", "label_name": "UMG", "content_provider_name": "UMG",
        "major_label": "Universal", "artist_country": "Colombia", "region": "LATAM",
        "main_language": "es",
    }]
    path = tmp_path / f"fuente_{chart_date}.xlsx"
    pd.DataFrame(filas).to_excel(path, sheet_name="Consulta1", index=False)
    return str(path)


def _seed_vacio():
    import io
    chart_csv = io.StringIO()
    pd.DataFrame(columns=["anio", "semana", "mes", "country_code", "banda", "conteo_universal"]).to_csv(chart_csv, index=False)
    chart_csv.seek(0)
    ms_csv = io.StringIO()
    pd.DataFrame(columns=["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"]).to_csv(ms_csv, index=False)
    ms_csv.seek(0)
    return chart_csv, ms_csv


def _semanas_por_fecha():
    df = history.cargar_ms_band_label_weekly()
    pares = df[["semana", "chart_date"]].drop_duplicates().sort_values("semana")
    return list(zip(pares["semana"], pares["chart_date"]))


def test_recargar_pone_en_su_lugar_la_semana_que_faltaba(tmp_path):
    # Reproduce el caso real: se pierde el 20-ago, se siguen cargando las
    # siguientes y quedan corridas (33 -> 27-ago, 34 -> 03-sep).
    history.seed_historico()
    f20 = _fuente(tmp_path, "2026-08-20")
    f27 = _fuente(tmp_path, "2026-08-27")
    f03 = _fuente(tmp_path, "2026-09-03")

    main.main(["--fuente", f27, "--semana", "35"])
    main.main(["--fuente", f03, "--semana", "36"])
    assert [f for _, f in _semanas_por_fecha()] == ["2026-08-27", "2026-09-03"]

    recargar_semanas.main([f20, f27, f03])

    # Ahora están las tres, numeradas en orden de fecha y sin huecos.
    pares = _semanas_por_fecha()
    assert [f for _, f in pares] == ["2026-08-20", "2026-08-27", "2026-09-03"]
    assert [int(s) for s, _ in pares] == [1, 2, 3]


def test_recargar_no_duplica_lo_que_ya_venia_en_la_siembra(tmp_path):
    history.seed_historico()
    f20 = _fuente(tmp_path, "2026-08-20")
    main.main(["--fuente", f20, "--semana", "34"])

    recargar_semanas.main([f20])   # el mismo archivo otra vez

    pares = _semanas_por_fecha()
    assert [f for _, f in pares] == ["2026-08-20"]


def test_vaciar_historico_no_borra_el_cache_de_spotify(tmp_path):
    history.seed_historico()
    main.main(["--fuente", _fuente(tmp_path, "2026-08-20"), "--semana", "34"])

    conn = history._conectar()
    conn.execute("INSERT OR REPLACE INTO spotify_isrc_cache (isrc, track_id) VALUES (?, ?)",
                 ("ISRC1", "abc123"))
    conn.commit()
    conn.close()

    history.vaciar_historico()

    assert history.cargar_ms_band_label_weekly().empty
    conn = history._conectar()
    try:
        fila = conn.execute("SELECT track_id FROM spotify_isrc_cache WHERE isrc = 'ISRC1'").fetchone()
    finally:
        conn.close()
    assert fila is not None and fila[0] == "abc123"
