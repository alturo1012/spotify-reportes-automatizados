"""Pruebas de src/spotify_release_dates.py — usan una base SQLite temporal
(nunca tocan data/history/universal_data.db de verdad) y un cliente FALSO
en vez de la API real de Spotify (no hay red disponible en pruebas, y no
queremos depender de credenciales reales para que los tests sean
deterministas).

Corre con: pytest tests/test_spotify_release_dates.py -v
"""
import pandas as pd
import pytest

from src import config, history, spotify_release_dates


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


def test_isrc_no_encontrado_queda_cacheado_y_la_corrida_siguiente_no_revienta(tmp_path):
    # REGRESIÓN del bug real que rompía el .exe: cuando Spotify no encuentra
    # un ISRC, se cachea con track_id NULL. En la corrida SIGUIENTE, pandas
    # lee ese NULL como NaN (float) y el `sorted(...)` de los track_ids
    # reventaba con "'<' not supported between instances of 'float' and
    # 'str'" -- o sea, la primera corrida "envenenaba" la caché y todas las
    # siguientes fallaban, dejando la columna vacía para siempre.
    cliente = ClienteFalso(
        id_por_isrc={"ISRC_SI": "trackX", "ISRC_NO": None},
        fecha_por_id={"trackX": "2021-07-07"},
    )
    isrcs = pd.Series(["ISRC_SI", "ISRC_NO"])

    primera = spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=cliente)
    assert list(primera) == ["2021-07-07", None]

    # Segunda corrida: ahora todo sale de la caché (incluido el "no
    # encontrado"), que es justo donde antes explotaba.
    cliente_que_no_debe_usarse = ClienteFalso(id_por_isrc={}, fecha_por_id={})
    segunda = spotify_release_dates.resolver_fechas_lanzamiento(
        isrcs, cliente=cliente_que_no_debe_usarse,
    )
    assert list(segunda) == ["2021-07-07", None]
    # Y no se vuelve a gastar una llamada a la API por el ISRC no encontrado.
    assert cliente_que_no_debe_usarse.llamadas_busqueda == []


def test_isrc_con_espacios_o_numericos_no_rompe_el_orden(tmp_path):
    # La columna ISRC puede venir con tipos mezclados según la semana
    # (texto, celdas vacías, algún número) -- pandas la deja como `object` y
    # mezclar float con str rompía el sorted(...). Deben normalizarse.
    cliente = ClienteFalso(
        id_por_isrc={"ISRC_A": "trackA"}, fecha_por_id={"trackA": "2019-09-09"},
    )
    isrcs = pd.Series(["  ISRC_A  ", 12345, None, "", float("nan")])

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(isrcs, cliente=cliente)

    assert fechas.iloc[0] == "2019-09-09"  # se limpia y encuentra igual
    assert fechas.iloc[2] is None and fechas.iloc[3] is None and fechas.iloc[4] is None
    # El valor numérico se consulta como texto, sin romper nada.
    assert "12345" in cliente.llamadas_busqueda


# --- base externa release_date.db (del proyecto anterior del usuario) ---

def _crear_db_externa(tmp_path, filas, nombre="release_date.db"):
    """Crea una base con la MISMA forma que la real: track(uri, release_date),
    donde uri es el id de track de Spotify sin el prefijo spotify:track:
    (ver release_date_management.py del proyecto anterior)."""
    import sqlite3
    ruta = tmp_path / nombre
    conn = sqlite3.connect(ruta)
    conn.execute("CREATE TABLE IF NOT EXISTS track (uri TEXT PRIMARY KEY, release_date TEXT)")
    conn.executemany("INSERT INTO track (uri, release_date) VALUES (?, ?)", filas)
    conn.commit()
    conn.close()
    return ruta


def test_la_fecha_sale_de_la_db_externa_sin_llamar_a_la_api(tmp_path, monkeypatch):
    monkeypatch.setattr(
        config, "RELEASE_DATE_DB",
        _crear_db_externa(tmp_path, [("trackX", "2018-04-13")]),
    )
    cliente = ClienteFalso(id_por_isrc={"ISRC_A": "trackX"}, fecha_por_id={})

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(
        pd.Series(["ISRC_A"]), cliente=cliente,
    )

    assert list(fechas) == ["2018-04-13"]
    # el ISRC -> track_id sí necesita la API (la db externa no tiene ISRC)...
    assert cliente.llamadas_busqueda == ["ISRC_A"]
    # ...pero la fecha NO: salió de la base externa.
    assert cliente.llamadas_fecha == []


def test_solo_va_a_la_api_por_lo_que_no_esta_en_la_db_externa(tmp_path, monkeypatch):
    monkeypatch.setattr(
        config, "RELEASE_DATE_DB",
        _crear_db_externa(tmp_path, [("trackX", "2018-04-13")]),
    )
    cliente = ClienteFalso(
        id_por_isrc={"ISRC_A": "trackX", "ISRC_B": "trackY"},
        fecha_por_id={"trackY": "2024-09-09"},
    )

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(
        pd.Series(["ISRC_A", "ISRC_B"]), cliente=cliente,
    )

    assert list(fechas) == ["2018-04-13", "2024-09-09"]
    # a la API solo se le pidió trackY; trackX ya estaba en la base externa
    assert cliente.llamadas_fecha == [["trackY"]]


def test_lo_leido_de_la_db_externa_queda_en_la_cache_propia(tmp_path, monkeypatch):
    # Así, si mañana el archivo externo no está, la fecha sigue resuelta.
    ruta = _crear_db_externa(tmp_path, [("trackX", "2018-04-13")])
    monkeypatch.setattr(config, "RELEASE_DATE_DB", ruta)
    cliente = ClienteFalso(id_por_isrc={"ISRC_A": "trackX"}, fecha_por_id={})
    spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["ISRC_A"]), cliente=cliente)

    monkeypatch.setattr(config, "RELEASE_DATE_DB", tmp_path / "ya_no_existe.db")
    cliente2 = ClienteFalso(id_por_isrc={}, fecha_por_id={})
    fechas = spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["ISRC_A"]), cliente=cliente2)

    assert list(fechas) == ["2018-04-13"]
    assert cliente2.llamadas_busqueda == [] and cliente2.llamadas_fecha == []


def test_sin_db_externa_todo_sigue_funcionando_contra_la_api(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "RELEASE_DATE_DB", tmp_path / "no_existe.db")
    cliente = ClienteFalso(id_por_isrc={"ISRC_A": "trackX"}, fecha_por_id={"trackX": "2020-01-01"})

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["ISRC_A"]), cliente=cliente)

    assert list(fechas) == ["2020-01-01"]
    assert cliente.llamadas_fecha == [["trackX"]]


def test_db_externa_corrupta_o_sin_la_tabla_no_tumba_el_reporte(tmp_path, monkeypatch):
    rota = tmp_path / "rota.db"
    rota.write_text("esto no es una base de datos sqlite")
    monkeypatch.setattr(config, "RELEASE_DATE_DB", rota)
    cliente = ClienteFalso(id_por_isrc={"ISRC_A": "trackX"}, fecha_por_id={"trackX": "2020-01-01"})

    fechas = spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["ISRC_A"]), cliente=cliente)

    assert list(fechas) == ["2020-01-01"]  # cae a la API sin quejarse


def test_no_le_escribe_nada_a_la_db_externa(tmp_path, monkeypatch):
    # Es un archivo del usuario: se abre en modo solo lectura.
    ruta = _crear_db_externa(tmp_path, [("trackX", "2018-04-13")])
    antes = ruta.read_bytes()
    monkeypatch.setattr(config, "RELEASE_DATE_DB", ruta)
    cliente = ClienteFalso(id_por_isrc={"ISRC_A": "trackX", "ISRC_B": "trackY"},
                           fecha_por_id={"trackY": "2024-09-09"})
    spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["ISRC_A", "ISRC_B"]), cliente=cliente)
    assert ruta.read_bytes() == antes


# --- fechas parciales y plan B por nombre (revisión del 21/09/2026) ---

import datetime


@pytest.mark.parametrize("valor, esperado", [
    ("2020-03-15", datetime.date(2020, 3, 15)),
    ("2006-03", datetime.date(2006, 3, 1)),            # Spotify sin día
    ("2000", datetime.date(2000, 1, 1)),                # Spotify solo con año
    ("2006-01-01 00:00:00", datetime.date(2006, 1, 1)), # como lo devuelve SQLite
    (pd.Timestamp("2019-06-28"), datetime.date(2019, 6, 28)),
    (None, None),
    (float("nan"), None),
    ("sin fecha", None),
    ("2021-02-30", None),                               # fecha imposible
])
def test_a_fecha_completa_las_fechas_parciales_como_el_informe_oficial(valor, esperado):
    # El oficial muestra una fecha de la que solo se sabe el año como el 1
    # de enero ("El Teléfono" figura como 1-ene-06).
    assert spotify_release_dates.a_fecha(valor) == esperado


def _item(track_id, nombre, artistas, fecha):
    return {"id": track_id, "name": nombre,
            "artists": [{"name": a} for a in artistas],
            "album": {"release_date": fecha}}


def test_elegir_por_nombre_encuentra_la_cancion_exacta():
    items = [_item("t1", "Cuando Me Enamoro", ["Enrique Iglesias", "Juan Luis Guerra"], "2010-04-19")]
    assert spotify_release_dates.elegir_por_nombre(items, "Cuando Me Enamoro", "Enrique Iglesias") == "t1"


def test_elegir_por_nombre_ignora_remasterizado_y_tildes():
    items = [_item("t1", "Mi Historia Entre Tus Dedos - Remasterizado 2016",
                   ["Gianluca Grignani"], "1995-01-01")]
    assert spotify_release_dates.elegir_por_nombre(
        items, "Mi Historia Entre Tus Dedos - Remasterizado", "Gianluca Grignani") == "t1"
    items = [_item("t2", "Para No Verte Mas", ["La Mosca Tse-Tse"], "2000-01-01")]
    assert spotify_release_dates.elegir_por_nombre(
        items, "Para No Verte Más", "La Mosca Tse-Tse") == "t2"


def test_elegir_por_nombre_no_confunde_un_remix_ni_un_en_vivo_con_la_original():
    # Tienen otra fecha: mejor dejar la celda vacía que poner la equivocada.
    items = [_item("r", "Hawái - Remix", ["Maluma", "The Weeknd"], "2020-11-05"),
             _item("v", "Hawái - En Vivo", ["Maluma"], "2021-06-01")]
    assert spotify_release_dates.elegir_por_nombre(items, "Hawái", "Maluma") is None


def test_elegir_por_nombre_exige_el_mismo_artista():
    items = [_item("otro", "Cuando Me Enamoro", ["Otro Artista"], "2015-01-01")]
    assert spotify_release_dates.elegir_por_nombre(items, "Cuando Me Enamoro", "Enrique Iglesias") is None


def test_elegir_por_nombre_prefiere_el_album_original_al_recopilatorio():
    items = [_item("recop", "Rosa Pastel", ["Belanova"], "2012-05-01"),
             _item("orig", "Rosa Pastel", ["Belanova"], "2003-09-23")]
    assert spotify_release_dates.elegir_por_nombre(items, "Rosa Pastel", "Belanova") == "orig"


def test_elegir_por_nombre_acepta_duos_y_colaboraciones():
    items = [_item("d", "Rakata", ["Wisin & Yandel"], "2005-01-01")]
    assert spotify_release_dates.elegir_por_nombre(items, "Rakata", "Wisin & Yandel") == "d"
    items = [_item("c", "DAKITI", ["Bad Bunny", "Jhay Cortez"], "2020-10-30")]
    assert spotify_release_dates.elegir_por_nombre(items, "DAKITI", "Bad Bunny, Jhay Cortez") == "c"


class ClienteConPlanB(ClienteFalso):
    """Como ClienteFalso, pero además sabe buscar por nombre."""

    def __init__(self, id_por_isrc, fecha_por_id, id_por_nombre):
        super().__init__(id_por_isrc, fecha_por_id)
        self.id_por_nombre = id_por_nombre
        self.llamadas_nombre = []

    def buscar_track_id_por_nombre(self, titulo, artista):
        self.llamadas_nombre.append((titulo, artista))
        return self.id_por_nombre.get((titulo, artista))


def test_si_el_isrc_no_aparece_busca_por_nombre():
    # Caso real: clásicos de catálogo que salían sin fecha y había que
    # completar a mano.
    cliente = ClienteConPlanB(
        id_por_isrc={},                                          # el ISRC no aparece
        fecha_por_id={"t9": "2010-04-19"},
        id_por_nombre={("Cuando Me Enamoro", "Enrique Iglesias"): "t9"},
    )
    fechas = spotify_release_dates.resolver_fechas_lanzamiento(
        pd.Series(["USUM71000001"]), cliente=cliente,
        nombres={"USUM71000001": ("Cuando Me Enamoro", "Enrique Iglesias")},
    )
    assert list(fechas) == ["2010-04-19"]


def test_lo_que_se_encuentra_por_nombre_queda_en_la_cache():
    nombres = {"ISRCX": ("Rosa Pastel", "Belanova")}
    cliente = ClienteConPlanB({}, {"t1": "2003-09-23"}, {("Rosa Pastel", "Belanova"): "t1"})
    spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["ISRCX"]), cliente=cliente, nombres=nombres)

    # Segunda corrida: ya no hace falta buscar ni por ISRC ni por nombre.
    otro = ClienteConPlanB({}, {"t1": "2003-09-23"}, {})
    fechas = spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["ISRCX"]), cliente=otro, nombres=nombres)
    assert list(fechas) == ["2003-09-23"]
    assert otro.llamadas_busqueda == [] and otro.llamadas_nombre == []


def test_los_no_encontrados_de_antes_se_reintentan_por_nombre():
    # Un ISRC que ya había quedado cacheado como "no encontrado" (antes de
    # existir el plan B) no se quedaba vacío para siempre.
    sin_plan_b = ClienteFalso(id_por_isrc={}, fecha_por_id={})
    spotify_release_dates.resolver_fechas_lanzamiento(pd.Series(["VIEJO1"]), cliente=sin_plan_b)

    cliente = ClienteConPlanB({}, {"t7": "2000-01-01"}, {("Para No Verte Más", "La Mosca Tse-Tse"): "t7"})
    fechas = spotify_release_dates.resolver_fechas_lanzamiento(
        pd.Series(["VIEJO1"]), cliente=cliente,
        nombres={"VIEJO1": ("Para No Verte Más", "La Mosca Tse-Tse")},
    )
    assert list(fechas) == ["2000-01-01"]
    assert cliente.llamadas_busqueda == []           # el ISRC no se vuelve a buscar


def test_un_cliente_sin_plan_b_sigue_funcionando():
    # Compatibilidad: un cliente que no sabe buscar por nombre no rompe nada.
    cliente = ClienteFalso(id_por_isrc={}, fecha_por_id={})
    fechas = spotify_release_dates.resolver_fechas_lanzamiento(
        pd.Series(["ISRCZ"]), cliente=cliente, nombres={"ISRCZ": ("X", "Y")})
    assert list(fechas) == [None]
