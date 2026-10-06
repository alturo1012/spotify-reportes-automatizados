"""Resuelve la fecha de lanzamiento de cada canción vía la API de Spotify, a
partir del ISRC que ya trae la fuente BQ (columna "ISRC", ver
config.SOURCE_COLUMNS) -- usado para la columna "Fecha de Lanzamiento" del
listado de canciones de "Resumen Total" (ver chart_semanal.py), que se
había dejado pendiente a propósito por esto mismo.

Adaptado a partir de 3 archivos que el usuario ya tenía de otro proyecto
(spotify_api.py, release_date_management.py, y el patrón de búsqueda por
ISRC de PlayListAPI.py/linkingISRCWithURI.py -- ver la nota en
claude/plan_fusion_paso_a_paso.md sobre por qué no se integraron esos 3
archivos completos). Cambios respecto al código original:

- Credenciales desde variables de entorno (SPOTIFY_CLIENT_ID /
  SPOTIFY_CLIENT_SECRET, vía un archivo .env -- ver .env.example) en vez de
  texto plano en el código.
- Caché en la MISMA base SQLite del proyecto (universal_data.db, ver
  history.conectar()) en vez de un archivo aparte -- dos tablas nuevas:
    - `spotify_isrc_cache` (isrc -> track_id): resuelto con una búsqueda
      por ISRC (1 llamada a la API por canción nueva -- Spotify no ofrece
      búsqueda por ISRC en lote).
    - `spotify_release_date_cache` (track_id -> release_date): resuelto en
      lotes de hasta 50 vía `sp.tracks(ids)`.
  Ambas cachés son permanentes: una vez resuelta una canción no se vuelve a
  consultar la API por ella, así que solo la primera vez que aparece una
  canción nueva cuesta una llamada -- las semanas siguientes son rápidas
  porque la mayoría ya está en caché (las canciones se repiten semana a
  semana).
- Antes de preguntarle a la API por una fecha, se consulta la base externa
  `data/release_date.db` del proyecto anterior del usuario (opcional, ver
  config.RELEASE_DATE_DB y _fechas_desde_db_externa). El orden de búsqueda
  de una fecha queda así: caché propia -> release_date.db -> API.

  OJO con el alcance: esa base está indexada por URI de Spotify, no por
  ISRC, así que solo ahorra llamadas en el paso "track_id -> fecha" (el
  barato, que va en lotes de 50). El paso "ISRC -> track_id" (el caro, una
  llamada por canción nueva) sigue yendo a la API, porque no existe ninguna
  tabla que relacione ISRC con URI.
- Reintento automático ante 429 (demasiadas solicitudes), respetando el
  header Retry-After.
- Si no hay credenciales configuradas, o la API falla por cualquier motivo,
  esto NO debe tumbar la generación del reporte completo -- ver
  chart_semanal.agregar_fecha_lanzamiento, que atrapa cualquier excepción
  de este módulo y deja la columna en blanco para esa corrida en vez de
  fallar. La próxima corrida lo vuelve a intentar.
"""
import datetime
import os
import re
import sqlite3
import time
import unicodedata
from pathlib import Path

import pandas as pd

from . import config, history

try:
    from dotenv import load_dotenv
    # OJO con el .exe (PyInstaller): `load_dotenv()` sin argumentos busca el
    # .env relativo al código, y en el .exe el código vive en una carpeta
    # temporal (_MEIxxxx) que se borra al cerrar -- ahí nunca está el .env
    # del usuario, que queda junto al ejecutable. Por eso se apunta primero
    # y explícitamente a config.ROOT_DIR, que ya resuelve "la carpeta del
    # .exe" cuando está empaquetado (ver config._calcular_root_dir; es el
    # mismo problema que ya se había corregido en el Paso 7 para la base del
    # histórico). Bug real: con el .exe la columna "Fecha Lzto" salía vacía
    # porque las credenciales nunca se llegaban a cargar.
    load_dotenv(config.ROOT_DIR / ".env")
    # Además el comportamiento normal (repo / carpeta actual), para cuando
    # se corre con `python -m src.main`.
    load_dotenv()
except ImportError:  # python-dotenv es opcional -- si no está, se sigue
    pass            # confiando en variables de entorno ya exportadas.


def _texto_o_none(valor):
    """Devuelve el valor como texto limpio, o None si no es utilizable
    (None, NaN, cadena vacía, "nan"/"none" en cualquier combinación).

    BUG REAL que arregla (el .exe generaba el reporte con la columna
    "Fecha Lzto" vacía y el aviso decía: "'<' not supported between
    instances of 'float' and 'str'"):

    Cuando Spotify no encuentra un ISRC, se guarda en la caché
    `spotify_isrc_cache` con `track_id = NULL` -- correcto, así no se
    vuelve a consultar por él. Pero al leer esa caché de vuelta, pandas
    convierte ese NULL en `NaN`, que es un **float**. Y `if tid:` NO lo
    filtra, porque en Python `bool(float("nan"))` es `True`. Ese NaN se
    colaba en el `sorted({...})` de track_ids junto a los ids de texto y
    reventaba al comparar float con str.

    O sea: la PRIMERA corrida guardaba en caché los ISRC no encontrados, y
    a partir de la segunda TODAS fallaban. Por eso no se veía en las
    pruebas (caché siempre vacía) ni en la primera corrida.

    Se aplica también a los ISRC de entrada: en otras semanas esa columna
    puede venir con celdas vacías o numéricas (pandas la deja como `object`
    con floats adentro), que romperían igual el `sorted(...)`.
    """
    if valor is None:
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nan", "none", "nat"}:
        return None
    return texto


def _fechas_desde_db_externa(track_ids: list) -> dict:
    """Busca las fechas de lanzamiento de esos track_ids en la base externa
    `release_date.db` (ver config.RELEASE_DATE_DB) y devuelve
    {track_id: fecha} solo con los que encontró.

    Esa base viene del proyecto anterior del usuario y tiene la tabla
    `track (uri, release_date)`. OJO con lo que NO es: está indexada por
    URI de Spotify, no por ISRC, así que solo sirve para el segundo paso
    (track_id -> fecha). El primero (ISRC -> track_id) sigue necesitando la
    API, porque no hay ninguna tabla que relacione ISRC con URI.

    Es opcional y nunca puede tumbar el reporte: si el archivo no existe, si
    no tiene la tabla `track`, o si está corrupto, devuelve {} y el flujo
    sigue contra la API como siempre.

    Se abre en modo SOLO LECTURA a propósito (`mode=ro`): es un archivo del
    usuario y este proyecto no le escribe nada.
    """
    ruta = Path(config.RELEASE_DATE_DB)
    if not track_ids or not ruta.exists():
        return {}

    # Los ids pueden estar guardados pelados ("4cOdK2wGLET...") o con el
    # prefijo completo ("spotify:track:4cOdK2wGLET..."). release_date_management.py
    # los guarda pelados (recorta con uri[14:]), pero se buscan las dos
    # formas por si alguna fila quedó con el prefijo.
    formas = {}
    for tid in track_ids:
        formas[tid] = tid
        formas[f"spotify:track:{tid}"] = tid

    encontradas = {}
    try:
        conn = sqlite3.connect(f"file:{ruta}?mode=ro", uri=True)
    except sqlite3.Error:
        return {}
    try:
        claves = list(formas)
        # SQLite limita la cantidad de parámetros por consulta (~999).
        for i in range(0, len(claves), 500):
            lote = claves[i:i + 500]
            marcador = ",".join("?" * len(lote))
            filas = conn.execute(
                f"SELECT uri, release_date FROM track WHERE uri IN ({marcador})", lote
            ).fetchall()
            for uri, fecha in filas:
                fecha = _texto_o_none(fecha)
                if fecha is not None:
                    encontradas[formas[uri]] = fecha
    except sqlite3.Error as e:
        print(
            f"Aviso: no se pudo leer {ruta.name} ({e}). Se resuelven las fechas "
            "contra la API de Spotify, como antes."
        )
        return {}
    finally:
        conn.close()
    return encontradas


def _crear_tablas(conn) -> None:
    conn.execute(
        """CREATE TABLE IF NOT EXISTS spotify_isrc_cache (
            isrc TEXT PRIMARY KEY,
            track_id TEXT
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS spotify_release_date_cache (
            track_id TEXT PRIMARY KEY,
            release_date TEXT
        )"""
    )


_FECHA_INICIAL = re.compile(r"^\s*(\d{4})(?:-(\d{1,2}))?(?:-(\d{1,2}))?")


def a_fecha(valor):
    """Convierte lo que venga como fecha de lanzamiento en un `date` de
    verdad, o None si no hay nada utilizable.

    Spotify a veces no sabe el día, o ni siquiera el mes
    (`release_date_precision`), y devuelve "2006" o "2006-03". El informe
    oficial muestra esas fechas como el primer día del período -- por
    ejemplo "El Teléfono / Héctor 'El Father'" figura como 1-ene-06 --, así
    que se completa igual: año solo -> 1 de enero; año-mes -> día 1.

    Acepta también fechas ya armadas (datetime, Timestamp) y el formato con
    hora que devuelve SQLite ("2006-01-01 00:00:00").
    """
    if valor is None:
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime.datetime)):
        return valor.date()
    if isinstance(valor, datetime.date):
        return valor
    coincidencia = _FECHA_INICIAL.match(str(valor))
    if not coincidencia:
        return None
    anio, mes, dia = coincidencia.groups()
    try:
        return datetime.date(int(anio), int(mes or 1), int(dia or 1))
    except ValueError:
        return None


# --- Búsqueda por nombre, cuando el ISRC no aparece en Spotify ------------

_ENTRE_PARENTESIS = re.compile(r"[\(\[][^\)\]]*[\)\]]")
# Palabras que describen la EDICIÓN y no la canción: con o sin ellas es la
# misma grabación, así que no deben impedir la coincidencia. Ojo: NO están
# "remix", "en vivo", "live", "acoustic" -- esas sí son otra grabación, con
# otra fecha.
_PALABRAS_DE_EDICION = re.compile(
    r"\b(remaster(ed|izado|izada)?|remasterizad[oa]|version|versi[oó]n|\d{4})\b"
)


def _normalizar(texto) -> str:
    """Minúsculas, sin tildes, sin lo que va entre paréntesis ni signos."""
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c)).lower()
    texto = _ENTRE_PARENTESIS.sub(" ", texto)
    texto = re.sub(r"[^a-z0-9ñ ]+", " ", texto)
    return " ".join(texto.split())


def _titulo_comparable(titulo) -> str:
    return " ".join(_PALABRAS_DE_EDICION.sub(" ", _normalizar(titulo)).split())


def _artista_principal(artista) -> str:
    """El primer artista de "A, B", "A & B", "A feat. B", "A x B"."""
    if not isinstance(artista, str):
        return ""
    primero = re.split(r",|&| feat\.? | ft\.? | x ", artista, maxsplit=1, flags=re.IGNORECASE)[0]
    return _normalizar(primero)


def _coincide_artista(artistas_resultado, artista) -> bool:
    """¿Alguno de los artistas del resultado es el artista buscado?

    Vale si coincide con el artista principal, con el texto completo (los
    dúos como "Wisin & Yandel" figuran como UN artista en Spotify), o si
    aparece como palabra completa dentro del texto de la fuente ("Bad Bunny,
    Jhay Cortez" contiene "jhay cortez"). El largo mínimo evita que un
    nombre de dos letras coincida con cualquier cosa.
    """
    completo = _normalizar(artista)
    principal = _artista_principal(artista)
    for nombre in artistas_resultado:
        n = _normalizar(nombre)
        if not n:
            continue
        if n == principal or n == completo:
            return True
        if len(n) >= 4 and f" {n} " in f" {completo} ":
            return True
    return False


def elegir_por_nombre(items: list, titulo, artista):
    """De los resultados de una búsqueda de Spotify por nombre, el track_id
    que corresponde a (titulo, artista), o None si ninguno coincide.

    Es estricto a propósito: el nombre tiene que ser el mismo (sin contar
    paréntesis, tildes ni "Remasterizado 2016") y el artista principal tiene
    que estar entre los artistas del resultado. Una fecha equivocada es
    peor que una celda vacía -- una versión en vivo o un remix tienen otra
    fecha, y con este criterio no se confunden con la original.

    Si coinciden varios (la misma canción en el álbum original y en un
    recopilatorio posterior), se queda con el de álbum MÁS ANTIGUO: la
    fecha de lanzamiento es la del original.
    """
    buscado = _titulo_comparable(titulo)
    if not buscado or not _normalizar(artista):
        return None
    candidatos = []
    for item in items or []:
        if not item or _titulo_comparable(item.get("name")) != buscado:
            continue
        nombres_artistas = [a.get("name") for a in item.get("artists") or []]
        if not _coincide_artista(nombres_artistas, artista):
            continue
        fecha = a_fecha((item.get("album") or {}).get("release_date"))
        candidatos.append((fecha or datetime.date.max, item.get("id")))
    if not candidatos:
        return None
    candidatos.sort(key=lambda par: par[0])
    return candidatos[0][1]


class SpotifyReleaseDateClient:
    """Envuelve spotipy (búsqueda por ISRC + fechas de lanzamiento en
    lote), con reintento automático ante 429. Requiere SPOTIFY_CLIENT_ID y
    SPOTIFY_CLIENT_SECRET en el entorno (ver .env.example)."""

    def __init__(self, client_id: str = None, client_secret: str = None):
        client_id = client_id or os.environ.get("SPOTIFY_CLIENT_ID")
        client_secret = client_secret or os.environ.get("SPOTIFY_CLIENT_SECRET")
        if not client_id or not client_secret:
            raise RuntimeError(
                "Faltan las credenciales de Spotify -- definí SPOTIFY_CLIENT_ID y "
                "SPOTIFY_CLIENT_SECRET (ver .env.example) para poder resolver fechas "
                "de lanzamiento."
            )
        import spotipy
        from spotipy.oauth2 import SpotifyClientCredentials

        self._spotipy = spotipy
        self.sp = spotipy.Spotify(
            auth_manager=SpotifyClientCredentials(client_id, client_secret)
        )

    def _con_reintento(self, func, *args, **kwargs):
        intentos = 0
        while True:
            try:
                return func(*args, **kwargs)
            except self._spotipy.SpotifyException as e:
                if e.http_status == 429 and intentos < 5:
                    espera = 5
                    if getattr(e, "headers", None):
                        espera = int(e.headers.get("Retry-After", 5))
                    time.sleep(espera + 1)
                    intentos += 1
                    continue
                raise

    def buscar_track_id_por_isrc(self, isrc: str):
        """Un ISRC -> un track_id de Spotify, o None si no se encontró.
        1 llamada por ISRC (Spotify no tiene búsqueda por ISRC en lote)."""
        resultado = self._con_reintento(self.sp.search, f"isrc:{isrc}", type="track", limit=1)
        items = resultado.get("tracks", {}).get("items", [])
        return items[0]["id"] if items else None

    def buscar_track_id_por_nombre(self, titulo, artista):
        """Plan B cuando el ISRC no aparece: busca por nombre y artista y
        se queda con el resultado que coincide de verdad (ver
        elegir_por_nombre). None si no hay ninguno confiable."""
        titulo_limpio = " ".join(_ENTRE_PARENTESIS.sub(" ", str(titulo or "")).split())
        principal = re.split(r",|&| feat\.? | ft\.? | x ", str(artista or ""), maxsplit=1,
                             flags=re.IGNORECASE)[0].strip()
        if not titulo_limpio or not principal:
            return None
        consulta = f'track:"{titulo_limpio}" artist:"{principal}"'
        resultado = self._con_reintento(self.sp.search, consulta, type="track", limit=10)
        items = resultado.get("tracks", {}).get("items", [])
        return elegir_por_nombre(items, titulo, artista)

    def fechas_de_lanzamiento(self, track_ids: list) -> dict:
        """Lote de hasta 50 track_ids -> {track_id: release_date}."""
        fechas = {}
        for i in range(0, len(track_ids), 50):
            lote = track_ids[i:i + 50]
            tracks = self._con_reintento(self.sp.tracks, lote)["tracks"]
            for track_id, track in zip(lote, tracks):
                fechas[track_id] = track["album"]["release_date"] if track else None
        return fechas


def resolver_fechas_lanzamiento(isrcs: pd.Series, cliente: SpotifyReleaseDateClient = None,
                                nombres: dict = None) -> pd.Series:
    """Dada una Series de códigos ISRC, devuelve una Series alineada (mismo
    índice) con la fecha de lanzamiento de cada uno ("YYYY-MM-DD", o a veces
    solo "YYYY"/"YYYY-MM" si Spotify no tiene el día/mes exacto -- se deja
    tal cual viene, sin forzar un formato de fecha). None donde no se pudo
    resolver (ISRC vacío, no encontrado en Spotify, o track sin álbum).

    Usa una caché permanente en universal_data.db (ver _crear_tablas) --
    solo consulta la API de Spotify por los ISRC que todavía no había visto.
    `cliente` es opcional: si no se pasa, se crea un SpotifyReleaseDateClient
    real (credenciales desde el entorno) la primera vez que hace falta --
    útil para inyectar un doble de prueba sin tocar la API de verdad.
    """
    # Normalizados a texto (o None) -- ver _texto_o_none: la columna puede
    # traer vacíos o valores numéricos según la semana, y mezclarlos rompía
    # el sorted(...) de abajo.
    #
    # Se arma como lista y se fuerza dtype=object a propósito, en vez de
    # usar `isrcs.map(...)`: con el dtype de texto nuevo de pandas, `.map`
    # infiere el tipo del resultado y vuelve a convertir los None en NaN
    # (float), que es justo lo que estamos tratando de evitar acá. Forzando
    # object, los None se quedan como None en cualquier versión de pandas.
    valores_norm = [_texto_o_none(valor) for valor in isrcs]
    isrcs_norm = pd.Series(valores_norm, index=isrcs.index, dtype=object)
    isrcs_unicos = sorted({isrc for isrc in valores_norm if isrc is not None})
    if not isrcs_unicos:
        return pd.Series([None] * len(isrcs), index=isrcs.index)

    conn = history.conectar()
    try:
        _crear_tablas(conn)

        marcador = ",".join("?" * len(isrcs_unicos))
        cache_isrc = pd.read_sql_query(
            f"SELECT isrc, track_id FROM spotify_isrc_cache WHERE isrc IN ({marcador})",
            conn, params=isrcs_unicos,
        )
        # _texto_o_none en el track_id: un ISRC ya buscado y NO encontrado
        # queda cacheado con track_id NULL, que pandas devuelve como NaN
        # (ver _texto_o_none). Se guarda como None para poder distinguir
        # "ya lo busqué y no está" (está en el dict, con valor None) de
        # "todavía no lo busqué" (no está en el dict) -- así no se vuelve a
        # gastar una llamada a la API por él en cada corrida.
        track_id_por_isrc = {
            isrc: _texto_o_none(track_id)
            for isrc, track_id in zip(cache_isrc["isrc"], cache_isrc["track_id"])
        }
        isrcs_faltantes = [isrc for isrc in isrcs_unicos if isrc not in track_id_por_isrc]

        if isrcs_faltantes:
            cliente = cliente or SpotifyReleaseDateClient()
            nuevas_filas = []
            for isrc in isrcs_faltantes:
                track_id = _texto_o_none(cliente.buscar_track_id_por_isrc(isrc))
                track_id_por_isrc[isrc] = track_id
                nuevas_filas.append((isrc, track_id))
            conn.executemany(
                "INSERT OR REPLACE INTO spotify_isrc_cache (isrc, track_id) VALUES (?, ?)",
                nuevas_filas,
            )
            conn.commit()

        # Plan B para los ISRC que Spotify no encuentra: buscar por nombre y
        # artista (ver elegir_por_nombre). Es el caso de varios clásicos de
        # catálogo -- "Para No Verte Más", "Cuando Me Enamoro" -- que salían
        # sin fecha y había que completar a mano.
        #
        # Incluye los que YA estaban cacheados como "no encontrado": esos se
        # reintentan por nombre en cada corrida hasta que aparezcan (son muy
        # pocos, un puñado de llamadas). Si aparecen, se guarda el track_id
        # en la caché y ya no se vuelven a buscar.
        sin_track = [
            isrc for isrc in isrcs_unicos
            if track_id_por_isrc.get(isrc) is None and nombres and isrc in nombres
        ]
        if sin_track:
            if cliente is None:
                try:
                    cliente = SpotifyReleaseDateClient()
                except RuntimeError:
                    # Sin credenciales no hay plan B, pero lo ya resuelto
                    # tiene que seguir saliendo: no se aborta nada.
                    cliente = None
            buscar = getattr(cliente, "buscar_track_id_por_nombre", None)
            if buscar is not None:
                recuperados = []
                for isrc in sin_track:
                    titulo, artista = nombres[isrc]
                    track_id = _texto_o_none(buscar(titulo, artista))
                    if track_id is not None:
                        track_id_por_isrc[isrc] = track_id
                        recuperados.append((isrc, track_id))
                if recuperados:
                    conn.executemany(
                        "INSERT OR REPLACE INTO spotify_isrc_cache (isrc, track_id) VALUES (?, ?)",
                        recuperados,
                    )
                    conn.commit()

        # OJO: el filtro tiene que ser `is not None`, no `if tid` -- ver
        # _texto_o_none: un NaN pasaría el `if tid` (bool(nan) es True) y
        # rompería este sorted() al comparar float con str.
        track_ids_unicos = sorted({tid for tid in track_id_por_isrc.values() if tid is not None})
        fecha_por_track_id = {}
        if track_ids_unicos:
            marcador_tid = ",".join("?" * len(track_ids_unicos))
            cache_fecha = pd.read_sql_query(
                f"SELECT track_id, release_date FROM spotify_release_date_cache "
                f"WHERE track_id IN ({marcador_tid})",
                conn, params=track_ids_unicos,
            )
            # Mismo cuidado que con los track_id: una fecha cacheada como
            # NULL (track sin álbum) vuelve como NaN y no debe escribirse
            # como "nan" en la celda del Excel.
            fecha_por_track_id = {
                tid: _texto_o_none(fecha)
                for tid, fecha in zip(cache_fecha["track_id"], cache_fecha["release_date"])
            }

        track_ids_faltantes = [tid for tid in track_ids_unicos if tid not in fecha_por_track_id]

        # Antes de gastar una llamada a la API, se mira la base externa
        # release_date.db del proyecto anterior (ver _fechas_desde_db_externa).
        # Lo que salga de ahí se copia a nuestra caché, así queda resuelto
        # para siempre aunque después el archivo externo no esté.
        if track_ids_faltantes:
            desde_db_externa = _fechas_desde_db_externa(track_ids_faltantes)
            if desde_db_externa:
                fecha_por_track_id.update(desde_db_externa)
                conn.executemany(
                    "INSERT OR REPLACE INTO spotify_release_date_cache (track_id, release_date) VALUES (?, ?)",
                    list(desde_db_externa.items()),
                )
                conn.commit()
                track_ids_faltantes = [
                    tid for tid in track_ids_faltantes if tid not in desde_db_externa
                ]

        # Solo lo que no estaba ni en la caché ni en la base externa va a la API.
        if track_ids_faltantes:
            cliente = cliente or SpotifyReleaseDateClient()
            nuevas_fechas = cliente.fechas_de_lanzamiento(track_ids_faltantes)
            fecha_por_track_id.update(nuevas_fechas)
            conn.executemany(
                "INSERT OR REPLACE INTO spotify_release_date_cache (track_id, release_date) VALUES (?, ?)",
                list(nuevas_fechas.items()),
            )
            conn.commit()
    finally:
        conn.close()

    fecha_por_isrc = {
        isrc: fecha_por_track_id.get(track_id_por_isrc.get(isrc))
        for isrc in isrcs_unicos
    }
    # Se mapea sobre los ISRC ya normalizados (no sobre los originales),
    # para que las claves coincidan aunque la fuente los traiga con
    # espacios de más o como número. dtype=object por el mismo motivo de
    # arriba: que los None sigan siendo None y no se conviertan en NaN.
    return pd.Series(
        [fecha_por_isrc.get(isrc) if isrc is not None else None for isrc in isrcs_norm],
        index=isrcs.index, dtype=object,
    )
