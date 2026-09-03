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
- Reintento automático ante 429 (demasiadas solicitudes), respetando el
  header Retry-After.
- Si no hay credenciales configuradas, o la API falla por cualquier motivo,
  esto NO debe tumbar la generación del reporte completo -- ver
  chart_semanal.agregar_fecha_lanzamiento, que atrapa cualquier excepción
  de este módulo y deja la columna en blanco para esa corrida en vez de
  fallar. La próxima corrida lo vuelve a intentar.
"""
import os
import time

import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:  # python-dotenv es opcional -- si no está, se sigue
    pass            # confiando en variables de entorno ya exportadas.

from . import history


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

    def fechas_de_lanzamiento(self, track_ids: list) -> dict:
        """Lote de hasta 50 track_ids -> {track_id: release_date}."""
        fechas = {}
        for i in range(0, len(track_ids), 50):
            lote = track_ids[i:i + 50]
            tracks = self._con_reintento(self.sp.tracks, lote)["tracks"]
            for track_id, track in zip(lote, tracks):
                fechas[track_id] = track["album"]["release_date"] if track else None
        return fechas


def resolver_fechas_lanzamiento(isrcs: pd.Series, cliente: SpotifyReleaseDateClient = None) -> pd.Series:
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
    isrcs_unicos = sorted(isrcs.dropna().unique())
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
        track_id_por_isrc = dict(zip(cache_isrc["isrc"], cache_isrc["track_id"]))
        isrcs_faltantes = [isrc for isrc in isrcs_unicos if isrc not in track_id_por_isrc]

        if isrcs_faltantes:
            cliente = cliente or SpotifyReleaseDateClient()
            nuevas_filas = []
            for isrc in isrcs_faltantes:
                track_id = cliente.buscar_track_id_por_isrc(isrc)
                track_id_por_isrc[isrc] = track_id
                nuevas_filas.append((isrc, track_id))
            conn.executemany(
                "INSERT OR REPLACE INTO spotify_isrc_cache (isrc, track_id) VALUES (?, ?)",
                nuevas_filas,
            )
            conn.commit()

        track_ids_unicos = sorted({tid for tid in track_id_por_isrc.values() if tid})
        fecha_por_track_id = {}
        if track_ids_unicos:
            marcador_tid = ",".join("?" * len(track_ids_unicos))
            cache_fecha = pd.read_sql_query(
                f"SELECT track_id, release_date FROM spotify_release_date_cache "
                f"WHERE track_id IN ({marcador_tid})",
                conn, params=track_ids_unicos,
            )
            fecha_por_track_id = dict(zip(cache_fecha["track_id"], cache_fecha["release_date"]))

        track_ids_faltantes = [tid for tid in track_ids_unicos if tid not in fecha_por_track_id]
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
    return isrcs.map(fecha_por_isrc)
