"""Persistencia histórica: acumula semana a semana lo necesario para calcular
YTD (Market Share) y la serie histórica completa (Chart Semanal - Resumen
Total).

Ver claude/mapeo_logica_plantillas.md y claude/plan_fusion_paso_a_paso.md
(Paso 2) para el diseño completo y por qué se decidió así.

Tres tablas:

- `chart_band_weekly`: conteo de tracks Universal por banda/país/semana
  (agregado). Alimenta "Resumen Total" del Reporte_Chart_Top_Semanal.
- `ms_label_weekly`: streams Top 200 por sello/país/semana (agregado).
  Alimenta el "% Market Share" del Reporte_MS_TOP200.
- `chart_track_weekly`: track por track (posición, artista, canción) por
  país/semana -- a diferencia de las otras dos, NO es agregada. Existe para
  no perder el detalle de "Detalle Tracks"/listado de canciones semana a
  semana (esas pestañas del reporte solo muestran la semana que se acaba de
  cargar, sin histórico -- esta tabla es lo que permite, más adelante,
  consultar qué traía cualquier semana pasada sin tener que guardar el
  Excel de cada semana aparte). Empezó a llenarse a partir de la semana en
  que se agregó (no tiene sembrado retroactivo de semanas anteriores a
  eso -- ver "Nota" en claude/plan_fusion_paso_a_paso.md sobre qué tomaría
  rellenar semanas viejas si hiciera falta).

Las dos primeras se sembraron UNA VEZ con `seed_historico()` a partir de los
CSV ya extraídos de los reportes/plantillas reales (ver `extraer_historico.py`
en el Project — no forma parte del repo, es un script de un solo uso), y las
tres se extienden semana a semana con `append_semana_chart()` /
`append_semana_ms()` / `append_semana_tracks()` usando el mismo método de
cálculo, para que el historial quede continuo entre lo viejo y lo nuevo.

La numeración de "semana" es secuencial por año (1, 2, 3... desde la
primera semana cargada de ese año) — igual que las plantillas originales,
NO es semana ISO. Al hacer append, la próxima semana de un año = la semana
máxima ya guardada de ese año + 1. Esto asume que las semanas se cargan en
orden cronológico, una por una (el flujo normal de uso semanal).
"""
from pathlib import Path
import sqlite3
import pandas as pd

from . import config

DB_PATH = config.ROOT_DIR / "data" / "history" / "universal_data.db"

SEED_DIR = config.ROOT_DIR / "data" / "history" / "seed"
SEED_CHART_CSV = SEED_DIR / "seed_chart_band_weekly.csv"
SEED_MS_CSV = SEED_DIR / "seed_ms_label_weekly.csv"
SEED_MS_BANDAS_CSV = SEED_DIR / "seed_ms_band_label_weekly.csv"


def conectar() -> sqlite3.Connection:
    """Punto de entrada público a la misma base SQLite del proyecto
    (universal_data.db), para que otros módulos (ej. spotify_release_dates.py)
    puedan agregar sus propias tablas de caché ahí en vez de abrir un
    archivo aparte -- todo el histórico y las cachés quedan en un solo
    archivo, más fácil de respaldar/mover."""
    return _conectar()


def _conectar() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS chart_band_weekly (
            anio INTEGER NOT NULL,
            semana INTEGER NOT NULL,
            mes TEXT,
            country_code TEXT NOT NULL,
            banda INTEGER NOT NULL,
            conteo_universal INTEGER NOT NULL,
            PRIMARY KEY (anio, semana, country_code, banda)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS ms_label_weekly (
            anio INTEGER NOT NULL,
            semana INTEGER NOT NULL,
            country_code TEXT NOT NULL,
            label_group TEXT NOT NULL,
            streams_top200 REAL NOT NULL,
            chart_date TEXT,
            PRIMARY KEY (anio, semana, country_code, label_group)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS chart_track_weekly (
            anio INTEGER NOT NULL,
            semana INTEGER NOT NULL,
            mes TEXT,
            chart_date TEXT,
            country_code TEXT NOT NULL,
            position INTEGER NOT NULL,
            artist TEXT,
            song_name TEXT,
            region TEXT,
            stream_count REAL,
            label_group TEXT,
            label_name TEXT,
            PRIMARY KEY (anio, semana, country_code, position)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS ms_band_label_weekly (
            anio INTEGER NOT NULL,
            semana INTEGER NOT NULL,
            chart_date TEXT,
            country_code TEXT NOT NULL,
            banda INTEGER NOT NULL,
            label_group TEXT NOT NULL,
            pct_streams REAL NOT NULL,
            PRIMARY KEY (anio, semana, country_code, banda, label_group)
        )
        """
    )
    return conn


def seed_historico(
    chart_csv: Path = None,
    ms_csv: Path = None,
    ms_bandas_csv: Path = None,
) -> None:
    """Carga UNA VEZ el histórico ya extraído de los reportes/plantillas
    reales. Es seguro correrlo más de una vez: usa INSERT OR IGNORE, así que
    no duplica filas si ya estaban cargadas.

    Cada parámetro en None usa el CSV de siembra que le corresponde
    (SEED_*_CSV). Se resuelven ACÁ ADENTRO y no como valor por defecto de la
    firma a propósito: los valores por defecto de Python se fijan cuando se
    define la función, así que un `monkeypatch.setattr(history, "SEED_...")`
    en los tests no tendría ningún efecto (bug real: los tests seguían
    tragándose el CSV real de 171.850 filas pese al monkeypatch).

    Si el archivo de un CSV no existe, esa tabla se omite y las demás se
    siembran igual. Es lo que permite a los tests sembrar un histórico
    controlado sin arrastrar el histórico real por banda/sello.
    """
    chart_csv = SEED_CHART_CSV if chart_csv is None else chart_csv
    ms_csv = SEED_MS_CSV if ms_csv is None else ms_csv
    ms_bandas_csv = SEED_MS_BANDAS_CSV if ms_bandas_csv is None else ms_bandas_csv

    chart_df = pd.read_csv(chart_csv)
    ms_df = pd.read_csv(ms_csv)

    conn = _conectar()
    try:
        conn.executemany(
            """INSERT OR IGNORE INTO chart_band_weekly
               (anio, semana, mes, country_code, banda, conteo_universal)
               VALUES (?, ?, ?, ?, ?, ?)""",
            chart_df[
                ["anio", "semana", "mes", "country_code", "banda", "conteo_universal"]
            ].itertuples(index=False, name=None),
        )
        conn.executemany(
            """INSERT OR IGNORE INTO ms_label_weekly
               (anio, semana, country_code, label_group, streams_top200, chart_date)
               VALUES (?, ?, ?, ?, ?, ?)""",
            ms_df[
                ["anio", "semana", "country_code", "label_group", "streams_top200", "chart_date"]
            ].itertuples(index=False, name=None),
        )
        if Path(ms_bandas_csv).exists():
            bandas_df = pd.read_csv(ms_bandas_csv)
            conn.executemany(
                """INSERT OR IGNORE INTO ms_band_label_weekly
                   (anio, semana, chart_date, country_code, banda, label_group, pct_streams)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                bandas_df[
                    ["anio", "semana", "chart_date", "country_code", "banda", "label_group", "pct_streams"]
                ].itertuples(index=False, name=None),
            )
        conn.commit()
    finally:
        conn.close()


def _proxima_semana(conn: sqlite3.Connection, tabla: str, anio: int) -> int:
    row = conn.execute(f"SELECT MAX(semana) FROM {tabla} WHERE anio = ?", (anio,)).fetchone()
    maximo = row[0]
    return 1 if maximo is None else maximo + 1


def _validar_una_sola_semana(df_semana: pd.DataFrame) -> pd.Timestamp:
    fechas = df_semana["chart_date"].unique()
    if len(fechas) != 1:
        raise ValueError(
            f"Se esperaba un DataFrame de una sola semana, llegaron {len(fechas)} fechas distintas: {fechas}"
        )
    return pd.Timestamp(fechas[0])


def append_semana_chart(df_semana: pd.DataFrame) -> None:
    """Calcula el conteo de tracks Universal por banda/país para la semana
    que trae `df_semana` (el DataFrame que devuelve `load_data.load_source`,
    ya con `country_code` y `label_group` normalizado) y lo agrega a
    `chart_band_weekly`, continuando la numeración de semana.
    """
    fecha = _validar_una_sola_semana(df_semana)
    anio = fecha.year
    mes = config.MESES_ES[fecha.month]

    conn = _conectar()
    try:
        semana = _proxima_semana(conn, "chart_band_weekly", anio)
        filas = []
        for country_code, grupo_pais in df_semana.groupby("country_code"):
            universal = grupo_pais[grupo_pais["label_group"] == "Universal"]
            for banda in config.BANDAS_CHART:
                conteo = int((universal["position"] <= banda).sum())
                filas.append((anio, semana, mes, country_code, banda, conteo))

        conn.executemany(
            """INSERT OR REPLACE INTO chart_band_weekly
               (anio, semana, mes, country_code, banda, conteo_universal)
               VALUES (?, ?, ?, ?, ?, ?)""",
            filas,
        )
        conn.commit()
    finally:
        conn.close()


# Las plantillas originales (y por lo tanto el histórico sembrado por
# seed_historico) guardan streams en MILLONES, no en el valor absoluto que
# trae `stream_count` de la fuente BQ. Verificado 1:1 contra
# "BASE Informe Sportify charts Semana 24": suma cruda de stream_count
# Universal Top200 Colombia = 12,012,244 vs. valor sembrado = 12.012244 ->
# factor exacto de 1,000,000. Hay que aplicar el mismo factor acá para que
# lo nuevo sea comparable con lo histórico.
FACTOR_ESCALA_STREAMS = 1_000_000


def append_semana_ms(df_semana: pd.DataFrame) -> None:
    """Calcula streams Top 200 por sello/país para la semana que trae
    `df_semana` y lo agrega a `ms_label_weekly`, continuando la numeración
    de semana. Los streams se guardan en millones (ver FACTOR_ESCALA_STREAMS)
    para que sean comparables con el histórico sembrado.
    """
    fecha = _validar_una_sola_semana(df_semana)
    anio = fecha.year
    top200 = df_semana[df_semana["position"] <= 200]

    conn = _conectar()
    try:
        semana = _proxima_semana(conn, "ms_label_weekly", anio)
        filas = []
        for country_code, grupo_pais in top200.groupby("country_code"):
            streams_por_label = grupo_pais.groupby("label_group")["stream_count"].sum()
            for label in config.LABEL_GROUPS_MS:
                streams_millones = float(streams_por_label.get(label, 0.0)) / FACTOR_ESCALA_STREAMS
                filas.append((anio, semana, country_code, label, streams_millones, fecha.date().isoformat()))

        conn.executemany(
            """INSERT OR REPLACE INTO ms_label_weekly
               (anio, semana, country_code, label_group, streams_top200, chart_date)
               VALUES (?, ?, ?, ?, ?, ?)""",
            filas,
        )
        conn.commit()
    finally:
        conn.close()


def append_semana_ms_bandas(df_semana: pd.DataFrame) -> None:
    """Calcula, para la semana que trae `df_semana`, el % de streams de cada
    sello dentro de cada banda (config.BANDAS_MARKET_SHARE) de cada país, y
    lo agrega a `ms_band_label_weekly` -- la tabla que alimenta las pestañas
    individuales de país del Reporte_MS_TOP200.

    Es el mismo cálculo que hacían las sub-tablas "Streams (%) TOP N" de la
    plantilla original, y el mismo que este proyecto venía haciendo al vuelo
    desde `chart_track_weekly`. Se pasó a tabla propia para poder sembrar de
    una vez los años de historia que ya traía el reporte real (2021 en
    adelante), que track por track no se podían reconstruir.
    """
    fecha = _validar_una_sola_semana(df_semana)
    anio = fecha.year
    fecha_str = fecha.date().isoformat()

    df = df_semana.copy()
    if "stream_count" not in df.columns:
        df["stream_count"] = 0.0
    df["stream_count"] = df["stream_count"].fillna(0.0)

    conn = _conectar()
    try:
        semana = _proxima_semana(conn, "ms_band_label_weekly", anio)
        filas = []
        for country_code, grupo_pais in df.groupby("country_code"):
            for banda in config.BANDAS_MARKET_SHARE:
                grupo_banda = grupo_pais[grupo_pais["position"] <= banda]
                total = grupo_banda["stream_count"].sum()
                streams_por_label = grupo_banda.groupby("label_group")["stream_count"].sum()
                for label in config.LABEL_GROUPS_MS:
                    streams_label = float(streams_por_label.get(label, 0.0))
                    pct = streams_label / total if total else 0.0
                    filas.append((anio, semana, fecha_str, country_code, banda, label, pct))

        conn.executemany(
            """INSERT OR REPLACE INTO ms_band_label_weekly
               (anio, semana, chart_date, country_code, banda, label_group, pct_streams)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            filas,
        )
        conn.commit()
    finally:
        conn.close()


def cargar_ms_band_label_weekly() -> pd.DataFrame:
    conn = _conectar()
    try:
        return pd.read_sql_query(
            "SELECT * FROM ms_band_label_weekly "
            "ORDER BY country_code, anio, semana, banda, label_group",
            conn,
        )
    finally:
        conn.close()


def append_semana_tracks(df_semana: pd.DataFrame) -> None:
    """Guarda el detalle track por track (posición, artista, canción, país)
    de la semana que trae `df_semana` en `chart_track_weekly`, continuando
    su propia numeración de semana (independiente de `chart_band_weekly`,
    igual que `ms_label_weekly` -- en la práctica quedan sincronizadas
    porque siempre se llaman juntas en la misma corrida, ver
    chart_semanal.generar_reporte).

    A diferencia de `append_semana_chart` (que agrega/cuenta), acá se guarda
    una fila por cada track de Top 200 tal cual viene de la fuente -- hasta
    200 x 17 países = ~3400 filas por semana.
    """
    fecha = _validar_una_sola_semana(df_semana)
    anio = fecha.year
    mes = config.MESES_ES[fecha.month]
    fecha_str = fecha.date().isoformat()

    df = df_semana.copy()
    for columna_opcional in ("region", "stream_count", "label_group", "label_name"):
        if columna_opcional not in df.columns:
            df[columna_opcional] = None

    top200 = df[df["position"] <= 200]

    conn = _conectar()
    try:
        semana = _proxima_semana(conn, "chart_track_weekly", anio)
        filas = [
            (
                anio, semana, mes, fecha_str, fila.country_code, int(fila.position),
                fila.artist, fila.song_name, fila.region, fila.stream_count,
                fila.label_group, fila.label_name,
            )
            for fila in top200.itertuples(index=False)
        ]

        conn.executemany(
            """INSERT OR REPLACE INTO chart_track_weekly
               (anio, semana, mes, chart_date, country_code, position, artist,
                song_name, region, stream_count, label_group, label_name)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            filas,
        )
        conn.commit()
    finally:
        conn.close()


def cargar_chart_track_weekly() -> pd.DataFrame:
    conn = _conectar()
    try:
        return pd.read_sql_query(
            "SELECT * FROM chart_track_weekly ORDER BY anio, semana, country_code, position", conn
        )
    finally:
        conn.close()


def cargar_tracks_de_semana(anio: int, semana: int) -> pd.DataFrame:
    """Detalle track por track de una semana puntual ya guardada -- para
    reconstruir "Detalle Tracks"/el listado de canciones de una semana
    pasada sin tener que volver a abrir el Excel de ese momento.
    """
    df = cargar_chart_track_weekly()
    return df[(df["anio"] == anio) & (df["semana"] == semana)]


def cargar_chart_band_weekly() -> pd.DataFrame:
    conn = _conectar()
    try:
        return pd.read_sql_query(
            "SELECT * FROM chart_band_weekly ORDER BY anio, semana, country_code, banda", conn
        )
    finally:
        conn.close()


def cargar_ms_label_weekly() -> pd.DataFrame:
    conn = _conectar()
    try:
        return pd.read_sql_query(
            "SELECT * FROM ms_label_weekly ORDER BY anio, semana, country_code, label_group", conn
        )
    finally:
        conn.close()


TABLAS_CON_FECHA = ("ms_label_weekly", "chart_track_weekly", "ms_band_label_weekly")


def semana_ya_cargada(chart_date) -> bool:
    """True si esta fecha ya fue guardada antes en el histórico, mirando la
    columna `chart_date` de TODAS las tablas que la tienen
    (ver TABLAS_CON_FECHA).

    Existe para que `main.py` pueda correrse dos veces por error con el
    mismo archivo fuente (por ejemplo, si el proceso se interrumpió a la
    mitad) sin duplicar la semana: como `_proxima_semana()` siempre calcula
    "la siguiente" sin mirar si la fecha ya estaba, hace falta este chequeo
    aparte antes de llamar a los `append_semana_*`.

    Antes solo miraba `ms_label_weekly`, que alcanzaba cuando las tablas se
    sembraban juntas y avanzaban a la par. Dejó de alcanzar al sembrar
    `ms_band_label_weekly` hasta la semana 33 mientras `ms_label_weekly`
    seguía en la 24: una fecha podía estar guardada en una tabla y no en la
    otra, y volver a cargarla la habría duplicado con otro número de semana.
    """
    fecha_str = pd.Timestamp(chart_date).date().isoformat()
    conn = _conectar()
    try:
        for tabla in TABLAS_CON_FECHA:
            fila = conn.execute(
                f"SELECT 1 FROM {tabla} WHERE chart_date = ? LIMIT 1", (fecha_str,)
            ).fetchone()
            if fila is not None:
                return True
        return False
    finally:
        conn.close()


def query_ytd_ms(anio: int, hasta_semana: int) -> pd.DataFrame:
    """Filas de `ms_label_weekly` para un año, hasta cierta semana
    (inclusive) — el rango que necesita el cálculo YTD de Market Share
    (mismo número de semanas comparado entre año actual y anterior).
    """
    df = cargar_ms_label_weekly()
    return df[(df["anio"] == anio) & (df["semana"] <= hasta_semana)]
