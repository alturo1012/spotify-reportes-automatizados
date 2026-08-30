"""Persistencia histórica: acumula semana a semana lo necesario para calcular
YTD (Market Share) y la serie histórica completa (Chart Semanal - Resumen
Total).

Dos tablas, ambas ya agregadas (no fila-por-fila / track-por-track):

- `chart_band_weekly`: conteo de tracks Universal por banda/país/semana.
  Alimenta "Resumen Total" del Reporte_Chart_Top_Semanal.
- `ms_label_weekly`: streams Top 200 por sello/país/semana.
  Alimenta el "% Market Share" del Reporte_MS_TOP200.

Ambas se siembran UNA VEZ con `seed_historico()` a partir de los CSV ya
extraídos de los reportes/plantillas reales (ver `extraer_historico.py` en
el Project — no forma parte del repo, es un script de un solo uso), y luego
se extienden semana a semana con `append_semana_chart()` /
`append_semana_ms()` usando el mismo método de cálculo, para que el
historial quede continuo entre lo viejo y lo nuevo.

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
    return conn


def seed_historico(chart_csv: Path = SEED_CHART_CSV, ms_csv: Path = SEED_MS_CSV) -> None:
    """Carga UNA VEZ el histórico ya extraído de los reportes/plantillas
    reales. Es seguro correrlo más de una vez: usa INSERT OR IGNORE, así que
    no duplica filas si ya estaban cargadas.
    """
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
    mes = fecha.strftime("%B").upper()

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


def query_ytd_ms(anio: int, hasta_semana: int) -> pd.DataFrame:
    """Filas de `ms_label_weekly` para un año, hasta cierta semana
    (inclusive) — el rango que necesita el cálculo YTD de Market Share
    (mismo número de semanas comparado entre año actual y anterior).
    """
    df = cargar_ms_label_weekly()
    return df[(df["anio"] == anio) & (df["semana"] <= hasta_semana)]
