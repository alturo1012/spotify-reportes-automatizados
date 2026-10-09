"""Paquete de instalación: todo lo que necesita otro equipo, en un solo zip.

    python -m scripts.empaquetar_instalacion
    python -m scripts.empaquetar_instalacion --salida "D:/Entregas"

Por qué existe: la siembra que está en GitHub llega solo hasta la semana 33
de 2026 (Spotify), la 35 (BMAT) y julio de 2026 (mensual). Un equipo
instalado desde el repo arranca con el histórico incompleto, y completarlo
obliga a volver a cargar cada semana en orden. Con este zip el otro equipo
arranca con la base al día, tal como está en el equipo donde se generó.

Qué lleva, dentro de una carpeta `ReportesSpotifyLatam/`:
- el ejecutable (hay que haber corrido build.bat antes);
- `data/history/universal_data.db`: una copia consistente de la base actual;
- `data/history/seed/`: la siembra, por si alguna vez hay que reconstruir;
- `data/release_date.db` y `data/bmat_titularidad_compartida.xlsx`;
- `data/output/` y `data/raw/` vacías;
- `.env.example` y `LEEME_INSTALACION.txt` con los pasos y hasta dónde llega
  el histórico.

Qué NO lleva: los respaldos, el registro de errores, las preferencias de
este equipo (rutas que en el otro no existen) y el `.env` con las
credenciales de Spotify, salvo que se pida explícitamente.
"""
import logging
import sqlite3
import tempfile
import zipfile
from datetime import date
from pathlib import Path

from . import bmat_calculo, config, history, mensual
from .version import VERSION

log = logging.getLogger(__name__)

NOMBRE_EXE = "ReportesSpotifyLatam.exe"
CARPETA = "ReportesSpotifyLatam"


def _estado_historico() -> list:
    lineas = []
    ultima = history.ultima_semana_cargada()
    lineas.append("- Spotify semanal: " + (
        f"semana {ultima[1]} de {ultima[0]} (fecha de corte {str(ultima[2])[:10]})"
        if ultima else "vacío"))
    bmat = bmat_calculo.ultima_semana()
    lineas.append("- BMAT: " + (f"semana {bmat[1]} de {bmat[0]}" if bmat else "vacío"))
    meses = mensual.meses_cargados()
    lineas.append("- Mensual: " + (
        f"{config.MESES_ES[meses[-1][1]].lower()} de {meses[-1][0]}" if meses else "vacío"))
    return lineas


def _leeme(estado: list, con_credenciales: bool) -> str:
    credenciales = (
        "El .env con las credenciales de Spotify ya viene incluido."
        if con_credenciales else
        "Copia .env.example como .env y escribe SPOTIFY_CLIENT_ID y SPOTIFY_CLIENT_SECRET.\n"
        "   Sin eso, los reportes salen igual pero las fechas de lanzamiento nuevas quedan vacías."
    )
    return "\r\n".join([
        f"Reportes Spotify Latam - versión {VERSION}",
        f"Paquete de instalación generado el {date.today():%d/%m/%Y}",
        "",
        "HISTÓRICO INCLUIDO (llega hasta):",
        *estado,
        "",
        "INSTALACIÓN",
        "1. Descomprime esta carpeta donde quieras (por ejemplo en Documentos).",
        f"2. {credenciales}",
        f"3. Abre {NOMBRE_EXE}. Arriba de la ventana debe decir la última semana",
        "   cargada que aparece en la lista de arriba.",
        "",
        "IMPORTANTE",
        f"- {NOMBRE_EXE} tiene que quedarse junto a la carpeta data y al archivo .env.",
        "  Si lo mueves, mueve los tres juntos.",
        "- El histórico vive en data/history/universal_data.db. El aplicativo guarda",
        "  respaldos solos en data/history/respaldos/, pero copia la carpeta",
        "  data/history a otro lugar (Drive, disco externo) de vez en cuando.",
        "- Si algo falla, el detalle queda en data/logs/aplicativo.log.",
        "- Desde esta instalación, carga las semanas en orden, empezando por la",
        "  siguiente a la que aparece arriba.",
        "",
    ])


def _copiar_base(destino: Path) -> None:
    """Copia consistente (API de respaldo de SQLite) y verificada."""
    fuente = sqlite3.connect(history.DB_PATH)
    try:
        copia = sqlite3.connect(destino)
        try:
            fuente.backup(copia)
            resultado = copia.execute("PRAGMA integrity_check").fetchone()[0]
        finally:
            copia.close()
    finally:
        fuente.close()
    if resultado != "ok":
        raise RuntimeError(f"La copia de la base no pasó la verificación de integridad: {resultado}")


def generar(salida=None, con_credenciales: bool = False) -> tuple:
    """Arma el zip y devuelve `(ruta, lineas_de_estado)`.

    Falla con un mensaje claro si falta el ejecutable o la base: un paquete
    sin cualquiera de los dos no sirve en el otro equipo.
    """
    raiz = Path(config.ROOT_DIR)
    exe = raiz / NOMBRE_EXE
    if not exe.exists():
        raise FileNotFoundError(
            f"No encontré {NOMBRE_EXE} en {raiz}. Corre build.bat primero y vuelve a intentar.")
    if not Path(history.DB_PATH).exists():
        raise FileNotFoundError(
            f"No encontré la base del histórico ({history.DB_PATH}). Sin ella el paquete "
            "no tendría sentido: genera al menos una semana o siembra el histórico.")
    env = raiz / ".env"
    if con_credenciales and not env.exists():
        raise FileNotFoundError(f"Pediste incluir las credenciales, pero no existe {env}.")

    estado = _estado_historico()
    salida = Path(salida or config.OUTPUT_DIR)
    salida.mkdir(parents=True, exist_ok=True)
    destino = salida / f"ReportesSpotifyLatam_v{VERSION}_instalacion_{date.today():%Y-%m-%d}.zip"
    parcial = destino.with_suffix(".tmp")

    log.info("Paquete de instalación: %s (credenciales=%s)", destino, con_credenciales)
    with tempfile.TemporaryDirectory() as tmp:
        base = Path(tmp) / "universal_data.db"
        _copiar_base(base)
        with zipfile.ZipFile(parcial, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(exe, f"{CARPETA}/{NOMBRE_EXE}")
            z.write(base, f"{CARPETA}/data/history/universal_data.db")
            semillas = raiz / "data" / "history" / "seed"
            for archivo in sorted(semillas.glob("*")) if semillas.is_dir() else []:
                if archivo.is_file():
                    z.write(archivo, f"{CARPETA}/data/history/seed/{archivo.name}")
            for opcional in (config.RELEASE_DATE_DB, raiz / "data" / "bmat_titularidad_compartida.xlsx"):
                if Path(opcional).exists():
                    z.write(opcional, f"{CARPETA}/data/{Path(opcional).name}")
            if (raiz / ".env.example").exists():
                z.write(raiz / ".env.example", f"{CARPETA}/.env.example")
            if con_credenciales:
                z.write(env, f"{CARPETA}/.env")
            for vacia in ("data/output/", "data/raw/"):
                z.writestr(f"{CARPETA}/{vacia}", "")
            z.writestr(f"{CARPETA}/LEEME_INSTALACION.txt", _leeme(estado, con_credenciales))
    parcial.replace(destino)
    log.info("Paquete listo: %s (%.1f MB)", destino, destino.stat().st_size / 1e6)
    return destino, estado
