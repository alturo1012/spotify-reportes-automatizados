"""Respaldo automático de la base del histórico (`universal_data.db`).

La base es lo único del proyecto que no se puede reconstruir por completo:
la siembra llega hasta la semana 33 de 2026 (Spotify) y la 35 (BMAT), y todo
lo cargado después vive solo ahí. Ya se perdió una semana una vez (el 20 de
agosto de 2026) al volver a sembrar encima.

Por eso, antes de cada proceso que escribe en la base, se guarda una copia
comprimida en `data/history/respaldos/`:

    universal_data_2026-10-09_153012_semanal.zip

Se conservan las últimas `config.RESPALDOS_A_CONSERVAR`; las más viejas se
borran solas. Cada copia pesa unos 12 MB (la base, sin comprimir, ~33 MB).

Para restaurar: cerrar el aplicativo, descomprimir el .zip elegido y poner
el `universal_data.db` que trae en `data/history/`, reemplazando el actual.

Un respaldo que falla NUNCA detiene la generación de reportes: se anota en
el registro de errores y el proceso sigue. Perder un respaldo es mucho menos
grave que no poder generar la semana.
"""
import logging
import sqlite3
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from . import config, history

log = logging.getLogger(__name__)

PREFIJO = "universal_data_"


def carpeta_respaldos() -> Path:
    """Al lado de la base, para que viajen juntas si se copia `data/`. Se
    calcula en cada llamada (y no como constante) porque las pruebas
    redirigen `history.DB_PATH` a una carpeta temporal."""
    return Path(history.DB_PATH).parent / "respaldos"


def listar() -> list:
    """Los respaldos existentes, del más viejo al más nuevo."""
    carpeta = carpeta_respaldos()
    if not carpeta.is_dir():
        return []
    # El nombre empieza con la fecha y hora, así que el orden alfabético es
    # el cronológico.
    return sorted(carpeta.glob(f"{PREFIJO}*.zip"))


def _limpiar_motivo(motivo: str) -> str:
    limpio = "".join(c if c.isalnum() else "-" for c in (motivo or "").lower()).strip("-")
    return limpio or "manual"


def _depurar(conservar: int) -> None:
    sobrantes = listar()[:-conservar] if conservar > 0 else []
    for viejo in sobrantes:
        try:
            viejo.unlink()
        except OSError:
            log.warning("No se pudo borrar el respaldo viejo %s", viejo, exc_info=True)


def respaldar(motivo: str, conservar: int = None):
    """Guarda una copia comprimida de la base y devuelve su ruta, o None si
    no había base (instalación nueva) o si la copia falló.

    `motivo` va en el nombre del archivo (semanal, mensual, bmat...), para
    saber qué proceso vino después de cada copia.
    """
    conservar = config.RESPALDOS_A_CONSERVAR if conservar is None else conservar
    origen = Path(history.DB_PATH)
    if not origen.exists():
        log.info("Sin respaldo: todavía no existe la base %s", origen)
        return None

    try:
        carpeta = carpeta_respaldos()
        carpeta.mkdir(parents=True, exist_ok=True)
        sello = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        destino = carpeta / f"{PREFIJO}{sello}_{_limpiar_motivo(motivo)}.zip"
        n = 2
        while destino.exists():   # dos respaldos en el mismo segundo
            destino = carpeta / f"{PREFIJO}{sello}_{_limpiar_motivo(motivo)}_{n}.zip"
            n += 1

        # Copia con la API de respaldo de SQLite y no con un copiar archivo:
        # así la copia es consistente aunque la base esté abierta.
        with tempfile.TemporaryDirectory() as tmp:
            copia = Path(tmp) / "universal_data.db"
            fuente = sqlite3.connect(origen)
            try:
                dst = sqlite3.connect(copia)
                try:
                    fuente.backup(dst)
                finally:
                    dst.close()
            finally:
                fuente.close()
            parcial = destino.with_suffix(".tmp")
            with zipfile.ZipFile(parcial, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(copia, "universal_data.db")
            parcial.replace(destino)   # nunca queda un .zip a medias con nombre válido
    except Exception:  # noqa: BLE001 -- un respaldo fallido no debe tumbar el proceso
        log.exception("No se pudo respaldar la base antes de: %s", motivo)
        return None

    log.info("Respaldo guardado: %s", destino.name)
    _depurar(conservar)
    return destino
