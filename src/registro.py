"""Registro de errores y actividad en `data/logs/aplicativo.log`.

El .exe se empaqueta sin consola (--windowed): todo lo que el programa
imprime se pierde, y si algo falla solo queda el mensaje en pantalla, que
desaparece al cerrarlo. Con este archivo, ante cualquier problema basta con
pedir `data/logs/aplicativo.log` para ver qué pasó, cuándo y con qué
archivo, incluido el detalle técnico completo del error.

Qué se anota:
- cada proceso que se corre (semanal, mensual, BMAT...) con sus datos de
  entrada y dónde quedaron los archivos;
- los avisos que se le mostraron al usuario;
- los respaldos de la base;
- los errores, con el traceback completo.

El archivo rota solo: al llegar a 1 MB pasa a `aplicativo.log.1` y se
conservan 5 anteriores, así que nunca ocupa más de ~6 MB.
"""
import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import config

ARCHIVO = "aplicativo.log"
TAMANO_MAXIMO = 1_000_000
ARCHIVOS_ANTERIORES = 5
FORMATO = "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s"

_MARCA = "_registro_aplicativo"


def carpeta_logs() -> Path:
    return Path(config.LOG_DIR)


def ruta_log() -> Path:
    return carpeta_logs() / ARCHIVO


def configurar() -> Path:
    """Activa el registro en archivo. Se puede llamar varias veces (desde la
    ventana y desde cada script): solo agrega el manejador la primera vez,
    o si la carpeta de logs cambió (pasa en las pruebas).

    Si no se puede crear el archivo (carpeta de solo lectura, por ejemplo),
    sigue sin registro en vez de fallar.
    """
    raiz = logging.getLogger()
    destino = ruta_log()
    for h in list(raiz.handlers):
        if getattr(h, _MARCA, False):
            if Path(h.baseFilename) == destino.resolve():
                return destino
            raiz.removeHandler(h)
            h.close()
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        manejador = RotatingFileHandler(destino, maxBytes=TAMANO_MAXIMO,
                                        backupCount=ARCHIVOS_ANTERIORES, encoding="utf-8")
    except OSError:
        return destino
    manejador.setFormatter(logging.Formatter(FORMATO))
    manejador.setLevel(logging.INFO)
    setattr(manejador, _MARCA, True)
    raiz.addHandler(manejador)
    if raiz.level > logging.INFO or raiz.level == logging.NOTSET:
        raiz.setLevel(logging.INFO)
    return destino


def registrar_excepciones_no_capturadas() -> None:
    """Para errores que se escapan de todo try: quedan en el archivo antes de
    que el programa se cierre."""
    anterior = sys.excepthook

    def _hook(tipo, valor, tb):
        logging.getLogger("aplicativo").critical("Error no controlado", exc_info=(tipo, valor, tb))
        anterior(tipo, valor, tb)

    sys.excepthook = _hook
