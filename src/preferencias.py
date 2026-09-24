"""Preferencias del usuario que sobreviven entre corridas.

Hoy solo guarda en qué carpeta dejar los reportes. Va en un JSON al lado de
la base (`data/preferencias.json`) en vez de en la base misma para que se
pueda abrir y corregir con el bloc de notas si hace falta, y para que
borrar el histórico no se lleve la configuración.

Si el archivo no está, está vacío o quedó corrupto, se usan los valores por
defecto: nunca hace fallar la generación de un reporte.
"""
import json
from pathlib import Path

from . import config

ARCHIVO = config.ROOT_DIR / "data" / "preferencias.json"

CARPETA_SALIDA = "carpeta_salida"
CARPETA_SALIDA_BMAT = "carpeta_salida_bmat"


def cargar() -> dict:
    try:
        with open(ARCHIVO, encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, dict) else {}
    except (OSError, ValueError):
        return {}


def guardar(clave: str, valor) -> None:
    """Deja `clave` con ese valor y respeta lo demás que hubiera."""
    datos = cargar()
    datos[clave] = str(valor) if valor is not None else None
    try:
        ARCHIVO.parent.mkdir(parents=True, exist_ok=True)
        with open(ARCHIVO, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except OSError:
        pass   # no poder guardar la preferencia no debe romper nada


def carpeta_salida(clave: str = CARPETA_SALIDA) -> Path:
    """La carpeta elegida la última vez, o data/output si no hay ninguna
    (o si la guardada ya no existe: por ejemplo un disco de red desconectado
    o una carpeta que borraron)."""
    guardada = cargar().get(clave)
    if guardada:
        ruta = Path(guardada)
        if ruta.is_dir():
            return ruta
    return config.OUTPUT_DIR


def recordar_carpeta_salida(carpeta, clave: str = CARPETA_SALIDA) -> None:
    guardar(clave, Path(carpeta))
