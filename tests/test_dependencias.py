"""Las versiones de las dependencias quedan fijas y coherentes.

Sin esto, `pip install` en un equipo nuevo instala la última versión que
exista ese día, y el .exe puede comportarse distinto al validado.
"""
import re
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent


def _versiones(archivo):
    versiones = {}
    for linea in (RAIZ / archivo).read_text(encoding="utf-8").splitlines():
        linea = linea.split("#", 1)[0].strip()
        if not linea:
            continue
        m = re.fullmatch(r"([A-Za-z0-9_.\-]+)==([A-Za-z0-9_.\-]+)", linea)
        assert m, f"{archivo}: '{linea}' no tiene versión exacta (usa ==)"
        versiones[m.group(1).lower().replace("_", "-")] = m.group(2)
    return versiones


def test_requirements_tiene_versiones_exactas():
    assert _versiones("requirements.txt")


def test_el_lock_trae_las_mismas_versiones_que_requirements():
    directas = _versiones("requirements.txt")
    lock = _versiones("requirements-lock.txt")
    for paquete, version in directas.items():
        assert lock.get(paquete) == version, f"{paquete}: {version} en requirements, {lock.get(paquete)} en el lock"


def test_el_lock_incluye_lo_necesario_para_armar_y_probar():
    lock = _versiones("requirements-lock.txt")
    assert "pyinstaller" in lock and "pytest" in lock


def test_no_quedan_dependencias_que_no_se_usan():
    for archivo in ("requirements.txt", "requirements-lock.txt"):
        assert "google-cloud-bigquery" not in _versiones(archivo)


def test_build_bat_instala_desde_el_lock():
    texto = (RAIZ / "build.bat").read_text(encoding="utf-8")
    assert "pip install -r requirements-lock.txt" in texto
    assert "pip install pyinstaller" not in texto
