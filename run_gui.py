"""Punto de entrada para empaquetar la GUI con PyInstaller (ver build.bat).

No lo uses directo con Python para probar la GUI normalmente -- para eso usa
`python -m src.gui` desde la raíz del repo. Este archivo existe solo porque
PyInstaller necesita un script FUERA del paquete `src` como punto de
entrada; `src/gui.py` sigue usando imports relativos (`from . import
config`), que funcionan igual una vez que Python importa `src.gui` como
parte del paquete `src` (que es justo lo que hace la línea de abajo).
"""
from src.gui import main

if __name__ == "__main__":
    main()
