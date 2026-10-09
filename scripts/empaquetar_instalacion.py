"""Arma el zip para instalar el aplicativo en otro equipo, con la base al día.

    python -m scripts.empaquetar_instalacion
    python -m scripts.empaquetar_instalacion --salida "D:/Entregas"
    python -m scripts.empaquetar_instalacion --con-credenciales

Antes: correr build.bat, para que el ejecutable sea el de la versión actual.

--con-credenciales mete también el .env con las credenciales de Spotify de
ESTE equipo. Por defecto no va: lo normal es que el otro equipo use las
suyas. Si lo usas, no compartas el zip por un canal público.

Ver src/paquete.py para el detalle de qué lleva y qué no.
"""
import argparse
import sys

from src import paquete, registro


def main(argv=None):
    p = argparse.ArgumentParser(description="Arma el paquete de instalación con la base al día")
    p.add_argument("--salida", help="Carpeta donde dejar el zip (por defecto data/output)")
    p.add_argument("--con-credenciales", action="store_true",
                   help="Incluir el .env con las credenciales de Spotify de este equipo")
    args = p.parse_args(argv)
    registro.configurar()
    try:
        ruta, estado = paquete.generar(args.salida, con_credenciales=args.con_credenciales)
    except (FileNotFoundError, RuntimeError) as e:
        sys.exit(f"No se armó el paquete: {e}")
    print(f"Paquete listo: {ruta} ({ruta.stat().st_size / 1e6:.1f} MB)")
    print("El histórico que lleva llega hasta:")
    for linea in estado:
        print(f"  {linea}")
    if args.con_credenciales:
        print("OJO: incluye el .env con las credenciales de Spotify de este equipo.")
    return ruta


if __name__ == "__main__":
    main()
