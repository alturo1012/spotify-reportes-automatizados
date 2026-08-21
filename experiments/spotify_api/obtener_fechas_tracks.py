"""Prototipo: obtener fechas de tracks/canciones vía la API de Spotify.

Este archivo es el espacio para pegar y probar el código que te compartan en
la reunión, de forma aislada de `src/` (que es la lógica ya validada de los
dos reportes semanales). Mientras esto esté en `experiments/`, no lo importa
ni lo usa nada de `src/` — es solo para prototipar y correr pruebas sueltas.

Cómo probarlo localmente:

    python -m experiments.spotify_api.obtener_fechas_tracks

Si el código que te den usa credenciales de la API de Spotify (client id /
secret), NO las pegues directo en este archivo. Ponlas en un `.env` local
(basado en `.env.example`, que sí se sube a git) y cárgalas con
`python-dotenv` o `os.environ`. El `.env` real debe quedar en
`.gitignore` para no subir credenciales al repo.

TODO: reemplazar este contenido con el código real que te compartan.
"""


def obtener_fechas_tracks(track_ids):
    """Placeholder: dado un listado de IDs/ISRC de tracks, devuelve su fecha
    de lanzamiento (release date) según la API de Spotify.

    Reemplazar con la implementación real. Firma tentativa — ajústala según
    cómo venga el código de la reunión (puede que trabaje con ISRC en vez de
    track_id de Spotify, por ejemplo).
    """
    raise NotImplementedError("Pegar aquí el código real de la reunión")


if __name__ == "__main__":
    # Prueba manual rápida mientras se integra el código real.
    ejemplo_ids = ["<track_id_o_isrc_de_prueba>"]
    print(obtener_fechas_tracks(ejemplo_ids))
