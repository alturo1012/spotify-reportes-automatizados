# experiments/spotify_api

Espacio aislado para probar el código de la API de Spotify que te compartan
(fechas de lanzamiento de tracks), sin tocar la lógica ya validada de
`src/` (los dos reportes semanales — Chart Top Semanal y Market Share).

## Por qué separado de `src/`

`src/` es la base de lógica de negocio de los reportes, ya mapeada 1:1
contra las plantillas reales (ver `claude/mapeo_logica_plantillas.md` en el
Project). Meter código nuevo y todavía no probado ahí mezclaría cosas
distintas. `experiments/` es el lugar para probar, romper y ajustar el
código de la API de Spotify hasta que funcione como esperas; solo cuando ya
esté probado y tengas claro para qué se usa dentro de los reportes, se migra
a `src/` (o se deja como un módulo aparte si termina siendo un proceso
independiente).

## Cómo usarlo

1. Pega el código que te den en `obtener_fechas_tracks.py` (o crea otro
   archivo dentro de esta misma carpeta si son varios scripts).
2. Si necesita credenciales de la API de Spotify:
   - Copia `.env.example` a `.env` (este último NO se sube a git).
   - Pon ahí tu `SPOTIPY_CLIENT_ID` / `SPOTIPY_CLIENT_SECRET` (o los nombres
     que use el código real).
3. Instala lo que haga falta. Si usan `spotipy` (la librería más común para
   esto en Python), agrégalo a `requirements-experiments.txt` en la raíz del
   repo y corre `pip install -r requirements-experiments.txt`.
4. Corre pruebas sueltas con:
   ```bash
   python -m experiments.spotify_api.obtener_fechas_tracks
   ```
5. Pruebas más formales van en `tests/experiments/test_obtener_fechas_tracks.py`.

## Antes de integrarlo a los reportes

Cuando ya funcione, conviene aclarar (ver la nota que dejamos pendiente):
¿para qué se van a usar estas fechas dentro de `Reporte_Chart_Top_Semanal` o
`Reporte_MS_TOP200`? Ninguno de los dos usa hoy una fecha de lanzamiento por
canción — si se vuelve un dato real del pipeline, hay que sumarlo a
`SOURCE_COLUMNS` en `src/config.py` y a `history.py` (ver
`claude/plan_fusion_paso_a_paso.md`, Paso 2).
