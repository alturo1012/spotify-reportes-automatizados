# Spotify Reportes Automatizados

Automatización en Python de la generación semanal/mensual de los reportes de
charts y market share de Spotify Latam, a partir de la fuente de datos de BigQuery.

## Objetivo

Reemplazar el proceso manual en Excel por scripts en Python que, a partir de la
fuente de datos (`Fuente_de_datos_BQ_Spotify...xlsx`), generen automáticamente:

- **Reporte_Chart_Top_Semanal_Spotify_Latam** (pestañas `Resumen Total` y `Detalle Tracks`)
- **Reporte_MS_MS_TOP_200_Spotify** (Market Share YTD, pestaña `% Market Share` + una pestaña por país)

## Estructura del repositorio

```
spotify-reportes-automatizados/
├── data/
│   ├── raw/          # Fuente de datos BQ (xlsx/csv) — NO se sube a git
│   └── output/        # Reportes generados — NO se sube a git
├── src/
│   ├── __init__.py
│   ├── config.py       # Rutas, nombres de columnas, países, constantes
│   ├── load_data.py    # Carga y limpieza de la fuente BQ
│   ├── chart_semanal.py    # Genera Reporte_Chart_Top_Semanal
│   ├── market_share.py     # Genera Reporte_MS_MS_TOP_200
│   └── main.py          # Orquesta la generación de ambos reportes
├── tests/
├── requirements.txt
├── .gitignore
└── README.md
```

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate     # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

1. Coloca la fuente de datos más reciente en `data/raw/`.
2. Ejecuta:

```bash
python -m src.main --semana 25 --fuente data/raw/Fuente_de_datos_BQ_Spotify.xlsx
```

3. Los reportes se generan en `data/output/`.

## Estado del proyecto

En desarrollo. Próximos pasos:
- [ x] Mapear 1:1 las fórmulas/lógica de las plantillas actuales (`PLANTILLA_SEMANAL`, `PLANTILLA_MES`) a Python
- [ ] Validar reportes generados contra los reportes manuales existentes semana a semana
- [ ] Programar la ejecución automática (GitHub Actions o tarea programada local)
