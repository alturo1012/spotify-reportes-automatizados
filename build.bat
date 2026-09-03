@echo off
echo ================================================
echo  Empaquetando ReportesSpotifyLatam.exe
echo ================================================
echo.

if not exist ".venv\Scripts\activate.bat" (
    echo No se encontro el entorno virtual .venv en esta carpeta.
    echo Primero sigue los pasos de configuracion del proyecto ^(crear y activar
    echo .venv, instalar requirements.txt^) antes de correr este script.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

echo Instalando PyInstaller ^(si ya esta instalado, esto es rapido^)...
pip install pyinstaller >nul 2>&1

echo.
echo Generando el ejecutable...
rem --hidden-import: spotipy y dotenv se importan de forma "perezosa" (dentro
rem de una funcion / en un try), y PyInstaller no siempre los detecta solo.
rem Sin esto el .exe puede quedar sin la parte de Spotify, y la columna
rem "Fecha Lzto" saldria siempre vacia.
pyinstaller --onefile --windowed --name ReportesSpotifyLatam --distpath . --workpath build_pyinstaller --specpath build_pyinstaller --hidden-import spotipy --hidden-import dotenv run_gui.py

echo.
if exist "ReportesSpotifyLatam.exe" (
    echo Listo. El ejecutable quedo en esta misma carpeta: ReportesSpotifyLatam.exe
    echo.
    echo IMPORTANTE: no muevas ese .exe a otra carpeta. Necesita quedarse junto
    echo a la carpeta "data" para poder leer y guardar el historico y los reportes.
    echo Si en algun momento lo mueves, copia tambien la carpeta "data" junto a el.
) else (
    echo Algo fallo generando el ejecutable. Revisa los mensajes de arriba.
)
echo.
pause
