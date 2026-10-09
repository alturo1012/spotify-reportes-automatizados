"""Versión del aplicativo.

Se muestra en el título de la ventana y queda en el registro cada vez que
se abre, para saber qué ejecutable tiene cada equipo: el .exe es una copia
del código del momento en que se corrió build.bat y no se actualiza solo.

Al publicar una versión nueva: subir este número, fusionar a main, crear el
tag con el mismo número (git tag -a v1.2 ...) y volver a correr build.bat.
"""
VERSION = "1.2"
