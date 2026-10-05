@echo off
REM %~dp0 es la carpeta de este .bat: asi el proyecto se puede mover de
REM carpeta, de disco o de PC sin editar el lanzador (misma regla que paths.py).
cd /d "%~dp0"
start "" pythonw "Panel de Automatizaciones.py"
