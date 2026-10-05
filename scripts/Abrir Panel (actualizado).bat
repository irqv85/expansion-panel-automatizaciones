@echo off
REM Trae la ultima version del repo y abre el panel.
REM
REM %~dp0 es la carpeta de este .bat, asi que el proyecto se puede mover de
REM carpeta, de disco o de PC sin editar nada (misma regla que paths.py).
REM
REM El panel SIEMPRE abre, actualice o no. Un lanzador que se queda trabado
REM porque no hay internet, porque git no esta instalado o porque hay cambios
REM locales sin guardar seria peor que no actualizar: te deja sin panel justo
REM cuando lo necesitas. Por eso cada fallo de git solo imprime un aviso y
REM sigue de largo.

setlocal
cd /d "%~dp0.."

where git >nul 2>&1
if errorlevel 1 (
    echo [aviso] git no esta en el PATH: se abre la version que ya tienes.
    goto :abrir
)

echo Buscando actualizaciones...
git fetch --quiet origin 2>nul
if errorlevel 1 (
    echo [aviso] no se pudo contactar GitHub: se abre la version que ya tienes.
    goto :abrir
)

REM --ff-only a proposito: si hay cambios locales que no estan en el remoto,
REM preferimos NO tocarlos y avisar, antes que hacer un merge a ciegas en un
REM lanzador que el usuario abre con doble clic sin leer.
git merge --ff-only origin/main --quiet 2>nul
if errorlevel 1 (
    echo [aviso] tienes cambios locales sin subir: no se actualizo.
    echo         Resuelvelo con git cuando puedas. El panel abre igual.
) else (
    echo Al dia.
)

:abrir
cd /d "%~dp0"
start "" pythonw "Panel de Automatizaciones.py"
endlocal
