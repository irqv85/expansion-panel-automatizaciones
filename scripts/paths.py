"""Rutas del proyecto GB Claude, resueltas sin depender del sistema operativo.

Antes cada script traia sus rutas absolutas escritas a mano
("/home/usuario/Documentos/proyecto/..."), asi que mover la carpeta -- o
cambiar de PC -- obligaba a editar 36 archivos. Aqui la raiz se deduce de
la ubicacion de este archivo (scripts/paths.py, es decir la raiz es el
directorio padre), asi que el proyecto se puede mover a cualquier disco o
usuario sin tocar una sola linea.

Uso desde scripts/:

    from paths import GB_CLAUDE, DOWNLOADS, FORECAST_DIR

Uso desde una subcarpeta mas profunda (por ejemplo
Claude/Scheduled/analiza-churn-batch/):

    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))
    from paths import GB_CLAUDE, DOWNLOADS
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

# ---------------------------------------------------------------- raiz

SCRIPTS_DIR = Path(__file__).resolve().parent
GB_CLAUDE = SCRIPTS_DIR.parent

IS_WINDOWS = sys.platform == "win32"


# ------------------------------------------------------------ descargas

def _carpeta_descargas():
    """Ubicacion real de la carpeta de descargas del usuario.

    En Windows no se puede asumir "~/Downloads": el usuario puede haberla
    movido a otro disco, y el nombre localizado varia. La API de Known
    Folders devuelve la ruta configurada de verdad; si falla por cualquier
    razon se prueban los nombres habituales en Windows y en Linux (donde la
    carpeta se llamaba "Descargas").
    """
    if IS_WINDOWS:
        try:
            import ctypes
            import ctypes.wintypes

            # FOLDERID_Downloads
            guid = "{374DE290-123F-4565-9164-39C4925E467B}"

            class GUID(ctypes.Structure):
                _fields_ = [
                    ("Data1", ctypes.wintypes.DWORD),
                    ("Data2", ctypes.wintypes.WORD),
                    ("Data3", ctypes.wintypes.WORD),
                    ("Data4", ctypes.c_byte * 8),
                ]

            folder_id = GUID()
            ctypes.windll.ole32.CLSIDFromString(
                ctypes.c_wchar_p(guid), ctypes.byref(folder_id))
            buf = ctypes.c_wchar_p()
            if ctypes.windll.shell32.SHGetKnownFolderPath(
                    ctypes.byref(folder_id), 0, None, ctypes.byref(buf)) == 0:
                ruta = Path(buf.value)
                ctypes.windll.ole32.CoTaskMemFree(buf)
                if ruta.exists():
                    return ruta
        except Exception:
            pass

    home = Path.home()
    for nombre in ("Downloads", "Descargas"):
        candidato = home / nombre
        if candidato.exists():
            return candidato
    return home / "Downloads"


DOWNLOADS = _carpeta_descargas()


# ------------------------------------------------- carpetas del proyecto

EMAILS_VENDEDORES_DIR = GB_CLAUDE / "Emails Vendedores"
FORECAST_DIR = GB_CLAUDE / "Forecast"
METRICAS_DIR = GB_CLAUDE / "Metricas Ejecutivas"
METRICAS_MENSUALES_DIR = GB_CLAUDE / "Metricas Mensuales"
COMP_PLAN_DIR = GB_CLAUDE / "Planes de compensacion"
PRESENTACIONES_DIR = GB_CLAUDE / "Reportes Claude - Presentation"
MANUALES_DIR = GB_CLAUDE / "Presentaciones - Manuales - Claude"
FW_CADENCE_DIR = GB_CLAUDE / "FW-Cadence-Generator"
DASHBOARD_DIR = GB_CLAUDE / "Dashboard SMB-MM"
MENTORING_DIR = GB_CLAUDE / "Mentoring"
CALCULATOR_DIR = GB_CLAUDE / "Gb Calculator"

CHURN_DIR = GB_CLAUDE / "Churn Accounts"
CHURN_INFORMES_DIR = CHURN_DIR / "informes"

GBS_CHECKER_DIR = GB_CLAUDE / "GBS Deal Checker"
GBS_CHECKER_INFORMES_DIR = GBS_CHECKER_DIR / "informes"
GBS_MAKER_DIR = GB_CLAUDE / "GBS Deal Maker"
GBS_MAKER_BORRADORES_DIR = GBS_MAKER_DIR / "borradores"
GBS_ORG_MAKER_DIR = GB_CLAUDE / "GBS Organization Maker"
GBS_ORG_MAKER_BORRADORES_DIR = GBS_ORG_MAKER_DIR / "borradores"

RUTINAS_DIR = GB_CLAUDE / "Rutinas"
LOGS_DIR = RUTINAS_DIR / "logs"

SCHEDULED_DIR = GB_CLAUDE / "Claude" / "Scheduled"
CHURN_BATCH_DIR = SCHEDULED_DIR / "analiza-churn-batch"

ENV_FILE = GB_CLAUDE / ".env"


def skill_file(nombre):
    """Ruta al SKILL.md de una skill programada, por nombre de carpeta."""
    return SCHEDULED_DIR / nombre / "SKILL.md"


def marker_file(nombre):
    """Marcador .last-success-<nombre> que evita repetir una rutina el mismo dia."""
    return LOGS_DIR / f".last-success-{nombre}"


# ----------------------------------------------------------- utilidades

def open_folder(ruta):
    """Abre una carpeta en el explorador de archivos del sistema.

    Reemplaza las llamadas directas a "xdg-open", que solo existe en Linux.
    """
    ruta = Path(ruta)
    ruta.mkdir(parents=True, exist_ok=True)
    if IS_WINDOWS:
        os.startfile(str(ruta))  # noqa: S606  (ruta propia, no viene de entrada externa)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(ruta)])
    else:
        subprocess.Popen(["xdg-open", str(ruta)])


def claude_bin():
    """Ruta al ejecutable del CLI de Claude, o None si no esta instalado.

    En Ubuntu estaba fijo en ~/.local/bin/claude. En Windows se instala en
    varios sitios segun el metodo (npm global, instalador nativo), asi que
    se busca en el PATH primero y luego en las ubicaciones habituales.
    """
    encontrado = shutil.which("claude")
    if encontrado:
        return Path(encontrado)

    # shutil.which mira el PATH de ESTE proceso: si se heredo de un shell
    # abierto antes de instalar el CLI, no lo encuentra aunque este bien
    # instalado. De ahi la busqueda directa en las rutas de instalacion.
    home = Path.home()
    local = home / "AppData" / "Local"
    candidatos = [
        local / "Microsoft" / "WinGet" / "Links" / "claude.exe",
        home / ".local" / "bin" / "claude.exe",
        home / ".local" / "bin" / "claude",
        home / "AppData" / "Roaming" / "npm" / "claude.cmd",
        local / "Programs" / "claude" / "claude.exe",
    ]
    # Instalacion por winget: la carpeta del paquete lleva un sufijo de hash,
    # asi que hay que buscarla por patron.
    paquetes = local / "Microsoft" / "WinGet" / "Packages"
    if paquetes.is_dir():
        candidatos.extend(sorted(paquetes.glob("Anthropic.ClaudeCode*/claude.exe")))

    for candidato in candidatos:
        if candidato.exists():
            return candidato
    return None


CLAUDE_BIN = claude_bin()


if __name__ == "__main__":
    print(f"GB_CLAUDE  = {GB_CLAUDE}")
    print(f"DOWNLOADS  = {DOWNLOADS}")
    print(f"CLAUDE_BIN = {CLAUDE_BIN or 'NO ENCONTRADO'}")
    print()
    faltantes = []
    for nombre, valor in sorted(globals().items()):
        if nombre.endswith("_DIR") and isinstance(valor, Path):
            estado = "ok " if valor.exists() else "NO EXISTE"
            print(f"  [{estado}] {nombre:32} {valor}")
            if not valor.exists():
                faltantes.append(nombre)
    if faltantes:
        print(f"\nCarpetas que aun no existen ({len(faltantes)}): se crean al primer uso.")
