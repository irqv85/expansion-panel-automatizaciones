"""
Panel grafico unico para las automatizaciones de esta carpeta. Cada seccion
tiene su propio semaforo:

  - Reportes para vendedores (Sofia): revisa "Reportes Claude - Presentation".
      verde = nada pendiente | amarillo = archivo nuevo | rojo = pendiente hace rato
      (skill "sofia-schedule-vendedores". Desde el 27-ago-2026 tiene botones
      "Calcular plan" -> "Ver plan" -> "Enviar": ambos pasos disparan Claude
      Code sin supervision -- headless -- porque calcular el plan ya
      requiere cruzar calendario y disponibilidad real por Outlook via MCP,
      algo a lo que este panel Python no tiene acceso directo. El clic en
      "Enviar" ES la aprobacion del plan, no hay otra confirmacion en el
      medio)
  - Calidad CRM (Reportes Next Step): revisa que los Excel de vTiger en
    Descargas esten frescos (fecha de hoy) y que no haya HTML sin enviar en
    "Emails Vendedores".
      verde = todo al dia | amarillo = nuevo por enviar | rojo = desactualizado
      o sin enviar hace rato | gris = faltan archivos fuente en Descargas
      (el envio es "Calcular envío" -> "Enviar", los dos pasos con el
      script local enviar_reportes_next_step.py: el primero calcula el plan
      con --plan-json sin tocar nada, el segundo manda los correos por
      Outlook de escritorio via COM con el HTML ADJUNTO. El clic en
      "Enviar" ES la aprobacion del plan, no hay otra confirmacion en el
      medio. "Generar" sigue igual, la regeneracion desde vTiger no se toco.
      Historia, para no volver a dar la vuelta: el envio por Thunderbird
      -compose se retiro el 26-ago-2026 porque salia duplicado, y como
      reemplazo se paso a Claude Code + MCP de M365, que no soporta
      adjuntos y por eso subia los reportes a SharePoint y mandaba solo un
      enlace. El 9-sep-2026 el responsable comercial pidio volver al adjunto: se vuelve al
      script local, pero sobre Outlook COM -- mail_helper.send_mail llama
      mail.Send() una sola vez por correo, no es el mecanismo de
      Thunderbird que duplicaba)
  - Reasignación de orgs (Sofia) (nueva 27-ago-2026): skill
    "deteccion-reasignacion-orgs-sofia", un solo boton "Detectar y
    reasignar" -- sin plan/confirmacion, igual que el cron diario
    (Rutinas/run-deteccion-reasignacion-orgs-sofia.sh): detecta menciones
    de vendedores sobre organizaciones de Sofia, publica el comentario en
    vTiger y reasigna el Assigned To directo cuando se cumplen las 3
    condiciones del skill.
      verde = el cron de hoy ya corrio con exito | amarillo = la ultima
      corrida exitosa no es de hoy | gris = sin corridas registradas
      (el semaforo solo lee el marcador que escribe el cron, no consulta
      Outlook/vTiger de nuevo -- eso requiere Claude via MCP)
  - Análisis de Churn (nueva 27-ago-2026): skill "analiza-churn-batch",
    boton "Analizar churn" -- tambien sin plan/confirmacion, porque es de
    solo lectura hacia vTiger (no escribe nada, no manda correos).
    Identifica organizaciones con churn aplicable, corre el analisis
    forense de "analiza-churn" en paralelo (Workflow) por cuenta sin
    informe previo, verifica cada veredicto, y arma un Excel resumen en
    Churn Accounts/informes/ (boton "Abrir carpeta" al lado para revisarlo).
    Puede tardar varios minutos.
      verde = ya hay un resumen de hoy | amarillo = el ultimo resumen no
      es de hoy | gris = todavia no hay ningun resumen (el semaforo solo
      lee la fecha del ultimo Excel guardado, no consulta vTiger de nuevo)
  - GBS Deal Checker (28-ago-2026): boton "Chequear GBS" -- de solo lectura
    hacia vTiger, igual que Churn. Revisa que cada deal ABIERTO del equipo
    (Sales Stage distinto de Closed Won/Closed Lost) tenga un "Deals GBS
    Link" en el ultimo DIQ de Descargas, y que ese link responda (no un
    404/410 directo -- eso en SharePoint suele significar que el documento
    detras ya no existe, sin importar quien lo abra). Corre local
    (gbs_deal_checker.py, sin Claude ni MCP) y guarda un Excel resumen en
    GBS Deal Checker/informes/ (boton "Abrir carpeta" al lado). Puede
    tardar 1-2 min por los chequeos de red.
      verde = ya hay un chequeo de hoy | amarillo = el ultimo chequeo no es
      de hoy | gris = todavia no hay ningun chequeo (el semaforo solo lee
      la fecha del ultimo Excel guardado, no vuelve a chequear los links)

REDIMENSIONAMIENTO (28-ago-2026, pedido explicito de el responsable comercial): Reasignacion de
orgs y Analisis de Churn comparten una sola fila (una al lado de la otra,
no cada una a ancho completo) -- la version anterior sumaba una fila entera
por tarjeta y la ventana quedaba mas alta que la pantalla, sin poder
maximizarse ni redimensionarse bien. Ademas, TODAS las filas de tarjetas
(incluida Forecast, que antes tenia weight=0 fijo) llevan weight=1 -- el
espacio que gana o pierde la ventana se reparte proporcional entre todas,
no solo entre las dos primeras.
  - Vacaciones: verde = equipo completo | amarillo = 1 de vacaciones |
    rojo = 2 o mas de vacaciones (capacidad reducida)

Los semaforos se refrescan solos cada minuto (y con el boton "Actualizar").
Interfaz horizontal con el sistema de diseno de GB Advisors (Inter, magenta
#EA018B, gradiente oscuro de marca y la pastilla de 10 px como firma).

Ejecutar con: python "Panel de Automatizaciones.py"  (o el .bat)
"""
import importlib.util
import json
import os
import re
import subprocess
import sys
import threading
import webbrowser
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk

# Cargar .env con credenciales de vTiger
from load_env import load_env
load_env()

from paths import (
    CHURN_INFORMES_DIR,
    COMP_PLAN_DIR,
    DOWNLOADS,
    EMAILS_VENDEDORES_DIR,
    FORECAST_DIR,
    FW_CADENCE_DIR,
    GB_CLAUDE,
    GBS_CHECKER_INFORMES_DIR,
    GBS_MAKER_DIR,
    GBS_ORG_MAKER_DIR,
    METRICAS_DIR,
    PRESENTACIONES_DIR,
    RUTINAS_DIR,
    claude_bin,
    marker_file,
    open_folder,
    skill_file,
)

SCRIPT_DIR = Path(__file__).resolve().parent
PYTHON = sys.executable
PYTHONW = PYTHON.replace("python.exe", "pythonw.exe")

# El panel corre bajo pythonw (sin consola), asi que cuando lanza un hijo que
# SI es aplicacion de consola -- claude.exe -- Windows le abre una ventana de
# cmd propia, que aparece y se queda mientras dura la corrida. Con
# CREATE_NO_WINDOW el hijo nace sin consola y su salida sigue llegando por las
# tuberias de siempre (9-sep-2026, la ventana que salia al dar "Analizar
# churn"). En Linux la constante no existe: el flag queda en 0.
SIN_VENTANA = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0

PENDING_STALE_HOURS = 2  # mas de esto sin enviar -> semaforo rojo
AUTO_REFRESH_MS = 60_000
SOFIA_PLAN_FILE = SCRIPT_DIR / "plan_sofia.json"
SOFIA_SKILL_FILE = skill_file("sofia-schedule-vendedores")
NEXT_STEP_PLAN_FILE = SCRIPT_DIR / "plan_envio_next_step.json"
REASIGNA_SKILL_FILE = skill_file("deteccion-reasignacion-orgs-sofia")
REASIGNA_MARKER_FILE = marker_file("deteccion-reasignacion-orgs-sofia")
# Las dos rutinas que Rutinas/Registrar-Tareas.ps1 declara, en el mismo orden
# en que corren de manana. El panel las expone tambien a mano (tarjeta
# "Rutinas automaticas (Sofia)"): tras migrar de Ubuntu a Windows el
# 9-sep-2026 las tareas quedaron sin registrar en el Programador, y sin un
# boton no habia forma de dispararlas desde aca.
RUTINAS_SOFIA = (
    ("daily-client-presentation-prep", "Presentaciones diarias", "08:00"),
    ("deteccion-reasignacion-orgs-sofia", "Reasignación de orgs", "08:30"),
)
RUNNER_RUTINAS = RUTINAS_DIR / "Run-Rutina.ps1"
# El deck de cliente no vive en Claude/Scheduled: es un .skill empaquetado
# (un ZIP) en la raiz del proyecto, igual que gb-advisors-design.skill.
DECK_SKILL_FILE = GB_CLAUDE / "gb-deck-cliente.skill"
# Prefijo del modulo Accounts en la API de vtiger: el id del registro que va
# en la URL (9244350) es el mismo numero, pero la API lo pide como 3x9244350.
VTIGER_ACCOUNTS_PREFIX = "3x"


def id_cuenta_vtiger(texto):
    """Saca el id de cuenta de lo que el responsable comercial pegue: el link de la ficha, el id
    crudo o el numero suelto. Devuelve "3x<N>" o None si no parece un id.

    Se acepta cualquier forma de URL en vez de calcar un formato concreto,
    porque vtiger ha usado varias (index.php?...&record=N, rutas /Accounts/N)
    y una expresion que solo entienda una fallaria en silencio el dia que
    cambie. Se busca el numero largo de la cadena, que en todas ellas es el id
    del registro (23-sep-2026).

    Ojo: esto NO valida que la cuenta exista. De eso se encarga el prompt, que
    tiene que resolver el id contra vtiger antes de generar nada."""
    if not texto:
        return None
    texto = texto.strip()
    m = re.fullmatch(r"(?:3x)?(\d{3,})", texto)
    if m:
        return VTIGER_ACCOUNTS_PREFIX + m.group(1)
    # record=N / recordId=N mandan sobre cualquier otro numero de la URL.
    m = re.search(r"record(?:Id)?=(\d{3,})", texto, re.IGNORECASE)
    if m:
        return VTIGER_ACCOUNTS_PREFIX + m.group(1)
    numeros = re.findall(r"(?<![\dx])(\d{4,})", texto)
    if len(numeros) == 1:
        return VTIGER_ACCOUNTS_PREFIX + numeros[0]
    if numeros:
        # Varios candidatos: el ultimo suele ser el del registro abierto.
        return VTIGER_ACCOUNTS_PREFIX + numeros[-1]
    return None


CHURN_SKILL_FILE = skill_file("analiza-churn-batch")
CHURN_RESUMEN_RE = re.compile(r"^resumen-churn-(\d{4}-\d{2}-\d{2})\.xlsx$")
GBS_CHECKER_RE = re.compile(r"^gbs-checker-(\d{4}-\d{2}-\d{2})\.xlsx$")
GBS_NOTIFY_SKILL_FILE = skill_file("notificar-gbs-deal-checker")
GBS_NOTIFY_PLAN_FILE = SCRIPT_DIR / "plan_notificacion_gbs.json"

GBS_MAKER_SKILL_FILE = skill_file("gbs-deal-maker")
GBS_MAKER_PLAN_FILE = SCRIPT_DIR / "plan_creacion_gbs_deals.json"
GBS_MAKER_BORRADORES_DIR = GBS_MAKER_DIR / "borradores"
GBS_MAKER_DESPLEGADOS_DIR = GBS_MAKER_DIR / "desplegados"

GBS_ORG_MAKER_SKILL_FILE = skill_file("gbs-organization-maker")
GBS_ORG_MAKER_PLAN_FILE = SCRIPT_DIR / "plan_creacion_gbs_orgs.json"
GBS_ORG_MAKER_BORRADORES_DIR = GBS_ORG_MAKER_DIR / "borradores"
GBS_ORG_MAKER_DESPLEGADOS_DIR = GBS_ORG_MAKER_DIR / "desplegados"
# Cuantas organizaciones entran en cada corrida de "Generar" (pedido de el responsable comercial,
# 28-ago-2026): son ~150 pendientes y cada una consume creditos de Claude, asi
# que se procesan por tandas. La eleccion se guarda para la proxima corrida.
GBS_ORG_LIMITE_FILE = SCRIPT_DIR / "state_gbs_org_limite.json"
GBS_ORG_LIMITE_MAX = 100
GBS_ORG_LIMITE_DEFAULT = 25

# Se resuelve en tiempo de ejecucion (PATH primero, luego las ubicaciones
# habituales del instalador). En Ubuntu estaba fijo en ~/.local/bin/claude;
# en Windows depende del metodo de instalacion, y si no esta instalado
# CLAUDE_BIN queda en None y run_claude_headless lo reporta como error claro
# en vez de fallar con "archivo no encontrado".
CLAUDE_BIN = claude_bin()
# Limite duro para las corridas headless de Claude Code (ver
# run_claude_headless). Sin esto, si una sesion se queda colgada (ej.
# esperando algo que nunca llega, como paso el 31-ago-2026 con el envio de
# Sofia) el panel se queda "Ejecutando Claude Code..." para siempre, sin
# avisar ni matar el proceso. 10 min alcanza de sobra para las tareas mas
# pesadas (generar/desplegar varios borradores) sin ser tan largo como para
# tapar un cuelgue real.
CLAUDE_HEADLESS_TIMEOUT_SECONDS = 30 * 60
# El analisis de churn no cabe en esos 30 minutos: corre un agente forense por
# cuenta sin informe previo, mas un verificador por cada uno, y cada informe
# son ~35 KB de .md mas su .html. La corrida del 14-sep-2026 tardo mas de dos
# horas (primer informe 14:43, Excel 16:10) -- desde el panel se habria matado
# a los 30 minutos y nunca habria llegado a escribir el Excel, que es
# justamente lo que el responsable comercial reportaba. 4 horas deja margen sin tapar un cuelgue
# real de verdad.
CHURN_TIMEOUT_SECONDS = 4 * 60 * 60
FORECAST_MANIFEST = SCRIPT_DIR / "ultima_generacion_forecast.json"
# Debe calzar con CATEGORY_OPTIONS/SEGMENT_OPTIONS/CAT_LABELS de
# generar_reporte_forecast.py (no se importa ese modulo solo para esto).
FORECAST_CATEGORY_OPTIONS = ["CX", "EX", "D42", "AI", "PF", "Other"]
FORECAST_CATEGORY_LABELS = {"CX": "CX", "EX": "EX", "D42": "Device 42", "AI": "Freddy/AI",
                            "PF": "Payment Frequency", "Other": "Other"}
FORECAST_SEGMENT_OPTIONS = ["New Business", "Expansion"]
FW_CADENCE_SKILL_FILE = skill_file("fw-cadence-deck")
FW_CADENCE_RE = re.compile(r"^Freshworks-CX-Pipeline-(\d{4}-\d{2})-\d{2}\.html$")

# ---- colores de semaforo (el resto de los tokens de marca va en la seccion UI) ----
COLOR_GREEN = "#1E9E5A"
COLOR_YELLOW = "#E0A100"
COLOR_RED = "#D2362B"
COLOR_GRAY = "#9AA0A6"


def _load_module(filename, modname):
    path = SCRIPT_DIR / filename
    spec = importlib.util.spec_from_file_location(modname, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


try:
    sofia = _load_module("Sofia - Schedule.py", "sofia_schedule")
except Exception:
    # Depende de win32com/Outlook COM (calendarios compartidos, free/busy) --
    # no portado a Ubuntu todavia. La rotacion de Sofia queda deshabilitada
    # (ver revisar_plan_sofia) hasta que se migre aparte.
    sofia = None
gen_ns = _load_module("generar_reportes_next_step.py", "gen_next_step")
send_ns = _load_module("enviar_reportes_next_step.py", "send_next_step")

NEXT_STEP_REQUIRED = ["orgiq", "oa", "diq", "far", "created", "farmconver",
                      "churn", "farmnoprod"]


# ---------------- logica de los semaforos ----------------

def check_sofia_status():
    try:
        state = sofia.load_state()
        processed = set(state.get("processed_files", []))
        # Una carpeta de salida que todavia no se creo no es un error: es el
        # estado normal de una instalacion nueva. Antes el semaforo mostraba un
        # WinError 3 con la ruta cruda, que parecia una rotura (5-oct-2026).
        if not sofia.WATCH_FOLDER.exists():
            return COLOR_GRAY, "Sin presentaciones generadas todavía."
        current_files = [p for p in sofia.WATCH_FOLDER.iterdir() if p.is_file()]
        pending = [p for p in current_files if p.name not in processed]
    except Exception as exc:
        return COLOR_GRAY, f"No se pudo revisar la carpeta: {exc}"

    if not pending:
        return COLOR_GREEN, "Todo enviado. No hay archivos pendientes."

    now = datetime.now()
    ages_hours = [
        (now - datetime.fromtimestamp(p.stat().st_mtime)).total_seconds() / 3600
        for p in pending
    ]
    max_age = max(ages_hours)
    nombres = ", ".join(p.name for p in pending)

    if max_age < PENDING_STALE_HOURS:
        return COLOR_YELLOW, f"Archivo(s) nuevo(s) detectado(s) ({len(pending)}): {nombres}"
    return COLOR_RED, f"Sin enviar hace mas de {PENDING_STALE_HOURS}h ({len(pending)}): {nombres}"


def check_next_step_status():
    try:
        files = gen_ns.find_latest_files()
    except Exception as exc:
        return COLOR_GRAY, f"No se pudo revisar Descargas: {exc}"

    missing = [c for c in NEXT_STEP_REQUIRED if c not in files]
    stale = gen_ns.check_stale_files(files, datetime.now().date())

    try:
        state = send_ns.load_state()
        processed = set(state.get("processed_files", []))
        current = [p for p in send_ns.WATCH_FOLDER.iterdir() if p.is_file()] if send_ns.WATCH_FOLDER.exists() else []
        pending = [p for p in current if p.name not in processed]
    except Exception:
        pending = []

    reasons = []
    color = COLOR_GREEN

    if missing:
        color = COLOR_GRAY
        reasons.append(f"Faltan {len(missing)} archivo(s) fuente de vTiger en Descargas")

    if stale:
        color = COLOR_RED
        nombres = ", ".join(p.name for _, p, _ in stale)
        reasons.append(f"Reportes de vTiger desactualizados: {nombres}")

    if pending:
        now = datetime.now()
        ages = [(now - datetime.fromtimestamp(p.stat().st_mtime)).total_seconds() / 3600 for p in pending]
        max_age = max(ages)
        nombres = ", ".join(p.name for p in pending)
        if max_age >= PENDING_STALE_HOURS:
            color = COLOR_RED
            reasons.append(f"Sin enviar hace mas de {PENDING_STALE_HOURS}h: {nombres}")
        elif color == COLOR_GREEN:
            color = COLOR_YELLOW
            reasons.append(f"Nuevo(s) por enviar: {nombres}")

    if not reasons:
        return COLOR_GREEN, "Todo al dia: fuentes frescas y nada pendiente por enviar."
    return color, " | ".join(reasons)


def marcar_corrida_hoy(nombre_skill):
    """Deja el marcador .last-success-<skill> con la fecha de hoy.

    Hasta el 29-sep-2026 ese archivo SOLO lo escribia Rutinas/Run-Rutina.ps1,
    asi que una corrida lanzada desde el panel terminaba bien y el semaforo se
    quedaba igual: seguia mostrando la fecha de la ultima corrida programada.
    el responsable comercial lo reporto con "Detectar y reasignar", que llevaba en 2026-09-14
    quince dias despues. El panel dispara exactamente el mismo trabajo, asi que
    tambien tiene que marcarlo.

    Se escribe sin BOM a proposito: PowerShell lo agregaba con
    Set-Content -Encoding utf8 y el semaforo, que lee con Python, comparaba
    "\ufeff2026-09-29" contra la fecha de hoy y nunca coincidia."""
    try:
        destino = marker_file(nombre_skill)
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text(datetime.now().strftime("%Y-%m-%d"), encoding="utf-8")
        return True, None
    except Exception as exc:
        return False, str(exc)


def check_reasigna_orgs_status():
    """Semaforo de reasignacion de orgs: a diferencia de las otras tarjetas,
    esta no puede saber si hay menciones nuevas sin consultar Outlook/vTiger
    (acceso que solo tiene Claude via MCP) -- asi que solo refleja si el cron
    diario (Rutinas/run-deteccion-reasignacion-orgs-sofia.sh) corrio hoy con
    exito, leyendo el mismo marcador que el propio cron escribe."""
    if not REASIGNA_MARKER_FILE.exists():
        return COLOR_GRAY, "Sin corridas registradas todavía."
    try:
        # utf-8-sig y no utf-8: PowerShell escribe el marcador con BOM y
        # Python no lo quita solo, asi que la comparacion con hoy fallaba
        # y el semaforo se quedaba amarillo (9-sep-2026).
        marcado = REASIGNA_MARKER_FILE.read_text(encoding="utf-8-sig").strip()
    except Exception as exc:
        return COLOR_GRAY, f"No se pudo leer el marcador: {exc}"
    hoy = datetime.now().strftime("%Y-%m-%d")
    if marcado == hoy:
        # Ya no dice "automática": desde el 29-sep-2026 el marcador tambien lo
        # escribe el boton del panel, asi que la corrida pudo ser a mano.
        return COLOR_GREEN, (f"Corrida de hoy ({hoy}) completada. "
                             "Dale a 'Detectar y reasignar' para revisar de nuevo ahora.")
    return COLOR_YELLOW, f"Última corrida exitosa: {marcado} (no es de hoy)."


def check_rutinas_sofia_status():
    """Semaforo de las rutinas programadas: lee los mismos marcadores
    .last-success-<skill> que escribe Run-Rutina.ps1, sin consultar nada.
    Verde solo si las DOS salieron bien hoy; si falta alguna, el texto dice
    cual, porque el caso normal tras la migracion es que ninguna corra."""
    hoy = datetime.now().strftime("%Y-%m-%d")
    al_dia, atrasadas, sin_correr = [], [], []
    for skill, etiqueta, _hora in RUTINAS_SOFIA:
        marcador = marker_file(skill)
        if not marcador.exists():
            sin_correr.append(etiqueta)
            continue
        try:
            marcado = marcador.read_text(encoding="utf-8-sig").strip()
        except Exception:
            sin_correr.append(etiqueta)
            continue
        if marcado == hoy:
            al_dia.append(etiqueta)
        else:
            atrasadas.append(f"{etiqueta} ({marcado})")

    if len(al_dia) == len(RUTINAS_SOFIA):
        return COLOR_GREEN, f"Las 2 rutinas corrieron hoy ({hoy})."
    partes = []
    if atrasadas:
        partes.append("última corrida: " + ", ".join(atrasadas))
    if sin_correr:
        partes.append("sin corridas registradas: " + ", ".join(sin_correr))
    detalle = "; ".join(partes)
    if not al_dia:
        return COLOR_YELLOW, f"Ninguna corrió hoy — {detalle}."
    return COLOR_YELLOW, f"Corrió hoy: {', '.join(al_dia)} — {detalle}."


def check_validar_deals_status():
    """Semaforo de validacion de deals: es un botón manual sin cron, así que solo
    devuelve un status informativo."""
    return COLOR_GRAY, "Dale a 'Validar deals' para revisar ahora."


def check_churn_status():
    """Semaforo de churn: la skill 'analiza-churn-batch' es de solo lectura
    y no corre por cron, asi que el semaforo solo refleja la fecha del
    ultimo Excel resumen-churn-YYYY-MM-DD.xlsx guardado -- no valida si hay
    churn nuevo sin analizar (para eso hay que consultar vTiger, que solo
    puede hacer Claude via MCP, no este panel)."""
    if not CHURN_INFORMES_DIR.exists():
        return COLOR_GRAY, "Carpeta de informes de churn no encontrada."
    fechas = []
    for p in CHURN_INFORMES_DIR.iterdir():
        m = CHURN_RESUMEN_RE.match(p.name)
        if m:
            fechas.append(m.group(1))
    if not fechas:
        return COLOR_GRAY, "Sin resumen de churn generado todavía."
    ultima = max(fechas)
    hoy = datetime.now().strftime("%Y-%m-%d")
    if ultima == hoy:
        return COLOR_GREEN, f"Resumen de hoy ({hoy}) ya generado."
    return COLOR_YELLOW, (f"Último resumen: {ultima} (no es de hoy). "
                          "Dale a 'Analizar churn' para correrlo de nuevo.")


def check_gbs_checker_status():
    """Semaforo del GBS Deal Checker: de solo lectura hacia vTiger (no
    escribe nada, no manda correos), igual que Churn -- el semaforo solo
    refleja la fecha del ultimo Excel gbs-checker-YYYY-MM-DD.xlsx guardado,
    no vuelve a revisar los links por su cuenta (eso tarda 1-2 min por los
    chequeos de red)."""
    if not GBS_CHECKER_INFORMES_DIR.exists():
        return COLOR_GRAY, "Carpeta de informes no encontrada."
    fechas = []
    for p in GBS_CHECKER_INFORMES_DIR.iterdir():
        m = GBS_CHECKER_RE.match(p.name)
        if m:
            fechas.append(m.group(1))
    if not fechas:
        return COLOR_GRAY, "Sin chequeo de GBS generado todavía."
    ultima = max(fechas)
    hoy = datetime.now().strftime("%Y-%m-%d")
    if ultima == hoy:
        return COLOR_GREEN, f"Chequeo de hoy ({hoy}) ya generado."
    return COLOR_YELLOW, (f"Último chequeo: {ultima} (no es de hoy). "
                          "Dale a 'Chequear GBS' para correrlo de nuevo.")


def _check_gbs_maker_status(borradores_dir, desplegados_dir, nombre_boton_generar):
    """Logica compartida del semaforo de GBS Deal Maker / GBS Organization
    Maker: amarillo si hay borradores en la carpeta de revision esperando
    que el responsable comercial los mire y despliegue (o borre los que no quiere); verde si no
    hay ninguno pendiente y ya se desplegó algo alguna vez; gris si todavía
    no se generó nada. No vuelve a generar ni desplegar nada por su cuenta,
    solo lee las carpetas."""
    pendientes = 0
    if borradores_dir.exists():
        pendientes = sum(1 for p in borradores_dir.iterdir()
                         if p.is_file() and p.suffix == ".docx")
    if pendientes:
        return COLOR_YELLOW, (f"{pendientes} borrador(es) esperando revisión "
                              "('Abrir carpeta') antes de 'Desplegar'.")
    if desplegados_dir.exists() and any(desplegados_dir.iterdir()):
        return COLOR_GREEN, f"Sin borradores pendientes. Dale a '{nombre_boton_generar}' para revisar de nuevo."
    return COLOR_GRAY, f"Sin GBS generado todavía. Dale a '{nombre_boton_generar}'."


def check_gbs_maker_status():
    return _check_gbs_maker_status(GBS_MAKER_BORRADORES_DIR, GBS_MAKER_DESPLEGADOS_DIR, "Generar")


def check_gbs_org_maker_status():
    return _check_gbs_maker_status(GBS_ORG_MAKER_BORRADORES_DIR, GBS_ORG_MAKER_DESPLEGADOS_DIR, "Generar")


def check_fw_cadence_status():
    """Semaforo del deck de cadencia Freshworks: verde si ya se genero uno
    este mes, amarillo si el mas reciente es de un mes anterior, gris si
    nunca se genero. Es mensual (no diario, a diferencia de Churn/Metricas)
    y de solo lectura hacia vTiger -- no corre por cron, asi que solo
    refleja la fecha del ultimo HTML guardado en FW-Cadence-Generator/."""
    if not FW_CADENCE_DIR.exists():
        return COLOR_GRAY, "Carpeta del generador no encontrada."
    meses = []
    for p in FW_CADENCE_DIR.iterdir():
        m = FW_CADENCE_RE.match(p.name)
        if m:
            meses.append((m.group(1), p))
    if not meses:
        return COLOR_GRAY, "Sin deck generado todavía."
    mes_actual = datetime.now().strftime("%Y-%m")
    ultimo_mes, ultimo_archivo = max(meses, key=lambda t: t[0])
    if ultimo_mes == mes_actual:
        generado = datetime.fromtimestamp(ultimo_archivo.stat().st_mtime)
        return COLOR_GREEN, f"Deck de {mes_actual} generado el {generado:%d/%m %H:%M}."
    return COLOR_YELLOW, (f"Último deck: {ultimo_mes} (no es de este mes). "
                          f"Dale a 'Generar' para el de {mes_actual}.")


def check_comp_plan_status():
    """Semaforo del plan de compensacion: verde si el archivo del trimestre
    se genero hoy, amarillo si es de otro dia, gris si no existe."""
    hoy = datetime.now()
    q = (hoy.month - 1) // 3 + 1
    nombre = f"Plan Compensacion MID Growth {hoy.year}Q{q}.xlsx"
    archivo = COMP_PLAN_DIR / nombre

    if not archivo.exists():
        return COLOR_GRAY, f"Sin generar para {hoy.year}Q{q}. Dale a 'Generar / actualizar'."

    generado = datetime.fromtimestamp(archivo.stat().st_mtime)
    if generado.date() == hoy.date():
        return COLOR_GREEN, f"{hoy.year}Q{q} actualizado hoy a las {generado:%H:%M}."
    return COLOR_YELLOW, (
        f"{hoy.year}Q{q} generado el {generado:%d/%m/%Y %H:%M}. "
        f"Vuelve a generarlo para incluir deals cerrados despues de esa fecha."
    )


MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def check_metricas_status():
    """Semaforo del reporte ejecutivo del trimestre en curso: verde si se
    genero hoy, amarillo si es de otro dia, gris si no existe."""
    hoy = datetime.now()
    q = (hoy.month - 1) // 3 + 1
    archivo = METRICAS_DIR / f"Metricas Mid-SMB {hoy.year}Q{q}.xlsx"

    if not archivo.exists():
        return COLOR_GRAY, f"Sin generar para {hoy.year}Q{q}. Dale a 'Generar reporte del Q'."

    generado = datetime.fromtimestamp(archivo.stat().st_mtime)
    if generado.date() == hoy.date():
        return COLOR_GREEN, f"{hoy.year}Q{q} al dia, generado hoy a las {generado:%H:%M}."
    return COLOR_YELLOW, (f"{hoy.year}Q{q} generado el {generado:%d/%m/%Y %H:%M}. "
                          f"Regeneralo para incluir lo cerrado desde entonces.")


def check_forecast_status():
    """Semaforo del Forecast: verde si se genero hoy, amarillo si es de otro
    momento, gris si nunca se genero."""
    if not FORECAST_MANIFEST.exists():
        return COLOR_GRAY, "Sin generar. Elegí el mes y dale a 'Generar'."
    try:
        with open(FORECAST_MANIFEST, "r", encoding="utf-8") as f:
            m = json.load(f)
        generado = datetime.fromisoformat(m["generado"])
    except Exception as exc:
        return COLOR_GRAY, f"No se pudo leer el último Forecast: {exc}"
    mes_txt = ", ".join(m.get("periodo", [])) or "?"
    resumen = f"{m.get('upside_count', 0)} upside, {m.get('commit_count', 0)} commit"
    if generado.date() == datetime.now().date():
        return COLOR_GREEN, f"Forecast de {mes_txt} generado hoy a las {generado:%H:%M} ({resumen})."
    return COLOR_YELLOW, (f"Forecast de {mes_txt} generado el {generado:%d/%m/%Y %H:%M} ({resumen}). "
                          f"Regeneralo para incluir cambios recientes.")


def find_pending_next_step_files():
    """Reportes de Next Step (por vendedor + general) generados pero
    todavia no enviados, en el orden en que se enviarian."""
    state = send_ns.load_state()
    processed = set(state.get("processed_files", []))
    if not send_ns.WATCH_FOLDER.exists():
        return []
    current = [p for p in send_ns.WATCH_FOLDER.iterdir() if p.is_file()]
    pending = [p for p in current if p.name not in processed]
    pending.sort(key=lambda p: p.name)
    return pending


def _fecha_reporte_next_step(nombre):
    """Fecha que lleva el nombre del reporte, o None si no es un reporte."""
    m = send_ns.VENDOR_FILE_RE.match(nombre)
    texto = m.group(2) if m else None
    if texto is None:
        m = send_ns.GENERAL_FILE_RE.match(nombre)
        texto = m.group(1) if m else None
    if texto is None:
        return None
    try:
        return datetime.strptime(texto, "%d-%m-%Y").date()
    except ValueError:
        return None


NEXT_STEP_MANIFEST = SCRIPT_DIR / "ultima_generacion_next_step.json"


def leer_manifiesto_next_step():
    if not NEXT_STEP_MANIFEST.exists():
        return None
    try:
        with open(NEXT_STEP_MANIFEST, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def fuentes_next_step_mas_nuevas():
    """(hay_fuentes_nuevas, momento_generacion). Compara los Excel de vTiger en
    Descargas contra la ultima generacion, para no mostrar un resumen viejo
    despues de bajar reportes frescos."""
    manifiesto = leer_manifiesto_next_step()
    if not manifiesto:
        return True, None
    try:
        generado = datetime.fromisoformat(manifiesto["generado"])
    except Exception:
        return True, None
    try:
        fuentes = gen_ns.find_latest_files()
    except Exception:
        return False, generado
    for p in fuentes.values():
        if datetime.fromtimestamp(p.stat().st_mtime) > generado:
            return True, generado
    return False, generado


def find_latest_next_step_files():
    """Reportes de la ultima generacion, se hayan enviado o no.

    Se apoya en el manifiesto que deja el generador, porque ni la lista de
    enviados ni la fecha del nombre alcanzan: el nombre solo lleva el dia (al
    regenerar sale igual y ya figura como enviado) y un vendedor sin alertas no
    produce archivo, de modo que el de una corrida anterior se queda en la
    carpeta y pareceria vigente."""
    manifiesto = leer_manifiesto_next_step()
    if manifiesto:
        archivos = [send_ns.WATCH_FOLDER / n for n in manifiesto.get("archivos", [])]
        archivos = sorted((p for p in archivos if p.exists()), key=lambda p: p.name)
        try:
            generado = datetime.fromisoformat(manifiesto["generado"]).date()
        except Exception:
            generado = None
        if archivos:
            return archivos, generado, manifiesto.get("sin_alertas", [])

    # Respaldo para generaciones anteriores al manifiesto: se agrupa por la
    # fecha del nombre.
    if not send_ns.WATCH_FOLDER.exists():
        return [], None, []
    fechados = []
    for p in send_ns.WATCH_FOLDER.iterdir():
        if not p.is_file():
            continue
        fecha = _fecha_reporte_next_step(p.name)
        if fecha:
            fechados.append((fecha, p))
    if not fechados:
        return [], None, []
    ultima = max(f for f, _ in fechados)
    archivos = sorted((p for f, p in fechados if f == ultima), key=lambda p: p.name)
    return archivos, ultima, []


def _extract_body(html_text):
    m = re.search(r"<body[^>]*>(.*)</body>", html_text, re.DOTALL | re.IGNORECASE)
    return m.group(1) if m else html_text


_ROW_RE = re.compile(r"<tr\b.*?</tr>", re.DOTALL | re.IGNORECASE)
_HEADER_ROW_RE = re.compile(r"<th\b", re.IGNORECASE)


def _esc_html(v):
    return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _esc_attr(v):
    return _esc_html(v).replace('"', "&quot;")


_BLOQUE_RE = re.compile(r'<h2 class="bloque">(.*?)</h2>(.*?)(?=<h2 class="bloque">|\Z)',
                        re.DOTALL | re.IGNORECASE)
_TABLE_RE = re.compile(r"<table>(.*?)</table>", re.DOTALL | re.IGNORECASE)

# Orden y metadata de las 6 categorias que arma generar_reportes_next_step.py
# (build_vendor_html): Deals/Farmings x Next Step/Producto/Atencion. El orden
# de las tuplas debe coincidir exactamente con el orden en que build_vendor_html
# arma las tablas dentro de cada bloque -- se emparejan por posicion, no por
# titulo (varios sub-titulos se repiten entre Deals y Farmings).
CATEGORIAS_ORDEN = ("deals_next_step", "deals_producto", "deals_atencion",
                    "farming_next_step", "farming_producto", "farming_atencion")
_CATEGORIA_INFO = {
    "deals_next_step": ("Deals -- alertas de Next Step",
                        ["Vendedor", "Deal / Organización", "MRR", "Next Step", "ECD", "Corregir"]),
    "deals_producto": ("Deals -- alertas de producto",
                       ["Vendedor", "Deal / Organización", "MRR", "Producto", "Service Type", "Corregir"]),
    "deals_atencion": ("Deals -- sin atención",
                       ["Vendedor", "Deal / Organización", "MRR", "Última actividad", "Corregir"]),
    "farming_next_step": ("Farmings -- alertas de Next Step",
                         ["Vendedor", "Farming / Organización", "Etapa", "Next Step", "Corregir"]),
    "farming_producto": ("Farmings -- sin producto",
                        ["Vendedor", "Farming / Organización", "Etapa", "Corregir"]),
    "farming_atencion": ("Farmings -- sin atención",
                        ["Vendedor", "Farming / Organización", "Etapa", "Última actividad", "Corregir"]),
}


def _identificar_categoria(bloque_titulo, header_row_html):
    """Identifica la categoria de una tabla por el contenido de su fila de
    encabezado, no por su posicion -- build_vendor_html omite por completo
    la tabla de una categoria sin filas (no deja un hueco), asi que si la
    PRIMERA categoria de un bloque esta vacia y la segunda no, emparejar por
    posicion corre las tablas un lugar y las mezcla."""
    if bloque_titulo == "Deals":
        if "Next Step" in header_row_html:
            return "deals_next_step"
        if "Producto" in header_row_html:
            return "deals_producto"
        if "actividad" in header_row_html:
            return "deals_atencion"
    elif bloque_titulo == "Farmings":
        if "Next Step" in header_row_html:
            return "farming_next_step"
        if "actividad" in header_row_html:
            return "farming_atencion"
        return "farming_producto"  # unica sin "Next Step" ni "actividad" en el encabezado
    return None


def _extract_category_rows(html_text):
    """Lee la estructura Deals/Farmings x Next Step/Producto/Atencion que
    arma build_vendor_html y devuelve {categoria: [<tr> de datos, sin header]}."""
    body = _extract_body(html_text)
    result = {c: [] for c in CATEGORIAS_ORDEN}
    for m in _BLOQUE_RE.finditer(body):
        titulo = re.sub(r"<.*?>", "", m.group(1)).strip()
        if titulo not in ("Deals", "Farmings"):
            continue
        for tabla in _TABLE_RE.finditer(m.group(2)):
            filas = list(_ROW_RE.finditer(tabla.group(1)))
            if not filas:
                continue
            header_row = filas[0].group(0)
            clave = _identificar_categoria(titulo, header_row)
            if clave is None:
                continue
            result[clave] = [
                r.group(0) for r in filas
                if not _HEADER_ROW_RE.search(r.group(0))
            ]
    return result


def _tag_vendor_row(row_html, vendor_name):
    """Le agrega al <tr> una celda con el nombre del vendedor + un
    data-vendor para poder filtrar con JS en la tabla consolidada."""
    vendor_esc = _esc_html(vendor_name)
    vendor_attr = _esc_attr(vendor_name)
    return re.sub(
        r"^<tr\b[^>]*>", f'<tr data-vendor="{vendor_attr}"><td>{vendor_esc}</td>', row_html, count=1,
    )


def build_consolidated_preview(pending_paths):
    """Resumen de Calidad de CRM: arriba el Account Book Attention general
    (% del equipo), despues un total por categoria, y abajo las 4 tablas
    (Deals/Farmings x Next Step/Producto) con todos los vendedores juntos y
    un filtro desplegable por vendedor."""
    # Si hay varios dias sin enviar, se acumulan copias viejas (mismo
    # vendedor o el general varias veces). Solo interesa la mas reciente de
    # cada uno para el resumen -- las anteriores quedan superadas por esa.
    latest_vendor_file = {}  # vendor_name -> (fecha, path)
    latest_general_file = None  # (fecha, path)
    other_files = []

    for p in pending_paths:
        m_vendor = send_ns.VENDOR_FILE_RE.match(p.name)
        m_general = send_ns.GENERAL_FILE_RE.match(p.name)
        if m_vendor:
            vendor_name, date_str = m_vendor.group(1), m_vendor.group(2)
            fecha = datetime.strptime(date_str, "%d-%m-%Y")
            best = latest_vendor_file.get(vendor_name)
            if best is None or fecha > best[0]:
                latest_vendor_file[vendor_name] = (fecha, p)
        elif m_general:
            date_str = m_general.group(1)
            fecha = datetime.strptime(date_str, "%d-%m-%Y")
            if latest_general_file is None or fecha > latest_general_file[0]:
                latest_general_file = (fecha, p)
        else:
            other_files.append(p)

    general_sections = []
    if latest_general_file:
        p = latest_general_file[1]
        html = p.read_text(encoding="utf-8")
        general_sections.append(
            f'<div style="max-width:900px;margin:0 auto 8px;padding:0 24px;'
            f'font-family:Segoe UI, sans-serif;font-size:11px;color:#5B5F6B;">{p.name}</div>'
            f'<div style="margin-bottom:24px;">{_extract_body(html)}</div>'
        )
    for p in other_files:
        general_sections.append(_extract_body(p.read_text(encoding="utf-8")))

    rows_by_categoria = {c: [] for c in CATEGORIAS_ORDEN}
    vendor_names = []
    for vendor_name, (_, p) in sorted(latest_vendor_file.items()):
        vendor_names.append(vendor_name)
        extraido = _extract_category_rows(p.read_text(encoding="utf-8"))
        for cat in CATEGORIAS_ORDEN:
            rows_by_categoria[cat].extend(_tag_vendor_row(r, vendor_name) for r in extraido[cat])

    general_html = "\n".join(general_sections)
    total_por_categoria = {c: len(rows_by_categoria[c]) for c in CATEGORIAS_ORDEN}
    total_general = sum(total_por_categoria.values())

    if vendor_names:
        options = "".join(
            f'<option value="{_esc_attr(v)}">{_esc_html(v)}</option>' for v in sorted(vendor_names)
        )
        tiles_html = "".join(
            f'''<div style="background:#FFFFFF;border-radius:16px;padding:16px 20px;flex:1;min-width:160px;">
              <div style="font-size:24px;font-weight:600;color:#12112C;">{total_por_categoria[c]}</div>
              <div style="font-size:12px;color:#5B5F6B;margin-top:2px;">{_CATEGORIA_INFO[c][0]}</div>
            </div>'''
            for c in CATEGORIAS_ORDEN
        )
        resumen_html = f"""
        <div style="max-width:900px;margin:0 auto;padding:0 24px;font-family:Segoe UI, sans-serif;">
          <h2 style="font-size:20px;color:#12112C;margin:24px 0 12px;">Calidad de CRM -- resumen general</h2>
          <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;">{tiles_html}</div>
          <div style="margin-bottom:16px;">
            <label style="font-size:12px;color:#5B5F6B;margin-right:8px;">Filtrar por vendedor:</label>
            <select id="vendor-filter" onchange="filterVendor(this.value)"
                    style="padding:6px 10px;border-radius:8px;border:1px solid #DEE2E6;font-family:Segoe UI, sans-serif;font-size:13px;">
              <option value="">Todos ({total_general} registro(s))</option>
              {options}
            </select>
            <span id="filter-count" style="font-size:12px;color:#5B5F6B;margin-left:10px;"></span>
          </div>
        </div>
        """

        tablas_html = []
        for cat in CATEGORIAS_ORDEN:
            filas = rows_by_categoria[cat]
            if not filas:
                continue
            titulo, encabezados = _CATEGORIA_INFO[cat]
            ths = "".join(
                f'<th style="text-align:left;font-size:11px;text-transform:uppercase;'
                f'color:#495057;background:#F1F3F5;padding:12px 14px;">{h}</th>'
                for h in encabezados
            )
            tablas_html.append(f"""
            <div style="max-width:900px;margin:0 auto 32px;padding:0 24px;font-family:Segoe UI, sans-serif;">
              <h3 style="font-size:15px;color:#5B5F6B;margin:0 0 10px;">{titulo} ({len(filas)})</h3>
              <div style="background:#FFFFFF;border-radius:24px;padding:8px 24px;">
                <table style="width:100%;border-collapse:collapse;font-size:13px;" class="crm-table" data-cat="{cat}">
                  <tr>{ths}</tr>
                  {''.join(filas)}
                </table>
              </div>
            </div>
            """)

        table_section = resumen_html + "\n".join(tablas_html) + """
        <script>
        function filterVendor(value) {
          var rows = document.querySelectorAll('.crm-table tr[data-vendor]');
          var shown = 0;
          rows.forEach(function(row) {
            var visible = (value === '' || row.dataset.vendor === value);
            row.style.display = visible ? '' : 'none';
            if (visible) shown++;
          });
          document.getElementById('filter-count').textContent =
            value === '' ? '' : (shown + ' registro(s)');
        }
        </script>
        """
    else:
        table_section = "<p style='padding:24px;font-family:Segoe UI, sans-serif;color:#5B5F6B;'>Sin deals ni farmings pendientes con alerta.</p>"

    combined_body = general_html + table_section
    if not pending_paths:
        combined_body = "<p style='padding:40px;font-family:Segoe UI, sans-serif;'>Nada pendiente.</p>"

    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8">
<title>Calidad de CRM -- resumen (vista previa)</title>{gen_ns.STYLE}</head>
<body>{combined_body}</body></html>"""


# ---------------- UI (sistema de diseno GB Advisors) ----------------

import ctypes
import platform
import shutil
import tkinter.font as tkfont

BRAND_DIR = SCRIPT_DIR / "brand"

# Tokens de marca (colors_and_type.css)
GBA_MAGENTA = "#EA018B"
GBA_MAGENTA_700 = "#B00467"
GBA_MAGENTA_100 = "#FFEBF7"
GBA_INK = "#12112C"
GBA_50 = "#F8F9FA"
GBA_100 = "#F1F3F5"
GBA_200 = "#E9ECEF"
GBA_300 = "#DEE2E6"
GBA_500 = "#ADB5BD"
GBA_700 = "#495057"
GBA_WHITE = "#FFFFFF"

BG_APP = GBA_50
CARD_BG = GBA_WHITE
CARD_BORDER = GBA_200

RADIUS = 16      # un solo rung de radio para toda la composicion
RULE_H = 10      # la pastilla magenta de 10 px, motivo de marca
PAD = 16
GAP = 16
PADX = 24
CONSOLE_MIN_W = 380
CARD_MIN_W = 340

_FONT_FAMILY = "Segoe UI"
_FONT_MEDIUM = "Segoe UI"


def activar_dpi():
    """Declara la app conscientes del DPI antes de crear la ventana. Sin esto
    Windows estira la ventana como bitmap al moverla entre monitores con
    escalados distintos, y el texto se ve borroso o encimado."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def icono_de_ventana(root, ruta_ico):
    """Cuelga el .ico de la ventana en la medida exacta que pide Windows.

    Hace falta ademas de iconbitmap(): aquel solo deja el icono de CLASE, y la
    barra de tareas pide el de VENTANA. Medido el 5-oct-2026, WM_GETICON
    devolvia 0 en las tres variantes, asi que Windows estiraba lo que encontraba
    y el icono salia borroso.

    Las medidas se le preguntan al sistema en vez de fijar 32: con la pantalla
    al 125% o 150% la barra pide 40 o 48 px, y un valor fijo reproduciria el
    mismo defecto un escalon mas arriba.
    """
    if os.name != "nt" or not ruta_ico.exists():
        return False
    try:
        import ctypes

        u = ctypes.windll.user32
        IMAGE_ICON, LR_LOADFROMFILE = 1, 0x0010
        WM_SETICON, ICON_SMALL, ICON_BIG = 0x0080, 0, 1
        SM_CXICON, SM_CYICON, SM_CXSMICON, SM_CYSMICON = 11, 12, 49, 50

        hwnd = int(root.wm_frame(), 16)
        medidas = (
            (ICON_BIG, u.GetSystemMetrics(SM_CXICON), u.GetSystemMetrics(SM_CYICON)),
            (ICON_SMALL, u.GetSystemMetrics(SM_CXSMICON), u.GetSystemMetrics(SM_CYSMICON)),
        )
        # Los handles se guardan en el root: si Python los recolecta, Windows
        # destruye el icono y la ventana se queda sin el.
        root._iconos_win32 = []
        for cual, ancho, alto in medidas:
            h = u.LoadImageW(None, str(ruta_ico), IMAGE_ICON, ancho, alto,
                             LR_LOADFROMFILE)
            if not h:
                continue
            root._iconos_win32.append(h)
            u.SendMessageW(hwnd, WM_SETICON, cual, h)
        return bool(root._iconos_win32)
    except Exception:
        return False


def identidad_barra_tareas():
    """Declara un AppUserModelID propio antes de crear la ventana. Sin esto
    Windows agrupa la ventana bajo pythonw.exe y la barra de tareas muestra
    el icono de Python, no el de GB Advisors, aunque iconbitmap haya puesto
    bien el icono del titulo (9-sep-2026, pedido de el responsable comercial). El ID es una
    cadena arbitraria pero estable: si cambia, Windows lo trata como otra
    aplicacion y pierde el anclado a la barra."""
    if platform.system() != "Windows":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "GBAdvisors.PanelAutomatizaciones"
        )
    except Exception:
        pass


def load_brand_fonts():
    """Carga Inter para esta sesion. En Windows es un registro privado del
    proceso (FR_PRIVATE, no instala nada); en Linux no existe ese mecanismo,
    asi que se copia una vez a ~/.local/share/fonts y se refresca fontconfig.
    Si falla, se queda con Segoe UI, el fallback del token."""
    global _FONT_FAMILY, _FONT_MEDIUM
    carpeta = BRAND_DIR / "fonts"
    if not carpeta.exists():
        return
    if platform.system() == "Windows":
        try:
            gdi = ctypes.WinDLL("gdi32")
            for ttf in carpeta.glob("*.ttf"):
                gdi.AddFontResourceExW(ctypes.c_wchar_p(str(ttf.resolve())), 0x10, 0)
        except Exception:
            return
    else:
        _instalar_fuentes_linux(carpeta)
    disponibles = set(tkfont.families())
    if "Inter 18pt" in disponibles:
        _FONT_FAMILY = "Inter 18pt"
        _FONT_MEDIUM = "Inter 18pt Medium" if "Inter 18pt Medium" in disponibles else "Inter 18pt"


def _instalar_fuentes_linux(carpeta):
    """Copia los .ttf de marca a la carpeta de fuentes del usuario (si aun no
    estan) y refresca fontconfig para que Tk las vea sin reiniciar sesion."""
    destino = Path.home() / ".local" / "share" / "fonts" / "GB-Advisors"
    try:
        destino.mkdir(parents=True, exist_ok=True)
        copiado = False
        for ttf in carpeta.glob("*.ttf"):
            objetivo = destino / ttf.name
            if not objetivo.exists():
                shutil.copy2(ttf, objetivo)
                copiado = True
        if copiado:
            subprocess.run(["fc-cache", "-f", str(destino)], check=False,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def f_reg(size):
    return (_FONT_FAMILY, size)


def f_med(size):
    return (_FONT_MEDIUM, size)


def f_mono(size):
    return ("Consolas", size)


def alto_linea(fuente):
    """Alto real de una linea, en pixeles. Toda la geometria se calcula con
    esto en vez de pixeles fijos: con 'tk scaling' distinto de 1 una fuente de
    tamano 12 puede ocupar 21 px, y las coordenadas fijas se enciman."""
    try:
        return tkfont.Font(family=fuente[0], size=fuente[1]).metrics("linespace")
    except Exception:
        return int(fuente[1] * 1.6)


VENTANAS_STATE_FILE = SCRIPT_DIR / "state_ventanas.json"


def _cargar_geometrias_guardadas():
    if not VENTANAS_STATE_FILE.exists():
        return {}
    try:
        with open(VENTANAS_STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar_geometria_ventana(clave, tamano):
    datos = _cargar_geometrias_guardadas()
    if datos.get(clave) == tamano:
        return
    datos[clave] = tamano
    try:
        with open(VENTANAS_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def recordar_tamano(ventana, clave, tamano_por_defecto):
    """Aplica el tamano que el usuario dejo la ultima vez que movio/estiro
    esta ventana (si existe uno guardado), y de ahi en adelante graba
    cualquier resize para la proxima corrida -- pedido explicito de el responsable comercial
    (28-ago-2026): el panel no debe volver a su tamano por defecto en cada
    arranque. Solo se guarda ancho x alto, no la posicion en pantalla (para
    no abrir fuera de la pantalla si cambia el monitor/resolucion)."""
    guardadas = _cargar_geometrias_guardadas()
    ventana.geometry(guardadas.get(clave, tamano_por_defecto))

    pendiente = {"id": None}

    def _al_cambiar(evento):
        if evento.widget is not ventana:
            return
        if pendiente["id"] is not None:
            ventana.after_cancel(pendiente["id"])
        pendiente["id"] = ventana.after(
            500,
            lambda: _guardar_geometria_ventana(
                clave, f"{ventana.winfo_width()}x{ventana.winfo_height()}"),
        )

    ventana.bind("<Configure>", _al_cambiar, add="+")


def _round_rect(canvas, x1, y1, x2, y2, r, **kw):
    r = max(2, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, splinesteps=36, **kw)


class RoundedCard(tk.Frame):
    """Tarjeta de superficie elevada que se redibuja al cambiar de tamano, de
    modo que sigue viendose bien si mueves o estiras la ventana."""

    def __init__(self, parent, bg=CARD_BG, border=CARD_BORDER, pad=PAD):
        super().__init__(parent, bg=BG_APP)
        self._bg, self._border, self._pad = bg, border, pad
        # width/height en 1: un Canvas sin tamano explicito pide 378x265 y esa
        # medida se impone sobre la rejilla, desbordando las filas.
        self.canvas = tk.Canvas(self, bg=BG_APP, highlightthickness=0, bd=0,
                                width=1, height=1)
        self.canvas.pack(fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=bg)
        self._shape = None
        self._win = self.canvas.create_window(pad, pad, anchor="nw", window=self.body)
        self.canvas.bind("<Configure>", self._redibujar)

    def _redibujar(self, evento):
        w, h = evento.width, evento.height
        if w < 8 or h < 8:
            return
        if self._shape is not None:
            self.canvas.delete(self._shape)
        self._shape = _round_rect(self.canvas, 1, 1, w - 1, h - 1, RADIUS,
                                  fill=self._bg, outline=self._border)
        self.canvas.tag_lower(self._shape)
        self.canvas.itemconfig(self._win, width=max(w - self._pad * 2, 1),
                               height=max(h - self._pad * 2, 1))


class PillButton(tk.Canvas):
    """Boton plano de marca. Su alto sale de la metrica de la fuente, asi que
    no se recorta ni se encima con otro escalado de pantalla."""

    def __init__(self, parent, text, command, variant="secondary", surface=CARD_BG,
                font_size=10, pad_x=30, pad_y=16):
        self.variant = variant
        self.command = command
        self.enabled = True
        self._font = f_med(font_size)
        fuente = tkfont.Font(family=self._font[0], size=self._font[1])
        ancho = fuente.measure(text) + pad_x
        alto = fuente.metrics("linespace") + pad_y
        super().__init__(parent, width=ancho, height=alto, bg=surface,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._ancho, self._alto = ancho, alto
        self._shape = _round_rect(self, 1, 1, ancho - 1, alto - 1, RADIUS,
                                  fill=self._fill(), outline=self._outline())
        self._label = self.create_text(ancho / 2, alto / 2, text=text,
                                       font=self._font, fill=self._fg())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _fill(self, hover=False):
        if not self.enabled:
            return GBA_100 if self.variant != "ghost" else ""
        if self.variant == "primary":
            return GBA_MAGENTA_700 if hover else GBA_MAGENTA
        if self.variant == "ghost":
            return "#2A2851" if hover else ""
        return GBA_200 if hover else GBA_100

    def _outline(self):
        if self.variant == "ghost":
            return "#4A4878" if self.enabled else "#33315C"
        return "" if self.variant == "primary" else GBA_300

    def _fg(self):
        if not self.enabled:
            return GBA_500
        return GBA_WHITE if self.variant in ("primary", "ghost") else GBA_INK

    def _repaint(self, hover=False):
        self.itemconfig(self._shape, fill=self._fill(hover), outline=self._outline())
        self.itemconfig(self._label, fill=self._fg())

    def _on_enter(self, _):
        if self.enabled:
            self._repaint(True)

    def _on_leave(self, _):
        self._repaint(False)

    def _on_click(self, _):
        if self.enabled and self.command:
            self.command()

    def set_enabled(self, valor):
        self.enabled = bool(valor)
        self.config(cursor="hand2" if self.enabled else "arrow")
        self._repaint(False)


class ToggleChip(tk.Canvas):
    """Chip on/off dibujado igual que PillButton (mismo radio, misma fuente,
    mismo magenta), para elegir entre opciones sin recurrir a un
    ttk.Combobox/tk.Checkbutton nativo -- esos rompen el estilo de marca del
    resto del panel."""

    def __init__(self, parent, text, command=None, active=False, surface=CARD_BG):
        self.command = command
        self.active = active
        self._font = f_med(10)
        fuente = tkfont.Font(family=self._font[0], size=self._font[1])
        ancho = fuente.measure(text) + 26
        alto = fuente.metrics("linespace") + 14
        super().__init__(parent, width=ancho, height=alto, bg=surface,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._ancho, self._alto = ancho, alto
        self._shape = _round_rect(self, 1, 1, ancho - 1, alto - 1, RADIUS,
                                  fill=self._fill(), outline=self._outline())
        self._label = self.create_text(ancho / 2, alto / 2, text=text,
                                       font=self._font, fill=self._fg())
        self.bind("<Enter>", lambda e: self._repaint(True))
        self.bind("<Leave>", lambda e: self._repaint(False))
        self.bind("<Button-1>", self._on_click)

    def _fill(self, hover=False):
        if self.active:
            return GBA_MAGENTA_700 if hover else GBA_MAGENTA
        return GBA_200 if hover else GBA_100

    def _outline(self):
        return "" if self.active else GBA_300

    def _fg(self):
        return GBA_WHITE if self.active else GBA_INK

    def _repaint(self, hover=False):
        self.itemconfig(self._shape, fill=self._fill(hover), outline=self._outline())
        self.itemconfig(self._label, fill=self._fg())

    def _on_click(self, _):
        self.set_active(not self.active)
        if self.command:
            self.command(self.active)

    def set_active(self, valor):
        self.active = bool(valor)
        self._repaint(False)


class QuarterStepper(tk.Frame):
    """Selector de trimestre con flechas, hermano de MonthStepper.

    Va de TRIMESTRES_ATRAS atras hasta el trimestre en curso: hacia adelante no
    tiene sentido, porque no hay datos de un trimestre que no ha empezado
    (1-oct-2026)."""

    def __init__(self, parent, surface=CARD_BG, on_change=None):
        super().__init__(parent, bg=surface)
        self._max = trimestre_actual()
        self._min = trimestre_desplazado(*self._max, -TRIMESTRES_ATRAS)
        self._anio, self._q = self._max
        self.on_change = on_change

        self._atras = PillButton(self, "\u2039", lambda: self._mover(-1),
                                 "secondary", surface=surface)
        self._atras.pack(side="left")
        self._lbl = tk.Label(self, text="", font=f_med(10), bg=surface, fg=GBA_INK,
                             width=9, anchor="center")
        self._lbl.pack(side="left", padx=4)
        self._adelante = PillButton(self, "\u203a", lambda: self._mover(1),
                                    "secondary", surface=surface)
        self._adelante.pack(side="left")
        self._pintar()

    def _mover(self, pasos):
        nuevo = trimestre_desplazado(self._anio, self._q, pasos)
        if not (self._min <= nuevo <= self._max):
            return
        self._anio, self._q = nuevo
        self._pintar()
        if self.on_change:
            self.on_change(self.get())

    def _pintar(self):
        actual = (self._anio, self._q) == self._max
        self._lbl.config(text=self.get() + (" •" if actual else ""))
        self._atras.set_enabled(trimestre_desplazado(self._anio, self._q, -1) >= self._min)
        self._adelante.set_enabled(trimestre_desplazado(self._anio, self._q, 1) <= self._max)

    def get(self):
        return trimestre_clave(self._anio, self._q)

    def es_actual(self):
        return (self._anio, self._q) == self._max


class MonthStepper(tk.Frame):
    """Selector de un mes con flechas, en vez de un desplegable. El panel no
    tiene ningun widget de lista y un ttk.Combobox rompe el estilo de marca;
    treinta casillas tampoco cabian. Con dos flechas y la etiqueta del mes se
    elige cualquier mes del rango sin construir un desplegable entero
    (16-sep-2026)."""

    def __init__(self, parent, clave_inicial, surface=BG_APP, on_change=None):
        super().__init__(parent, bg=surface)
        hoy = datetime.now()
        self._hoy = hoy
        self._min = mes_desplazado(hoy.year, hoy.month, -FORECAST_MESES_ATRAS)
        self._max = mes_desplazado(hoy.year, hoy.month, FORECAST_MESES_ADELANTE)
        self._anio, self._mes = mes_desde_clave(clave_inicial, hoy)
        self.on_change = on_change

        self._btn_atras = PillButton(self, "\u2039", lambda: self._mover(-1),
                                     "secondary", surface=surface)
        self._btn_atras.pack(side="left")
        self._lbl = tk.Label(self, text="", font=f_med(11), bg=surface, fg=GBA_INK,
                             width=16, anchor="center")
        self._lbl.pack(side="left", padx=6)
        self._btn_adelante = PillButton(self, "\u203a", lambda: self._mover(1),
                                        "secondary", surface=surface)
        self._btn_adelante.pack(side="left")
        self._pintar()

    def _mover(self, pasos):
        nuevo = mes_desplazado(self._anio, self._mes, pasos)
        if not (self._min <= nuevo <= self._max):
            return
        self._anio, self._mes = nuevo
        self._pintar()
        if self.on_change:
            self.on_change(self.get())

    def _pintar(self):
        self._lbl.config(text=mes_etiqueta(self.get(), self._hoy))
        self._btn_atras.set_enabled(mes_desplazado(self._anio, self._mes, -1) >= self._min)
        self._btn_adelante.set_enabled(mes_desplazado(self._anio, self._mes, 1) <= self._max)

    def get(self):
        return mes_clave(self._anio, self._mes)


class Checkbox(tk.Canvas):
    """Checkbox real (caja + tilde) dibujado a mano, para listas de opciones
    donde un chip tipo boton se confunde con una accion -- esto se lee como
    lo que es, una casilla para marcar, sin caer en el tk.Checkbutton nativo
    (aspecto Windows 98)."""

    BOX = 16

    def __init__(self, parent, text, command=None, checked=False, surface=CARD_BG):
        self.command = command
        self.checked = checked
        self._font = f_reg(10)
        fuente = tkfont.Font(family=self._font[0], size=self._font[1])
        alto = max(self.BOX + 4, fuente.metrics("linespace") + 4)
        ancho = self.BOX + 8 + fuente.measure(text) + 4
        super().__init__(parent, width=ancho, height=alto, bg=surface,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._y0 = (alto - self.BOX) / 2
        self._box = _round_rect(self, 1, self._y0, self.BOX + 1, self._y0 + self.BOX, 5,
                                fill=self._fill(), outline=self._outline())
        x1, y1 = 4, self._y0 + self.BOX * 0.55
        x2, y2 = self.BOX * 0.42 + 1, self._y0 + self.BOX - 3
        x3, y3 = self.BOX - 2, self._y0 + 3
        self._check = self.create_line(x1, y1, x2, y2, x3, y3, fill=GBA_WHITE, width=2,
                                       capstyle="round", joinstyle="round",
                                       state="normal" if checked else "hidden")
        self.create_text(self.BOX + 9, alto / 2, anchor="w", text=text,
                         font=self._font, fill=GBA_INK)
        self.bind("<Button-1>", self._on_click)

    def _fill(self):
        return GBA_MAGENTA if self.checked else CARD_BG

    def _outline(self):
        return "" if self.checked else GBA_300

    def _on_click(self, _):
        self.set_checked(not self.checked)
        if self.command:
            self.command(self.checked)

    def set_checked(self, valor):
        self.checked = bool(valor)
        self.itemconfig(self._box, fill=self._fill(), outline=self._outline())
        self.itemconfig(self._check, state="normal" if self.checked else "hidden")


class StatusCard(RoundedCard):
    """Punto de semaforo, titulo, detalle y acciones.

    El detalle vive en un contenedor de alto fijo (LINEAS_DETALLE lineas) para
    que un mensaje largo nunca se derrame sobre los botones: se recorta con
    elipsis y el texto completo siempre queda en la consola de actividad."""

    LINEAS_DETALLE = 3

    def __init__(self, parent, titulo, horizontal=False):
        super().__init__(parent)
        self.horizontal = horizontal
        self.body.columnconfigure(0, weight=1)
        h_texto = alto_linea(f_reg(9))
        lineas = 2 if horizontal else self.LINEAS_DETALLE

        cab = self._cab = tk.Frame(self.body, bg=CARD_BG)
        self.dot = tk.Canvas(cab, width=12, height=12, bg=CARD_BG,
                             highlightthickness=0, bd=0)
        self._dot_id = self.dot.create_oval(1, 1, 11, 11, fill=GBA_500, outline="")
        self.dot.pack(side="left", pady=(alto_linea(f_med(12)) // 3, 0))
        tk.Label(cab, text=titulo, font=f_med(12), bg=CARD_BG, fg=GBA_INK,
                 anchor="w").pack(side="left", padx=(9, 0))

        self.caja_detalle = tk.Frame(self.body, bg=CARD_BG, height=h_texto * lineas)
        self.caja_detalle.pack_propagate(False)
        self.detalle = tk.Label(self.caja_detalle, text="Revisando…", font=f_reg(9),
                                bg=CARD_BG, fg=GBA_700, justify="left", anchor="nw")
        self.detalle.pack(fill="both", expand=True, anchor="nw")
        self.button_row = tk.Frame(self.body, bg=CARD_BG)

        if horizontal:
            cab.grid(row=0, column=0, sticky="ew")
            self.caja_detalle.grid(row=1, column=0, sticky="new", pady=(6, 0))
            self.button_row.grid(row=0, column=1, rowspan=2, sticky="e", padx=(GAP, 0))
        else:
            cab.grid(row=0, column=0, sticky="ew")
            self.caja_detalle.grid(row=1, column=0, sticky="new", pady=(6, 6))
            self.button_row.grid(row=2, column=0, sticky="sw")
        self.body.rowconfigure(1, weight=1)

        self._lineas = lineas
        self._lineas_visibles = lineas
        self._wrap = 0
        self._texto = "Revisando…"
        self.body.bind("<Configure>", self._ajustar_wrap)

    def _ajustar_wrap(self, evento):
        # Ojo: cget("wraplength") devuelve un Tcl_Obj, no un int; comparar
        # contra el rompia el callback y el texto quedaba sin envolver.
        reserva = self.button_row.winfo_reqwidth() + GAP if self.horizontal else 0
        ancho = max(evento.width - reserva, 120)
        if abs(self._wrap - ancho) > 4:
            self._wrap = ancho
            self.detalle.config(wraplength=ancho)
            self._pintar_texto()

        # El detalle se ajusta al alto que sobra. Si se deja fijo y la tarjeta
        # queda corta, el texto se dibuja encima de los botones.
        h_linea = alto_linea(f_reg(9))
        ocupado = self._cab.winfo_reqheight() + 12
        if not self.horizontal:
            ocupado += self.button_row.winfo_reqheight()
        libre = evento.height - ocupado
        alto = max(h_linea, min(h_linea * self._lineas, libre))
        if abs(self.caja_detalle.winfo_reqheight() - alto) > 2:
            self.caja_detalle.config(height=alto)
            self._lineas_visibles = max(1, alto // h_linea)
            self._pintar_texto()

    def _pintar_texto(self):
        self.detalle.config(text=self._recortar(self._texto))

    def _recortar(self, texto):
        """Deja el texto en las lineas que caben, con elipsis si sobra."""
        if not self._wrap:
            return texto
        try:
            fuente = tkfont.Font(family=f_reg(9)[0], size=f_reg(9)[1])
            ancho_char = max(fuente.measure("abcdefghijklmnopqrstuvwxyz") / 26, 1)
        except Exception:
            return texto
        cabe = int(self._wrap / ancho_char) * self._lineas_visibles
        return texto if len(texto) <= cabe else texto[:max(cabe - 1, 8)].rstrip() + "…"

    def set_status(self, color, texto):
        self.dot.itemconfig(self._dot_id, fill=color)
        self._texto = texto
        self._pintar_texto()


class Hero(tk.Frame):
    """Banda superior sobre el gradiente oscuro de marca, con la pastilla
    magenta de 10 px como firma. Se redibuja al cambiar de ancho."""

    def __init__(self, parent, on_refresh, on_vacaciones=None):
        self.h_titulo = alto_linea(f_med(21))
        self.h_sub = alto_linea(f_reg(10))
        alto = PAD + self.h_titulo + 8 + RULE_H + 10 + self.h_sub + PAD
        super().__init__(parent, bg=GBA_INK, height=alto)
        self.pack_propagate(False)
        self._alto = alto
        self.canvas = tk.Canvas(self, bg=GBA_INK, highlightthickness=0, bd=0,
                                width=1, height=1)
        self.canvas.pack(fill="both", expand=True)

        # El equipo manda: antes estaba escrito a mano y quedaba mintiendo en
        # cuanto el panel se usaba para otro equipo (5-oct-2026).
        _eq = gen_ns.NOMBRE_PANEL
        self._contexto = f"{_eq} · GB Advisors" if _eq else "GB Advisors"
        self._img = None
        self._fondo_src = None
        fondo = BRAND_DIR / "gradient-dark.png"
        if fondo.exists():
            try:
                from PIL import Image
                self._fondo_src = Image.open(fondo).convert("RGB")
            except Exception:
                self._fondo_src = None

        self._logo_img = None
        self._logo_w = 0
        logo = BRAND_DIR / "brand-mark-white.png"
        if logo.exists():
            try:
                from PIL import Image, ImageTk
                src = Image.open(logo).convert("RGBA")
                logo_h = int(self.h_titulo * 0.9)
                logo_w = max(int(logo_h * src.width / src.height), 1)
                self._logo_img = ImageTk.PhotoImage(
                    src.resize((logo_w, logo_h), Image.LANCZOS))
                self._logo_w = logo_w
            except Exception:
                self._logo_img = None

        titulo_x = PADX + (self._logo_w + 10 if self._logo_img else 0)
        self._logo_id = None
        if self._logo_img is not None:
            self._logo_id = self.canvas.create_image(
                PADX, 0, anchor="w", image=self._logo_img)
        self._img_id = None
        self._titulo_id = self.canvas.create_text(
            titulo_x, 0, anchor="w", text="Panel de automatizaciones",
            font=f_med(21), fill=GBA_WHITE)
        self._pill_id = None
        self.sello = self.canvas.create_text(PADX, 0, anchor="w", text=self._contexto,
                                             font=f_reg(10), fill="#C9C7E0")
        botones = tk.Frame(self, bg=GBA_INK)
        if on_vacaciones is not None:
            self.btn_vacaciones = PillButton(botones, "Vacaciones", on_vacaciones, "ghost", surface=GBA_INK)
            self.btn_vacaciones.pack(side="left", padx=(0, 8))
        self.btn_refresh = PillButton(botones, "Actualizar", on_refresh, "ghost", surface=GBA_INK)
        self.btn_refresh.pack(side="left")
        self._btn_id = self.canvas.create_window(0, 0, anchor="e", window=botones)

        self._pendiente = None
        self.canvas.bind("<Configure>", self._on_configure)

    def _on_configure(self, evento):
        w, h = evento.width, evento.height
        y_titulo = PAD + self.h_titulo / 2
        y_pill = PAD + self.h_titulo + 8
        y_sello = y_pill + RULE_H + 10 + self.h_sub / 2

        if self._logo_id is not None:
            self.canvas.coords(self._logo_id, PADX, y_titulo)
        titulo_x = PADX + (self._logo_w + 10 if self._logo_img is not None else 0)
        self.canvas.coords(self._titulo_id, titulo_x, y_titulo)
        self.canvas.coords(self.sello, PADX, y_sello)
        self.canvas.coords(self._btn_id, w - PADX, h / 2)
        if self._pill_id is not None:
            self.canvas.delete(self._pill_id)
        self._pill_id = _round_rect(self.canvas, PADX, y_pill, PADX + 64, y_pill + RULE_H,
                                    RULE_H / 2, fill=GBA_MAGENTA, outline="")
        if self._logo_id is not None:
            self.canvas.tag_raise(self._logo_id)
        self.canvas.tag_raise(self._titulo_id)
        self.canvas.tag_raise(self.sello)

        # El gradiente se reescala con retardo: es una imagen de 1920x1080 y
        # rehacerla en cada pixel de arrastre trabaria la ventana.
        if self._pendiente:
            self.after_cancel(self._pendiente)
        self._pendiente = self.after(120, lambda: self._pintar_fondo(w, h))

    def _pintar_fondo(self, w, h):
        self._pendiente = None
        if not self._fondo_src or w < 8 or h < 8:
            return
        try:
            from PIL import Image, ImageTk
            src = self._fondo_src
            escala = max(w / src.width, 1)
            alto_escalado = max(int(src.height * escala), h)
            im = src.resize((max(w, 1), alto_escalado), Image.LANCZOS)
            centro = int(alto_escalado * 0.62)
            top = max(0, min(centro - h // 2, alto_escalado - h))
            self._img = ImageTk.PhotoImage(im.crop((0, top, w, top + h)))
        except Exception:
            return
        if self._img_id is not None:
            self.canvas.delete(self._img_id)
        self._img_id = self.canvas.create_image(0, 0, anchor="nw", image=self._img)
        self.canvas.tag_lower(self._img_id)

    def set_sello(self, texto):
        self.canvas.itemconfig(self.sello, text=f"{self._contexto} · {texto}")


REASON_LABELS = {
    "ya_asignado": "ya está en la invitación",
    "turno_reunion": "por turno (libre a esa hora)",
    "turno_sin_reunion": "por turno (sin reunión)",
}


def format_plan_lines(plan):
    """Arma el plan calculado por Sofia como lineas de texto planas (para la
    consola de actividad), en vez de una ventana aparte."""
    entries = plan.get("entries", [])
    blocked = plan.get("blocked", [])
    lineas = [f"Revisar antes de enviar -- se enviarían {len(entries)} asignación(es):"]

    if plan.get("vacaciones"):
        lineas.append("  De vacaciones (excluidos): " + ", ".join(plan["vacaciones"]))
    if plan.get("fuera_de_horario"):
        lineas.append("  Aviso: hay asignaciones fuera de horario laboral (lun-vie 8am-6pm).")
    for b in blocked:
        lineas.append(f"  Pendiente {b['file_name']}: {b['motivo']}")

    for e in entries:
        fecha = ""
        if e.get("meeting_start"):
            try:
                fecha = datetime.fromisoformat(e["meeting_start"]).strftime("%d/%m %H:%M")
            except ValueError:
                fecha = e["meeting_start"]
        reunion = (f"{fecha} {e['meeting_subject']}".strip()
                   if e.get("meeting_subject") else "sin reunión")
        envio = "invitación + archivo" if e.get("send_invite") else "solo archivo"
        motivo = REASON_LABELS.get(e["reason"], e["reason"])
        lineas.append(f"  - {e['company']} -> {e['seller_name']} ({motivo}). "
                      f"Reunión: {reunion}. Envío: {envio}. Archivo: {e['file_name']}")

    if entries:
        lineas.append("Cada asignación con reunión se envía en 2 correos aparte: "
                      "1) la invitación sola, sin adjunto; 2) el archivo de presentación.")
        lineas.append("Revisa la lista de arriba y dale a 'Enviar' para mandarlo tal cual.")
    else:
        lineas.append("No hay nada que enviar con este plan.")
    return lineas


def format_next_step_plan_lines(plan):
    """Arma el plan de envio de Next Step (calculado por
    enviar_reportes_next_step.py --plan-json) como lineas de texto planas
    para la consola de actividad -- mismo criterio que format_plan_lines
    para Sofia, pero con el esquema {"items": [...]} de este plan."""
    items = plan.get("items", [])
    a_enviar = [i for i in items if i.get("accion") == "enviar"]
    a_omitir = [i for i in items if i.get("accion") == "omitir"]

    if not items:
        return ["No hay reportes nuevos por enviar."]

    lineas = [f"Plan de envío -- se mandarían {len(a_enviar)} correo(s):"]
    for item in a_enviar:
        if item.get("tipo") == "general":
            destino = f"{len(item.get('destinatarios', []))} vendedores"
        else:
            destino = item.get("email", "?")
        lineas.append(f"  - {item['archivo']} -> {destino} ({item.get('asunto', '')})")
    for item in a_omitir:
        lineas.append(f"  - {item['archivo']}: OMITIR ({item.get('motivo', '')})")
    if plan.get("desactualizados"):
        lineas.append("  Desactualizado(s), se ignoran: " + ", ".join(plan["desactualizados"]))

    if a_enviar:
        lineas.append("Revisa la lista de arriba y dale a 'Enviar' para mandarlo tal cual.")
    else:
        lineas.append("No hay nada que enviar con este plan (todo omitido o vacío).")
    return lineas


# El forecast se manda a Freshworks, asi que las etiquetas van en ingles.
FORECAST_MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"]
# Hasta donde llega el selector de meses. Hacia atras para poder rehacer un
# forecast de un mes ya cerrado, hacia adelante para planificar: antes del
# 16-sep-2026 solo se podia elegir el mes actual y el proximo, y el responsable comercial pidio
# poder abrir la ventana a mas tiempo.
FORECAST_MESES_ATRAS = 6
FORECAST_MESES_ADELANTE = 24


def mes_clave(anio, mes):
    return f"{anio:04d}-{mes:02d}"


def mes_desde_clave(clave, hoy=None):
    """Convierte 'YYYY-MM' en (anio, mes). Acepta tambien las dos claves
    relativas antiguas por si quedo alguna guardada de una version previa."""
    hoy = hoy or datetime.now()
    if clave == "actual":
        return hoy.year, hoy.month
    if clave == "proximo":
        return (hoy.year + 1, 1) if hoy.month == 12 else (hoy.year, hoy.month + 1)
    anio, mes = clave.split("-")
    return int(anio), int(mes)


def mes_etiqueta(clave, hoy=None):
    anio, mes = mes_desde_clave(clave, hoy)
    return f"{FORECAST_MONTH_NAMES[mes - 1]} {anio}"


# Hasta cuantos trimestres atras se puede pedir un reporte. El 1-oct-2026,
# primer dia de Q4, el responsable comercial se encontro con que el plan de compensacion y las
# metricas salian en cero y sin forma de mirar Q3: los scripts ya aceptaban
# --trimestre, pero el panel siempre pedia el actual.
TRIMESTRES_ATRAS = 8


def trimestre_actual(hoy=None):
    hoy = hoy or datetime.now()
    return hoy.year, (hoy.month - 1) // 3 + 1


def trimestre_desplazado(anio, q, pasos):
    total = anio * 4 + (q - 1) + pasos
    return total // 4, total % 4 + 1


def trimestre_clave(anio, q):
    return f"{anio}Q{q}"


def mes_desplazado(anio, mes, pasos):
    total = anio * 12 + (mes - 1) + pasos
    return total // 12, total % 12 + 1


def meses_del_rango(desde, hasta, hoy=None):
    """Lista de claves 'YYYY-MM' entre ambos extremos, inclusive. Si vienen
    al reves se ordenan, igual que hace el generador."""
    a = mes_desde_clave(desde, hoy)
    b = mes_desde_clave(hasta, hoy)
    if a > b:
        a, b = b, a
    claves = []
    anio, mes = a
    while (anio, mes) <= b:
        claves.append(mes_clave(anio, mes))
        anio, mes = mes_desplazado(anio, mes, 1)
    return claves


class ForecastOptionsWindow(tk.Toplevel):
    """Que mes(es)/categorias/segmento entran en el Forecast antes de
    generarlo -- el filtro queda aplicado en los datos del archivo, no solo
    escondido en el HTML, porque este reporte se manda a Freshworks. Todo
    con checkboxes reales (Checkbox), no chips tipo boton -- esto es una
    lista de opciones para marcar, no una fila de acciones."""

    def __init__(self, panel, desde_activo, hasta_activo, categorias_activas,
                segmentos_activos, reps_activos, servicios_activo, on_apply):
        super().__init__(panel.root)
        self.on_apply = on_apply
        self.title("Forecast options")
        self.configure(bg=BG_APP)
        self.transient(panel.root)
        self.resizable(False, False)
        self.lift()
        self.grab_set()

        pad = 20

        def _seccion(texto):
            tk.Label(self, text=texto, font=f_med(13), bg=BG_APP,
                     fg=GBA_INK).pack(anchor="w", padx=pad, pady=(pad, 8))

        def _fila():
            f = tk.Frame(self, bg=BG_APP)
            f.pack(anchor="w", padx=pad, pady=(0, 4))
            return f

        _seccion("Month range (both ends included)")
        fila_mes = _fila()
        tk.Label(fila_mes, text="From", font=f_reg(10), bg=BG_APP,
                 fg=GBA_700).pack(side="left", padx=(0, 8))
        self._desde = MonthStepper(fila_mes, desde_activo, surface=BG_APP,
                                   on_change=lambda _v: self._pintar_rango())
        self._desde.pack(side="left", padx=(0, 18))
        tk.Label(fila_mes, text="to", font=f_reg(10), bg=BG_APP,
                 fg=GBA_700).pack(side="left", padx=(0, 8))
        self._hasta = MonthStepper(fila_mes, hasta_activo, surface=BG_APP,
                                   on_change=lambda _v: self._pintar_rango())
        self._hasta.pack(side="left")
        self._lbl_rango = tk.Label(self, text="", font=f_reg(10), bg=BG_APP,
                                   fg=GBA_700)
        self._lbl_rango.pack(anchor="w", padx=pad, pady=(6, 0))

        _seccion("Categories")
        cats_frame = tk.Frame(self, bg=BG_APP)
        cats_frame.pack(anchor="w", padx=pad)
        self._cat_boxes = {}
        for i, clave in enumerate(FORECAST_CATEGORY_OPTIONS):
            box = Checkbox(cats_frame, FORECAST_CATEGORY_LABELS[clave], None,
                          checked=(clave in categorias_activas), surface=BG_APP)
            box.grid(row=i // 3, column=i % 3, padx=(0, 18), pady=(0, 8), sticky="w")
            self._cat_boxes[clave] = box

        _seccion("Segment")
        seg_frame = tk.Frame(self, bg=BG_APP)
        seg_frame.pack(anchor="w", padx=pad)
        self._seg_boxes = {}
        for i, etiqueta in enumerate(FORECAST_SEGMENT_OPTIONS):
            box = Checkbox(seg_frame, etiqueta, None,
                          checked=(etiqueta in segmentos_activos), surface=BG_APP)
            box.grid(row=0, column=i, padx=(0, 18), sticky="w")
            self._seg_boxes[etiqueta] = box

        _seccion("Rep")
        rep_frame = tk.Frame(self, bg=BG_APP)
        rep_frame.pack(anchor="w", padx=pad)
        self._rep_boxes = {}
        for i, rep in enumerate(gen_ns.TEAM_OWNERS):
            box = Checkbox(rep_frame, rep, None,
                          checked=(rep in reps_activos), surface=BG_APP)
            box.grid(row=i // 2, column=i % 2, padx=(0, 18), pady=(0, 8), sticky="w")
            self._rep_boxes[rep] = box

        _seccion("Services")
        fila_serv = _fila()
        self._servicios_box = Checkbox(fila_serv, "Include GB Services", None,
                                       checked=servicios_activo, surface=BG_APP)
        self._servicios_box.pack(side="left")

        fila_botones = tk.Frame(self, bg=BG_APP)
        fila_botones.pack(anchor="e", padx=pad, pady=(18, pad))
        PillButton(fila_botones, "Cancel", self.destroy, "secondary", surface=BG_APP).pack(side="left")
        PillButton(fila_botones, "Apply", self._aplicar, "primary", surface=BG_APP).pack(side="left", padx=(8, 0))

    def _pintar_rango(self):
        """Cuantos meses abarca el rango elegido, para que se vea de una que
        'From September 2026 to December 2026' son cuatro meses."""
        meses = meses_del_rango(self._desde.get(), self._hasta.get())
        if len(meses) == 1:
            texto = f"1 month: {mes_etiqueta(meses[0])}"
        else:
            texto = (f"{len(meses)} months: {mes_etiqueta(meses[0])} "
                     f"\u2192 {mes_etiqueta(meses[-1])}")
        self._lbl_rango.config(text=texto)

    def _aplicar(self):
        categorias = {k for k, box in self._cat_boxes.items() if box.checked}
        segmentos = {k for k, box in self._seg_boxes.items() if box.checked}
        reps = {k for k, box in self._rep_boxes.items() if box.checked}
        servicios = self._servicios_box.checked
        # Se normaliza aqui: si el usuario dejo el "hasta" antes del "desde",
        # se ordenan en vez de generar un forecast vacio.
        desde, hasta = sorted([self._desde.get(), self._hasta.get()])
        self.on_apply(desde, hasta, categorias, segmentos, reps, servicios)
        self.destroy()


def cargar_limite_gbs_org():
    """Cuantas organizaciones entran por corrida, tal como quedo la ultima
    vez (se guarda para no tener que elegirlo de nuevo en cada arranque)."""
    if GBS_ORG_LIMITE_FILE.exists():
        try:
            with open(GBS_ORG_LIMITE_FILE, "r", encoding="utf-8") as f:
                valor = int(json.load(f).get("limite", GBS_ORG_LIMITE_DEFAULT))
            return max(1, min(valor, GBS_ORG_LIMITE_MAX))
        except Exception:
            pass
    return GBS_ORG_LIMITE_DEFAULT


def guardar_limite_gbs_org(valor):
    try:
        with open(GBS_ORG_LIMITE_FILE, "w", encoding="utf-8") as f:
            json.dump({"limite": int(valor)}, f)
    except Exception:
        pass


class CantidadOrgsWindow(tk.Toplevel):
    """Cuantas organizaciones procesa la proxima corrida de "Generar" del
    GBS Organization Maker. Existe porque el lote completo son ~150 y cada
    una consume creditos de Claude: conviene ir por tandas en vez de
    dispararlas todas de una (pedido de el responsable comercial, 28-ago-2026). El tope duro es
    GBS_ORG_LIMITE_MAX; el planificador ademas ordena la cola por cantidad
    de deals, asi que una tanda corta se lleva las que mas insumo tienen,
    no las primeras alfabeticamente."""

    PRESETS = (5, 10, 25, 50, 100)

    def __init__(self, panel, valor_actual, on_apply):
        super().__init__(panel.root)
        self.on_apply = on_apply
        self.title("Cantidad de organizaciones")
        self.configure(bg=BG_APP)
        self.transient(panel.root)
        self.resizable(False, False)
        self.lift()
        self.grab_set()

        pad = 20
        tk.Label(self, text="¿Cuántas organizaciones procesar por corrida?", font=f_med(13),
                 bg=BG_APP, fg=GBA_INK).pack(anchor="w", padx=pad, pady=(pad, 4))
        tk.Label(self, text=f"Cada organización consume créditos de Claude. Máximo "
                            f"{GBS_ORG_LIMITE_MAX} por corrida.\nEntran primero las que más "
                            "deals tienen (las que no tienen ninguno no generan nada).",
                 font=f_reg(9), bg=BG_APP, fg=GBA_700,
                 justify="left").pack(anchor="w", padx=pad, pady=(0, 14))

        fila_entry = tk.Frame(self, bg=BG_APP)
        fila_entry.pack(anchor="w", padx=pad)
        self.var = tk.StringVar(value=str(valor_actual))
        self.entry = tk.Entry(fila_entry, textvariable=self.var, font=f_reg(13), width=5,
                              bg=GBA_WHITE, fg=GBA_INK, relief="flat", justify="center",
                              highlightthickness=1, highlightbackground=GBA_300,
                              highlightcolor=GBA_MAGENTA)
        self.entry.pack(side="left", ipady=7)
        tk.Label(fila_entry, text="organización(es)", font=f_reg(10), bg=BG_APP,
                 fg=GBA_700).pack(side="left", padx=(10, 0))

        fila_presets = tk.Frame(self, bg=BG_APP)
        fila_presets.pack(anchor="w", padx=pad, pady=(12, 0))
        for n in self.PRESETS:
            PillButton(fila_presets, str(n), lambda v=n: self._set_valor(v),
                       surface=BG_APP, font_size=9, pad_x=16, pad_y=9).pack(side="left", padx=(0, 6))

        self.aviso = tk.Label(self, text="", font=f_reg(9), bg=BG_APP, fg=GBA_MAGENTA_700)
        self.aviso.pack(anchor="w", padx=pad, pady=(10, 0))

        fila_botones = tk.Frame(self, bg=BG_APP)
        fila_botones.pack(anchor="e", padx=pad, pady=(12, pad))
        PillButton(fila_botones, "Cancelar", self.destroy, "secondary",
                   surface=BG_APP).pack(side="left")
        PillButton(fila_botones, "Aplicar", self._aplicar, "primary",
                   surface=BG_APP).pack(side="left", padx=(8, 0))

        self.entry.focus_set()
        self.entry.select_range(0, "end")
        self.bind("<Return>", lambda e: self._aplicar())
        self.bind("<Escape>", lambda e: self.destroy())

    def _set_valor(self, valor):
        self.var.set(str(valor))
        self.aviso.config(text="")

    def _aplicar(self):
        texto = self.var.get().strip()
        if not texto.isdigit() or not (1 <= int(texto) <= GBS_ORG_LIMITE_MAX):
            self.aviso.config(text=f"Escribí un número entre 1 y {GBS_ORG_LIMITE_MAX}.")
            return
        self.on_apply(int(texto))
        self.destroy()


class SofiaAIWindow(tk.Toplevel):
    """Ventana 'Sofia - AI': agrupa los Reportes para vendedores y la
    Reasignación de orgs de Sofia detrás de un solo botón en la grilla
    principal, en vez de una tarjeta cada uno (pedido de el responsable comercial, 28-ago-2026 --
    separar las opciones en un menu previo). Se crea una sola vez al
    arrancar el panel y se oculta/muestra con withdraw()/deiconify() -- nunca
    se destruye -- para que sus tarjetas sigan vivas y el refresh_status() de
    siempre las siga actualizando aunque la ventana este cerrada."""

    def __init__(self, panel):
        super().__init__(panel.root)
        self.panel = panel
        self.title("Sofia - AI")
        self.configure(bg=BG_APP)
        self.transient(panel.root)
        self.minsize(480, 780)
        recordar_tamano(self, "sofia_ai", "560x820")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        contenedor = tk.Frame(self, bg=BG_APP)
        contenedor.pack(fill="both", expand=True, padx=20, pady=(20, 0))
        contenedor.columnconfigure(0, weight=1)
        for fila in range(4):
            contenedor.rowconfigure(fila, weight=1)

        self.sofia_card = StatusCard(contenedor, "Reportes para vendedores (Sofia)")
        self.sofia_card.grid(row=0, column=0, sticky="nsew", pady=(0, 16))
        self.reasigna_card = StatusCard(contenedor, "Reasignación de orgs (Sofia)")
        self.reasigna_card.grid(row=1, column=0, sticky="nsew", pady=(0, 16))
        self.validar_reasigna_card = StatusCard(contenedor, "Validar creación de deals")
        self.validar_reasigna_card.grid(row=2, column=0, sticky="nsew", pady=(0, 16))
        self.rutinas_card = StatusCard(contenedor, "Rutinas automáticas (Sofia)")
        self.rutinas_card.grid(row=3, column=0, sticky="nsew")

        fila_botones = tk.Frame(self, bg=BG_APP)
        fila_botones.pack(anchor="e", padx=20, pady=20)
        PillButton(fila_botones, "Cerrar", self.withdraw, "secondary",
                  surface=BG_APP).pack(side="left")

        self.withdraw()

    def abrir(self):
        self.deiconify()
        self.lift()
        self.focus_force()


class VtigerAIWindow(tk.Toplevel):
    """Ventana 'vTiger Análisis': agrupa Calidad CRM, Métricas del
    trimestre, Análisis de Churn, GBS Deal Checker, GBS Deal Maker y GBS
    Organization Maker detrás de un solo botón en la grilla principal --
    mismo patrón que SofiaAIWindow (pedido de el responsable comercial, 28-ago-2026: separar las
    opciones en un menu previo). Se crea una sola vez al arrancar el panel y
    se oculta/muestra con withdraw()/deiconify() -- nunca se destruye --
    para que sus tarjetas sigan vivas y el refresh_status() de siempre las
    siga actualizando aunque la ventana este cerrada."""

    def __init__(self, panel):
        super().__init__(panel.root)
        self.panel = panel
        self.title("vTiger Análisis")
        self.configure(bg=BG_APP)
        self.transient(panel.root)
        # Grilla de 3x2 (28-ago-2026, pedido de el responsable comercial: "cada segmento al lado
        # del otro", la version anterior de 2x2 con 4 tarjetas ya quedaba
        # bien -- al sumar las 2 tarjetas de GBS Maker se paso a 3 columnas
        # en vez de agregar una tercera fila, para seguir ancha y no alta.
        self.minsize(980, 560)
        recordar_tamano(self, "vtiger_ai", "1220x680")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        contenedor = tk.Frame(self, bg=BG_APP)
        contenedor.pack(fill="both", expand=True, padx=20, pady=(20, 0))
        contenedor.columnconfigure(0, weight=1)
        contenedor.columnconfigure(1, weight=1)
        contenedor.columnconfigure(2, weight=1)
        contenedor.rowconfigure(0, weight=1)
        contenedor.rowconfigure(1, weight=1)

        self.ns_card = StatusCard(contenedor, "Calidad CRM")
        self.ns_card.grid(row=0, column=0, sticky="nsew", padx=(0, 16), pady=(0, 16))
        self.met_card = StatusCard(contenedor, "Métricas del trimestre")
        self.met_card.grid(row=0, column=1, sticky="nsew", padx=(0, 16), pady=(0, 16))
        self.churn_card = StatusCard(contenedor, "Análisis de Churn")
        self.churn_card.grid(row=0, column=2, sticky="nsew", pady=(0, 16))
        self.gbs_card = StatusCard(contenedor, "GBS Deal Checker")
        self.gbs_card.grid(row=1, column=0, sticky="nsew", padx=(0, 16))
        self.gbs_maker_card = StatusCard(contenedor, "GBS Deal Maker")
        self.gbs_maker_card.grid(row=1, column=1, sticky="nsew", padx=(0, 16))
        self.gbs_org_maker_card = StatusCard(contenedor, "GBS Organization Maker")
        self.gbs_org_maker_card.grid(row=1, column=2, sticky="nsew")

        fila_botones = tk.Frame(self, bg=BG_APP)
        fila_botones.pack(anchor="e", padx=20, pady=20)
        PillButton(fila_botones, "Cerrar", self.withdraw, "secondary",
                  surface=BG_APP).pack(side="left")

        self.withdraw()

    def abrir(self):
        self.deiconify()
        self.lift()
        self.focus_force()


class FreshworksWindow(tk.Toplevel):
    """Ventana 'Freshworks': agrupa Forecast y Cadence Generator detrás de
    un solo botón en la grilla principal -- mismo patrón que SofiaAIWindow /
    VtigerAIWindow (pedido de el responsable comercial, 28-ago-2026). Se crea una sola vez al
    arrancar el panel y se oculta/muestra con withdraw()/deiconify() -- nunca
    se destruye -- para que sus tarjetas sigan vivas y el refresh_status() de
    siempre las siga actualizando aunque la ventana este cerrada."""

    def __init__(self, panel):
        super().__init__(panel.root)
        self.panel = panel
        self.title("Freshworks")
        self.configure(bg=BG_APP)
        self.transient(panel.root)
        self.minsize(520, 520)
        recordar_tamano(self, "freshworks", "620x700")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        contenedor = tk.Frame(self, bg=BG_APP)
        contenedor.pack(fill="both", expand=True, padx=20, pady=(20, 0))
        contenedor.columnconfigure(0, weight=1)
        contenedor.rowconfigure(0, weight=1)
        contenedor.rowconfigure(1, weight=1)

        self.forecast_card = StatusCard(contenedor, "Forecast")
        self.forecast_card.grid(row=0, column=0, sticky="nsew", pady=(0, 16))
        self.cadence_card = StatusCard(contenedor, "Cadence Generator")
        self.cadence_card.grid(row=1, column=0, sticky="nsew")

        fila_botones = tk.Frame(self, bg=BG_APP)
        fila_botones.pack(anchor="e", padx=20, pady=20)
        PillButton(fila_botones, "Cerrar", self.withdraw, "secondary",
                  surface=BG_APP).pack(side="left")

        self.withdraw()

    def abrir(self):
        self.deiconify()
        self.lift()
        self.focus_force()


class Panel:
    def __init__(self, root):
        self.root = root
        load_brand_fonts()

        # El equipo va en el titulo: con dos paneles instalados, las ventanas
        # se veian iguales y no habia forma de saber cual era cual.
        _nombre = gen_ns.NOMBRE_PANEL
        root.title(f"{_nombre} — Panel de automatizaciones" if _nombre
                   else "Panel de automatizaciones — GB Advisors")
        root.configure(bg=BG_APP)
        icono_ico = SCRIPT_DIR / "gb_icon.ico"
        icono_png = SCRIPT_DIR / "gb_icon.png"
        # En Windows manda el .ico y nada mas: trae las seis medidas
        # (16 a 256) y Windows elige la exacta para el titulo, la barra de
        # tareas y Alt+Tab. Va con default= para que las Toplevel (Forecast
        # options, Sofia AI, etc.) lo hereden sin repetir la llamada.
        # El PNG queda solo para Linux: es de 256 px y una sola medida, asi
        # que si se aplicara en Windows por iconphoto -- que corre despues y
        # sobreescribe a iconbitmap -- el icono de 16 px saldria del
        # reescalado de esos 256 y se veria sucio (9-sep-2026).
        icono_ok = False
        if platform.system() == "Windows" and icono_ico.exists():
            try:
                root.iconbitmap(default=str(icono_ico))
                icono_ok = True
            except Exception:
                pass
        # Fuera de Windows iconbitmap no sirve (X11/Linux lo rechaza con
        # TclError porque .ico no es un bitmap XBM): ahi hay que entregarle
        # un PhotoImage via iconphoto. Se guarda en self._icon_img Y en
        # root._icon_img (dos referencias, por si alguna se pierde) porque
        # si el PhotoImage de Python se recolecta, Tk borra la imagen y la
        # ventana se queda sin icono aunque iconphoto haya corrido bien.
        icono_fuente = icono_png if icono_png.exists() else icono_ico
        if not icono_ok and icono_fuente.exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(icono_fuente)
                self._icon_img = ImageTk.PhotoImage(img)
                root._icon_img = self._icon_img
                root.iconphoto(True, self._icon_img)
            except Exception:
                pass

        # Ademas del icono de clase que deja iconbitmap, se le cuelga a la
        # VENTANA el suyo en la medida exacta: es el que mira la barra de
        # tareas. update_idletasks primero, porque wm_frame necesita que la
        # ventana ya exista de verdad.
        if icono_ico.exists():
            root.update_idletasks()
            icono_de_ventana(root, icono_ico)

        self._setup_style()

        # Alturas derivadas de la metrica real de la fuente, no de pixeles
        # fijos: asi la ventana se ve igual con cualquier escalado de pantalla.
        h_titulo = alto_linea(f_med(12))
        h_texto = alto_linea(f_reg(9))
        h_boton = alto_linea(f_med(10)) + 16
        # Con holgura: sin ella el contenido calza exacto y el detalle termina
        # dibujandose sobre los botones.
        card_h = PAD * 2 + h_titulo + 8 + h_texto * 3 + 10 + h_boton + 6
        # Grilla de 2 filas x 2 columnas desde el 28-ago-2026: los 3 grupos
        # (Sofia - AI, vTiger Analisis, Freshworks) agrupan varias tarjetas
        # cada uno detras de un solo boton lanzador (ver SofiaAIWindow /
        # VtigerAIWindow / FreshworksWindow), y Plan de compensacion queda
        # solo -- los 4 son tarjetas simples de un solo boton en la grilla
        # principal, altura card_h normal (Forecast, que necesitaba una fila
        # mas alta, ahora vive dentro de FreshworksWindow, no aca).
        # La tarjeta de Presentacion de cliente no necesita las 3 lineas de
        # texto de estado de las demas (no tiene semaforo), pero si el alto de
        # su fila de controles, que lleva un campo de texto. Sin reservarle
        # alto propio la fila solo recibia las sobras y la tarjeta quedaba de
        # 1 pixel: presente en la grilla pero invisible (23-sep-2026).
        deck_h = PAD * 2 + h_titulo + 10 + h_boton + 14
        grid_h = card_h * 2 + deck_h + GAP * 2

        root.rowconfigure(1, weight=1)
        root.columnconfigure(0, weight=1)

        self.hero = Hero(root, self.refresh_status, self.abrir_vacaciones)
        self.hero.grid(row=0, column=0, sticky="ew")

        cuerpo = tk.Frame(root, bg=BG_APP)
        cuerpo.grid(row=1, column=0, sticky="nsew", padx=PADX, pady=PADX)
        cuerpo.columnconfigure(0, weight=3, minsize=CARD_MIN_W)
        cuerpo.columnconfigure(1, weight=3, minsize=CARD_MIN_W)
        cuerpo.columnconfigure(2, weight=2, minsize=CONSOLE_MIN_W)
        # 2 filas con weight=1: asi el espacio extra al agrandar la ventana
        # (o maximizarla) se reparte entre las 4 tarjetas de la grilla.
        cuerpo.rowconfigure(0, weight=1, minsize=card_h)
        cuerpo.rowconfigure(1, weight=1, minsize=card_h)
        # Fila 2: "Presentación de cliente" a lo ancho de las dos columnas. Va
        # ancha porque lleva un campo de texto para pegar el link, que en media
        # columna quedaria inservible. weight=0 para que al agrandar la ventana
        # el espacio extra siga yendo a las 4 tarjetas de arriba, pero con
        # minsize propio: sin el, la fila se quedaba sin alto (23-sep-2026).
        cuerpo.rowconfigure(2, weight=0, minsize=deck_h)

        def celda_grid(widget, fila, col, **kw):
            widget.grid(row=fila, column=col, sticky="nsew",
                        padx=(0, GAP), pady=(0, GAP), **kw)

        def filas_botones(card):
            """Dos filas de botones apiladas dentro de card.button_row, para
            tarjetas con mas botones de los que entran en una sola fila sin
            desbordar la tarjeta (Sofia y Calidad CRM, 28-ago-2026: con 3-4
            botones en una columna angosta -- la mitad del ancho del panel,
            no a ancho completo como Forecast -- el ultimo boton quedaba
            empujado fuera de la tarjeta y no se veia)."""
            fila_a = tk.Frame(card.button_row, bg=CARD_BG)
            fila_a.pack(anchor="w")
            fila_b = tk.Frame(card.button_row, bg=CARD_BG)
            fila_b.pack(anchor="w", pady=(6, 0))
            return fila_a, fila_b

        # --- Sofia - AI (28-ago-2026) ---
        # Agrupa "Reportes para vendedores" y "Reasignacion de orgs" detras
        # de un solo boton lanzador en la grilla, en vez de una tarjeta cada
        # uno -- pedido explicito de el responsable comercial (separar las opciones en un menu
        # previo). La ventana en si (con las dos tarjetas reales, sus
        # semaforos y sus botones) vive en SofiaAIWindow; aca solo se crea y
        # se guardan referencias a sus tarjetas para que el resto del codigo
        # de abajo (botones, refresh_status) siga igual que antes.
        self.sofia_ai_window = SofiaAIWindow(self)
        self.sofia_ai_card = StatusCard(cuerpo, "Sofia - AI")
        celda_grid(self.sofia_ai_card, 0, 0)
        PillButton(self.sofia_ai_card.button_row, "Abrir",
                  self.sofia_ai_window.abrir, "primary").pack(side="left")

        # --- vTiger Análisis (28-ago-2026) ---
        # Agrupa Calidad CRM, Metricas del trimestre y Analisis de Churn
        # detras de un solo boton lanzador -- mismo patron que Sofia - AI.
        # La ventana en si vive en VtigerAIWindow; aca solo se crea y se
        # guardan referencias a sus tarjetas.
        self.vtiger_ai_window = VtigerAIWindow(self)
        self.vtiger_ai_card = StatusCard(cuerpo, "vTiger Análisis")
        celda_grid(self.vtiger_ai_card, 0, 1)
        PillButton(self.vtiger_ai_card.button_row, "Abrir",
                  self.vtiger_ai_window.abrir, "primary").pack(side="left")

        # --- Sofia ---
        # Migrado a Claude Code (skill "sofia-schedule-vendedores"), agosto 2026:
        # este flujo dependia de win32com/Outlook COM y no corre en Ubuntu.
        # Botones reales de vuelta el 27-ago-2026, mismo patron "Calcular
        # plan -> Enviar" que Calidad CRM -- pero aca AMBOS pasos disparan
        # Claude Code sin supervision (headless), no solo el envio: calcular
        # el plan requiere cruzar el calendario de Outlook y la
        # disponibilidad de cada vendedor (find_meeting_availability), algo
        # a lo que este panel Python no tiene acceso directo (solo Claude,
        # via MCP de M365). Ver calcular_plan_sofia / confirmar_envio_plan.
        # Vive dentro de SofiaAIWindow desde el 28-ago-2026 (ver arriba), ya
        # no directo en la grilla del panel.
        self.sofia_card = self.sofia_ai_window.sofia_card
        self._plan_pendiente = None
        sofia_fila_a, sofia_fila_b = filas_botones(self.sofia_card)
        # Fila C: el envio, separado de los dos pasos previos.
        sofia_fila_c = tk.Frame(self.sofia_card.button_row, bg=CARD_BG)
        sofia_fila_c.pack(anchor="w", pady=(6, 0))

        # ORDEN: 1) Revisar reuniones  2) Calcular plan  3) Ver plan  4) Enviar.
        # Antes "Calcular plan" iba primero y "Revisar reuniones" despues, que es
        # justo al reves de como funciona: revisar reuniones es lo que GENERA los
        # decks de las reuniones nuevas, asi que un plan calculado antes no puede
        # verlos. El 28-sep-2026 el plan se calculo a las 10:56 y el deck de
        # Latin American School aparecio a las 10:59: quedo fuera del plan y esa
        # reunion no se envio. Ademas de reordenarlos, "Revisar reuniones" encadena
        # el calculo del plan al terminar, para que no dependa de que alguien
        # recuerde el orden.
        self.btn_revisar_reuniones = PillButton(sofia_fila_a, "1. Revisar reuniones",
                                                self.revisar_reuniones_sofia, "primary")
        self.btn_revisar_reuniones.pack(side="left")
        self.btn_calc_sofia = PillButton(sofia_fila_a, "2. Calcular plan",
                                         self.revisar_plan_sofia)
        self.btn_calc_sofia.pack(side="left", padx=(8, 0))
        self.btn_ver_resumen_sofia = PillButton(sofia_fila_b, "3. Ver plan",
                                                self.ver_resumen_sofia)
        self.btn_ver_resumen_sofia.pack(side="left")
        self.btn_ver_resumen_sofia.set_enabled(False)
        self.btn_confirm_sofia = PillButton(sofia_fila_c, "4. Enviar",
                                            self.confirmar_envio_plan, "primary")
        self.btn_confirm_sofia.pack(side="left")
        self.btn_confirm_sofia.set_enabled(False)

        # --- Calidad CRM (antes "Next Step") ---
        # El envio vuelve a ser local y con adjunto desde el 9-sep-2026
        # (pedido de el responsable comercial): corre enviar_reportes_next_step.py, que manda por
        # Outlook de escritorio via COM adjuntando el HTML. Antes pasaba por
        # Claude Code + MCP de M365, que no soporta adjuntos y por eso subia
        # los reportes a SharePoint y mandaba solo un enlace.
        # Ese paso por Claude se habia adoptado el 26-ago-2026 al retirar
        # Thunderbird -compose, que duplicaba los correos (mismo Message-ID,
        # segundos de diferencia en los Sent Items reales); la duplicacion
        # era de Thunderbird al enviar, no del panel, y Outlook COM no
        # comparte ese mecanismo: mail_helper.send_mail llama mail.Send() una
        # sola vez por correo.
        # Vive dentro de VtigerAIWindow desde el 28-ago-2026, ya no directo
        # en la grilla del panel (ver bloque "vTiger Análisis" mas arriba).
        self.ns_card = self.vtiger_ai_window.ns_card
        ns_fila_a, ns_fila_b = filas_botones(self.ns_card)
        self.btn_gen_ns = PillButton(
            ns_fila_a, "Generar",
            self.generar_reportes_ns)
        self.btn_gen_ns.pack(side="left")
        self.btn_review_ns = PillButton(ns_fila_a, "Ver resumen",
                                        self.revisar_resumen_next_step)
        self.btn_review_ns.pack(side="left", padx=(8, 0))
        # Flujo secuencial (igual que Forecast): Calcular envío -> se
        # habilita Enviar solo si el plan tiene algo por mandar. Los dos
        # pasos son el mismo script local (--plan-json para calcular, sin
        # argumentos para enviar), asi que lo que sale es exactamente el plan
        # que se mostro. El clic en "Enviar" ES la confirmacion explicita;
        # por eso Enviar solo se habilita despues de calcular un plan
        # fresco, nunca de entrada.
        self._plan_ns_pendiente = None
        self.btn_calc_envio_ns = PillButton(ns_fila_b, "Calcular envío",
                                            self.calcular_envio_ns)
        self.btn_calc_envio_ns.pack(side="left")
        self.btn_enviar_ns = PillButton(ns_fila_b, "Enviar",
                                        self.enviar_plan_ns, "primary")
        self.btn_enviar_ns.pack(side="left", padx=(8, 0))
        self.btn_enviar_ns.set_enabled(False)

        # --- Freshworks (28-ago-2026) ---
        # Agrupa Forecast y Cadence Generator detras de un solo boton
        # lanzador -- mismo patron que Sofia - AI / vTiger Analisis. La
        # ventana en si vive en FreshworksWindow; aca solo se crea y se
        # guardan referencias a sus tarjetas.
        # La ventana se construye SIEMPRE aunque la tarjeta no se muestre: mas
        # abajo se toman referencias a sus tarjetas (forecast_card,
        # cadence_card) y sin ella habria que desmontar medio archivo.
        self.freshworks_window = FreshworksWindow(self)
        if gen_ns.MOSTRAR_FRESHWORKS:
            self.freshworks_card = StatusCard(cuerpo, "Freshworks")
            celda_grid(self.freshworks_card, 1, 1)
            PillButton(self.freshworks_card.button_row, "Abrir",
                       self.freshworks_window.abrir, "primary").pack(side="left")

        # --- Presentacion de cliente a demanda (23-sep-2026) ---
        # Aparte de la rutina de Sofia a proposito: aquella busca reuniones en
        # el calendario y respeta la regla de no duplicar de 7 dias; esta la
        # pide el responsable comercial para una cuenta concreta y la genera aunque exista una
        # reciente. Y va por ID de registro, no por nombre: la rutina resuelve
        # la cuenta con LIKE '%nombre%' y se queda con el mejor parecido, que
        # con nombres similares puede ser la cuenta equivocada.
        self.deck_card = StatusCard(cuerpo, "Presentación de cliente")
        self.deck_card.grid(row=2, column=0, columnspan=2, sticky="nsew",
                            padx=(0, GAP), pady=(0, GAP))
        # StatusCard ancla su button_row al oeste (sticky="sw"), que es lo
        # correcto para una fila de botones: no se estira. Aqui si hace falta,
        # porque el campo del link tiene que quedarse con el ancho sobrante en
        # vez de un tamano fijo. Se cambia solo en ESTA tarjeta (23-sep-2026).
        self.deck_card.button_row.grid_configure(sticky="sew")
        fila_deck = tk.Frame(self.deck_card.button_row, bg=CARD_BG)
        fila_deck.pack(anchor="w", fill="x", expand=True)
        tk.Label(fila_deck, text="Link de vTiger:", font=f_reg(10), bg=CARD_BG,
                 fg=GBA_700).pack(side="left", padx=(0, 8))
        # Los botones se anclan a la derecha y el campo se queda con lo que
        # sobre: con el campo a ancho fijo, "Abrir carpeta" quedaba empujado
        # fuera de la tarjeta y no se veia (23-sep-2026).
        self.btn_deck_abrir = PillButton(fila_deck, "Abrir carpeta",
                                         lambda: open_folder(PRESENTACIONES_DIR))
        self.btn_deck_abrir.pack(side="right", padx=(8, 0))
        self.btn_deck = PillButton(fila_deck, "Generar presentación",
                                   self.generar_deck_cliente, "primary")
        self.btn_deck.pack(side="right", padx=(8, 0))
        self.var_deck = tk.StringVar()
        self.entry_deck = tk.Entry(fila_deck, textvariable=self.var_deck, font=f_reg(11),
                                   width=20, bg=GBA_WHITE, fg=GBA_INK, relief="flat",
                                   highlightthickness=1, highlightbackground=GBA_300,
                                   highlightcolor=GBA_MAGENTA)
        self.entry_deck.pack(side="left", fill="x", expand=True, ipady=6)
        self.entry_deck.bind("<Return>", lambda _e: self.generar_deck_cliente())

        # --- Plan de compensacion ---
        self.comp_card = StatusCard(cuerpo, "Plan de compensación")
        celda_grid(self.comp_card, 1, 0)
        # El selector va ANTES del boton: primero se elige el trimestre y luego
        # se genera, que es el orden en que se lee.
        self.sel_comp = QuarterStepper(self.comp_card.button_row)
        self.sel_comp.pack(side="left", padx=(0, 10))
        self.btn_comp = PillButton(self.comp_card.button_row, "Generar",
                                   self.generar_plan_compensacion, "primary")
        self.btn_comp.pack(side="left")
        self.btn_comp_abrir = PillButton(self.comp_card.button_row, "Abrir carpeta",
                                         self.abrir_carpeta_compensacion)
        self.btn_comp_abrir.pack(side="left", padx=(8, 0))

        # --- Metricas del trimestre ---
        # Vive dentro de VtigerAIWindow desde el 28-ago-2026, ya no directo
        # en la grilla del panel (ver bloque "vTiger Análisis" mas arriba).
        self.met_card = self.vtiger_ai_window.met_card
        self.sel_met = QuarterStepper(self.met_card.button_row)
        self.sel_met.pack(side="left", padx=(0, 10))
        self.btn_met = PillButton(self.met_card.button_row, "Generar",
                                  self.generar_metricas_trimestre, "primary")
        self.btn_met.pack(side="left")
        self.btn_met_abrir = PillButton(self.met_card.button_row, "Abrir carpeta",
                                        self.abrir_carpeta_metricas)
        self.btn_met_abrir.pack(side="left", padx=(8, 0))

        # --- Forecast: mes actual/proximo, Upside y Commit ---
        # (Vacaciones ya no va aca -- ahora es el boton "Vacaciones" del Hero.)
        # Vive dentro de FreshworksWindow desde el 28-ago-2026, ya no directo
        # en la grilla del panel (ver bloque "Freshworks" mas arriba).
        self.forecast_card = self.freshworks_window.forecast_card

        # Fila de controles (mes + servicios GB) entre el detalle y los
        # botones -- StatusCard no trae esta fila por defecto, asi que se
        # corre el button_row un lugar para hacerle espacio.
        controles_forecast = tk.Frame(self.forecast_card.body, bg=CARD_BG)
        self.forecast_card.button_row.grid_forget()
        controles_forecast.grid(row=2, column=0, sticky="w", pady=(0, 6))
        self.forecast_card.button_row.grid(row=3, column=0, sticky="sw")

        # Rango de meses del forecast, como claves "YYYY-MM". Arranca en el
        # mes actual, que era el unico comportamiento posible antes del
        # 16-sep-2026.
        _hoy_fc = datetime.now()
        self._forecast_desde = mes_clave(_hoy_fc.year, _hoy_fc.month)
        self._forecast_hasta = self._forecast_desde
        self._forecast_categorias = set(FORECAST_CATEGORY_OPTIONS)
        self._forecast_segmentos = set(FORECAST_SEGMENT_OPTIONS)
        self._forecast_reps = set(gen_ns.TEAM_OWNERS)
        self._forecast_incluir_servicios = True

        self.lbl_forecast_opciones = tk.Label(
            controles_forecast, text="", font=f_reg(9), bg=CARD_BG, fg=GBA_700,
            justify="left", anchor="w",
        )
        self.lbl_forecast_opciones.pack(side="left", padx=(0, 10))
        # Botones compactos (font_size/pad chicos): la fila de Forecast tiene
        # 5 botones, con el tamano normal de PillButton se ve enorme.
        FCOMPACT = dict(font_size=9, pad_x=16, pad_y=9)
        self.btn_forecast_opciones = PillButton(controles_forecast, "Options…",
                                                self.abrir_opciones_forecast, **FCOMPACT)
        self.btn_forecast_opciones.pack(side="left")
        self._actualizar_resumen_forecast()

        # Flujo secuencial: Options -> Generar -> Ver informe -> Enviar.
        # Recien abierto (o justo despues de un Enviar), solo Options esta
        # habilitado -- hay que pasar por el ciclo completo otra vez antes
        # de poder volver a mandar algo.
        self.btn_forecast_gen = PillButton(self.forecast_card.button_row, "Generar",
                                           self.generar_forecast, "primary", **FCOMPACT)
        self.btn_forecast_gen.pack(side="left")
        self.btn_forecast_gen.set_enabled(False)
        self.btn_forecast_ver = PillButton(self.forecast_card.button_row, "Ver informe",
                                           self.ver_informe_forecast, **FCOMPACT)
        self.btn_forecast_ver.pack(side="left", padx=(8, 0))
        self.btn_forecast_ver.set_enabled(False)
        self.btn_forecast_enviar = PillButton(self.forecast_card.button_row, "Enviar",
                                              self.enviar_forecast, **FCOMPACT)
        self.btn_forecast_enviar.pack(side="left", padx=(8, 0))
        self.btn_forecast_enviar.set_enabled(False)
        self.btn_forecast_abrir = PillButton(self.forecast_card.button_row, "Abrir carpeta",
                                             self.abrir_carpeta_forecast, **FCOMPACT)
        self.btn_forecast_abrir.pack(side="left", padx=(8, 0))

        # --- Cadence Generator (28-ago-2026) ---
        # Deck mensual de la cadencia con Freshworks (pipeline por vendedor),
        # generador fijo build_deck.py + template.html en FW-Cadence-
        # Generator/ (skill "fw-cadence-deck", ver Claude/Scheduled/). De
        # solo lectura hacia vTiger (no escribe nada, no manda correos) --
        # igual que Reasignacion de orgs y Churn, un solo boton sin plan/
        # confirmacion. Vive dentro de FreshworksWindow desde que se creo.
        self.cadence_card = self.freshworks_window.cadence_card
        self.btn_cadence_gen = PillButton(self.cadence_card.button_row, "Generar",
                                          self.ejecutar_cadence_generator, "primary")
        self.btn_cadence_gen.pack(side="left")
        self.btn_cadence_abrir = PillButton(self.cadence_card.button_row, "Abrir carpeta",
                                            self.abrir_carpeta_cadence)
        self.btn_cadence_abrir.pack(side="left", padx=(8, 0))

        # --- Reasignacion de organizaciones de Sofia (nueva 27-ago-2026) ---
        # Sin patron plan/confirmacion: el skill "deteccion-reasignacion-
        # orgs-sofia" ya corre desatendido dos veces al dia (cron) y publica
        # el comentario + reasigna directo cuando se cumplen sus 3
        # condiciones (mencion positiva, org de Sofia, reunion pasada
        # confirmada) -- este boton es la misma corrida, a demanda, con el
        # mismo nivel de supervision (ninguna en el medio). El semaforo
        # refleja el marcador que ya escribe el cron, no una consulta nueva.
        # Vive dentro de SofiaAIWindow desde el 28-ago-2026, ya no directo en
        # la grilla del panel (ver bloque "Sofia - AI" mas arriba).
        self.reasigna_card = self.sofia_ai_window.reasigna_card
        self.btn_reasigna = PillButton(self.reasigna_card.button_row, "Detectar y reasignar",
                                       self.ejecutar_reasigna_orgs, "primary")
        self.btn_reasigna.pack(side="left")

        # --- Validar creación de deals ---
        self.validar_reasigna_card = self.sofia_ai_window.validar_reasigna_card
        self.btn_validar_deals = PillButton(self.validar_reasigna_card.button_row, "Validar deals",
                                            self.ejecutar_validar_deals, "primary")
        self.btn_validar_deals.pack(side="left")

        # --- Rutinas automaticas (Sofia) ---
        # Disparan la MISMA rutina que la Tarea Programada, con el mismo
        # Run-Rutina.ps1, no un prompt aparte: asi la corrida a mano escribe el
        # log y el marcador diario que leen los semaforos. Van con -Force
        # porque el script se salta la corrida si el marcador dice que ya
        # salio bien hoy, y desde un boton eso pareceria que no hace nada.
        self.rutinas_card = self.sofia_ai_window.rutinas_card
        self.btn_rutina_prep = PillButton(
            self.rutinas_card.button_row, RUTINAS_SOFIA[0][1],
            lambda: self.run_rutina(RUTINAS_SOFIA[0][0]), "primary")
        self.btn_rutina_prep.pack(side="left")
        self.btn_rutina_reasigna = PillButton(
            self.rutinas_card.button_row, RUTINAS_SOFIA[1][1],
            lambda: self.run_rutina(RUTINAS_SOFIA[1][0]))
        self.btn_rutina_reasigna.pack(side="left", padx=(8, 0))

        # --- Analisis de Churn ---
        # Skill "analiza-churn-batch": de solo lectura hacia vTiger (no
        # escribe nada, no manda correos) -- por eso un solo boton sin
        # plan/confirmacion. Usa el tool Workflow para correr el analisis
        # forense de varias cuentas en paralelo, asi que puede tardar
        # bastante mas que los otros botones. Vive dentro de VtigerAIWindow
        # desde el 28-ago-2026, ya no directo en la grilla del panel.
        self.churn_card = self.vtiger_ai_window.churn_card
        self.btn_churn = PillButton(self.churn_card.button_row, "Analizar churn",
                                    self.ejecutar_analisis_churn, "primary")
        self.btn_churn.pack(side="left")
        self.btn_churn_desplegar = PillButton(self.churn_card.button_row, "Desplegar churn",
                                              self.ejecutar_churn_desplegar)
        self.btn_churn_desplegar.pack(side="left", padx=(8, 0))
        self.btn_churn_abrir = PillButton(self.churn_card.button_row, "Abrir carpeta",
                                          self.abrir_carpeta_churn)
        self.btn_churn_abrir.pack(side="left", padx=(8, 0))

        # --- GBS Deal Checker (28-ago-2026) ---
        # "Chequear GBS" es de solo lectura hacia vTiger (no escribe nada,
        # no manda correos) -- revisa que cada deal ABIERTO del equipo tenga
        # un "Deals GBS Link" valido en el ultimo DIQ de Descargas. Corre
        # local (gbs_deal_checker.py, sin Claude ni MCP) porque solo
        # necesita el Excel ya descargado y chequeos HTTP; puede tardar
        # 1-2 min por la parte de red.
        # "Notificar" (28-ago-2026, pedido de el responsable comercial) SI escribe: publica en
        # vTiger una mencion real al vendedor de cada deal con GBS roto,
        # pidiendole que lo revise y avise cuando lo arregle. Sigue el mismo
        # patron "Calcular -> Enviar" que Sofia/Calidad CRM: primero corre
        # planificar_notificacion_gbs.py (local, sin MCP) para ver a quien
        # le falta avisar todavia -- se salta a quien ya se le notifico y no
        # ha corregido, asi no se le manda el mismo aviso cada vez que se
        # aprieta el boton -- y solo si el plan trae algo nuevo dispara
        # Claude Code sin supervision (headless) para publicar la mencion
        # real via MCP de vTiger, que este panel no tiene. El clic en
        # "Notificar" ES la aprobacion, no hay confirmacion en el medio.
        # Vive dentro de VtigerAIWindow, ver bloque "vTiger Análisis".
        self.gbs_card = self.vtiger_ai_window.gbs_card
        gbs_fila_a, gbs_fila_b = filas_botones(self.gbs_card)
        self.btn_gbs_checker = PillButton(gbs_fila_a, "Chequear GBS",
                                          self.ejecutar_gbs_checker, "primary")
        self.btn_gbs_checker.pack(side="left")
        self.btn_gbs_notificar = PillButton(gbs_fila_a, "Notificar",
                                            self.ejecutar_notificar_gbs)
        self.btn_gbs_notificar.pack(side="left", padx=(8, 0))
        self.btn_gbs_abrir = PillButton(gbs_fila_b, "Abrir carpeta",
                                        self.abrir_carpeta_gbs)
        self.btn_gbs_abrir.pack(side="left")

        # --- GBS Deal Maker (28-ago-2026) ---
        # Redacta y arma en Word (Goals | Barriers | Solutions) el GBS de
        # cada deal abierto del equipo que hoy no tiene NINGUN GBS Link --
        # el contenido sale UNICAMENTE de los pain points reales que ya
        # estan escritos en los comentarios de ese deal en vTiger (skill
        # "gbs-deal-maker", seccion GENERAR: sin inventar nada, si no hay
        # pain points registrados no genera el documento). "Generar" guarda
        # el .docx LOCAL en GBS Deal Maker/borradores/, sin tocar vTiger ni
        # SharePoint -- para que el responsable comercial pueda abrir y revisar cada Word antes
        # de subir nada (boton "Abrir carpeta"; los que borre ahi quedan
        # fuera del despliegue, esa es la aprobacion). "Desplegar" sube a
        # SharePoint (biblioteca de equipo "sites/sales", no OneDrive
        # personal de nadie -- para no repetir la falla que rompio ~26 GBS
        # Link cuando esas personas salieron de la empresa) cada borrador
        # que siga en la carpeta, y guarda ese link en el campo GBS Link
        # del deal en vTiger (seccion DESPLEGAR del mismo skill). Ambos
        # pasos disparan Claude Code sin supervision porque necesitan MCP
        # de vTiger/SharePoint. El clic en "Desplegar" ES la aprobacion.
        self.gbs_maker_card = self.vtiger_ai_window.gbs_maker_card
        gbsmk_fila_a, gbsmk_fila_b = filas_botones(self.gbs_maker_card)
        self.btn_gbsmk_generar = PillButton(gbsmk_fila_a, "Generar",
                                            self.ejecutar_gbs_maker_generar)
        self.btn_gbsmk_generar.pack(side="left")
        self.btn_gbsmk_abrir = PillButton(gbsmk_fila_a, "Abrir carpeta",
                                          self.abrir_carpeta_gbs_maker)
        self.btn_gbsmk_abrir.pack(side="left", padx=(8, 0))
        self.btn_gbsmk_desplegar = PillButton(gbsmk_fila_b, "Desplegar",
                                              self.ejecutar_gbs_maker_desplegar, "primary")
        self.btn_gbsmk_desplegar.pack(side="left")

        # --- GBS Organization Maker (28-ago-2026) ---
        # Mismo mecanismo que GBS Deal Maker (skill "gbs-organization-
        # maker") pero a nivel de organizacion: toma las organizaciones del
        # equipo sin GBS Link (pueden ser muchas mas que deals -- el skill
        # usa el tool Workflow para redactarlas en paralelo, igual que
        # Churn), redacta Goals/Barriers/Solutions solo con pain points
        # reales de los comentarios de la organizacion, y guarda el .docx
        # en GBS Organization Maker/borradores/ para revisar antes de
        # desplegar. "Desplegar" sube a SharePoint y guarda el link en el
        # campo GBS Link de la organizacion en vTiger.
        # A diferencia de Deal Maker, esta tarjeta tiene un control de
        # CANTIDAD (boton "Cantidad…", ver CantidadOrgsWindow): son ~150
        # organizaciones pendientes y cada una consume creditos, asi que se
        # procesan por tandas -- el numero elegido se le pasa al
        # planificador como --limite y se recuerda entre corridas.
        self.gbs_org_maker_card = self.vtiger_ai_window.gbs_org_maker_card

        # Fila de control (cantidad) entre el detalle y los botones, mismo
        # patron que Forecast: StatusCard no la trae, hay que correr el
        # button_row un lugar para hacerle espacio.
        controles_gbsom = tk.Frame(self.gbs_org_maker_card.body, bg=CARD_BG)
        self.gbs_org_maker_card.button_row.grid_forget()
        controles_gbsom.grid(row=2, column=0, sticky="w", pady=(0, 6))
        self.gbs_org_maker_card.button_row.grid(row=3, column=0, sticky="sw")

        self._gbs_org_limite = cargar_limite_gbs_org()
        self.lbl_gbsom_limite = tk.Label(controles_gbsom, text="", font=f_reg(9),
                                         bg=CARD_BG, fg=GBA_700)
        self.lbl_gbsom_limite.pack(side="left", padx=(0, 8))
        self.btn_gbsom_cantidad = PillButton(controles_gbsom, "Cantidad…",
                                             self.abrir_cantidad_gbs_org,
                                             font_size=9, pad_x=16, pad_y=9)
        self.btn_gbsom_cantidad.pack(side="left")
        self._actualizar_label_gbs_org()

        gbsom_fila_a, gbsom_fila_b = filas_botones(self.gbs_org_maker_card)
        self.btn_gbsom_generar = PillButton(gbsom_fila_a, "Generar",
                                            self.ejecutar_gbs_org_maker_generar)
        self.btn_gbsom_generar.pack(side="left")
        self.btn_gbsom_abrir = PillButton(gbsom_fila_a, "Abrir carpeta",
                                          self.abrir_carpeta_gbs_org_maker)
        self.btn_gbsom_abrir.pack(side="left", padx=(8, 0))
        self.btn_gbsom_desplegar = PillButton(gbsom_fila_b, "Desplegar",
                                              self.ejecutar_gbs_org_maker_desplegar, "primary")
        self.btn_gbsom_desplegar.pack(side="left")

        # --- Consola de actividad ---
        consola = RoundedCard(cuerpo, pad=14)
        consola.grid(row=0, column=2, rowspan=4, sticky="nsew")
        consola.body.rowconfigure(1, weight=1)
        consola.body.columnconfigure(0, weight=1)

        cab = tk.Frame(consola.body, bg=CARD_BG)
        cab.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(cab, text="Actividad", font=f_med(12), bg=CARD_BG,
                 fg=GBA_INK).pack(side="left")
        tk.Label(cab, text="salida de los scripts", font=f_reg(9), bg=CARD_BG,
                 fg=GBA_500).pack(side="left", padx=(8, 0))

        # height=1: un Text sin altura explicita pide 24 lineas (~388 px) y
        # deforma la rejilla. Crece con sticky, no con su tamano pedido.
        self.output = tk.Text(consola.body, state="disabled", font=f_mono(9), height=1,
                              bg=CARD_BG, fg=GBA_700, relief="flat", bd=0,
                              highlightthickness=0, wrap="word", padx=0, pady=8)
        barra = ttk.Scrollbar(consola.body, orient="vertical", command=self.output.yview)
        self.output.configure(yscrollcommand=barra.set)
        self.output.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        barra.grid(row=1, column=1, sticky="ns", pady=(8, 0))
        self.output.tag_configure("ok", foreground="#1F5C2E")
        self.output.tag_configure("aviso", foreground=GBA_MAGENTA_700)

        self._buttons = [
            self.btn_gen_ns, self.btn_review_ns,
            self.btn_comp, self.btn_comp_abrir,
            self.btn_met, self.btn_met_abrir,
            self.btn_forecast_opciones, self.btn_forecast_abrir,
            self.btn_churn_abrir, self.btn_churn_desplegar, self.btn_cadence_abrir, self.btn_gbs_abrir,
            self.btn_gbsmk_abrir, self.btn_gbsom_abrir, self.btn_gbsom_cantidad,
        ]
        # Sofia y Calidad CRM (27-ago-2026) tienen de vuelta el patron
        # "Calcular -> Ver -> Enviar" (igual que Forecast), con sus botones
        # excluidos a proposito de self._buttons -- su estado (habilitado/
        # deshabilitado) lo maneja cada flujo a mano, no el bloqueo global.
        # En Calidad CRM el calculo del plan es un script Python local
        # (no necesita MCP); en Sofia, hasta el calculo del plan necesita
        # Claude (cruza calendario y disponibilidad real por Outlook). El
        # envio de ambos SIEMPRE dispara Claude Code sin supervision
        # (headless, --dangerously-skip-permissions) porque solo el tiene
        # acceso a SharePoint/Outlook via MCP. Ver run_claude_headless.

        # Tamano inicial y minimo calculados, no fijos a mano.
        ancho_min = PADX * 2 + CARD_MIN_W * 2 + CONSOLE_MIN_W + GAP * 2
        alto_min = self.hero._alto + PADX * 2 + grid_h
        root.minsize(ancho_min, alto_min)
        recordar_tamano(root, "principal", f"{max(ancho_min, 1360)}x{alto_min}")

        self.log("Listo. Los semáforos se actualizan solos cada minuto.")
        self.refresh_status()
        self._schedule_auto_refresh()

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background=CARD_BG, fieldbackground=CARD_BG, foreground=GBA_INK,
                        rowheight=26, font=f_reg(9), borderwidth=0)
        style.configure("Treeview.Heading", background=GBA_100, foreground=GBA_700,
                        font=f_med(9), relief="flat", borderwidth=0)
        style.map("Treeview.Heading", background=[("active", GBA_200)])
        style.map("Treeview", background=[("selected", GBA_MAGENTA_100)],
                  foreground=[("selected", GBA_INK)])
        style.configure("Vertical.TScrollbar", background=GBA_200, troughcolor=CARD_BG,
                        bordercolor=CARD_BG, arrowcolor=GBA_700, relief="flat")

    # ---- consola ----

    def log(self, texto):
        tag = ()
        bajo = texto.lower()
        if "error" in bajo or "aviso" in bajo or "alerta" in bajo:
            tag = ("aviso",)
        elif "listo" in bajo or "enviado" in bajo:
            tag = ("ok",)
        self.output.configure(state="normal")
        self.output.insert("end", texto + "\n", tag)
        # Sin auto-scroll: el usuario controla el scroll manualmente (28-ago-2026).
        # Esto evita que el scroll se mueva solo durante una corrida con muchos logs.
        self.output.configure(state="disabled")

    # ---- semaforos ----

    def refresh_status(self):
        resultados = {}
        for card, fn in (
            (self.sofia_card, check_sofia_status),
            (self.ns_card, check_next_step_status),
            (self.comp_card, check_comp_plan_status),
            (self.met_card, check_metricas_status),
            (self.forecast_card, check_forecast_status),
            (self.reasigna_card, check_reasigna_orgs_status),
            (self.validar_reasigna_card, check_validar_deals_status),
            (self.rutinas_card, check_rutinas_sofia_status),
            (self.churn_card, check_churn_status),
            (self.gbs_card, check_gbs_checker_status),
            (self.gbs_maker_card, check_gbs_maker_status),
            (self.gbs_org_maker_card, check_gbs_org_maker_status),
            (self.cadence_card, check_fw_cadence_status),
        ):
            color, texto = fn()
            card.set_status(color, texto)
            resultados[card] = (color, texto)

        # Botones lanzadores (Sofia - AI, vTiger Análisis, 28-ago-2026): cada
        # uno toma el semaforo mas urgente de las tarjetas que agrupa (rojo >
        # amarillo > gris > verde), para que se note desde la grilla
        # principal si algo adentro necesita atencion sin abrir la ventana.
        prioridad = {COLOR_RED: 3, COLOR_YELLOW: 2, COLOR_GRAY: 1, COLOR_GREEN: 0}

        def peor_color(*colores):
            return max(colores, key=lambda c: prioridad[c])

        color_sofia, texto_sofia = resultados[self.sofia_card]
        color_reasigna, texto_reasigna = resultados[self.reasigna_card]
        color_validar, texto_validar = resultados[self.validar_reasigna_card]
        self.sofia_ai_card.set_status(
            peor_color(color_sofia, color_reasigna, color_validar),
            f"Reportes vendedores: {texto_sofia} · Reasignación de orgs: {texto_reasigna} · Validar deals: {texto_validar}")

        color_ns, texto_ns = resultados[self.ns_card]
        color_met, texto_met = resultados[self.met_card]
        color_churn, texto_churn = resultados[self.churn_card]
        color_gbs, texto_gbs = resultados[self.gbs_card]
        color_gbsmk, texto_gbsmk = resultados[self.gbs_maker_card]
        color_gbsom, texto_gbsom = resultados[self.gbs_org_maker_card]
        self.vtiger_ai_card.set_status(
            peor_color(color_ns, color_met, color_churn, color_gbs, color_gbsmk, color_gbsom),
            f"Calidad CRM: {texto_ns} · Métricas: {texto_met} · Churn: {texto_churn} · "
            f"GBS: {texto_gbs} · GBS Deal Maker: {texto_gbsmk} · GBS Org Maker: {texto_gbsom}")

        # La tarjeta puede no existir: con mostrar_freshworks en false la ventana
        # se construye igual (hay referencias a sus tarjetas) pero no se muestra.
        if getattr(self, "freshworks_card", None) is not None:
            color_forecast, texto_forecast = resultados[self.forecast_card]
            color_cadence, texto_cadence = resultados[self.cadence_card]
            self.freshworks_card.set_status(
                peor_color(color_forecast, color_cadence),
                f"Forecast: {texto_forecast} · Cadence Generator: {texto_cadence}")

        self.hero.set_sello(f"Actualizado {datetime.now():%H:%M}")

    def _schedule_auto_refresh(self):
        self.refresh_status()
        self.root.after(AUTO_REFRESH_MS, self._schedule_auto_refresh)

    # ---- acciones ----

    def revisar_plan_sofia(self):
        if not SOFIA_SKILL_FILE.exists():
            self.log(f"No se encontró el skill en {SOFIA_SKILL_FILE}. Abortando.")
            return
        self._plan_pendiente = None
        self.btn_ver_resumen_sofia.set_enabled(False)
        self.btn_confirm_sofia.set_enabled(False)
        self.btn_calc_sofia.set_enabled(False)
        self.log("Calculando qué se repartiría y a quién (cruza calendarios y "
                 "disponibilidad real por Outlook, puede tardar 1-2 min)…")
        prompt = (
            f"Sigue el skill en {SOFIA_SKILL_FILE} (ignora el frontmatter YAML), pero SOLO "
            "la sección \"=== FLUJO PRINCIPAL (calcular plan) ===\" -- NO ejecutes la sección "
            "\"=== EJECUCIÓN ===\", no subas nada a SharePoint, no reenvíes ninguna invitación, "
            "no mandes ningún correo. Esta corrida es solo para calcular el plan.\n\n"
            f"En vez de mostrarle el plan a el responsable comercial en una tabla y esperar confirmación, escríbelo "
            f"a {SOFIA_PLAN_FILE} como JSON con este esquema exacto:\n"
            '{"entries": [{"file_name": "<nombre del html>", "company": "<empresa>", '
            '"seller_name": "<nombre del vendedor>", "reason": "ya_asignado|turno_reunion|'
            'turno_sin_reunion", "meeting_start": "<ISO8601 o null si no hay reunión>", '
            '"meeting_subject": "<asunto o null>", "send_invite": true si hay reunión '
            'asociada (false si no)}], "blocked": [{"file_name": "<nombre>", "motivo": '
            '"<por qué quedó bloqueado>"}], "vacaciones": ["<vendedores ignorados por '
            'vacaciones>"]}\n\n'
            "Si un archivo tiene más de un vendedor asociado (varios ya en la invitación), "
            "agrega una entrada por cada vendedor con el mismo file_name. No escribas nada "
            "más que ese JSON en el archivo. Termina con un resumen breve de una línea."
        )
        self.run_claude_headless(prompt, on_done=self._plan_sofia_calculado)

    def _plan_sofia_calculado(self, code):
        self.btn_calc_sofia.set_enabled(True)
        if code != 0:
            self.log("No se pudo calcular el plan. Revisa el detalle de arriba.")
            return
        try:
            with open(SOFIA_PLAN_FILE, "r", encoding="utf-8") as f:
                plan = json.load(f)
        except Exception as exc:
            self.log(f"No se pudo leer el plan: {exc}")
            return
        self._plan_pendiente = plan
        entries = plan.get("entries", [])
        self.log(f"Plan calculado: {len(entries)} asignación(es). "
                 f"Dale a 'Ver plan' para ver el detalle.")
        self.btn_ver_resumen_sofia.set_enabled(True)
        self.btn_confirm_sofia.set_enabled(bool(entries))

    def ver_resumen_sofia(self):
        plan = getattr(self, "_plan_pendiente", None)
        if plan is None:
            self.log("Todavía no hay un plan calculado. Dale a 'Calcular plan' primero.")
            return
        for linea in format_plan_lines(plan):
            self.log(linea)

    def confirmar_envio_plan(self):
        plan = getattr(self, "_plan_pendiente", None)
        if not plan or not plan.get("entries"):
            self.log("No hay un plan pendiente por confirmar. Dale a 'Calcular plan' primero.")
            return
        self.log("Plan aprobado. Enviando por Outlook con la presentación adjunta…")
        self._plan_pendiente = None
        self.btn_calc_sofia.set_enabled(False)
        self.btn_ver_resumen_sofia.set_enabled(False)
        self.btn_confirm_sofia.set_enabled(False)
        # Script local en vez de Claude Code (26-sep-2026). Antes el reparto
        # pasaba por el conector de M365, que no admite adjuntos, asi que subia
        # el HTML a SharePoint y mandaba un enlace; y buscaba la invitacion como
        # CORREO de Sofia con asunto exacto, lo que la perdia en cuanto Outlook
        # le anteponia "Provisional:". Ahora el correo sale de este Outlook con
        # el archivo adjunto y la invitacion se reenvia desde la cita del
        # calendario, buscada por hora de inicio.
        self.run_script("enviar_presentaciones_sofia.py", [],
                        on_done=self._tras_enviar_sofia)

    def _tras_enviar_sofia(self, code):
        if code != 0:
            self.log("El reparto terminó con error. Revisa el detalle de arriba y vuelve a "
                     "'Calcular plan' antes de reintentar.")
            return
        self.log("Reparto terminado. Dale a 'Calcular plan' de nuevo si querés confirmar que no "
                 "quedó nada pendiente.")

    def revisar_reuniones_sofia(self):
        prep_skill = skill_file("daily-client-presentation-prep")
        if not prep_skill.exists():
            self.log(f"No se encontró el skill en {prep_skill}. Abortando.")
            self.btn_revisar_reuniones.set_enabled(True)
            return
        self.log("Generando presentaciones de reuniones de Sofía (hoy + próximo día hábil)…")
        self.btn_revisar_reuniones.set_enabled(False)
        prompt = (
            f"Lee el skill completo en {prep_skill} (ignora el frontmatter YAML). "
            "Sigue EXACTAMENTE todos los pasos del 1 al 10 de ese skill: "
            "busca reuniones de Sofia Garcia en el calendario de el responsable comercial para hoy y el próximo día hábil, "
            "genera las presentaciones HTML usando el generador fijo (_generador-deck/build_deck.py), "
            f"y guarda los archivos en {PRESENTACIONES_DIR}.\n\n"
            "Respeta las exclusiones (monday.com), la ventana de 7 días para no duplicar, "
            "y todas las reglas de nombre de archivo, validación y entrega indicadas en el skill.\n\n"
            "Termina con un resumen claro: qué presentaciones se generaron y dónde, "
            "cuáles se saltaron (por monday.com o duplicado reciente) y cualquier error."
        )
        self.run_claude_headless(prompt, on_done=self._tras_revisar_reuniones_sofia)

    def _tras_revisar_reuniones_sofia(self, code):
        self.btn_revisar_reuniones.set_enabled(True)
        if code != 0:
            self.log("Error al recuperar reuniones. Verifica la autenticación con Outlook/Microsoft 365.")
            return
        self.log("✓ Reuniones revisadas. Calculando el plan con los decks que acaban de "
                 "generarse…")
        # Encadenado a proposito: si el plan se calcula antes de este paso, no ve
        # los decks recien generados y la reunion nueva se queda sin enviar
        # (28-sep-2026, caso Latin American School).
        self.revisar_plan_sofia()

    def ejecutar_reasigna_orgs(self):
        if not REASIGNA_SKILL_FILE.exists():
            self.log(f"No se encontró el skill en {REASIGNA_SKILL_FILE}. Abortando.")
            return
        self.btn_reasigna.set_enabled(False)
        self.log("Revisando menciones de vTiger sobre organizaciones de Sofía (Claude Code, "
                 "sin supervisión, puede tardar unos minutos)…")
        prompt = (
            f"Sigue el skill en {REASIGNA_SKILL_FILE} (ignora el frontmatter YAML) al pie de la "
            "letra, de principio a fin: detecta las menciones, valida las 3 condiciones, publica "
            "el comentario en vTiger con las menciones reales confirmadas, reasigna el Assigned "
            "To por API, y entrega al final la tabla tal como indica el skill."
        )
        self.run_claude_headless(prompt, on_done=self._tras_reasigna_orgs)

    def _tras_reasigna_orgs(self, code):
        self.btn_reasigna.set_enabled(True)
        if code != 0:
            self.log("La corrida terminó con error. Revisa el detalle de arriba.")
            return
        # Sin esto el semaforo se quedaba con la fecha de la ultima corrida
        # programada, aunque la de ahora hubiera salido bien (29-sep-2026).
        ok, err = marcar_corrida_hoy("deteccion-reasignacion-orgs-sofia")
        if not ok:
            self.log(f"Corrida terminada, pero no se pudo actualizar el marcador: {err}")
            return
        self.log("Corrida terminada.")

    def ejecutar_validar_deals(self):
        validar_skill = skill_file("validar-creacion-deals")
        if not validar_skill.exists():
            self.log(f"No se encontró el skill en {validar_skill}. Abortando.")
            return
        self.btn_validar_deals.set_enabled(False)
        self.log("Validando que los vendedores crearon los deals (Claude Code sin supervisión)…")
        prompt = (
            f"Lee el skill completo en {validar_skill} (ignora el frontmatter YAML) al pie de la letra. "
            "Detecta organizaciones reasignadas recientemente de Sofía a vendedores, valida si cada "
            "vendedor creó el deal correspondiente, y si no existe, menciona al vendedor en vTiger "
            "con un recordatorio comercial y serio (en inglés) sobre la importancia de atender al cliente.\n\n"
            "Sigue todos los pasos del PASO 1 al PASO 4 del skill. Termina con una tabla clara: "
            "organizaciones, vendedor, y estado (deal creado o mención enviada)."
        )
        self.run_claude_headless(prompt, on_done=self._tras_validar_deals)

    def _tras_validar_deals(self, code):
        self.btn_validar_deals.set_enabled(True)
        if code != 0:
            self.log("La validación terminó con error. Revisa el detalle de arriba.")
            return
        self.log("✓ Validación completada. Revisa la tabla arriba.")

    def ejecutar_analisis_churn(self):
        if not CHURN_SKILL_FILE.exists():
            self.log(f"No se encontró el skill en {CHURN_SKILL_FILE}. Abortando.")
            return
        self.btn_churn.set_enabled(False)
        self.log("Analizando churn (forense por cuenta en paralelo con varios agentes)…")
        self.log("IMPORTANTE: con cuentas nuevas por analizar esto tarda HORAS, no minutos "
                 "(la corrida del 14-sep-2026 tardó más de dos). El panel sigue respondiendo; "
                 "no cierres la ventana o la corrida se queda huérfana y no verás el resultado.")
        prompt = (
            f"Sigue el skill en {CHURN_SKILL_FILE} (ignora el frontmatter YAML) al pie de la "
            "letra, de principio a fin: identifica las organizaciones con churn aplicable "
            f"(cruzando los Excel ChurnAssets y OA más recientes en {DOWNLOADS}), corre "
            "el análisis forense de 'analiza-churn' en paralelo con el tool Workflow sobre las "
            "cuentas sin informe previo, verifica cada veredicto, arma el Excel resumen, y "
            "entrega al final la tabla tal como indica el skill."
        )
        self.run_claude_headless(prompt, on_done=self._tras_analisis_churn,
                                 timeout=CHURN_TIMEOUT_SECONDS)

    def _tras_analisis_churn(self, code):
        self.btn_churn.set_enabled(True)
        if code != 0:
            self.log("El análisis terminó con error. Revisa el detalle de arriba.")
            return
        self.log("Análisis de churn terminado.")

    def ejecutar_churn_desplegar(self):
        if not CHURN_SKILL_FILE.exists():
            self.log(f"No se encontró el skill en {CHURN_SKILL_FILE}. Abortando.")
            return
        self.btn_churn_desplegar.set_enabled(False)
        self.log("Desplegando análisis de churn a SharePoint (sube .html y actualiza links en Excel)…")
        prompt = (
            f"Sigue el skill en {CHURN_SKILL_FILE} (ignora el frontmatter YAML), pero SOLO la "
            "sección \"=== SECCIÓN DESPLEGAR (SÍ escribe: SharePoint) ===\" -- NO ejecutes la "
            "sección de análisis forense. Sube los archivos .html a SharePoint (sites/sales/Churn "
            "Accounts/) y actualiza los links en el Excel resumen, tal como indica el skill."
        )
        self.run_claude_headless(prompt, on_done=self._tras_churn_desplegar,
                                 timeout=CHURN_TIMEOUT_SECONDS)

    def _tras_churn_desplegar(self, code):
        self.btn_churn_desplegar.set_enabled(True)
        if code != 0:
            self.log("El despliegue de churn terminó con error. Revisa el detalle de arriba.")
            return
        self.log("Despliegue de churn terminado. Excel actualizado con links de SharePoint.")

    def ejecutar_gbs_checker(self):
        self.btn_gbs_checker.set_enabled(False)
        self.log("Revisando el GBS Link de los deals abiertos del equipo (chequeo de "
                 "red, puede tardar 1-2 min)…")
        self.run_script("gbs_deal_checker.py", [], on_done=self._tras_gbs_checker)

    def _tras_gbs_checker(self, code):
        self.btn_gbs_checker.set_enabled(True)
        if code != 0:
            self.log("El chequeo de GBS terminó con error. Revisa el detalle de arriba.")
            return
        self.log("Chequeo de GBS terminado.")

    def ejecutar_notificar_gbs(self):
        if not GBS_NOTIFY_SKILL_FILE.exists():
            self.log(f"No se encontró el skill en {GBS_NOTIFY_SKILL_FILE}. Abortando.")
            return
        self.btn_gbs_notificar.set_enabled(False)
        self.log("Calculando a quién le falta notificar el GBS roto…")
        self.run_script("planificar_notificacion_gbs.py", [],
                        on_done=self._plan_notificacion_gbs_calculado)

    def _plan_notificacion_gbs_calculado(self, code):
        if code != 0:
            self.btn_gbs_notificar.set_enabled(True)
            self.log("No se pudo calcular a quién notificar. ¿Corriste 'Chequear GBS' hoy?")
            return
        try:
            with open(GBS_NOTIFY_PLAN_FILE, "r", encoding="utf-8") as f:
                plan = json.load(f)
        except Exception as exc:
            self.btn_gbs_notificar.set_enabled(True)
            self.log(f"No se pudo leer el plan de notificación: {exc}")
            return
        items = plan.get("items", [])
        if not items:
            self.btn_gbs_notificar.set_enabled(True)
            self.log("Nada nuevo que notificar: no hay deals con GBS roto sin avisar todavía.")
            return
        self.log(f"Publicando la mención en vTiger para {len(items)} deal(es) nuevo(s) con GBS "
                 "roto (Claude Code, sin supervisión, puede tardar unos minutos)…")
        prompt = (
            f"Sigue el skill en {GBS_NOTIFY_SKILL_FILE} (ignora el frontmatter YAML) al pie de "
            f"la letra, de principio a fin, usando el plan ya calculado en "
            f"{GBS_NOTIFY_PLAN_FILE}: resuelve cada deal en vTiger, publica la mención real "
            "(EN INGLÉS) al vendedor asignado pidiéndole que revise el GBS y avise cuando lo "
            "corrija, valida que la mención se registró como mención real, actualiza el "
            "registro de notificados, y entrega al final la tabla tal como indica el skill."
        )
        self.run_claude_headless(prompt, on_done=self._tras_notificar_gbs)

    def _tras_notificar_gbs(self, code):
        self.btn_gbs_notificar.set_enabled(True)
        if code != 0:
            self.log("La notificación terminó con error. Revisa el detalle de arriba.")
            return
        self.log("Notificación a vendedores terminada.")

    # ---- GBS Deal Maker / GBS Organization Maker (28-ago-2026) ----
    # Misma logica para ambos, parametrizada: "Generar" corre primero un
    # script local (planificar_creacion_gbs_*.py) que arma el plan de que
    # deals/orgs necesitan un GBS nuevo, y si trae algo dispara Claude Code
    # headless siguiendo la seccion GENERAR del skill correspondiente --
    # eso solo escribe .docx locales en 'borradores/', nunca toca vTiger ni
    # SharePoint. "Desplegar" no calcula ningun plan (la aprobacion es que
    # el responsable comercial haya revisado la carpeta y borrado los que no quiere subir):
    # dispara directo la seccion DESPLEGAR del skill sobre lo que quede en
    # 'borradores/'.

    def _generar_gbs_maker(self, boton_generar, script_planificador, plan_file,
                           skill_file, seccion_log, on_plan_calculado, args=None):
        if not skill_file.exists():
            self.log(f"No se encontró el skill en {skill_file}. Abortando.")
            return
        boton_generar.set_enabled(False)
        self.log(f"Calculando qué {seccion_log} necesitan un GBS nuevo…")
        self.run_script(script_planificador, args or [], on_done=on_plan_calculado)

    def _desplegar_gbs_maker(self, boton_desplegar, skill_file, borradores_dir,
                             prompt_desplegar, on_done):
        if not skill_file.exists():
            self.log(f"No se encontró el skill en {skill_file}. Abortando.")
            return
        pendientes = [p for p in borradores_dir.iterdir() if p.suffix == ".docx"] \
            if borradores_dir.exists() else []
        if not pendientes:
            self.log("No hay ningún borrador en la carpeta para desplegar. "
                     "Dale a 'Generar' primero, o revisa 'Abrir carpeta'.")
            return
        boton_desplegar.set_enabled(False)
        self.log(f"Desplegando {len(pendientes)} borrador(es) a SharePoint y vTiger "
                 "(Claude Code, sin supervisión, puede tardar varios minutos)…")
        self.run_claude_headless(prompt_desplegar, on_done=on_done)

    def ejecutar_gbs_maker_generar(self):
        self._generar_gbs_maker(
            self.btn_gbsmk_generar, "planificar_creacion_gbs_deals.py", GBS_MAKER_PLAN_FILE,
            GBS_MAKER_SKILL_FILE, "deals", self._plan_gbs_maker_calculado)

    def _plan_gbs_maker_calculado(self, code):
        if code != 0:
            self.btn_gbsmk_generar.set_enabled(True)
            self.log("No se pudo calcular qué deals necesitan un GBS nuevo. "
                     "¿Corriste 'Chequear GBS' hoy?")
            return
        try:
            with open(GBS_MAKER_PLAN_FILE, "r", encoding="utf-8") as f:
                plan = json.load(f)
        except Exception as exc:
            self.btn_gbsmk_generar.set_enabled(True)
            self.log(f"No se pudo leer el plan: {exc}")
            return
        items = plan.get("items", [])
        if not items:
            self.btn_gbsmk_generar.set_enabled(True)
            self.log("Nada nuevo que generar: no hay deals abiertos sin GBS Link pendientes.")
            return
        self.log(f"Redactando el GBS de {len(items)} deal(es) a partir de sus comentarios en "
                 "vTiger (Claude Code, sin supervisión, puede tardar unos minutos)…")
        prompt = (
            f"Sigue el skill en {GBS_MAKER_SKILL_FILE} (ignora el frontmatter YAML), pero SOLO "
            f"la sección \"=== SECCIÓN GENERAR (solo local, NO toca vTiger ni SharePoint) ===\" "
            "-- NO ejecutes la sección DESPLEGAR en esta corrida. Usa el plan ya calculado en "
            f"{GBS_MAKER_PLAN_FILE}. Recuerda: el contenido de Goals/Barriers/Solutions sale "
            "ÚNICAMENTE de los pain points reales en los comentarios de vTiger de cada deal, "
            "nunca inventado; si un deal no tiene pain points identificables, no generes nada "
            "para ese deal."
        )
        self.run_claude_headless(prompt, on_done=self._tras_gbs_maker_generar)

    def _tras_gbs_maker_generar(self, code):
        self.btn_gbsmk_generar.set_enabled(True)
        if code != 0:
            self.log("La generación de GBS terminó con error. Revisa el detalle de arriba.")
            return
        self.log("GBS generados. Revísalos con 'Abrir carpeta' antes de 'Desplegar'.")

    def ejecutar_gbs_maker_desplegar(self):
        prompt = (
            f"Sigue el skill en {GBS_MAKER_SKILL_FILE} (ignora el frontmatter YAML), pero SOLO "
            f"la sección \"=== SECCIÓN DESPLEGAR (SÍ escribe: SharePoint + vTiger) ===\" -- NO "
            "ejecutes la sección GENERAR en esta corrida. Usa como entrada los .docx que estén "
            f"hoy en {GBS_MAKER_BORRADORES_DIR} cruzados con su manifiesto. Este despliegue fue "
            "aprobado por quien opera el panel al hacer clic en 'Desplegar' en el Panel de "
            "Automatizaciones -- los borradores que no se querían ya los borró de la carpeta "
            "antes de este clic, así que subí TODO lo que encuentres ahí sin pedir más "
            "confirmación."
        )
        self._desplegar_gbs_maker(self.btn_gbsmk_desplegar, GBS_MAKER_SKILL_FILE,
                                  GBS_MAKER_BORRADORES_DIR, prompt, self._tras_gbs_maker_desplegar)

    def _tras_gbs_maker_desplegar(self, code):
        self.btn_gbsmk_desplegar.set_enabled(True)
        if code != 0:
            self.log("El despliegue de GBS terminó con error. Revisa el detalle de arriba.")
            return
        self.log("Despliegue de GBS terminado.")

    def abrir_cantidad_gbs_org(self):
        CantidadOrgsWindow(self, self._gbs_org_limite, self._aplicar_cantidad_gbs_org)

    def _aplicar_cantidad_gbs_org(self, valor):
        self._gbs_org_limite = valor
        guardar_limite_gbs_org(valor)
        self._actualizar_label_gbs_org()
        self.log(f"GBS Organization Maker: la próxima corrida de 'Generar' procesará hasta "
                 f"{valor} organización(es), empezando por las que más deals tienen.")

    def _actualizar_label_gbs_org(self):
        self.lbl_gbsom_limite.config(text=f"Hasta {self._gbs_org_limite} por corrida")

    def ejecutar_gbs_org_maker_generar(self):
        self._generar_gbs_maker(
            self.btn_gbsom_generar, "planificar_creacion_gbs_orgs.py", GBS_ORG_MAKER_PLAN_FILE,
            GBS_ORG_MAKER_SKILL_FILE, "organizaciones", self._plan_gbs_org_maker_calculado,
            args=["--limite", str(self._gbs_org_limite)])

    def _plan_gbs_org_maker_calculado(self, code):
        if code != 0:
            self.btn_gbsom_generar.set_enabled(True)
            self.log("No se pudo calcular qué organizaciones necesitan un GBS nuevo.")
            return
        try:
            with open(GBS_ORG_MAKER_PLAN_FILE, "r", encoding="utf-8") as f:
                plan = json.load(f)
        except Exception as exc:
            self.btn_gbsom_generar.set_enabled(True)
            self.log(f"No se pudo leer el plan: {exc}")
            return
        items = plan.get("items", [])
        if not items:
            self.btn_gbsom_generar.set_enabled(True)
            self.log("Nada nuevo que generar: no hay organizaciones del equipo sin GBS Link "
                     "pendientes.")
            return
        self.log(f"Tanda de {len(items)} organización(es) (límite configurado: "
                 f"{self._gbs_org_limite}). Redactando su GBS a partir de sus "
                 "comentarios en vTiger (Claude Code, sin supervisión, en paralelo vía "
                 "Workflow -- puede tardar varios minutos con un lote así de grande)…")
        prompt = (
            f"Sigue el skill en {GBS_ORG_MAKER_SKILL_FILE} (ignora el frontmatter YAML), pero "
            f"SOLO la sección \"=== SECCIÓN GENERAR (solo local, NO toca vTiger ni SharePoint) "
            "===\" -- NO ejecutes la sección DESPLEGAR en esta corrida. Usa el plan ya calculado "
            f"en {GBS_ORG_MAKER_PLAN_FILE}, y usa el tool Workflow para procesarlas en paralelo "
            "tal como indica el skill. Recuerda: el contenido de Goals/Barriers/Solutions sale "
            "ÚNICAMENTE de los pain points reales en los comentarios de vTiger de cada "
            "organización, nunca inventado; si una organización no tiene pain points "
            "identificables, no generes nada para ella."
        )
        self.run_claude_headless(prompt, on_done=self._tras_gbs_org_maker_generar)

    def _tras_gbs_org_maker_generar(self, code):
        self.btn_gbsom_generar.set_enabled(True)
        if code != 0:
            self.log("La generación de GBS terminó con error. Revisa el detalle de arriba.")
            return
        self.log("GBS generados. Revísalos con 'Abrir carpeta' antes de 'Desplegar'.")

    def ejecutar_gbs_org_maker_desplegar(self):
        prompt = (
            f"Sigue el skill en {GBS_ORG_MAKER_SKILL_FILE} (ignora el frontmatter YAML), pero "
            f"SOLO la sección \"=== SECCIÓN DESPLEGAR (SÍ escribe: SharePoint + vTiger) ===\" -- "
            "NO ejecutes la sección GENERAR en esta corrida. Usa como entrada los .docx que "
            f"estén hoy en {GBS_ORG_MAKER_BORRADORES_DIR} cruzados con su manifiesto, y usa el "
            "tool Workflow para procesarlos en paralelo tal como indica el skill. Este "
            "despliegue fue aprobado por quien opera el panel al hacer clic en 'Desplegar' en el Panel "
            "de Automatizaciones -- los borradores que no se querían ya los borró de la "
            "carpeta antes de este clic, así que subí TODO lo que encuentres ahí sin pedir más "
            "confirmación."
        )
        self._desplegar_gbs_maker(self.btn_gbsom_desplegar, GBS_ORG_MAKER_SKILL_FILE,
                                  GBS_ORG_MAKER_BORRADORES_DIR, prompt,
                                  self._tras_gbs_org_maker_desplegar)

    def _tras_gbs_org_maker_desplegar(self, code):
        self.btn_gbsom_desplegar.set_enabled(True)
        if code != 0:
            self.log("El despliegue de GBS terminó con error. Revisa el detalle de arriba.")
            return
        self.log("Despliegue de GBS terminado.")

    def _recoger_notas_cadencia(self):
        """Mueve a FW-Cadence-Generator el notas-cadencia.json que el deck deja
        en Descargas al pulsar "Guardar notas".

        El deck es una pagina abierta desde el disco y el navegador no le deja
        escribir sobre la carpeta del proyecto, asi que lo descarga. Esto cierra
        el circuito sin que el responsable comercial tenga que mover el archivo a mano: el deck del
        mes siguiente sale con los blockers del mes anterior ya escritos
        (16-sep-2026)."""
        destino = FW_CADENCE_DIR / "notas-cadencia.json"
        candidatos = sorted(DOWNLOADS.glob("notas-cadencia*.json"),
                            key=lambda q: q.stat().st_mtime, reverse=True)
        if not candidatos:
            return
        origen = candidatos[0]
        if destino.exists() and destino.stat().st_mtime >= origen.stat().st_mtime:
            return  # lo que ya hay es igual de nuevo o mas: no se pisa
        try:
            cuantas = len(json.loads(origen.read_text(encoding="utf-8")))
        except Exception as exc:
            self.log(f"Se encontró {origen.name} en Descargas pero no se pudo leer ({exc}); "
                     "se ignora y el deck saldrá con las notas anteriores.")
            return
        try:
            shutil.copy2(origen, destino)
            origen.unlink()
            self.log(f"Notas de la cadencia recogidas de Descargas: {cuantas} nota(s). "
                     "El deck saldrá con ellas.")
        except Exception as exc:
            self.log(f"No se pudieron copiar las notas desde Descargas: {exc}")

    def generar_deck_cliente(self):
        """Genera el deck de 9 slides de UNA cuenta, a partir del link que el responsable comercial
        pega. El panel solo extrae el id; quien resuelve la cuenta y decide si
        existe es el prompt, que tiene acceso a vtiger."""
        crudo = self.var_deck.get().strip()
        if not crudo:
            self.log("Pegá el link de vTiger de la cuenta antes de generar.")
            return
        if not DECK_SKILL_FILE.exists():
            self.log(f"No se encontró el skill en {DECK_SKILL_FILE}. Abortando.")
            return
        cuenta_id = id_cuenta_vtiger(crudo)
        if not cuenta_id:
            self.log(f"No pude sacar un ID de cuenta de «{crudo[:60]}». Pegá el link de la "
                     "ficha de la organización en vTiger, o su ID (ej. 3x9244350).")
            return

        self.btn_deck.set_enabled(False)
        self.log(f"Generando la presentación de la cuenta {cuenta_id}…")
        prompt = (
            f"Genera el deck comercial de cliente para la cuenta de vTiger con id "
            f"EXACTAMENTE '{cuenta_id}'. el responsable comercial lo pidió a mano desde el panel pegando "
            f"este link: {crudo}\n\n"
            "PRIMERO, antes de generar nada: autentica contra vTiger y resuelve esa "
            f"cuenta con SELECT id, accountname FROM Accounts WHERE id = '{cuenta_id}'. "
            "Si no devuelve exactamente una cuenta, DETENTE y dilo claramente indicando "
            "el id que se intentó. NO busques por nombre ni elijas 'el más parecido': el "
            "id vino de un link, así que es exacto, y adivinar aquí significa generarle "
            "a el responsable comercial la presentación de otro cliente. Si sí existe, di de qué cuenta se "
            "trata antes de seguir.\n\n"
            f"Después descomprime el skill {DECK_SKILL_FILE} (es un ZIP) en una carpeta "
            "temporal, lee su SKILL.md completo y síguelo al pie de la letra para esa "
            "cuenta, usando el generador fijo (build_deck.py + template.html); no "
            "escribas el HTML a mano.\n\n"
            "Diferencias con la rutina automática de Sofía, que NO aplican aquí: no hay "
            "reunión de calendario que consultar, y la regla de no duplicar de 7 días NO "
            "se aplica -- si ya existe una presentación reciente de este cliente, "
            "genérala igual; para eso se pidió a mano.\n\n"
            f"Guarda el HTML en {PRESENTACIONES_DIR} con el nombre que indique el skill. "
            "Termina diciendo el nombre de la cuenta y la ruta del archivo generado."
        )
        self.run_claude_headless(prompt, on_done=self._tras_deck_cliente)

    def _tras_deck_cliente(self, code):
        self.btn_deck.set_enabled(True)
        if code != 0:
            self.log("La presentación terminó con error. Revisa el detalle de arriba.")
            return
        self.var_deck.set("")
        self.log("✓ Presentación generada. Se abre la carpeta.")
        open_folder(PRESENTACIONES_DIR)

    def ejecutar_cadence_generator(self):
        if not FW_CADENCE_SKILL_FILE.exists():
            self.log(f"No se encontró el skill en {FW_CADENCE_SKILL_FILE}. Abortando.")
            return
        self._recoger_notas_cadencia()
        self.btn_cadence_gen.set_enabled(False)
        self.log("Generando el deck de cadencia Freshworks (Claude Code, sin supervisión, "
                 "puede tardar unos minutos)…")
        prompt = (
            f"Sigue el skill en {FW_CADENCE_SKILL_FILE} (ignora el frontmatter YAML) al pie "
            "de la letra, de principio a fin: localiza los 2 reportes de vTiger en Descargas, "
            "arma el enriquecimiento (_vtiger.json), corre build_deck.py con template.html "
            "SIN reescribir el HTML a mano, y entrega al final la ruta del .html y el resumen "
            "tal como indica el skill."
        )
        self.run_claude_headless(prompt, on_done=self._tras_cadence_generator)

    def _tras_cadence_generator(self, code):
        self.btn_cadence_gen.set_enabled(True)
        if code != 0:
            self.log("La generación terminó con error. Revisa el detalle de arriba.")
            return
        self.log("Deck de cadencia Freshworks generado.")

    def abrir_carpeta_cadence(self):
        open_folder(FW_CADENCE_DIR)
        self.log(f"Carpeta abierta: {FW_CADENCE_DIR}")

    def generar_plan_compensacion(self):
        """Con el trimestre en curso no se pasa --trimestre: se deja el valor
        por defecto del script, que en Metricas significa "hasta la fecha". Solo
        se pasa cuando se pide un trimestre cerrado."""
        args = [] if self.sel_comp.es_actual() else ["--trimestre", self.sel_comp.get()]
        self.log(f"Plan de compensación de {self.sel_comp.get()}"
                 f"{' (en curso, hasta hoy)' if self.sel_comp.es_actual() else ''}…")
        self.run_script("Plan Compensacion.py", args)

    def generar_metricas_trimestre(self):
        args = [] if self.sel_met.es_actual() else ["--trimestre", self.sel_met.get()]
        self.log(f"Métricas de {self.sel_met.get()}"
                 f"{' (en curso, hasta hoy)' if self.sel_met.es_actual() else ''}…")
        self.run_script("Metricas Mes.py", args)

    def abrir_carpeta_compensacion(self):
        open_folder(COMP_PLAN_DIR)
        self.log(f"Carpeta abierta: {COMP_PLAN_DIR}")

    def abrir_carpeta_metricas(self):
        open_folder(METRICAS_DIR)
        self.log(f"Carpeta abierta: {METRICAS_DIR}")

    def abrir_carpeta_forecast(self):
        open_folder(FORECAST_DIR)
        self.log(f"Carpeta abierta: {FORECAST_DIR}")

    def abrir_carpeta_churn(self):
        open_folder(CHURN_INFORMES_DIR)
        self.log(f"Carpeta abierta: {CHURN_INFORMES_DIR}")

    def abrir_carpeta_gbs(self):
        open_folder(GBS_CHECKER_INFORMES_DIR)
        self.log(f"Carpeta abierta: {GBS_CHECKER_INFORMES_DIR}")

    def abrir_carpeta_gbs_maker(self):
        open_folder(GBS_MAKER_BORRADORES_DIR)
        self.log(f"Carpeta abierta: {GBS_MAKER_BORRADORES_DIR}")

    def abrir_carpeta_gbs_org_maker(self):
        open_folder(GBS_ORG_MAKER_BORRADORES_DIR)
        self.log(f"Carpeta abierta: {GBS_ORG_MAKER_BORRADORES_DIR}")

    def abrir_vacaciones(self):
        subprocess.Popen([PYTHON, str(SCRIPT_DIR / "Configurar Vacaciones.py")],
                         creationflags=SIN_VENTANA)
        self.log("Se abrió la ventana de Configurar vacaciones.")

    def abrir_opciones_forecast(self):
        ForecastOptionsWindow(self, self._forecast_desde, self._forecast_hasta,
                             self._forecast_categorias, self._forecast_segmentos,
                             self._forecast_reps, self._forecast_incluir_servicios,
                             self._aplicar_opciones_forecast)

    def _aplicar_opciones_forecast(self, desde, hasta, categorias, segmentos, reps, servicios):
        self._forecast_desde = desde
        self._forecast_hasta = hasta
        self._forecast_categorias = categorias
        self._forecast_segmentos = segmentos
        self._forecast_reps = reps
        self._forecast_incluir_servicios = servicios
        self._actualizar_resumen_forecast()
        self.log(f"Forecast options -- {self.lbl_forecast_opciones.cget('text')}")
        # Opciones nuevas -> hay que generar de nuevo antes de ver/enviar.
        self.btn_forecast_gen.set_enabled(True)
        self.btn_forecast_ver.set_enabled(False)
        self.btn_forecast_enviar.set_enabled(False)

    def _actualizar_resumen_forecast(self):
        meses = meses_del_rango(self._forecast_desde, self._forecast_hasta)
        meses_txt = (mes_etiqueta(meses[0]) if len(meses) == 1
                     else f"{mes_etiqueta(meses[0])} \u2192 {mes_etiqueta(meses[-1])} "
                          f"({len(meses)} months)")
        texto = (f"{meses_txt} · {len(self._forecast_categorias)}/{len(FORECAST_CATEGORY_OPTIONS)} "
                f"categories · {len(self._forecast_segmentos)}/{len(FORECAST_SEGMENT_OPTIONS)} segments · "
                f"{len(self._forecast_reps)}/{len(gen_ns.TEAM_OWNERS)} reps · "
                f"Services {'on' if self._forecast_incluir_servicios else 'off'}")
        self.lbl_forecast_opciones.config(text=texto)

    def generar_forecast(self):
        if not self._forecast_categorias or not self._forecast_segmentos or not self._forecast_reps:
            self.log("Elegí al menos una categoría, un segmento y un rep en 'Options…' antes de generar.")
            return
        args = [
            "--desde", self._forecast_desde,
            "--hasta", self._forecast_hasta,
            "--categorias", ",".join(sorted(self._forecast_categorias)),
            "--segmentos", ",".join(sorted(self._forecast_segmentos)),
            "--reps", ",".join(sorted(self._forecast_reps)),
        ]
        if not self._forecast_incluir_servicios:
            args.append("--excluir-servicios")
        self.btn_forecast_gen.set_enabled(False)
        self.btn_forecast_ver.set_enabled(False)
        self.btn_forecast_enviar.set_enabled(False)
        self.log(f"Generando Forecast ({self.lbl_forecast_opciones.cget('text')})…")
        self.run_script("generar_reporte_forecast.py", args, on_done=self._tras_generar_forecast)

    def _tras_generar_forecast(self, code):
        self.btn_forecast_gen.set_enabled(True)
        if code != 0:
            self.log("La generación del Forecast falló, revisa el detalle de arriba.")
            return
        self.btn_forecast_ver.set_enabled(True)
        self.btn_forecast_enviar.set_enabled(True)

    def ver_informe_forecast(self):
        if not FORECAST_MANIFEST.exists():
            self.log("Todavía no hay un Forecast generado. Dale a 'Generar' primero.")
            return
        try:
            with open(FORECAST_MANIFEST, "r", encoding="utf-8") as f:
                m = json.load(f)
            path = FORECAST_DIR / m["archivo"]
        except Exception as exc:
            self.log(f"No se pudo leer el último Forecast: {exc}")
            return
        if not path.exists():
            self.log(f"No se encontró el archivo del reporte: {path}")
            return
        webbrowser.open(path.resolve().as_uri())
        self.log(f"Forecast abierto en el navegador: {path.name}")

    def enviar_forecast(self):
        # Una vez enviado, solo queda activo 'Options' -- para volver a
        # mandar algo hay que pasar de nuevo por Options -> Generar -> Ver
        # informe, asi nunca se reenvia sin revisar que sigue vigente.
        self.btn_forecast_gen.set_enabled(False)
        self.btn_forecast_ver.set_enabled(False)
        self.btn_forecast_enviar.set_enabled(False)
        self.run_script("enviar_reporte_forecast.py", [])

    def revisar_resumen_next_step(self, ya_regenerado=False):
        # Si bajaste Excel nuevos, el resumen tiene que reflejarlos: se genera
        # primero y despues se abre, en un solo clic.
        if not ya_regenerado:
            try:
                hay_nuevos, generado = fuentes_next_step_mas_nuevas()
            except Exception:
                hay_nuevos, generado = False, None
            if hay_nuevos:
                cuando = (f"la última generación fue {generado:%d/%m %H:%M}"
                          if generado else "no hay una generación previa registrada")
                self.log(f"Los Excel de vTiger son más nuevos ({cuando}). "
                         f"Regenerando los reportes antes de mostrar el resumen…")
                if not self._limpiar_carpeta_next_step():
                    return
                self.run_script("generar_reportes_next_step.py", [],
                                on_done=lambda code: self._resumen_tras_generar(code))
                return

        try:
            archivos, fecha, sin_alertas = find_latest_next_step_files()
        except Exception as exc:
            self.log(f"No se pudo armar el resumen: {exc}")
            return

        if not archivos:
            self.log("Todavía no hay reportes de Calidad CRM generados. Dale a 'Generar' primero.")
            return

        try:
            html = build_consolidated_preview(archivos)
            preview_path = SCRIPT_DIR / "resumen_next_step_preview.html"
            preview_path.write_text(html, encoding="utf-8")
            webbrowser.open(preview_path.resolve().as_uri())
        except Exception as exc:
            self.log(f"No se pudo abrir el resumen: {exc}")
            return

        pendientes = {p.name for p in find_pending_next_step_files()}
        sin_enviar = [p.name for p in archivos if p.name in pendientes]
        estado = (f"{len(sin_enviar)} sin enviar" if sin_enviar
                  else "todos ya enviados; el resumen es solo para revisar")
        cuando = f" del {fecha:%d/%m/%Y}" if fecha else ""
        self.log(f"Resumen{cuando} abierto en el navegador "
                 f"({len(archivos)} reporte(s), {estado}): "
                 + ", ".join(p.name for p in archivos))
        if sin_alertas:
            self.log(f"Sin alertas en esta corrida (no generaron reporte): "
                     f"{', '.join(sin_alertas)}")

    def _resumen_tras_generar(self, code):
        if code != 0:
            self.log("La generación falló, así que no se abre el resumen.")
            return
        self.revisar_resumen_next_step(ya_regenerado=True)

    def _limpiar_carpeta_next_step(self):
        """Borra todos los reportes de 'Emails Vendedores' antes de generar
        de nuevo, asi nunca queda un archivo suelto de una corrida anterior
        mezclado con el resultado fresco (eso ya paso una vez: un reporte con
        formato viejo se le envio a un vendedor por error porque quedo
        tirado en la carpeta sin formar parte de la ultima generacion).
        Devuelve False si no se pudo limpiar (para abortar el resto del
        flujo en vez de generar sobre una carpeta a medio borrar)."""
        try:
            borrados = 0
            if send_ns.WATCH_FOLDER.exists():
                for p in send_ns.WATCH_FOLDER.iterdir():
                    if p.is_file():
                        p.unlink()
                        borrados += 1
            self.log(f"Se borraron {borrados} reporte(s) anterior(es) de 'Emails Vendedores'.")
            return True
        except Exception as exc:
            self.log(f"No se pudo limpiar 'Emails Vendedores': {exc}.")
            return False

    def generar_reportes_ns(self):
        if not self._limpiar_carpeta_next_step():
            return
        # Regenerar invalida cualquier plan de envio ya calculado (los
        # archivos podrian cambiar de contenido aunque no de nombre) --
        # obliga a recalcular antes de poder volver a habilitar "Enviar".
        self._plan_ns_pendiente = None
        self.btn_enviar_ns.set_enabled(False)
        self.run_script("generar_reportes_next_step.py", [])

    def calcular_envio_ns(self):
        self._plan_ns_pendiente = None
        self.btn_enviar_ns.set_enabled(False)
        self.btn_calc_envio_ns.set_enabled(False)
        self.log("Calculando qué reportes de Next Step faltan por enviar…")
        self.run_script("enviar_reportes_next_step.py", ["--plan-json", str(NEXT_STEP_PLAN_FILE)],
                        on_done=self._plan_ns_calculado)

    def _plan_ns_calculado(self, code):
        self.btn_calc_envio_ns.set_enabled(True)
        if code != 0:
            self.log("No se pudo calcular el plan de envío. Revisa el detalle de arriba.")
            return
        try:
            with open(NEXT_STEP_PLAN_FILE, "r", encoding="utf-8") as f:
                plan = json.load(f)
        except Exception as exc:
            self.log(f"No se pudo leer el plan: {exc}")
            return
        self._plan_ns_pendiente = plan
        for linea in format_next_step_plan_lines(plan):
            self.log(linea)
        hay_envios = any(item.get("accion") == "enviar" for item in plan.get("items", []))
        self.btn_enviar_ns.set_enabled(hay_envios)

    def enviar_plan_ns(self):
        plan = self._plan_ns_pendiente
        if not plan or not any(item.get("accion") == "enviar" for item in plan.get("items", [])):
            self.log("No hay un plan pendiente por enviar. Dale a 'Calcular envío' primero.")
            return
        self.log("Plan aprobado. Enviando por Outlook con el reporte adjunto…")
        self._plan_ns_pendiente = None
        self.btn_calc_envio_ns.set_enabled(False)
        self.btn_enviar_ns.set_enabled(False)
        # Sin argumentos = envio real. El script reclasifica la carpeta con la
        # misma find_pending() que uso para el plan, asi que manda exactamente
        # lo que se acaba de mostrar; y es el mismo script el que marca
        # processed_files, de ahi que el envio sea directo (send_mail) y no una
        # ventana de redaccion: con Display el estado diria "enviado" sin que
        # nadie hubiera dado clic todavia.
        self.run_script("enviar_reportes_next_step.py", [],
                        on_done=self._tras_enviar_ns)

    def _tras_enviar_ns(self, code):
        self.btn_calc_envio_ns.set_enabled(True)
        if code != 0:
            self.log("El envío terminó con error. Revisa el detalle de arriba (Outlook cerrado o sin "
                     "sesión, o un adjunto que no se encontró) y vuelve a 'Calcular envío' antes "
                     "de reintentar.")
            return
        self.log("✓ Envío completado. Dale a 'Calcular envío' de nuevo si querés confirmar que no "
                 "quedó nada pendiente.")

    # ---- ejecucion ----

    def set_buttons_enabled(self, enabled):
        for b in self._buttons:
            b.set_enabled(enabled)

    def run_script(self, filename, args, on_done=None):
        self.set_buttons_enabled(False)
        self.log(f"— Ejecutando {filename} {' '.join(args)} —")
        hilo = threading.Thread(target=self._run_script_thread,
                                args=(filename, args, on_done), daemon=True)
        hilo.start()

    def _run_script_thread(self, filename, args, on_done=None):
        script_path = SCRIPT_DIR / filename
        # Sin esto, el script hijo escribe su salida con el codepage ANSI
        # del sistema (cp1252) porque su stdout esta redirigido, no es una
        # consola real -- y como aca se decodifica como UTF-8, cualquier
        # acento salia como caracter de reemplazo.
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        try:
            proc = subprocess.run(
                [PYTHON, str(script_path), *args],
                cwd=str(SCRIPT_DIR),
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                # Normalmente PYTHON es pythonw y no hay consola que ocultar,
                # pero si el panel se arranca con python.exe (ej. desde una
                # terminal para ver errores) cada script abriria su ventana.
                creationflags=SIN_VENTANA,
                env=env,
            )
            output = (proc.stdout or "") + (proc.stderr or "")
            code = proc.returncode
        except Exception as exc:
            output = f"Error al ejecutar: {exc}"
            code = -1

        def finish():
            for line in output.splitlines():
                self.log(line)
            self.log(f"— {filename} terminó (código {code}) —")
            self.set_buttons_enabled(True)
            self.refresh_status()
            if on_done is not None:
                on_done(code)

        self.root.after(0, finish)

    def run_rutina(self, skill):
        """Corre a mano una de las rutinas programadas, con el mismo
        Run-Rutina.ps1 que usa el Programador de tareas. Se llama al .ps1 en
        vez de armar el prompt aca para que no haya dos definiciones de la
        misma rutina que se puedan desincronizar: el script resuelve la skill,
        el CLI y el prompt, y ademas deja el log y el marcador diario."""
        if not RUNNER_RUTINAS.exists():
            self.log(f"No se encontró el ejecutor de rutinas en {RUNNER_RUTINAS}. Abortando.")
            return
        if not skill_file(skill).exists():
            self.log(f"No se encontró el skill '{skill}'. Abortando, no se ejecuta nada.")
            return
        self.set_buttons_enabled(False)
        self.log(f"— Corriendo la rutina '{skill}' a mano (sin supervisión) —")
        self.log("Es la misma corrida que haría la tarea programada; puede tardar varios minutos.")
        hilo = threading.Thread(target=self._run_rutina_thread,
                                args=(skill,), daemon=True)
        hilo.start()

    def _run_rutina_thread(self, skill):
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-File", str(RUNNER_RUTINAS),
                 "-Skill", skill, "-Force"],
                cwd=str(GB_CLAUDE),
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                stdin=subprocess.DEVNULL, creationflags=SIN_VENTANA,
                env=env, timeout=CLAUDE_HEADLESS_TIMEOUT_SECONDS,
            )
            output = (proc.stdout or "") + (proc.stderr or "")
            code = proc.returncode
        except subprocess.TimeoutExpired as exc:
            parcial = (exc.stdout or "") + (exc.stderr or "")
            if isinstance(parcial, bytes):
                parcial = parcial.decode("utf-8", "replace")
            minutos = CLAUDE_HEADLESS_TIMEOUT_SECONDS // 60
            output = (f"{parcial}\nLa rutina no terminó en {minutos} minutos; se mató el "
                      "proceso. Revisa el log en Rutinas/logs antes de reintentar.")
            code = -1
        except Exception as exc:
            output = f"Error al ejecutar la rutina: {exc}"
            code = -1

        def finish():
            for line in output.splitlines():
                self.log(line)
            if code == 0:
                self.log(f"✓ Rutina '{skill}' terminada. El detalle completo queda en "
                         f"Rutinas/logs/{skill}.log")
            else:
                self.log(f"— La rutina '{skill}' terminó con error (código {code}) —")
            self.set_buttons_enabled(True)
            self.refresh_status()

        self.root.after(0, finish)

    def run_claude_headless(self, prompt, on_done=None,
                            timeout=CLAUDE_HEADLESS_TIMEOUT_SECONDS):
        """Corre Claude Code sin supervision (mismo mecanismo que los .sh de
        cron en Rutinas/: 'claude -p <prompt> --dangerously-skip-permissions'),
        para tareas que necesitan herramientas MCP (SharePoint, Outlook,
        vTiger) a las que este panel Python no tiene acceso directo. Solo
        se usa para ejecutar un plan que el responsable comercial ya aprobo con un clic (ej.
        'Enviar' en Calidad CRM) -- nunca para calcular o decidir nada."""
        self.set_buttons_enabled(False)
        self.log("— Ejecutando Claude Code (sin supervisión) —")
        hilo = threading.Thread(target=self._run_claude_thread,
                                args=(prompt, on_done, timeout), daemon=True)
        hilo.start()

    def _run_claude_thread(self, prompt, on_done=None,
                           timeout=CLAUDE_HEADLESS_TIMEOUT_SECONDS):
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        if CLAUDE_BIN is None:
            # Mejor un mensaje claro que un FileNotFoundError opaco: esto pasa
            # cuando el CLI de Claude no esta instalado en la maquina (caso
            # tipico despues de migrar de PC).
            def sin_cli():
                self.log("No se encontró el CLI de Claude Code en este equipo.")
                self.log("Instálalo y vuelve a intentar; el panel lo busca en el "
                         "PATH y en las rutas habituales del instalador.")
                self.log("— Claude terminó (código -1) —")
                self.set_buttons_enabled(True)
                self.refresh_status()
                if on_done is not None:
                    on_done(-1)
            self.root.after(0, sin_cli)
            return
        try:
            proc = subprocess.run(
                [str(CLAUDE_BIN), "-p", prompt, "--dangerously-skip-permissions"],
                cwd=str(SCRIPT_DIR.parent),
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                # Sin stdin cerrado, el CLI se queda 3 s esperando datos por la
                # entrada estandar y avisa "no stdin data received in 3s" antes
                # de seguir: el prompt ya va en -p, no hay nada que leer por ahi
                # (9-sep-2026).
                stdin=subprocess.DEVNULL,
                creationflags=SIN_VENTANA,
                env=env, timeout=timeout,
            )
            output = (proc.stdout or "") + (proc.stderr or "")
            code = proc.returncode
        except subprocess.TimeoutExpired as exc:
            # subprocess.run ya mato el proceso (y sus hijos directos) antes
            # de relanzar esto; solo falta avisar en vez de quedarnos
            # esperando para siempre.
            parcial = (exc.stdout or "") + (exc.stderr or "")
            minutos = timeout // 60
            output = (
                f"{parcial}\nClaude Code no respondio en {minutos} minutos; se "
                "mato el proceso. Puede haber quedado algo a medio enviar -- "
                "revisa el estado (ej. 'Calcular plan') antes de reintentar. "
                "Si los correos/SharePoint son muy lentos, puede aumentar el timeout en el código."
            )
            code = -1
        except Exception as exc:
            output = f"Error al ejecutar Claude: {exc}"
            code = -1

        def finish():
            for line in output.splitlines():
                self.log(line)
            self.log(f"— Claude terminó (código {code}) —")
            self.set_buttons_enabled(True)
            self.refresh_status()
            if on_done is not None:
                on_done(code)

        self.root.after(0, finish)


def main():
    activar_dpi()
    identidad_barra_tareas()
    # className fijo (no el "Tk" generico por defecto) para que la barra de
    # tareas/dock de Ubuntu pueda identificar esta ventana y mostrarle el
    # icono de GB Advisors -- ver Panel de Automatizaciones.desktop
    # (StartupWMClass debe calzar con el "Panel-automatizaciones" que Tk
    # arma a partir de este className: mayuscula inicial, resto igual).
    root = tk.Tk(className="panel-automatizaciones")
    panel = Panel(root)  # noqa: F841 (referencia viva; ver nota en __init__)
    root.mainloop()


if __name__ == "__main__":
    main()
