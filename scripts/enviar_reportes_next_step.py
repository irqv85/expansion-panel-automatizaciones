"""
Revisa la carpeta "Emails Vendedores" y envia por Outlook los reportes que
todavia no se han enviado:

  - "Next Step - {Vendedor}_{fecha}.html"  -> se envia a ese vendedor,
    con copia a Yeniree, adjuntando el HTML.
  - "Account Book Attention - General_{fecha}.html" -> se envia a los 4
    vendedores juntos, con copia a Yeniree. Solo se envia UNA vez por
    semana laboral (si ya se envio esta semana ISO, se omite aunque haya
    un archivo nuevo; el archivo igual se marca como procesado).

Requiere: Outlook de escritorio instalado y con sesion iniciada.

Ejecutar con: python enviar_reportes_next_step.py
Modo de prueba (no envia nada, solo muestra que haria): python enviar_reportes_next_step.py --dry-run

Modo plan (no envia nada, no toca el estado -- solo calcula que se enviaria y
lo escribe a un JSON para que otro proceso, ej. el panel o Claude via MCP de
M365, lo revise/ejecute despues): python enviar_reportes_next_step.py --plan-json <ruta.json>
"""
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

from paths import EMAILS_VENDEDORES_DIR as WATCH_FOLDER

SCRIPT_DIR = Path(__file__).resolve().parent
STATE_FILE = SCRIPT_DIR / "state_envio_next_step.json"
LOG_FILE = SCRIPT_DIR / "envio_next_step.log"
MANIFEST_FILE = SCRIPT_DIR / "ultima_generacion_next_step.json"

SELLERS = [
    {"name": "Vendedora 3", "email": "vendedora1@tuempresa.com"},
    {"name": "Vendedor 2", "email": "vendedor2@tuempresa.com"},
    {"name": "Vendedora 1", "email": "vendedora3@tuempresa.com"},
    {"name": "Vendedor 4", "email": "vendedor4@tuempresa.com"},
]
CC_EMAIL = "copia@tuempresa.com"

VENDOR_FILE_RE = re.compile(r"^Next Step - (.+)_(\d{2}-\d{2}-\d{4})\.html$")
GENERAL_FILE_RE = re.compile(r"^Account Book Attention - General_(\d{2}-\d{2}-\d{4})\.html$")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)


def load_state():
    if STATE_FILE.exists():
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"processed_files": [], "last_general_week": None}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def find_seller(name):
    return next((s for s in SELLERS if s["name"] == name), None)


def load_manifest_files():
    """Nombres de archivo de la ultima corrida de generar_reportes_next_step.py,
    o None si no hay manifiesto (version vieja de la carpeta, sin regenerar).

    Sin esto, un archivo viejo que quedo en la carpeta (de una corrida previa,
    con reglas o formato desactualizados) se enviaria igual con solo con que
    no estuviera todavia marcado como procesado -- eso ya paso una vez con un
    reporte de el vendedor de un formato anterior a "Calidad de CRM"."""
    if not MANIFEST_FILE.exists():
        return None
    try:
        with open(MANIFEST_FILE, "r", encoding="utf-8") as f:
            manifiesto = json.load(f)
        return set(manifiesto.get("archivos", []))
    except Exception:
        return None


def iso_week_key(d):
    year, week, _ = d.isocalendar()
    return f"{year}-W{week:02d}"


def build_vendor_mail(seller):
    first_name = seller["name"].split()[0]
    subject = f"Reporte de vTiger - {seller['name']}"
    body = (
        f"Hola {first_name}, te comparto el reporte con los campos de vTiger que requieren tu atención.\n\n"
        f"Avísame cuando estén corregidos.\n\n"
        f"Cualquier duda estoy a la orden,"
    )
    return subject, body


def build_general_mail():
    subject = "Account book attention - Reporte."
    body = (
        "Hola team,\n\n"
        "Un gusto saludarles.\n\n"
        "Les comparto el account book attention, se les recuerda que este porcentaje debe estar "
        "ajustado de acuerdo a lo conversado.\n\n"
        "Saludos."
    )
    return subject, body


def send_mail(outlook_app, to, cc, subject, body, attachment, dry_run):
    """Envia el reporte por Outlook, sin intervencion.

    El envio es directo (no abre ventana de redaccion) porque este script
    marca cada archivo como procesado en state_envio_next_step.json: si solo
    abriera la ventana, el estado diria "enviado" sin que nadie hubiera dado
    clic todavia. Para revisar antes de mandar existe --dry-run, y para el
    envio supervisado esta mail_helper.display_mail.
    """
    import mail_helper
    mail_helper.send_mail(
        to=to, cc=cc or None, subject=subject, body=body,
        attachment=attachment, dry_run=dry_run,
    )


def find_pending(files, processed, vigentes):
    """Separa los archivos de la carpeta en (vendor_files, general_files,
    stale), la misma clasificacion que usan tanto el envio real/dry-run como
    el modo --plan-json -- una sola fuente de verdad para que el plan que se
    le muestra a el responsable comercial sea exactamente lo que despues se ejecuta."""
    vendor_files = []
    general_files = []
    stale = []
    for path in files:
        if not path.is_file() or path.name in processed:
            continue
        m_vendor = VENDOR_FILE_RE.match(path.name)
        m_general = GENERAL_FILE_RE.match(path.name)
        if not m_vendor and not m_general:
            continue
        if vigentes is not None and path.name not in vigentes:
            stale.append(path.name)
            continue
        if m_vendor:
            vendor_files.append((path, m_vendor.group(1)))
        elif m_general:
            general_files.append(path)
    return vendor_files, general_files, stale


def write_plan_json(plan_path, now, vendor_files, general_files, stale, last_general_week):
    """Calcula el plan de envio (quien recibe que, y por que se omite el
    general si corresponde) SIN enviar nada ni tocar el estado, y lo escribe
    a JSON para que el panel lo muestre y, tras la aprobacion de el responsable comercial, otro
    proceso (Claude via MCP de M365, que es quien tiene acceso a SharePoint/
    Outlook) lo ejecute tal cual sin tener que recalcularlo."""
    items = []
    for path, vendor_name in vendor_files:
        seller = find_seller(vendor_name)
        if not seller:
            items.append({
                "archivo": path.name, "tipo": "vendedor", "vendedor": vendor_name,
                "accion": "omitir", "motivo": "vendedor no reconocido",
            })
            continue
        subject, _ = build_vendor_mail(seller)
        items.append({
            "archivo": path.name, "tipo": "vendedor", "vendedor": seller["name"],
            "email": seller["email"], "cc": CC_EMAIL, "asunto": subject,
            "accion": "enviar",
        })

    if general_files:
        latest_general = general_files[-1]
        current_week = iso_week_key(now)
        subject, _ = build_general_mail()
        if current_week == last_general_week:
            items.append({
                "archivo": latest_general.name, "tipo": "general",
                "destinatarios": [s["email"] for s in SELLERS], "cc": CC_EMAIL,
                "asunto": subject, "accion": "omitir",
                "motivo": f"ya se envio esta semana ({current_week})",
            })
        else:
            items.append({
                "archivo": latest_general.name, "tipo": "general",
                "destinatarios": [s["email"] for s in SELLERS], "cc": CC_EMAIL,
                "asunto": subject, "accion": "enviar",
            })
        # Los demas archivos generales de la corrida (si hubiera mas de uno,
        # no deberia pasar) se listan como omitidos para que igual queden
        # marcados como procesados al ejecutar -- ver SKILL.md paso 6.
        for path in general_files[:-1]:
            items.append({
                "archivo": path.name, "tipo": "general", "accion": "omitir",
                "motivo": "superado por una corrida de generacion mas reciente",
            })

    plan = {
        "generado": now.isoformat(),
        "items": items,
        "desactualizados": stale,
    }
    with open(plan_path, "w", encoding="utf-8") as f:
        json.dump(plan, f, indent=2, ensure_ascii=False)
    return plan


def main():
    dry_run = "--dry-run" in sys.argv
    plan_json_arg = None
    if "--plan-json" in sys.argv:
        idx = sys.argv.index("--plan-json")
        if idx + 1 >= len(sys.argv):
            print("Falta la ruta despues de --plan-json.")
            sys.exit(1)
        plan_json_arg = Path(sys.argv[idx + 1])

    now = datetime.now()
    logging.info("=" * 60)
    logging.info(f"Inicio de ejecucion: {now} (dry_run={dry_run}, plan_json={plan_json_arg})")

    if not WATCH_FOLDER.exists():
        print(f"No existe la carpeta: {WATCH_FOLDER}")
        return

    state = load_state()
    processed = set(state.get("processed_files", []))
    last_general_week = state.get("last_general_week")

    outlook_app = None  # mail_helper abre su propia instancia de Outlook; se conserva por firma

    files = sorted(WATCH_FOLDER.iterdir(), key=lambda p: p.stat().st_mtime)

    vigentes = load_manifest_files()
    if vigentes is None:
        print("AVISO: no se encontro el manifiesto de la ultima generacion; "
              "se enviara cualquier archivo pendiente (podria estar desactualizado).")
        logging.warning("Sin manifiesto: se envia sin filtrar por vigencia.")

    vendor_files, general_files, stale = find_pending(files, processed, vigentes)

    if stale:
        print(f"Archivo(s) desactualizado(s) en la carpeta, no forman parte de la "
              f"ultima generacion -- se omiten sin enviar: {', '.join(stale)}")
        logging.info(f"Omitidos por no estar en el manifiesto vigente: {', '.join(stale)}")

    if plan_json_arg is not None:
        # Siempre se escribe el plan (incluso vacio) aunque no haya nada
        # pendiente -- si no, un plan viejo de una corrida anterior quedaria
        # tirado en el archivo y alguien podria ejecutarlo por error creyendo
        # que sigue vigente.
        plan = write_plan_json(plan_json_arg, now, vendor_files, general_files, stale, last_general_week)
        if not plan["items"]:
            print("No hay reportes nuevos por enviar.")
        else:
            print(f"Plan de envio calculado ({len(plan['items'])} archivo(s)):")
            for item in plan["items"]:
                destino = item.get("email") or (
                    f"{len(item['destinatarios'])} vendedores" if "destinatarios" in item else "?"
                )
                if item["accion"] == "enviar":
                    print(f"  - {item['archivo']} -> {destino}: ENVIAR")
                else:
                    print(f"  - {item['archivo']} -> {destino}: OMITIR ({item.get('motivo', '')})")
        print(f"Guardado en {plan_json_arg}")
        logging.info(f"Plan escrito en {plan_json_arg} con {len(plan['items'])} item(s).")
        return

    if not vendor_files and not general_files:
        print("No hay reportes nuevos por enviar.")
        logging.info("No hay reportes nuevos.")
        return

    # --- reportes por vendedor ---
    for path, vendor_name in vendor_files:
        seller = find_seller(vendor_name)
        if not seller:
            logging.warning(f"No reconozco al vendedor '{vendor_name}' ({path.name}), se omite.")
            print(f"Vendedor no reconocido: '{vendor_name}' ({path.name}) - omitido.")
            continue
        subject, body = build_vendor_mail(seller)
        send_mail(outlook_app, seller["email"], CC_EMAIL, subject, body, path, dry_run)
        logging.info(f"Enviado '{path.name}' a {seller['name']} ({seller['email']}), CC {CC_EMAIL}")
        print(f"Enviado a {seller['name']}: {path.name}")
        processed.add(path.name)

    # --- reporte general (maximo 1 vez por semana laboral) ---
    if general_files:
        latest_general = general_files[-1]
        current_week = iso_week_key(now)
        if current_week == last_general_week:
            print(f"Account Book Attention ya se envió esta semana ({current_week}), se omite.")
            logging.info(f"General omitido, ya enviado esta semana ({current_week}).")
        else:
            to_all = ";".join(s["email"] for s in SELLERS)
            subject, body = build_general_mail()
            send_mail(outlook_app, to_all, CC_EMAIL, subject, body, latest_general, dry_run)
            logging.info(f"Enviado general '{latest_general.name}' a {to_all}, CC {CC_EMAIL}")
            print(f"Enviado general a todo el equipo: {latest_general.name}")
            last_general_week = current_week
        for path in general_files:
            processed.add(path.name)

    state["processed_files"] = sorted(processed)
    state["last_general_week"] = last_general_week
    if not dry_run:
        save_state(state)
    logging.info("Fin de ejecucion.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logging.exception("Error inesperado durante la ejecucion.")
        sys.exit(1)
