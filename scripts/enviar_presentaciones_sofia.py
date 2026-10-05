"""
Envia por Outlook de escritorio las presentaciones que la skill
"sofia-schedule-vendedores" ya repartio en plan_sofia.json, con el HTML
ADJUNTO, y reenvia la invitacion de la reunion cuando corresponde.

Por que existe (26-sep-2026, pedido de el responsable comercial tras una corrida fallida):

1. Los correos salian con un ENLACE a SharePoint en vez del archivo. Eso venia
   de agosto, cuando el envio se migro al conector de M365, que no soporta
   adjuntos en ningun flujo. Con Outlook COM si se puede adjuntar, igual que
   ya se hizo con los reportes de Calidad CRM el 9-sep-2026.

2. A Francisco no le llego la invitacion de su reunion. La skill la buscaba
   como CORREO: outlook_email_search(sender="sofia", query="<asunto>") y se
   quedaba solo con los de asunto EXACTO. Fallaba por dos motivos a la vez: la
   invitacion real de "Discovery session - Laptop Center CR - 09/29/2026"
   estaba en la bandeja como "Provisional: Discovery session - ..." (prefijo
   que pone Outlook a las tentativas), y su remitente no es Sofia sino el
   propio buzon. No casaba ni por asunto ni por remitente, asi que se salto en
   silencio. Aqui se busca la CITA en el calendario por fecha de inicio y
   organizador, que es lo que la reunion realmente es, y se reenvia esa.

Ejecutar con: python enviar_presentaciones_sofia.py
Modo de prueba (no envia nada, no toca el estado):
             python enviar_presentaciones_sofia.py --dry-run

Solo la invitacion, sin reenviar el correo con el adjunto (util cuando el
correo ya salio en una corrida anterior y lo unico que falta es la reunion):
             python enviar_presentaciones_sofia.py --solo-invitacion
Se puede acotar a una persona con --vendedor "Vendedor 2".
"""
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

from paths import PRESENTACIONES_DIR

SCRIPT_DIR = Path(__file__).resolve().parent
PLAN_FILE = SCRIPT_DIR / "plan_sofia.json"
STATE_FILE = SCRIPT_DIR / "state.json"

from equipo import CC_EMAIL, VENDEDORES  # noqa: F401

# Margen al buscar la cita por hora de inicio. El plan trae la hora con zona
# horaria y Outlook la devuelve en la del buzon; un cuarto de hora absorbe esa
# diferencia sin llegar a confundir dos reuniones distintas del mismo dia.
MARGEN_MINUTOS = 15
OL_FOLDER_CALENDAR = 9


def correo_de(nombre):
    if nombre in VENDEDORES:
        return VENDEDORES[nombre]
    objetivo = str(nombre or "").strip().lower()
    for n, c in VENDEDORES.items():
        if n.lower() == objetivo:
            return c
    return None


def cargar_plan():
    if not PLAN_FILE.exists():
        print(f"No existe {PLAN_FILE.name}. Corre 'Calcular plan' primero.")
        sys.exit(1)
    with open(PLAN_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def cargar_estado():
    if STATE_FILE.exists():
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"rotation_index": 0, "processed_files": []}


def guardar_estado(estado):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def cuerpo_correo(nombre_vendedor, empresa):
    primer_nombre = str(nombre_vendedor).split()[0]
    return (
        f"Hola {primer_nombre}\n\n"
        f"Te adjunto el archivo de presentación para tu reunión con {empresa}.\n\n"
        "Cualquier duda dejame saber."
    )


def _a_naive(dt):
    """Outlook compara en hora local del buzon y sin zona; el plan trae la
    hora con offset. Se quita la zona para poder restar las dos."""
    return dt.replace(tzinfo=None) if dt.tzinfo is not None else dt


def buscar_cita(inicio_iso, asunto):
    """Devuelve la cita del calendario que corresponde a la reunion, o None.

    Se busca por VENTANA DE TIEMPO y se desempata por asunto y organizador, en
    vez de exigir un asunto exacto: Outlook antepone prefijos ("Provisional:",
    "Cancelado:") segun el estado de la respuesta, y un match exacto se pierde
    la cita justo cuando esta sin confirmar."""
    import win32com.client as win32

    try:
        objetivo = _a_naive(datetime.fromisoformat(inicio_iso))
    except Exception:
        return None, "la fecha de inicio del plan no se pudo interpretar"

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    items = ns.GetDefaultFolder(OL_FOLDER_CALENDAR).Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    desde = objetivo - timedelta(minutes=MARGEN_MINUTOS)
    hasta = objetivo + timedelta(minutes=MARGEN_MINUTOS)
    filtro = (f"[Start] >= '{desde.strftime('%m/%d/%Y %H:%M %p')}' AND "
              f"[Start] <= '{hasta.strftime('%m/%d/%Y %H:%M %p')}'")

    candidatas = []
    for it in items.Restrict(filtro):
        try:
            candidatas.append((it.Subject or "", getattr(it, "Organizer", "") or "", it))
        except Exception:
            continue
    if not candidatas:
        return None, f"no hay ninguna cita en el calendario cerca de {objetivo:%d-%m-%Y %H:%M}"

    clave = str(asunto or "").strip().lower()

    def puntua(par):
        subj, org, _ = par
        s = subj.lower()
        return (
            2 if clave and clave in s else (1 if clave and s in clave else 0),
            1 if "sofia" in org.lower() else 0,
        )

    candidatas.sort(key=puntua, reverse=True)
    mejor = candidatas[0]
    if puntua(mejor) == (0, 0):
        return None, (f"hay {len(candidatas)} cita(s) a esa hora pero ninguna coincide con "
                      f"«{asunto}» ni la organiza Sofía")
    return mejor[2], None


# Carpetas donde puede estar la invitacion original, en este orden. Elementos
# eliminados va incluido a proposito: Outlook mueve ahi la convocatoria en
# cuanto se responde, asi que la invitacion de una reunion YA ACEPTADA vive en
# la papelera, no en la bandeja (comprobado el 28-sep-2026 con la de Laptop
# Center CR).
# Enviados NO se busca: una convocatoria ahi es algo que mando el responsable comercial, no lo que
# le mando Sofia. Sin esta exclusion se elegia su propio reenvio manual ("RV:
# Discovery session...") por ser el mas reciente, y se habria reenviado un
# reenvio (28-sep-2026).
CARPETAS_BUSQUEDA = (6, 3)  # Bandeja de entrada, Elementos eliminados
CLASE_MEETING_ITEM = 53
PREFIJOS = ("re:", "rv:", "fw:", "fwd:", "provisional:", "tentative:",
            "aceptado:", "accepted:", "cancelado:", "canceled:", "cancelled:")


def _sin_prefijos(texto):
    t = str(texto or "").strip().lower()
    hubo = True
    while hubo:
        hubo = False
        for pre in PREFIJOS:
            if t.startswith(pre):
                t = t[len(pre):].strip()
                hubo = True
    return t


def buscar_convocatoria(ns, asunto):
    """Devuelve la convocatoria original (MeetingItem) de esa reunion, o None.

    Es lo que hay que reenviar para que al vendedor le llegue una invitacion de
    verdad, con su enlace de Teams y su boton de unirse. NO sirve reenviar la
    cita del calendario: AppointmentItem solo ofrece ForwardAsVcal(), que manda
    un .vcs adjunto y no se parece en nada a una invitacion (28-sep-2026, se le
    mando asi a Francisco por error y hubo que rehacerlo a mano).

    Tampoco sirve buscar por asunto exacto: Outlook antepone prefijos segun el
    estado ("Provisional:", "RV:"), y ademas el "Provisional:" que aparece en
    Enviados es la RESPUESTA tentativa de el responsable comercial, no la convocatoria. Por eso se
    filtra por clase de mensaje y se compara el asunto ya sin prefijos."""
    clave = _sin_prefijos(asunto)
    if not clave:
        return None
    candidatas = []
    for fid in CARPETAS_BUSQUEDA:
        try:
            carpeta = ns.GetDefaultFolder(fid)
        except Exception:
            continue
        try:
            items = carpeta.Items
            items.Sort("[ReceivedTime]", True)
        except Exception:
            continue
        for it in items:
            try:
                if it.Class != CLASE_MEETING_ITEM:
                    continue
                if not str(getattr(it, "MessageClass", "")).startswith(
                        "IPM.Schedule.Meeting.Request"):
                    continue
                if _sin_prefijos(it.Subject) != clave:
                    continue
                candidatas.append(it)
            except Exception:
                continue
    if not candidatas:
        return None
    # La mas reciente: si la reunion se reprogramo, la ultima convocatoria es
    # la vigente.
    def cuando(it):
        for attr in ("ReceivedTime", "CreationTime"):
            try:
                return getattr(it, attr)
            except Exception:
                continue
        return None

    def es_reenvio(it):
        try:
            return str(it.Subject or "").strip().lower().startswith(("rv:", "fw:", "fwd:"))
        except Exception:
            return False

    # Primero la convocatoria original (asunto sin prefijo de reenvio) y, entre
    # las que empatan, la mas reciente: si la reunion se reprogramo, la ultima
    # convocatoria es la vigente.
    try:
        candidatas.sort(key=lambda it: (not es_reenvio(it), cuando(it)), reverse=True)
    except Exception:
        pass
    return candidatas[0]


def reenviar_invitacion(cita, destinatarios, dry_run):
    if dry_run:
        print(f"      [DRY-RUN] reenviaría la invitación «{cita.Subject}» a "
              f"{', '.join(destinatarios)}")
        return True, None
    try:
        fwd = cita.Forward()
    except Exception as exc:
        return False, f"no se pudo crear el reenvío: {exc}"
    try:
        # Los destinatarios van por Recipients, NO por .To: al reenviar una
        # convocatoria el resultado es otro MeetingItem, y ahi .To es de solo
        # lectura ("Property 'Forward.To' can not be set"). Ese fue el motivo de
        # que el 28-sep-2026 saliera el correo de Latin American School pero no
        # su invitacion. Con un MailItem (el caso del .vcs) si se podia, y por
        # eso no se habia notado.
        for direccion in destinatarios:
            fwd.Recipients.Add(direccion)
        if not fwd.Recipients.ResolveAll():
            sin_resolver = [r.Name for r in fwd.Recipients if not r.Resolved]
            return False, f"destinatario(s) sin resolver en Outlook: {', '.join(sin_resolver)}"
        fwd.Send()
        return True, None
    except Exception as exc:
        return False, str(exc)


def main():
    dry_run = "--dry-run" in sys.argv
    solo_invitacion = "--solo-invitacion" in sys.argv
    vendedor_filtro = None
    if "--vendedor" in sys.argv:
        i = sys.argv.index("--vendedor")
        if i + 1 < len(sys.argv):
            vendedor_filtro = sys.argv[i + 1].strip().lower()
    plan = cargar_plan()
    entradas = plan.get("entries", [])
    if not entradas:
        print("El plan no tiene entradas. Nada que enviar.")
        return

    import mail_helper

    estado = cargar_estado()
    procesados = list(estado.get("processed_files", []))
    enviados, fallidos, invitaciones, sin_invitacion = 0, 0, 0, []
    omitidos = []

    for e in entradas:
        archivo = PRESENTACIONES_DIR / e["file_name"]
        vendedor = e.get("seller_name")
        destino = correo_de(vendedor)
        empresa = e.get("company") or archivo.stem

        if vendedor_filtro and str(vendedor or "").strip().lower() != vendedor_filtro:
            continue
        print(f"  - {vendedor} ← {e['file_name']}")
        # Segunda barrera, por si llega un plan viejo o una version futura del
        # skill vuelve a repartir por turno lo que no tiene reunion. El
        # 25-sep-2026 salieron dos correos de mas por eso: eran decks sueltos
        # generados a mano desde la tarjeta "Presentacion de cliente", que caen
        # en la misma carpeta y se repartieron como si fueran reuniones. Un
        # archivo sin reunion no se manda.
        if not e.get("meeting_start"):
            print(f"      OMITIDO: no tiene reunión asociada "
                  f"(motivo del plan: {e.get('reason') or 'sin motivo'})")
            omitidos.append(f"{e['file_name']} ({vendedor}): sin reunión")
            continue
        if not destino:
            print(f"      OMITIDO: no reconozco al vendedor «{vendedor}»")
            fallidos += 1
            continue
        if not archivo.exists():
            print(f"      OMITIDO: no existe {archivo}")
            fallidos += 1
            continue

        if solo_invitacion:
            print("      (--solo-invitacion: no se reenvía el correo con el adjunto)")
        else:
            try:
                mail_helper.send_mail(
                    to=destino, cc=CC_EMAIL,
                    subject=f"Archivo de presentacion - {empresa}",
                    body=cuerpo_correo(vendedor, empresa),
                    attachment=archivo, dry_run=dry_run,
                )
                enviados += 1
            except Exception as exc:
                print(f"      ERROR al enviar: {exc}")
                fallidos += 1
                continue

        if e.get("send_invite"):
            import win32com.client as _w32
            ns = _w32.Dispatch("Outlook.Application").GetNamespace("MAPI")
            # Primero el asunto del plan; si no aparece, se usa el de la cita
            # del calendario, que es la fuente mas fiable del asunto real.
            asunto = e.get("meeting_subject")
            convocatoria = buscar_convocatoria(ns, asunto)
            motivo = None
            if convocatoria is None:
                cita, motivo = buscar_cita(e.get("meeting_start"), asunto)
                if cita is not None:
                    convocatoria = buscar_convocatoria(ns, cita.Subject)
                    motivo = (f"no se encontró la convocatoria original de «{cita.Subject}» "
                              "en bandeja de entrada ni en eliminados")
            if convocatoria is None:
                motivo = motivo or f"no se encontró la convocatoria de «{asunto}»"
                print(f"      SIN INVITACIÓN: {motivo}")
                sin_invitacion.append(f"{vendedor} ({empresa}): {motivo}")
            elif dry_run:
                print(f"      [DRY-RUN] reenviaría la invitación «{convocatoria.Subject}» "
                      f"({convocatoria.MessageClass}) a {destino}, {CC_EMAIL}")
                invitaciones += 1
            else:
                ok, err = reenviar_invitacion(convocatoria, [destino, CC_EMAIL], dry_run)
                if ok:
                    invitaciones += 1
                    print(f"      invitación reenviada: «{convocatoria.Subject}»")
                else:
                    print(f"      SIN INVITACIÓN: no se pudo reenviar ({err})")
                    sin_invitacion.append(f"{vendedor} ({empresa}): {err}")

        if e["file_name"] not in procesados:
            procesados.append(e["file_name"])

    if not dry_run and not solo_invitacion:
        estado["processed_files"] = procesados
        if "rotation_index" in plan:
            estado["rotation_index"] = plan["rotation_index"]
        guardar_estado(estado)

    print()
    print(f"Correos enviados: {enviados} | fallidos: {fallidos} | omitidos: {len(omitidos)}")
    print(f"Invitaciones reenviadas: {invitaciones}")
    if omitidos:
        print("Omitidos por no tener reunión asociada:")
        for o in omitidos:
            print(f"  - {o}")
    if sin_invitacion:
        print("Sin invitación (el correo SÍ salió):")
        for s in sin_invitacion:
            print(f"  - {s}")
    if dry_run:
        print("\n[DRY-RUN] No se envió nada y no se tocó el estado.")


if __name__ == "__main__":
    main()
