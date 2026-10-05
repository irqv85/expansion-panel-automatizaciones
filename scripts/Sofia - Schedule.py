"""
Revisa la carpeta de reportes por archivos nuevos. Para cada archivo nuevo:

1. Busca en el calendario de el responsable comercial una reunion organizada por Sofia Garcia
   (agente virtual) cuyo asunto coincida con el nombre de la empresa del
   archivo.
   a. Si la reunion ya tiene asignado (como asistente, o nombrado en el
      asunto) a uno o mas de los 4 vendedores -- reprogramacion, o Sofia
      ya reconocio el contexto -- se le reenvia la invitacion + el archivo
      (2 correos aparte) a CADA UNO de ellos, sin consumir turno de
      rotacion. (Aparecer como asistente en Outlook no garantiza que ya le
      llego la invitacion real, por eso siempre se reenvia.)
   b. Si la reunion no tiene ningun vendedor conocido asignado, se asigna
      al que este libre a la HORA DE LA REUNION (siguiendo el turno de
      rotacion), y se le envian los mismos 2 correos (invitacion + archivo).
2. Si no hay reunion de Sofia para ese archivo, se usa el flujo normal:
   se asigna al vendedor libre AHORA MISMO segun el turno de rotacion, y
   se le envia solo el archivo.

Otros asistentes de la reunion que no sean ninguno de los 4 vendedores
(ej. gente de otros equipos/roles) se ignoran por completo.

Cuando hay reunion se envian SIEMPRE 2 correos separados, en este orden:
  1) la invitacion de la reunion sola, sin adjunto (para que el vendedor la
     agregue a su calendario);
  2) el archivo de presentacion, en otro correo.

Los casos 1b y 2 avanzan el mismo turno compartido de rotacion. En todos
los casos se pone en copia a Yeniree Garcia, con el mismo asunto/cuerpo:
"Archivo de presentacion - {empresa}".

Vendedores de vacaciones (marcados con --vacaciones-agregar) se saltan por
completo: no se les asigna por turno ni aunque aparezcan como asistentes
en la reunion de Sofia.

Disponibilidad: se cruzan dos fuentes, porque a cada una se le escapa algo.
El Free/Busy no ve eventos marcados "Show As: Free" (tipico de asistentes
Optional) y el calendario compartido no muestra detalles de citas privadas.
Solo se considera libre si NINGUNA de las dos muestra algo a esa hora.

Requiere: Outlook de escritorio instalado y con sesion iniciada.

Uso:
  python "Sofia - Schedule.py"                   Calcula y envia de una vez
  python "Sofia - Schedule.py" --dry-run          Muestra que haria, sin enviar
  python "Sofia - Schedule.py" --plan-out p.json  Calcula el plan para revisarlo
                                                   (no envia nada)
  python "Sofia - Schedule.py" --execute-plan p.json
                                                  Envia exactamente ese plan ya
                                                   revisado (rapido: no recalcula)
  python "Sofia - Schedule.py" --vacaciones-agregar Francisco
  python "Sofia - Schedule.py" --vacaciones-quitar Francisco
  python "Sofia - Schedule.py" --vacaciones-listar
"""
import json
import logging
import re
import sys
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

# win32com solo existe en Windows y solo lo necesitan las funciones que
# hablan con Outlook. Se importa de forma diferida (mas abajo, donde se usa)
# para que este modulo se pueda importar en otros lados (p.ej. "Configurar
# Vacaciones.py") sin requerir pywin32.

# ---- Configuracion ----
from paths import PRESENTACIONES_DIR as WATCH_FOLDER

SCRIPT_DIR = Path(__file__).resolve().parent
STATE_FILE = SCRIPT_DIR / "state.json"
LOG_FILE = SCRIPT_DIR / "envio_vendedores.log"
VACATION_FILE = SCRIPT_DIR / "vacaciones.json"

SELLERS = [
    {"name": "Vendedora 3", "email": "vendedora1@tuempresa.com"},
    {"name": "Vendedor 2", "email": "vendedor2@tuempresa.com"},
    {"name": "Vendedora 1", "email": "vendedora3@tuempresa.com"},
    {"name": "Vendedor 4", "email": "vendedor4@tuempresa.com"},
]
CC_EMAIL = "copia@tuempresa.com"
SOFIA_MARKER = "sofia"  # texto que debe aparecer en el organizador de la reunion

BUSINESS_DAYS = {0, 1, 2, 3, 4}  # lunes=0 ... viernes=4
BUSINESS_START_HOUR = 8
BUSINESS_END_HOUR = 18  # tope: 8am-5pm normal, hasta 6pm si hace falta

FREEBUSY_INTERVAL_MIN = 30

olFolderCalendar = 9

_log_formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
_file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
_file_handler.setFormatter(_log_formatter)
_console_handler = logging.StreamHandler(sys.stdout)
_console_handler.setFormatter(_log_formatter)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(_file_handler)
logger.addHandler(_console_handler)


def load_state():
    if STATE_FILE.exists():
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"rotation_index": 0, "processed_files": []}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def load_vacation_emails():
    if VACATION_FILE.exists():
        with open(VACATION_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {e.lower() for e in data.get("de_vacaciones", [])}
    return set()


def save_vacation_emails(emails):
    with open(VACATION_FILE, "w", encoding="utf-8") as f:
        json.dump({"de_vacaciones": sorted(emails)}, f, indent=2, ensure_ascii=False)


def resolve_seller(identifier):
    """Busca un vendedor de SELLERS por correo (exacto) o por nombre/primer
    nombre (coincidencia parcial, sin tildes). None si no se reconoce."""
    ident_lower = (identifier or "").lower()
    ident_norm = _normalize(identifier)
    for seller in SELLERS:
        if seller["email"].lower() == ident_lower:
            return seller
    for seller in SELLERS:
        if ident_norm and ident_norm in _normalize(seller["name"]):
            return seller
    return None


def is_business_hours(now):
    return now.weekday() in BUSINESS_DAYS and BUSINESS_START_HOUR <= now.hour < BUSINESS_END_HOUR


SOFIA_SEARCH_DAYS_BACK = 1
SOFIA_SEARCH_DAYS_FORWARD = 45


def sofia_search_window(now):
    """Ventana de busqueda de reuniones de Sofia en el calendario propio.

    Se probo acotarla a 'hoy y el proximo dia habil' para acelerar
    'Revisar', pero rompio el caso real de archivos con reunion agendada
    con varios dias de anticipacion (ej. una reunion para dentro de una
    semana ya no se encontraba, y el archivo se mandaba sin invitacion).
    La demora de 'Revisar' no es este escaneo -- corre en un par de
    segundos, cacheado -- sino la consulta de disponibilidad por vendedor
    (is_free), que abre el calendario compartido de cada uno via COM."""
    inicio = now - timedelta(days=SOFIA_SEARCH_DAYS_BACK)
    fin = now + timedelta(days=SOFIA_SEARCH_DAYS_FORWARD)
    return inicio, fin


olResponseDeclined = 4


def _is_free_via_freebusy(outlook_ns, email, when):
    """Respaldo si no se pudo abrir el calendario compartido: usa el Show As
    de FreeBusy. Menos confiable -- un evento marcado 'Show As: Free' (comun
    en asistentes Optional) aparece como libre aunque la persona ya tenga esa
    reunion aceptada en su calendario."""
    recipient = outlook_ns.CreateRecipient(email)
    recipient.Resolve()
    if not recipient.Resolved:
        logging.warning(f"No se pudo resolver el destinatario: {email}")
        return False
    fb = recipient.FreeBusy(when, FREEBUSY_INTERVAL_MIN, True)
    if not fb:
        logging.warning(f"Sin datos de disponibilidad para: {email}")
        return False
    return fb[0] == "0"


RESTRICT_DATE_FMT = "%m/%d/%Y %I:%M %p"


def iter_calendar_items(items, window_start, window_end):
    """Recorre los eventos de un calendario dentro de la ventana dada.

    NO usa Items.Restrict: para ventanas anchas (dias/semanas) se comprobo
    que devuelve un subconjunto incompleto sin avisar -- ej. en una ventana
    de 45 dias con 28 eventos reales, Restrict trajo solo 8 y se perdio
    justo la reunion de Sofia que se buscaba. Como esto es el calendario
    PROPIO (no uno compartido via COM), recorrerlo entero es rapido; el
    costo real esta en abrir calendarios compartidos (ver is_free), no aca."""
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    for item in items:
        try:
            item_start = to_py_datetime(item.Start)
        except Exception:
            continue
        if item_start > window_end:
            break
        if item_start < window_start:
            continue
        yield item


_SHARED_CALENDAR_CACHE = {}


def get_shared_calendar(outlook_ns, email):
    """Carpeta de calendario compartido del vendedor (cacheada). None si no
    se puede abrir (sin permiso de calendario compartido)."""
    key = email.lower()
    if key in _SHARED_CALENDAR_CACHE:
        return _SHARED_CALENDAR_CACHE[key]

    calendar = None
    try:
        recipient = outlook_ns.CreateRecipient(email)
        recipient.Resolve()
        if recipient.Resolved:
            calendar = outlook_ns.GetSharedDefaultFolder(recipient, olFolderCalendar)
        else:
            logging.warning(f"No se pudo resolver el destinatario: {email}")
    except Exception as exc:
        logging.warning(f"No se pudo abrir el calendario compartido de {email}: {exc}")
        calendar = None

    _SHARED_CALENDAR_CACHE[key] = calendar
    return calendar


def find_overlapping_events(outlook_ns, email, slot_start, slot_end):
    """Eventos del calendario del vendedor que se cruzan con el rango dado.

    El filtro de cruce lo resuelve Outlook (Restrict), no Python: leer
    propiedades de eventos de un calendario compartido cuesta ~0.3s por
    propiedad, asi que se traen solo los pocos eventos que realmente choquen
    en vez de recorrer dias completos.

    Devuelve la lista de asuntos, o None si no se pudo consultar."""
    calendar = get_shared_calendar(outlook_ns, email)
    if calendar is None:
        return None

    try:
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        restricted = items.Restrict(
            f"[Start] < '{slot_end.strftime(RESTRICT_DATE_FMT)}' AND "
            f"[End] > '{slot_start.strftime(RESTRICT_DATE_FMT)}'"
        )
        subjects = []
        for item in restricted:
            if getattr(item, "ResponseStatus", None) == olResponseDeclined:
                continue
            subjects.append(getattr(item, "Subject", "") or "(sin asunto)")
            if len(subjects) >= 5:
                break
        return subjects
    except Exception as exc:
        logging.warning(f"No se pudo consultar el calendario de {email}: {exc}")
        return None


def is_free(outlook_ns, email, when, duration_minutes=FREEBUSY_INTERVAL_MIN):
    """Libre = ni el calendario compartido ni el Free/Busy muestran algo a esa
    hora. Se cruzan las dos fuentes porque cada una se le escapa algo: el
    Free/Busy no ve eventos marcados 'Show As: Free' (tipico de asistentes
    Optional), y el calendario compartido no muestra detalles de citas
    privadas."""
    # Primero el FreeBusy: es una sola llamada instantanea. Si ya dice
    # ocupado, no hace falta la consulta al calendario, que es lenta.
    if not _is_free_via_freebusy(outlook_ns, email, when):
        logging.info(f"{email} ocupado a las {when:%Y-%m-%d %H:%M} (segun Free/Busy).")
        return False

    slot_end = when + timedelta(minutes=duration_minutes)
    conflicts = find_overlapping_events(outlook_ns, email, when, slot_end)
    if conflicts is None:
        logging.warning(
            f"No se pudo leer el calendario de {email}; queda solo el FreeBusy "
            f"(Show As) como fuente, menos confiable."
        )
        return True

    if conflicts:
        logging.info(
            f"{email} ocupado a las {when:%Y-%m-%d %H:%M} (el Free/Busy no lo "
            f"mostraba): " + "; ".join(conflicts)
        )
        return False

    return True


def pick_seller(outlook_ns, rotation_index, when, vacation_emails=frozenset()):
    """Devuelve (idx, seller) del primer vendedor libre (y no de vacaciones)
    en 'when', empezando en rotation_index y siguiendo el orden de turno.
    (None, None) si nadie esta disponible."""
    for offset in range(len(SELLERS)):
        idx = (rotation_index + offset) % len(SELLERS)
        seller = SELLERS[idx]
        if seller["email"].lower() in vacation_emails:
            logging.info(f"{seller['name']} esta de vacaciones, se salta.")
            continue
        if is_free(outlook_ns, seller["email"], when):
            return idx, seller
        logging.info(f"{seller['name']} ocupado a las {when}, se intenta el siguiente.")
    return None, None


def slug_to_company_name(filename_stem):
    stem = re.sub(r"-\d{4}-\d{2}-\d{2}$", "", filename_stem)
    words = re.split(r"[-_]+", stem)
    return " ".join(w.capitalize() for w in words if w)


def _strip_accents(text):
    return "".join(
        c for c in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(c)
    )


def _normalize(text):
    return re.sub(r"[^a-z0-9]", "", _strip_accents((text or "").lower()))


SUBJECT_MATCH_STOPWORDS = {
    "de", "del", "la", "el", "los", "las", "y", "en", "para", "con", "a", "al",
}
SUBJECT_MATCH_MIN_WORD_LEN = 4


def subjects_match(company_name, event_subject):
    a, b = _normalize(company_name), _normalize(event_subject)
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    # El nombre de archivo a veces trae la razon social completa mas un
    # codigo/acronimo (ej. "junta-comunal-de-bella-vista-fqml"), pero la
    # reunion de Sofia solo usa el codigo en el asunto ("... | FQML"). Si
    # el match de frase completa falla, se acepta si comparten al menos una
    # palabra significativa (4+ letras, sin stopwords).
    words = [
        w for w in re.split(r"[^a-z0-9]+", _strip_accents(company_name.lower()))
        if len(w) >= SUBJECT_MATCH_MIN_WORD_LEN and w not in SUBJECT_MATCH_STOPWORDS
    ]
    return any(w in b for w in words)


def to_py_datetime(com_date):
    return datetime(
        com_date.year, com_date.month, com_date.day,
        com_date.hour, com_date.minute, com_date.second,
    )


_OWN_MEETINGS_CACHE = None  # None = todavia no se cargo; [] = se cargo y no hay ninguna


def load_sofia_meetings(outlook_ns, now):
    """Carga una sola vez (cacheado) las reuniones organizadas por Sofia en
    el calendario propio, dentro de la ventana de busqueda. El cache
    distingue "todavia no se cargo" de "se cargo y no hay ninguna": con una
    lista vacia como valor inicial, si no habia reuniones el chequeo
    `if cache:` era falso y volvia a recorrer el calendario completo para
    cada archivo pendiente."""
    global _OWN_MEETINGS_CACHE
    if _OWN_MEETINGS_CACHE is not None:
        return _OWN_MEETINGS_CACHE

    calendar = outlook_ns.GetDefaultFolder(olFolderCalendar)
    items = calendar.Items
    window_start, window_end = sofia_search_window(now)

    encontradas = []
    for item in iter_calendar_items(items, window_start, window_end):
        organizer_text = (getattr(item, "Organizer", "") or "").lower()
        if SOFIA_MARKER not in organizer_text:
            continue
        encontradas.append(item)

    if encontradas:
        logging.info(f"Reuniones de Sofia encontradas en el calendario: {len(encontradas)}.")
    else:
        logging.info("No hay reuniones de Sofia proximas en tu calendario.")

    _OWN_MEETINGS_CACHE = encontradas
    return _OWN_MEETINGS_CACHE


def find_sofia_meeting(outlook_ns, company_name, now):
    for item in load_sofia_meetings(outlook_ns, now):
        if subjects_match(company_name, getattr(item, "Subject", "")):
            return item
    return None


def build_subject_body(seller, company_name):
    first_name = seller["name"].split()[0]
    subject = f"Archivo de presentacion - {company_name}"
    body = (
        f"Hola {first_name}\n\n"
        f"Encontrarás adjunto el archivo de presentación para tu reunión con {company_name}\n\n"
        f"Cualquier duda dejame saber."
    )
    return subject, body


def send_file_only(outlook_app, seller, filepath, company_name, dry_run=False):
    subject, body = build_subject_body(seller, company_name)
    if dry_run:
        logging.info(
            f"[DRY-RUN] Se enviaria '{filepath.name}' a {seller['name']} "
            f"({seller['email']}), CC {CC_EMAIL} -- asunto: '{subject}'"
        )
        return
    mail = outlook_app.CreateItem(0)  # olMailItem
    mail.To = seller["email"]
    mail.CC = CC_EMAIL
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(str(filepath))
    mail.Send()
    logging.info(f"Enviado '{filepath.name}' a {seller['name']} ({seller['email']}), CC {CC_EMAIL}")


olFolderInbox = 6
olFolderDeletedItems = 3
olTo = 1
olCC = 2
MEETING_REQUEST_CLASS = "IPM.Schedule.Meeting.Request"


def _candidate_mail_folders(outlook_ns):
    """Carpetas donde puede estar la invitacion original de Sofia, en orden de
    preferencia. Outlook borra las solicitudes de reunion una vez respondidas,
    asi que Elementos eliminados es un lugar normal donde encontrarlas."""
    folders = []
    for folder_id in (olFolderInbox, olFolderDeletedItems):
        try:
            folder = outlook_ns.GetDefaultFolder(folder_id)
            folders.append(folder)
            for sub in folder.Folders:
                folders.append(sub)
        except Exception:
            continue
    return folders


def find_original_meeting_request(outlook_ns, meeting_subject):
    """Busca en el buzon la invitacion original (MeetingItem) que envio Sofia.

    Reenviar ESA es lo que produce una invitacion de calendario real para el
    vendedor (con Aceptar/Rechazar). La cita del calendario, en cambio, es un
    AppointmentItem y solo permite ForwardAsVcal(), que llega como un correo
    con un .vcs adjunto -- no como invitacion."""
    target = _normalize(meeting_subject)
    if not target:
        return None

    for folder in _candidate_mail_folders(outlook_ns):
        try:
            restricted = folder.Items.Restrict(f"[MessageClass] = '{MEETING_REQUEST_CLASS}'")
        except Exception:
            continue
        try:
            for item in restricted:
                sender = (getattr(item, "SenderName", "") or "").lower()
                if SOFIA_MARKER not in sender:
                    continue
                if _normalize(getattr(item, "Subject", "") or "") == target:
                    return item
        except Exception:
            continue
    return None


def send_meeting_invite(outlook_ns, meeting_item, meeting_subject, seller, dry_run=False):
    """Reenvia la invitacion de la reunion al vendedor, en un correo aparte y
    SIN adjuntar el archivo de presentacion.

    Primero intenta reenviar la invitacion original de Sofia (MeetingItem), que
    es la que llega como invitacion de calendario de verdad. Si ya no esta en el
    buzon, cae a ForwardAsVcal() de la cita, que llega como correo con la
    reunion adjunta en un .vcs."""
    forward_mail = None
    es_invitacion_real = False

    original = find_original_meeting_request(outlook_ns, meeting_subject or getattr(meeting_item, "Subject", ""))
    if original is not None:
        try:
            forward_mail = original.Forward()
            es_invitacion_real = True
        except Exception as exc:
            logging.warning(f"No se pudo reenviar la invitacion original: {exc}")

    if forward_mail is None:
        try:
            forward_mail = meeting_item.ForwardAsVcal()
            logging.warning(
                "No se encontro la invitacion original de Sofia en el buzon; se "
                "envia la reunion adjunta como .vcs (no llega como invitacion "
                "de calendario)."
            )
        except Exception as exc:
            logging.warning(
                f"No se pudo reenviar la invitacion de '{meeting_subject}': {exc}"
            )
            return False

    # Un MeetingItem reenviado no permite asignar .To/.CC (son de solo
    # lectura); hay que agregar los destinatarios por Recipients.
    try:
        forward_mail.Recipients.Add(seller["email"]).Type = olTo
        forward_mail.Recipients.Add(CC_EMAIL).Type = olCC
        forward_mail.Recipients.ResolveAll()
    except Exception as exc:
        logging.warning(f"No se pudieron agregar los destinatarios a la invitacion: {exc}")
        return False
    # Se deja el Subject/Body que genera Outlook al reenviar.

    if dry_run:
        logging.info(
            f"[DRY-RUN] Se reenviaria la invitacion de '{meeting_item.Subject}' "
            f"a {seller['name']} ({seller['email']}), CC {CC_EMAIL}"
        )
        return True

    forward_mail.Send()
    logging.info(
        f"Invitacion de '{meeting_item.Subject}' reenviada a {seller['name']} "
        f"({seller['email']}), CC {CC_EMAIL}"
    )
    return True


def _get_recipient_smtp(recipient):
    try:
        addr_entry = recipient.AddressEntry
        user_type = getattr(addr_entry, "AddressEntryUserType", None)
        if user_type in (0, 5):  # olExchangeUserAddressEntry, olExchangeRemoteUserAddressEntry
            exch_user = addr_entry.GetExchangeUser()
            if exch_user and exch_user.PrimarySmtpAddress:
                return exch_user.PrimarySmtpAddress.lower()
    except Exception:
        pass
    return (getattr(recipient, "Address", "") or "").lower()


def find_named_sellers(meeting_item, vacation_emails=frozenset()):
    """Devuelve la lista de TODOS los vendedores conocidos (sin duplicados,
    sin los de vacaciones) que ya aparecen como asistentes de la reunion, o
    mencionados en el asunto -- reprogramacion / Sofia ya reconocio el
    contexto. Puede haber mas de uno en la misma reunion. Otros asistentes
    que no esten en SELLERS (ej. otros roles/equipos) se ignoran. Lista
    vacia si no hay ninguna coincidencia utilizable."""
    found = []
    found_emails = set()

    def _add(seller):
        email = seller["email"].lower()
        if email in found_emails:
            return
        if email in vacation_emails:
            logging.info(
                f"{seller['name']} aparece asociado a la reunion pero esta de "
                f"vacaciones; se ignora esa asignacion."
            )
            return
        found_emails.add(email)
        found.append(seller)

    try:
        recipients = meeting_item.Recipients
        for i in range(1, recipients.Count + 1):
            r = recipients.Item(i)
            smtp = _get_recipient_smtp(r)
            name_norm = _normalize(getattr(r, "Name", "") or "")
            for seller in SELLERS:
                if smtp:
                    # Ya sabemos con certeza quien es este asistente (su
                    # correo real); no caer al respaldo por nombre, que
                    # podria confundirlo con un cliente que comparta primer
                    # nombre con un vendedor (ej. "Francisco Ferrante" del
                    # cliente vs. "Vendedor 2" del equipo).
                    matched = smtp == seller["email"].lower()
                else:
                    first_name_norm = _normalize(seller["name"].split()[0])
                    matched = bool(first_name_norm) and first_name_norm in name_norm
                if matched:
                    _add(seller)
    except Exception:
        logging.warning("No se pudo leer la lista de asistentes de la reunion.")

    subject_norm = _normalize(getattr(meeting_item, "Subject", ""))
    for seller in SELLERS:
        first_name_norm = _normalize(seller["name"].split()[0])
        if first_name_norm and first_name_norm in subject_norm:
            _add(seller)

    return found


REASON_LABELS = {
    "ya_asignado": "ya esta en la invitacion de Sofia",
    "turno_reunion": "por turno, libre a la hora de la reunion",
    "turno_sin_reunion": "por turno, sin reunion de Sofia asociada",
}


def find_new_files(processed):
    current_files = {p.name for p in WATCH_FOLDER.iterdir() if p.is_file()}
    return sorted(
        (WATCH_FOLDER / name for name in current_files if name not in processed),
        key=lambda p: p.stat().st_ctime,
    )


def build_plan(outlook_ns, now, state, vacation_emails, new_files):
    """Calcula QUE se enviaria y A QUIEN, sin enviar nada. El turno de
    rotacion avanza dentro del plan (no solo al enviar), para que la vista
    previa refleje el reparto real entre vendedores."""
    rotation_index = state.get("rotation_index", 0)
    entries = []
    blocked = []

    # 1a pasada: resolver la reunion de cada archivo (rapido, usa el cache
    # de reuniones de Sofia) para saber QUE horas hay que evaluar.
    resolved = []
    for filepath in new_files:
        company_name = slug_to_company_name(filepath.stem)
        meeting = find_sofia_meeting(outlook_ns, company_name, now)
        meeting_start = to_py_datetime(meeting.Start) if meeting is not None else None
        resolved.append((filepath, company_name, meeting, meeting_start))
        logging.info(
            f"'{filepath.name}' (empresa: '{company_name}') -> "
            + (f"reunion '{meeting.Subject}' el {meeting_start}" if meeting is not None else "sin reunion de Sofia")
        )

    # 2a pasada: asignar vendedor (turno + disponibilidad real a esa hora).
    for filepath, company_name, meeting, meeting_start in resolved:

        def _entry(seller, reason, meeting_obj=None, meeting_start=None):
            return {
                "file": str(filepath),
                "file_name": filepath.name,
                "company": company_name,
                "seller_name": seller["name"],
                "seller_email": seller["email"],
                "meeting_subject": (getattr(meeting_obj, "Subject", "") or "") if meeting_obj else None,
                "meeting_entry_id": getattr(meeting_obj, "EntryID", None) if meeting_obj else None,
                "meeting_start": meeting_start.isoformat() if meeting_start else None,
                "send_invite": meeting_obj is not None,
                "reason": reason,
            }

        if meeting is not None:
            named_sellers = find_named_sellers(meeting, vacation_emails)
            if named_sellers:
                # Ya hay vendedor(es) conocido(s) en la invitacion: no
                # consume turno de rotacion.
                for seller in named_sellers:
                    entries.append(_entry(seller, "ya_asignado", meeting, meeting_start))
                continue

            idx, seller = pick_seller(outlook_ns, rotation_index, meeting_start, vacation_emails)
            if seller is None:
                blocked.append({
                    "file_name": filepath.name,
                    "motivo": "Todos los vendedores disponibles estan ocupados a la hora de la reunion.",
                })
                continue
            entries.append(_entry(seller, "turno_reunion", meeting, meeting_start))
            rotation_index = (idx + 1) % len(SELLERS)
        else:
            idx, seller = pick_seller(outlook_ns, rotation_index, now, vacation_emails)
            if seller is None:
                blocked.append({
                    "file_name": filepath.name,
                    "motivo": "Todos los vendedores disponibles estan ocupados ahora.",
                })
                continue
            entries.append(_entry(seller, "turno_sin_reunion"))
            rotation_index = (idx + 1) % len(SELLERS)

    return {
        "generated_at": now.isoformat(),
        "fuera_de_horario": not is_business_hours(now),
        "rotation_index_after": rotation_index,
        "vacaciones": [s["name"] for s in SELLERS if s["email"].lower() in vacation_emails],
        "entries": entries,
        "blocked": blocked,
    }


def log_plan(plan):
    logging.info("-" * 60)
    if plan["vacaciones"]:
        logging.info("De vacaciones (no reciben envios): " + ", ".join(plan["vacaciones"]))
    if plan["fuera_de_horario"]:
        logging.info("AVISO: fuera de horario laboral (lun-vie 8am-6pm).")
    if not plan["entries"]:
        logging.info("Nada por enviar.")
    for e in plan["entries"]:
        correos = "2 correos (1: invitacion sola, 2: archivo)" if e["send_invite"] else "1 correo (archivo)"
        logging.info(
            f"{e['file_name']} -> {e['seller_name']} [{REASON_LABELS.get(e['reason'], e['reason'])}] :: {correos}"
        )
        if e["meeting_subject"]:
            logging.info(f"    reunion: '{e['meeting_subject']}' ({e['meeting_start']})")
    for b in plan["blocked"]:
        logging.info(f"PENDIENTE {b['file_name']}: {b['motivo']}")
    logging.info("-" * 60)


def execute_plan(outlook_app, outlook_ns, plan, dry_run=False):
    """Ejecuta un plan ya calculado. Por cada asignacion se envian correos
    SEPARADOS y en este orden: primero la invitacion de la reunion sola (sin
    adjunto, para que el vendedor la acepte en su calendario) y despues, en
    otro correo, el archivo de presentacion."""
    entries = plan.get("entries", [])
    if not entries:
        logging.info("El plan no tiene nada por enviar.")
        return

    attempted, failed = set(), set()

    for e in entries:
        filepath = Path(e["file"])
        seller = {"name": e["seller_name"], "email": e["seller_email"]}
        attempted.add(e["file_name"])
        try:
            if not filepath.exists():
                raise FileNotFoundError(f"El archivo ya no existe: {filepath}")

            # 1) La invitacion, sola y sin adjunto.
            if e.get("send_invite") and e.get("meeting_entry_id"):
                try:
                    meeting_item = outlook_ns.GetItemFromID(e["meeting_entry_id"])
                except Exception as exc:
                    meeting_item = None
                    logging.warning(
                        f"No se pudo recuperar la reunion '{e.get('meeting_subject')}' "
                        f"del calendario: {exc}. Se envia solo el archivo."
                    )
                if meeting_item is not None:
                    send_meeting_invite(
                        outlook_ns, meeting_item, e.get("meeting_subject"),
                        seller, dry_run=dry_run,
                    )

            # 2) El archivo, en un correo aparte.
            send_file_only(outlook_app, seller, filepath, e["company"], dry_run=dry_run)
        except Exception:
            failed.add(e["file_name"])
            logging.exception(f"Error enviando '{e['file_name']}' a {seller['name']}.")

    if dry_run:
        logging.info("[DRY-RUN] No se guardo ningun cambio de estado.")
        return

    state = load_state()
    processed = set(state.get("processed_files", []))
    processed |= (attempted - failed)
    state["processed_files"] = sorted(processed)
    state["rotation_index"] = plan.get("rotation_index_after", state.get("rotation_index", 0))
    save_state(state)

    if failed:
        logging.info(
            "Archivos que fallaron y se reintentaran en la proxima corrida: "
            + ", ".join(sorted(failed))
        )


def compute_plan(now=None):
    """Prepara todo lo necesario y devuelve (plan, outlook_app, outlook_ns)."""
    now = now or datetime.now()
    state = load_state()
    vacation_emails = load_vacation_emails()
    new_files = find_new_files(set(state.get("processed_files", [])))

    import win32com.client as win32
    outlook_app = win32.Dispatch("Outlook.Application")
    outlook_ns = outlook_app.GetNamespace("MAPI")

    if not new_files:
        plan = {
            "generated_at": now.isoformat(),
            "fuera_de_horario": not is_business_hours(now),
            "rotation_index_after": state.get("rotation_index", 0),
            "vacaciones": [s["name"] for s in SELLERS if s["email"].lower() in vacation_emails],
            "entries": [],
            "blocked": [],
        }
        return plan, outlook_app, outlook_ns

    logging.info(f"Archivos nuevos detectados ({len(new_files)}): " + ", ".join(p.name for p in new_files))
    plan = build_plan(outlook_ns, now, state, vacation_emails, new_files)
    return plan, outlook_app, outlook_ns


def main(dry_run=False):
    now = datetime.now()
    logging.info("=" * 60)
    logging.info(f"Inicio de ejecucion: {now}" + (" [DRY-RUN: solo muestra, no envia ni guarda]" if dry_run else ""))

    if not is_business_hours(now):
        logging.info("Fuera de horario laboral (lun-vie 8am-6pm). No se procesa nada.")
        return

    plan, outlook_app, outlook_ns = compute_plan(now)
    log_plan(plan)
    execute_plan(outlook_app, outlook_ns, plan, dry_run=dry_run)
    logging.info("Fin de ejecucion.")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Envio automatico de reportes a vendedores.")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Solo muestra que se enviaria y a quien, sin enviar nada ni guardar estado.",
    )
    parser.add_argument(
        "--vacaciones-agregar", metavar="VENDEDOR",
        help="Marca a un vendedor (nombre o correo) de vacaciones; no recibira envios hasta que se le quite.",
    )
    parser.add_argument(
        "--vacaciones-quitar", metavar="VENDEDOR",
        help="Quita a un vendedor de la lista de vacaciones.",
    )
    parser.add_argument(
        "--vacaciones-listar", action="store_true",
        help="Muestra quienes estan marcados de vacaciones.",
    )
    parser.add_argument(
        "--plan-out", metavar="ARCHIVO_JSON",
        help="Calcula el plan (que se enviaria y a quien) y lo guarda en ese JSON, sin enviar nada.",
    )
    parser.add_argument(
        "--execute-plan", metavar="ARCHIVO_JSON",
        help="Envia exactamente lo que dice ese plan JSON (generado antes con --plan-out).",
    )
    args = parser.parse_args()

    if args.vacaciones_listar:
        emails = load_vacation_emails()
        if not emails:
            print("Nadie esta marcado de vacaciones.")
        else:
            print("De vacaciones:")
            for s in SELLERS:
                if s["email"].lower() in emails:
                    print(f"  - {s['name']} ({s['email']})")
        sys.exit(0)

    if args.vacaciones_agregar:
        seller = resolve_seller(args.vacaciones_agregar)
        if not seller:
            print(f"No se reconoce al vendedor '{args.vacaciones_agregar}'.")
            sys.exit(1)
        emails = load_vacation_emails()
        emails.add(seller["email"].lower())
        save_vacation_emails(emails)
        print(f"{seller['name']} marcado de vacaciones. No recibira envios hasta que se le quite.")
        sys.exit(0)

    if args.vacaciones_quitar:
        seller = resolve_seller(args.vacaciones_quitar)
        if not seller:
            print(f"No se reconoce al vendedor '{args.vacaciones_quitar}'.")
            sys.exit(1)
        emails = load_vacation_emails()
        emails.discard(seller["email"].lower())
        save_vacation_emails(emails)
        print(f"{seller['name']} ya no esta de vacaciones.")
        sys.exit(0)

    if args.plan_out:
        try:
            logging.info("=" * 60)
            logging.info("Calculando plan para revision (no se envia nada)...")
            plan, _app, _ns = compute_plan()
            log_plan(plan)
            with open(args.plan_out, "w", encoding="utf-8") as f:
                json.dump(plan, f, indent=2, ensure_ascii=False)
            logging.info(f"Plan guardado en: {args.plan_out}")
        except Exception:
            logging.exception("Error calculando el plan.")
            sys.exit(1)
        sys.exit(0)

    if args.execute_plan:
        try:
            with open(args.execute_plan, "r", encoding="utf-8") as f:
                plan = json.load(f)
            logging.info("=" * 60)
            logging.info(f"Ejecutando plan confirmado desde: {args.execute_plan}")
            log_plan(plan)
            import win32com.client as win32
            outlook_app = win32.Dispatch("Outlook.Application")
            outlook_ns = outlook_app.GetNamespace("MAPI")
            execute_plan(outlook_app, outlook_ns, plan, dry_run=args.dry_run)
            logging.info("Fin de ejecucion del plan.")
        except Exception:
            logging.exception("Error ejecutando el plan.")
            sys.exit(1)
        sys.exit(0)

    try:
        main(dry_run=args.dry_run)
    except Exception:
        logging.exception("Error inesperado durante la ejecucion.")
        sys.exit(1)
